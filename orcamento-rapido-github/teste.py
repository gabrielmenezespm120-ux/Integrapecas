"""Teste local sem WhatsApp e sem gastar API: o Claude é trocado por um extrator falso."""
import hashlib, hmac, json, os, threading, time, urllib.error, urllib.request
from pathlib import Path
from http.server import ThreadingHTTPServer

os.environ.update(DRY_RUN="1", WHATSAPP_APP_SECRET="segredo", WHATSAPP_VERIFY_TOKEN="verif",
                  NUMEROS_AUTORIZADOS="5511999990000")
import app
from openpyxl import load_workbook

# 1. normalização e cruzamento com estoque
assert app.normalizar(" flt-1001 ") == "FLT1001"
assert app.num("1.234,50") == 1234.5 and app.num("") is None

itens = [{"codigo": "flt 1001", "descricao": "filtro oleo", "quantidade": 2},
         {"codigo": "", "descricao": "PARAFUSO M8", "quantidade": 10}]
caminho, avisos = app.montar_orcamento(itens, "5511999990000")
ws = load_workbook(caminho).active
assert ws["B9"].value == "flt 1001" and ws["C9"].value == "FILTRO DE OLEO MOTOR"
assert ws["E9"].value == 48.9 and ws["F9"].value == 12 and ws["G9"].value == "=D9*E9"
assert ws["C10"].value == "PARAFUSO M8" and ws["E10"].value is None
assert avisos == ["PARAFUSO M8: não achei no estoque"], avisos
caminho.unlink()

# 2. webhook: verificação, assinatura, número autorizado
app.extrair_itens = lambda blocos: {"itens": itens[:1], "observacao": ""}
srv = ThreadingHTTPServer(("127.0.0.1", 0), app.Servidor)
threading.Thread(target=srv.serve_forever, daemon=True).start()
url = f"http://127.0.0.1:{srv.server_port}/"

def post(payload, segredo="segredo"):
    corpo = json.dumps(payload).encode()
    sig = "sha256=" + hmac.new(segredo.encode(), corpo, hashlib.sha256).hexdigest()
    req = urllib.request.Request(url, data=corpo, headers={"X-Hub-Signature-256": sig})
    try:
        return urllib.request.urlopen(req).status
    except urllib.error.HTTPError as e:
        return e.code

assert urllib.request.urlopen(url + "?hub.mode=subscribe&hub.verify_token=verif&hub.challenge=42").read() == b"42"
try:
    urllib.request.urlopen(url + "?hub.mode=subscribe&hub.verify_token=errado&hub.challenge=42")
    raise AssertionError("aceitou verify_token errado")
except urllib.error.HTTPError as e:
    assert e.code == 403

def msg(de, mid):
    return {"entry": [{"changes": [{"value": {"messages": [
        {"id": mid, "from": de, "type": "text", "text": {"body": "2 filtro FLT-1001"}}]}}]}]}

saida = app.BASE / app.CFG["pasta_saida"]
antes = set(saida.glob("*.xlsx"))
assert post(msg("5511999990000", "a1"), segredo="falso") == 403
assert post(msg("5500000000000", "a2")) == 200          # número não autorizado: ignora
assert post(msg("5511999990000", "a3")) == 200
assert post(msg("5511999990000", "a3")) == 200          # reenvio da Meta: não duplica
time.sleep(1)
novos = set(saida.glob("*.xlsx")) - antes
assert len(novos) == 1, novos
for f in novos:
    f.unlink()

# 2b. site: login, página, orçamento por texto, estoque só para gerente
import base64
from urllib.parse import unquote
os.environ.update(VENDEDORES="joao:senha1,maria:senha2", GERENTES="maria")

def site(rota, corpo=None, login="joao:senha1", tipo="text/plain; charset=utf-8"):
    h = {"Content-Type": tipo}
    if login:
        h["Authorization"] = "Basic " + base64.b64encode(login.encode()).decode()
    req = urllib.request.Request(url.rstrip("/") + rota, data=corpo, headers=h)
    try:
        r = urllib.request.urlopen(req)
        return r.status, r.headers, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.headers, e.read()

assert site("/", login=None)[0] == 401
assert site("/", login="joao:errada")[0] == 401
st, _, html = site("/")
assert st == 200 and "Gerar orçamento".encode() in html
st, h, xlsx = site("/orcamento", "2 filtro FLT-1001".encode())
assert st == 200 and xlsx[:2] == b"PK", (st, xlsx[:200])
assert json.loads(unquote(h["X-Avisos"])) == [] and h["X-Itens"] == "1"
assert site("/orcamento", b"", tipo="image/png")[0] == 422       # nada enviado
st, _, msg = site("/orcamento", b"%PDF", tipo="application/zip")
assert st == 422 and "não suportado" in msg.decode(), msg
assert site("/estoque", b"x", tipo="text/csv")[0] == 403          # joao não é gerente
for f in set(saida.glob("*.xlsx")) - antes:
    f.unlink()
srv.shutdown()

# 3. áudio: baixa, transcreve e manda a transcrição ao extrator (whisper e Meta trocados por falsos)
recebido = []
app.baixar_midia = lambda media_id: b"ogg"
app.transcrever = lambda dados: "duas unidades do filtro FLT 1001"
app.extrair_itens = lambda blocos: recebido.extend(blocos) or {"itens": itens[:1], "observacao": ""}
app.processar({"id": "a4", "de": "5511999990000", "tipo": "audio", "texto": "",
               "media_id": "m1", "mime": "audio/ogg; codecs=opus", "arquivo": ""})
assert "duas unidades do filtro FLT 1001" in recebido[0]["text"], recebido
for f in set(saida.glob("*.xlsx")) - antes:
    f.unlink()

# 4. atualização de estoque pelo WhatsApp (numa cópia, sem tocar no demo)
import io, shutil, tempfile
from openpyxl import Workbook
tmp = Path(tempfile.mkdtemp())
shutil.copy(app.BASE / app.CFG["estoque_csv"], tmp / "estoque.csv")
app.BASE, app.CFG["estoque_csv"] = tmp, "estoque.csv"
app.CFG["cabecalhos_estoque"] = {"codigo": "Código", "preco": "Preço Venda"}
os.environ["NUMEROS_ESTOQUE"] = "5511888880000"
respostas = []
app.enviar_texto = lambda para, texto: respostas.append(texto)

def doc(de, nome, dados, legenda="estoque"):
    app.baixar_midia = lambda media_id: dados
    app.processar({"id": "e", "de": de, "tipo": "document", "texto": legenda,
                   "media_id": "m", "mime": "", "arquivo": nome})

doc("5511999990000", "e.csv", b"x")                         # vendedor comum: recusado
assert "não pode" in respostas[-1]
doc("5511888880000", "e.csv", b"Codigo;Descricao\nA1;X\n")  # faltam colunas: mantém o antigo
assert "NÃO atualizado" in respostas[-1] and "preco" in respostas[-1]
assert "FLT1001" in app.carregar_estoque()
wb = Workbook(); ws = wb.active
ws.append(["Código", "Descrição", "Preço Venda", "Estoque"])
ws.append(["NOV-1", "PECA NOVA", 1234.5, 3]); ws.append([None, "linha sem código", 1.0, 1])
buf = io.BytesIO(); wb.save(buf)
doc("5511888880000", "estoque.xlsx", buf.getvalue())
assert respostas[-1].startswith("Estoque atualizado: 1 produtos"), respostas[-1]
est = app.carregar_estoque()
assert list(est) == ["NOV1"] and app.num(est["NOV1"]["preco"]) == 1234.5
assert (tmp / "estoque.anterior.csv").exists()
shutil.rmtree(tmp)
print("OK: todos os testes passaram")
