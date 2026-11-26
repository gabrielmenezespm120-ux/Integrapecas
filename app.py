"""Orçamento automático: site (principal) e WhatsApp oficial (opcional).

O vendedor manda o pedido do cliente (texto, foto, PDF, planilha ou áudio) pelo site
ou encaminha para o número da empresa; o servidor lê os itens com o Claude, cruza com
o estoque e devolve o orçamento em .xlsx no modelo da empresa.

Variáveis de ambiente (nunca no código):
  ANTHROPIC_API_KEY       chave da API do Claude
  VENDEDORES              login do site: "joao:senha1,maria:senha2" (sem vírgula nem ":" na senha)
  GERENTES                logins do site que podem atualizar o estoque, ex.: "maria"
  --- só se usar o WhatsApp ---
  WHATSAPP_TOKEN          token permanente do app na Meta
  WHATSAPP_PHONE_ID       Phone number ID do número da empresa
  WHATSAPP_VERIFY_TOKEN   texto qualquer, igual ao cadastrado no webhook da Meta
  WHATSAPP_APP_SECRET     App Secret (valida a assinatura de cada chamada)
  NUMEROS_AUTORIZADOS     números dos vendedores, separados por vírgula (ex.: 5511999990000)
  NUMEROS_ESTOQUE         quem pode trocar o estoque/preços mandando um arquivo com a palavra ESTOQUE
                          (vazio = ninguém; normalmente só o gerente)
  GRAPH_VERSION           versão da Graph API (padrão v21.0; conferir a atual no painel da Meta)
  WHISPER_MODELO          modelo de transcrição de áudio (padrão small; base é mais leve e erra mais)
  DRY_RUN=1              não envia nada ao WhatsApp, só grava o .xlsx (teste local)
  PORTA                   porta HTTP (padrão 8080)
"""
import base64
import csv
import hashlib
import hmac
import io
import json
import mimetypes
import os
import re
import threading
import unicodedata
import urllib.request
import uuid
from datetime import date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

import anthropic
from openpyxl import load_workbook

BASE = Path(__file__).parent
CFG = json.loads((BASE / "config.json").read_text(encoding="utf-8"))
GRAPH = f"https://graph.facebook.com/{os.environ.get('GRAPH_VERSION', 'v21.0')}"
_whisper, _whisper_lock = None, threading.Lock()
_processadas = set()  # shortcut: dedup só em memória; se rodar em mais de um processo, gravar em disco/banco

PROMPT = """Você recebe um pedido de peças enviado por um cliente (texto, foto, print, PDF ou planilha).
Liste cada item pedido. Regras:
- codigo: o part number/código da peça exatamente como escrito; "" se não houver.
- descricao: o nome da peça, curto, em maiúsculas; sem repetir o código.
- quantidade: número pedido; se não vier, use 1 e diga em "observacao".
- Se um caractere estiver ambíguo (0/O, 1/I, 8/B), mantenha o mais provável e avise em "observacao".
- Em transcrição de áudio, códigos e números falados podem ter vindo errados: liste em "observacao" os que conferir.
- Ignore saudações, assinaturas e dados do cliente."""

SCHEMA = {
    "type": "object",
    "properties": {
        "itens": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "codigo": {"type": "string"},
                    "descricao": {"type": "string"},
                    "quantidade": {"type": "number"},
                },
                "required": ["codigo", "descricao", "quantidade"],
                "additionalProperties": False,
            },
        },
        "observacao": {"type": "string"},
    },
    "required": ["itens", "observacao"],
    "additionalProperties": False,
}


# ---------- segurança ----------
def assinatura_ok(corpo: bytes, cabecalho: str, segredo: str) -> bool:
    if not segredo or not cabecalho.startswith("sha256="):
        return False
    esperado = hmac.new(segredo.encode(), corpo, hashlib.sha256).hexdigest()
    return hmac.compare_digest(esperado, cabecalho[7:])


def autorizado(numero: str, variavel: str = "NUMEROS_AUTORIZADOS") -> bool:
    lista = {n.strip() for n in os.environ.get(variavel, "").split(",") if n.strip()}
    return numero in lista


# ---------- WhatsApp ----------
def extrair_mensagens(payload: dict) -> list[dict]:
    msgs = []
    for entry in payload.get("entry", []):
        for change in entry.get("changes", []):
            for m in change.get("value", {}).get("messages", []):
                tipo = m.get("type")
                anexo = m.get(tipo, {}) if tipo in ("image", "document", "audio") else {}
                msgs.append({
                    "id": m.get("id"), "de": m.get("from"), "tipo": tipo,
                    "texto": m.get("text", {}).get("body", "") or anexo.get("caption", ""),
                    "media_id": anexo.get("id"), "mime": anexo.get("mime_type", ""),
                    "arquivo": anexo.get("filename", ""),
                })
    return msgs


def _graph(url: str, dados: bytes | None = None, tipo: str = "application/json", cru: bool = False):
    req = urllib.request.Request(url, data=dados, headers={"Authorization": f"Bearer {os.environ['WHATSAPP_TOKEN']}"})
    if dados is not None:
        req.add_header("Content-Type", tipo)
    with urllib.request.urlopen(req, timeout=60) as r:
        corpo = r.read()
    return corpo if cru else json.loads(corpo)


def baixar_midia(media_id: str) -> bytes:
    info = _graph(f"{GRAPH}/{media_id}")
    return _graph(info["url"], cru=True)


def enviar_texto(para: str, texto: str):
    if os.environ.get("DRY_RUN"):
        print(f"[DRY_RUN] para {para}: {texto}")
        return
    corpo = {"messaging_product": "whatsapp", "to": para, "type": "text", "text": {"body": texto}}
    _graph(f"{GRAPH}/{os.environ['WHATSAPP_PHONE_ID']}/messages", json.dumps(corpo).encode())


def enviar_documento(para: str, caminho: Path, legenda: str):
    if os.environ.get("DRY_RUN"):
        print(f"[DRY_RUN] para {para}: {caminho.name} | {legenda}")
        return
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    sep = uuid.uuid4().hex
    partes = [
        f'--{sep}\r\nContent-Disposition: form-data; name="messaging_product"\r\n\r\nwhatsapp\r\n'.encode(),
        f'--{sep}\r\nContent-Disposition: form-data; name="type"\r\n\r\n{mime}\r\n'.encode(),
        f'--{sep}\r\nContent-Disposition: form-data; name="file"; filename="{caminho.name}"\r\n'
        f"Content-Type: {mime}\r\n\r\n".encode(),
        caminho.read_bytes(), f"\r\n--{sep}--\r\n".encode(),
    ]
    phone = os.environ["WHATSAPP_PHONE_ID"]
    midia = _graph(f"{GRAPH}/{phone}/media", b"".join(partes), f"multipart/form-data; boundary={sep}")
    corpo = {"messaging_product": "whatsapp", "to": para, "type": "document",
             "document": {"id": midia["id"], "filename": caminho.name, "caption": legenda}}
    _graph(f"{GRAPH}/{phone}/messages", json.dumps(corpo).encode())


# ---------- leitura do pedido ----------
def transcrever(dados: bytes) -> str:
    """Áudio do WhatsApp (ogg/opus) → texto, offline, com faster-whisper."""
    global _whisper
    with _whisper_lock:  # um modelo só na memória e uma transcrição por vez
        if _whisper is None:
            from faster_whisper import WhisperModel
            _whisper = WhisperModel(os.environ.get("WHISPER_MODELO", "small"), device="cpu", compute_type="int8")
        partes, _ = _whisper.transcribe(io.BytesIO(dados), language="pt", vad_filter=True)
        return " ".join(p.text.strip() for p in partes)


def conteudo_do_pedido(msg: dict) -> list[dict]:
    """Mensagem do WhatsApp → blocos de conteúdo para o Claude."""
    dados = baixar_midia(msg["media_id"]) if msg["media_id"] else None
    return blocos_de(dados, msg["mime"], msg["arquivo"], msg["texto"])


def blocos_de(dados: bytes | None, mime: str, arquivo: str, texto: str = "") -> list[dict]:
    """Arquivo e/ou texto do pedido → blocos de conteúdo para o Claude."""
    blocos = []
    if dados:
        mime = mime.split(";")[0].strip()
        if arquivo.lower().endswith((".opus", ".ogg", ".m4a", ".mp3", ".wav")):
            mime = "audio/" + arquivo.rsplit(".", 1)[1].lower()  # celular costuma mandar áudio sem tipo
        elif not mime or mime == "application/octet-stream":
            mime = mimetypes.guess_type(arquivo)[0] or ""
        if mime.startswith("audio/"):
            blocos.append({"type": "text", "text": "Transcrição automática de um áudio do cliente:\n" + transcrever(dados)})
        elif mime.startswith("image/"):
            blocos.append({"type": "image", "source": {"type": "base64", "media_type": mime,
                                                       "data": base64.standard_b64encode(dados).decode()}})
        elif mime == "application/pdf":
            blocos.append({"type": "document", "source": {"type": "base64", "media_type": mime,
                                                          "data": base64.standard_b64encode(dados).decode()}})
        elif arquivo.lower().endswith(".xlsx"):
            wb = load_workbook(io.BytesIO(dados), read_only=True, data_only=True)
            linhas = ["\t".join("" if c is None else str(c) for c in row) for ws in wb for row in ws.iter_rows(values_only=True)]
            blocos.append({"type": "text", "text": "Planilha do cliente:\n" + "\n".join(linhas)})
        elif mime.startswith("text/") or arquivo.lower().endswith(".csv"):
            blocos.append({"type": "text", "text": dados.decode("utf-8", errors="replace")})
        else:
            raise ValueError(f"tipo de arquivo não suportado: {mime or arquivo}")
    if texto:
        blocos.append({"type": "text", "text": texto})
    return blocos


def extrair_itens(blocos: list[dict]) -> dict:
    client = anthropic.Anthropic()
    resp = client.beta.messages.create(
        model="claude-opus-5-5",
        max_tokens=16000,
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        output_config={"effort": "low", "format": {"type": "json_schema", "schema": SCHEMA}},
        system=PROMPT,
        messages=[{"role": "user", "content": blocos}],
    )
    if resp.stop_reason == "refusal":
        raise ValueError("o modelo recusou o pedido")
    texto = next(b.text for b in resp.content if b.type == "text")
    return json.loads(texto)


# ---------- estoque e planilha ----------
def normalizar(codigo: str) -> str:
    return re.sub(r"[\s.\-/]", "", codigo).upper()


def carregar_estoque() -> dict:
    with open(BASE / CFG["estoque_csv"], encoding="utf-8-sig", newline="") as f:
        return {normalizar(r["codigo"]): r for r in csv.DictReader(f, delimiter=";")}


def num(valor: str):
    try:
        return float(valor.replace(".", "").replace(",", ".")) if valor else None
    except ValueError:
        return None


def _chave(texto) -> str:
    sem_acento = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return sem_acento.strip().lower()


def atualizar_estoque(dados: bytes, arquivo: str) -> int:
    """Troca o estoque pelo relatório enviado (CSV ou XLSX). Valida tudo antes de gravar."""
    if arquivo.lower().endswith(".xlsx"):
        ws = load_workbook(io.BytesIO(dados), read_only=True, data_only=True).active
        linhas = [list(r) for r in ws.iter_rows(values_only=True)]
    else:
        texto = dados.decode("utf-8-sig", errors="replace")
        dialeto = csv.Sniffer().sniff(texto.splitlines()[0], delimiters=";,\t")
        linhas = list(csv.reader(io.StringIO(texto), dialeto))
    if not linhas:
        raise ValueError("arquivo vazio")
    nomes = {k: _chave(v) for k, v in CFG.get("cabecalhos_estoque", {}).items()}
    cab = [_chave(c) for c in linhas[0]]
    faltando = [k for k in ("codigo", "descricao", "preco", "estoque") if nomes.get(k, k) not in cab]
    if faltando:
        raise ValueError(f"não achei as colunas: {', '.join(faltando)}")
    idx = {k: cab.index(nomes.get(k, k)) for k in ("codigo", "descricao", "preco", "estoque")}

    def br(v):  # número da planilha → texto no formato que num() lê (vírgula decimal)
        return f"{v:.2f}".replace(".", ",") if isinstance(v, float) else ("" if v is None else str(v))

    itens = [{k: br(r[i]) if i < len(r) else "" for k, i in idx.items()} for r in linhas[1:]]
    itens = [i for i in itens if i["codigo"].strip()]
    if not itens:
        raise ValueError("nenhum produto com código")
    destino = BASE / CFG["estoque_csv"]
    temp = destino.with_suffix(".novo")
    with open(temp, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, ["codigo", "descricao", "preco", "estoque"], delimiter=";")
        w.writeheader()
        w.writerows(itens)
    if destino.exists():
        os.replace(destino, destino.with_suffix(".anterior.csv"))  # guarda a versão anterior
    os.replace(temp, destino)
    return len(itens)


def montar_orcamento(itens: list[dict], para: str) -> tuple[Path, list[str]]:
    estoque, col = carregar_estoque(), CFG["colunas"]
    wb = load_workbook(BASE / CFG["modelo_xlsx"])
    ws = wb.active
    ws[CFG["celula_data"]] = date.today().strftime("%d/%m/%Y")
    avisos = []
    vagas = CFG["linha_final"] - CFG["linha_inicial"] + 1
    if len(itens) > vagas:
        avisos.append(f"pedido com {len(itens)} itens, o modelo só tem {vagas} linhas: o excedente ficou fora")
    for linha, item in zip(range(CFG["linha_inicial"], CFG["linha_final"] + 1), itens):
        cad = estoque.get(normalizar(item["codigo"])) if item["codigo"] else None
        preco = num(cad["preco"]) if cad else None
        ws[f"{col['codigo']}{linha}"] = item["codigo"]
        ws[f"{col['descricao']}{linha}"] = cad["descricao"] if cad else item["descricao"]
        ws[f"{col['qtd']}{linha}"] = item["quantidade"]
        ws[f"{col['estoque']}{linha}"] = num(cad["estoque"]) if cad else None
        if preco is not None:
            ws[f"{col['preco']}{linha}"] = preco
            ws[f"{col['total']}{linha}"] = f"={col['qtd']}{linha}*{col['preco']}{linha}"
        if not cad:
            avisos.append(f"{item['codigo'] or item['descricao']}: não achei no estoque")
    pasta = BASE / CFG["pasta_saida"]
    pasta.mkdir(exist_ok=True)
    saida = pasta / f"Orcamento_{date.today():%Y-%m-%d}_{re.sub(r'[^A-Za-z0-9]', '', para)}_{uuid.uuid4().hex[:4]}.xlsx"
    wb.save(saida)
    return saida, avisos


# ---------- fluxo ----------
def gerar_orcamento(blocos: list[dict], quem: str) -> tuple[Path | None, list[str], int]:
    """Pedido → (planilha, avisos, nº de itens). Sem itens: planilha None e o motivo nos avisos."""
    lido = extrair_itens(blocos)
    obs = [lido["observacao"]] if lido.get("observacao") else []
    if not lido["itens"]:
        return None, obs, 0
    caminho, avisos = montar_orcamento(lido["itens"], quem)
    return caminho, avisos + obs, len(lido["itens"])


def processar(msg: dict):
    de = msg["de"]
    try:
        if msg["tipo"] not in ("text", "image", "document", "audio"):
            return
        if msg["tipo"] == "document" and "ESTOQUE" in msg["texto"].upper():
            if not autorizado(de, "NUMEROS_ESTOQUE"):
                enviar_texto(de, "Este número não pode atualizar o estoque. Fale com o gerente.")
                return
            try:
                n = atualizar_estoque(baixar_midia(msg["media_id"]), msg["arquivo"])
            except (ValueError, csv.Error) as e:
                enviar_texto(de, f"Estoque NÃO atualizado ({e}). O estoque anterior continua valendo.")
                return
            enviar_texto(de, f"Estoque atualizado: {n} produtos. Os próximos orçamentos já usam esses valores.")
            return
        blocos = conteudo_do_pedido(msg)
        if not blocos:
            return
        caminho, avisos, n = gerar_orcamento(blocos, de[-4:])
        if not caminho:
            enviar_texto(de, "Não encontrei itens nessa mensagem. " + " ".join(avisos))
            return
        enviar_documento(de, caminho, f"{n} itens. Confira antes de enviar ao cliente.")
        if avisos:
            enviar_texto(de, "Conferir:\n- " + "\n- ".join(avisos))
    except Exception as e:  # o vendedor precisa saber que falhou, senão fica esperando
        print(f"erro na mensagem {msg['id']}: {e!r}")
        enviar_texto(de, "Não consegui montar esse orçamento. Tente de novo ou faça manualmente.")


def usuario(cabecalho: str) -> str | None:
    """Login do site (HTTP Basic). VENDEDORES = "joao:senha1,maria:senha2"."""
    # shortcut: senhas em texto na variável de ambiente; com muitos usuários, guardar hash num arquivo
    if not cabecalho.startswith("Basic "):
        return None
    try:
        nome, _, senha = base64.b64decode(cabecalho[6:]).decode().partition(":")
    except ValueError:
        return None
    for par in os.environ.get("VENDEDORES", "").split(","):
        n, _, s = par.strip().partition(":")
        if n and hmac.compare_digest(n.encode(), nome.encode()) and hmac.compare_digest(s.encode(), senha.encode()):
            return n
    return None


LIMITE = 25 * 1024 * 1024  # maior arquivo aceito pelo site (o WhatsApp limita documentos a 100 MB)


class Servidor(BaseHTTPRequestHandler):
    def do_GET(self):
        q = {k: v[0] for k, v in parse_qs(urlparse(self.path).query).items()}
        if "hub.mode" in q:  # verificação do webhook pela Meta
            ok = q.get("hub.mode") == "subscribe" and hmac.compare_digest(
                q.get("hub.verify_token", ""), os.environ.get("WHATSAPP_VERIFY_TOKEN", "\0"))
            self._responder(200 if ok else 403, q.get("hub.challenge", "") if ok else "")
        elif urlparse(self.path).path == "/" and self._logado():
            self._responder(200, (BASE / "pagina.html").read_text(encoding="utf-8"), "text/html; charset=utf-8")
        elif urlparse(self.path).path != "/":
            self._responder(404, "")

    def do_POST(self):
        rota = urlparse(self.path).path
        if rota in ("/orcamento", "/estoque"):
            self._site(rota)
        else:
            self._webhook()

    def _logado(self) -> str | None:
        nome = usuario(self.headers.get("Authorization", ""))
        if not nome:
            self.send_response(401)
            self.send_header("WWW-Authenticate", 'Basic realm="Orcamentos", charset="UTF-8"')
            self.end_headers()
        return nome

    def _site(self, rota: str):
        nome = self._logado()
        if not nome:
            return
        tamanho = int(self.headers.get("Content-Length", 0))
        if tamanho > LIMITE:
            self._responder(413, "Arquivo grande demais (máximo 25 MB).")
            return
        dados = self.rfile.read(tamanho)
        arquivo = unquote(self.headers.get("X-Arquivo", ""))
        try:
            if rota == "/estoque":
                if not autorizado(nome, "GERENTES"):
                    self._responder(403, "Só o gerente pode atualizar o estoque.")
                    return
                n = atualizar_estoque(dados, arquivo)
                self._responder(200, f"Estoque atualizado: {n} produtos.")
                return
            blocos = blocos_de(dados, self.headers.get("Content-Type", ""), arquivo)
            if not blocos:
                self._responder(422, "Mande um arquivo ou o texto do pedido.")
                return
            caminho, avisos, n = gerar_orcamento(blocos, nome)
        except (ValueError, csv.Error) as e:
            self._responder(422, f"Não deu certo: {e}")
            return
        except Exception as e:  # erro inesperado (rede, API): o vendedor precisa saber
            print(f"erro no site ({nome}): {e!r}")
            self._responder(500, "Não consegui montar esse orçamento. Tente de novo ou faça manualmente.")
            return
        if not caminho:
            self._responder(422, "Não encontrei itens nesse pedido. " + " ".join(avisos))
            return
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{caminho.name}"')
        self.send_header("X-Avisos", quote(json.dumps(avisos, ensure_ascii=False)))
        self.send_header("X-Itens", str(n))
        self.end_headers()
        self.wfile.write(caminho.read_bytes())

    def _webhook(self):
        corpo = self.rfile.read(min(int(self.headers.get("Content-Length", 0)), LIMITE))
        if not assinatura_ok(corpo, self.headers.get("X-Hub-Signature-256", ""), os.environ.get("WHATSAPP_APP_SECRET", "")):
            self._responder(403, "")
            return
        self._responder(200, "ok")  # a Meta reenvia se demorar: responde já e processa depois
        for msg in extrair_mensagens(json.loads(corpo or b"{}")):
            if msg["id"] in _processadas or not autorizado(msg["de"]):
                continue
            _processadas.add(msg["id"])
            threading.Thread(target=processar, args=(msg,), daemon=True).start()

    def _responder(self, status: int, texto: str, tipo: str = "text/plain; charset=utf-8"):
        self.send_response(status)
        self.send_header("Content-Type", tipo)
        self.end_headers()
        self.wfile.write(texto.encode())

    def log_message(self, *args):
        pass


if __name__ == "__main__":
    porta = int(os.environ.get("PORTA", 8080))
    print(f"Servidor no ar: http://localhost:{porta}")
    ThreadingHTTPServer(("0.0.0.0", porta), Servidor).serve_forever()
