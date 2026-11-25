# Orçamento rápido (site + WhatsApp opcional)

O vendedor manda o pedido do cliente (texto, foto/print, PDF, planilha ou áudio) e recebe o orçamento em .xlsx, no modelo da empresa, com preço e estoque.
- **Site (principal):** abre o endereço no celular ou PC, entra com usuário e senha, arrasta/cola o pedido e baixa a planilha. Sem número extra, sem Meta.
- **WhatsApp (opcional):** encaminha o pedido para um número da empresa na **API oficial do WhatsApp**, sem QR Code nem risco de banimento.

## Arquivos
- `app.py`: servidor (site + webhook do WhatsApp; lê o pedido com o Claude, cruza com o estoque, gera a planilha)
- `pagina.html`: a página do site
- `config.json`: modelo do cliente (qual planilha, quais colunas, onde fica o estoque)
- `demo/`: dados fictícios (`estoque.csv`, `modelo.xlsx` gerado por `criar_modelo.py`)
- `teste.py`: teste local, sem WhatsApp e sem gastar API

## Testar no PC (grátis)
```
.venv\Scripts\python.exe teste.py
```
Abrir o site no próprio PC (precisa de `ANTHROPIC_API_KEY` para ler pedidos de verdade; poucos centavos por pedido):
```
$env:VENDEDORES="joao:senha1"; $env:GERENTES="joao"; .venv\Scripts\python.exe app.py
```
e entrar em http://localhost:8080 com joao / senha1.

## Colocar no ar para um cliente (site)
1. **Servidor:** uma VPS pequena com Linux, **2 GB de RAM ou mais** por causa da transcrição de áudio (R$ 40–80/mês). Um servidor atende vários clientes (uma cópia da pasta e uma porta por cliente).
2. **Domínio e HTTPS:** um domínio (.com.br, ~R$ 40/ano) e o Caddy na frente, que pega o certificado sozinho: `caddy reverse-proxy --from cliente.seu-dominio.com.br --to :8080`. **Sem HTTPS a senha viaja aberta: nunca rodar o site sem ele.**
3. **Variáveis de ambiente** (lista no topo do `app.py`): `ANTHROPIC_API_KEY`, `VENDEDORES` (um login por vendedor) e `GERENTES`.
4. `pip install -r requirements.txt` e `python app.py` como serviço do systemd (reinicia sozinho).
5. Mandar o endereço e os logins para os vendedores.

## WhatsApp (opcional, para quem quiser)
1. **Conta da Meta (no nome do cliente):** developers.facebook.com → criar app "Business" → produto **WhatsApp** → cadastrar e verificar um número (não pode estar em uso no app WhatsApp comum).
2. **Token permanente:** Configurações do negócio → Usuários do sistema → token com `whatsapp_business_messaging` e `whatsapp_business_management`.
3. Variáveis `WHATSAPP_*`, `NUMEROS_AUTORIZADOS` e `NUMEROS_ESTOQUE`.
4. **Webhook na Meta:** URL `https://cliente.seu-dominio.com.br/`, o mesmo verify token, campo **messages**.

## Adaptar para cada cliente
- Trocar `modelo_xlsx` pelo modelo dele e ajustar linhas/colunas no `config.json`.
- Estoque: o gerente envia a exportação do ERP (CSV ou XLSX) pelo site ("Atualizar estoque e preços") ou pelo WhatsApp com a palavra **ESTOQUE** na legenda. Em `cabecalhos_estoque`, pôr o nome que o ERP dá a cada coluna (ex.: `"preco": "Preço Venda"`; maiúsculas e acentos não importam). Arquivo sem alguma das 4 colunas é recusado e o estoque anterior continua; a versão anterior fica em `estoque.anterior.csv`.

## Custos
- Claude: Opus 5.5 com esforço baixo; cerca de R$ 0,10–0,30 por pedido (estimativa, medir com pedidos reais).
- VPS, domínio e backup.
- WhatsApp (só se usar): respostas dentro da janela de 24 h aberta pelo vendedor. Conferir a tabela atual da Meta.

## Limites conhecidos
- Login do site: senhas na variável de ambiente e sem bloqueio por tentativas erradas; bom para poucos vendedores.
- Áudio: transcrito no próprio servidor (faster-whisper, modelo small, ~0,5 GB baixado no primeiro uso). Um áudio por vez; códigos falados vêm na lista de "conferir".
- WhatsApp: controle de mensagens repetidas só em memória; não rodar duas cópias.
- O estoque é o último relatório enviado, não o ERP em tempo real.
- Sem consulta a fornecedores (ver os termos de uso de cada portal antes de automatizar).
