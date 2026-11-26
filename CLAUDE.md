# FastQuote — contexto do projeto

Negócio próprio do usuário (fora do emprego atual): vender a distribuidoras de peças do ramo automotivo em geral (linha leve e pesada, máquinas, agrícola, empilhadeira) um sistema que lê o pedido do cliente e devolve o orçamento pronto em Excel, no modelo da distribuidora, com preço e estoque.

## Nome e site
- Marca: **FastQuote** (26/11/2026; substituiu QuoteDesk, que tinha conflito nos EUA). Em 26/11: nada no INPI; fastquote.com.br livre (domínio para começar); fastquote.com e getfastquote.com parados e à venda na GoDaddy; USPTO: marcas FASTQUOTE de software todas mortas, viva só "FARMERS FASTQUOTE" (classe 36, seguros); nenhum software de cotação chamado FastQuote. Ponto fraco: "fast quote" é expressão comum (Google difícil, INPI pode achar descritivo): pedir marca mista (nome + logo) e usar sempre a frase "orçamento automático de peças". Para o exterior, escolher depois um domínio internacional (ex.: fastquote.app) apontando para o mesmo sistema.
- Logo: folha de orçamento inclinada com raio e rastro de velocidade, animada (movimento contínuo) no topo do site. Arquivos em `layouts/marca/` (`logo.svg` parado, `logo-animado-loop.svg`, `logo-animado-entrada.svg`).
- Nome em inglês, mas o texto do site na língua de cada cliente (pt-BR primeiro).
- Site de apresentação: `layouts/a-catalogo.html` (os outros layouts são rascunhos descartados).

## Regras que não mudam
- Nada da empresa onde o usuário trabalha (Elson, peças de empilhadeira): nenhum código, conta, senha, catálogo ou cliente de lá. Não vender para concorrentes diretos dela.
- Chaves e senhas só em variáveis de ambiente ou `.env` local, nunca no repositório.
- Dados de um cliente nunca aparecem para outro (base compartilhada futura: só dados técnicos de peças).
- Modelo do Claude: manter `claude-opus-5-5` por enquanto (decisão do usuário); trocar só depois de medir com pedidos reais.
- Prioridades do usuário: praticidade para quem usa (qualquer pessoa, sem baixar arquivo) e gastar menos token sem perder qualidade.

## O que existe (testado só localmente, com o Claude simulado)
- `app.py`: um servidor que atende o **site** (canal principal, login HTTP Basic por vendedor) e o **WhatsApp oficial da Meta** (opcional). Lê texto, foto, PDF, planilha e áudio (faster-whisper offline), extrai itens com o Claude (saída em JSON schema), cruza com o estoque (`config.json` → CSV) e preenche o modelo .xlsx. O gerente atualiza o estoque mandando a exportação do ERP (site ou WhatsApp com "ESTOQUE"), validada antes de gravar.
- `pagina.html`: a página do site. `teste.py`: testes (rodar sempre depois de mudar algo). `LEIAME.md`: como rodar e colocar no ar.
- Docs na conta claude.ai do usuário: guia do cliente (https://claude.ai/artifact/WytSXj9camtCsea1CooT9f) e plano de lançamento (https://claude.ai/artifact/5KxrHd3SN7eRyfn8nzJK4a).

## Próximos passos combinados
1. **Compartilhar → Orçamento** direto do WhatsApp: PWA com share target (Android) e Atalho da Apple (iPhone).
2. **Busca de peça em níveis:** cadastro do cliente → base de conhecimento (cache das peças já confirmadas) → catálogos do cliente convertidos em tabela → pesquisa web (máx. 3 buscas por item).
3. **Veracidade:** código só aceito se estiver escrito na página aberta (conferido por código), fonte do fabricante ou 2 fontes; marcar Confirmado / Provável / Não encontrado com link; vendedor confirma em um toque e isso alimenta a base.
4. **Planos e contador de uso** (Essencial / Profissional / Avançado; franquia de orçamentos e de itens pesquisados).
5. Piloto com 1–2 distribuidoras medindo acerto, minutos economizados e custo por pedido.

## Ainda não verificado
Leitura real pelo Claude (falta chave da API), envio real pelo WhatsApp (falta conta Meta), áudio OGG real do WhatsApp, site no celular, exportação de ERP real.
