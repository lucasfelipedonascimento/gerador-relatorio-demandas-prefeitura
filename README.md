# Gerador de Solicitação – Conselho Popular

Formulário web que preenche o modelo `.docx` e devolve **Word + PDF** para download.
Sem banco de dados: tudo é gerado em memória e entregue na resposta.

## Estrutura
```
api/gerar.py        função serverless (Flask) – POST /api/gerar
lib/validators.py   validação (CPF, CEP, telefone…) e formatação
lib/docx_builder.py preenche o modelo e insere as fotos
lib/pdf_builder.py  gera o PDF lendo o .docx já preenchido
assets/             modelo .docx + fontes Carlito (equivalente ao Calibri)
public/index.html   formulário (máscaras, ViaCEP, compressão de fotos)
dev.py              servidor local
vercel.json
```

## Rodar localmente
```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python dev.py          # http://localhost:5000
```

## Publicar na Vercel
```bash
npm i -g vercel
vercel            # primeira vez (preview)
vercel --prod
```
Ou envie a pasta para o GitHub e importe o repositório em vercel.com/new (sem configurar build).

## Alterando o modelo
Edite `assets/Modelo-DOC-Solicitacao-Prefeitura.docx` mantendo os marcadores `{{...}}`
(mapeados em `CAMPOS` no `lib/docx_builder.py`) e os dois marcadores de imagem
`{{#picture demandaImagens}}` / `{{#picture solicitacaoPrint}}` (texto alternativo da imagem).
Para adicionar outra região, inclua-a no `<select>` do `index.html` e em `REGIOES_PERMITIDAS` (`lib/validators.py`).

## Limites da Vercel (já tratados)
- Corpo da requisição/resposta ≤ 4,5 MB → as fotos são reduzidas no navegador e reotimizadas no servidor; máx. 10 fotos.
- Não há Word/LibreOffice na Vercel → o PDF é montado com ReportLab a partir do conteúdo do .docx.
