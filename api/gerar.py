"""Função serverless (Vercel): recebe o formulário e devolve DOCX + PDF em base64.

Sem banco e sem armazenamento: tudo é gerado em memória e devolvido na resposta.
"""

import base64
import os
import re
import sys
import unicodedata
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask, jsonify, request, send_from_directory  # noqa: E402

from lib.docx_builder import gerar_docx, preparar_imagem  # noqa: E402
from lib.pdf_builder import docx_para_pdf  # noqa: E402
from lib.validators import validar_formulario  # noqa: E402

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024  # Limite de 5 MB

MAX_FOTOS = 10
LIMITE_RESPOSTA = 3_200_000  # bytes
PRESETS = [(1100, 72), (900, 62), (720, 52), (600, 45)]

# Diretório público para o front-end na Vercel
PUBLIC_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "public")
)


@app.get("/")
def serve_index():
  return send_from_directory(PUBLIC_DIR, "index.html")


@app.get("/<path:path>")
def serve_static(path):
  if os.path.exists(os.path.join(PUBLIC_DIR, path)):
    return send_from_directory(PUBLIC_DIR, path)
  return "Not Found", 404


def _slug(texto: str) -> str:
  texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
  return re.sub(r"[^A-Za-z0-9]+", "_", texto).strip("_")[:40] or "documento"


def _arquivo_enviado(f) -> bytes | None:
  if f is None or not f.filename:
    return None
  blob = f.read()
  return blob or None


@app.errorhandler(413)
def _muito_grande(_):
  return (
      jsonify(error="As fotos ultrapassam o limite de envio. Envie menos fotos."),
      413,
  )


@app.post("/gerar")
def gerar():
  dados, erros = validar_formulario(request.form)

  fotos_raw = [
      b for b in (_arquivo_enviado(f) for f in request.files.getlist("fotos")) if b
  ]
  if not fotos_raw:
    erros["fotos"] = "Envie ao menos uma foto da demanda."
  elif len(fotos_raw) > MAX_FOTOS:
    erros["fotos"] = f"Máximo de {MAX_FOTOS} fotos."
  print_raw = _arquivo_enviado(request.files.get("print"))

  if erros:
    return jsonify(errors=erros), 400

  resultado = None
  try:
    for max_px, qualidade in PRESETS:
      fotos = [preparar_imagem(b, max_px, qualidade) for b in fotos_raw]
      print_app = (
          preparar_imagem(print_raw, max_px, qualidade) if print_raw else None
      )
      docx_bytes = gerar_docx(dados, fotos, print_app)
      pdf_bytes = docx_para_pdf(docx_bytes, f"Solicitação - {dados['nome']}")
      resultado = (docx_bytes, pdf_bytes)
      if len(docx_bytes) + len(pdf_bytes) <= LIMITE_RESPOSTA:
        break
  except OSError:  # PIL: arquivo corrompido ou inválido
    return (
        jsonify(
            errors={"fotos": "Uma das imagens é inválida ou está corrompida."}
        ),
        400,
    )

  docx_bytes, pdf_bytes = resultado
  if len(docx_bytes) + len(pdf_bytes) > LIMITE_RESPOSTA * 1.3:
    return (
        jsonify(error="Arquivos muito grandes. Reduza a quantidade de fotos."),
        413,
    )

  hoje = datetime.now(timezone(timedelta(hours=-3))).strftime("%Y%m%d")
  base = f"Solicitacao_{_slug(dados['nome'])}_{hoje}"
  b64 = lambda b: base64.b64encode(b).decode("ascii")  # noqa: E731
  return jsonify(
      docx={"filename": f"{base}.docx", "base64": b64(docx_bytes)},
      pdf={"filename": f"{base}.pdf", "base64": b64(pdf_bytes)},
  )


@app.errorhandler(Exception)
def _erro(e):
  from werkzeug.exceptions import HTTPException

  if isinstance(e, HTTPException):
    return jsonify(error=e.description), e.code
  app.logger.exception("Falha ao gerar documento")
  return jsonify(error="Não foi possível gerar o documento. Tente novamente."), 500