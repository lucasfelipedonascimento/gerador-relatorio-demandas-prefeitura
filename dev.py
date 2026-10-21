"""Servidor local: serve public/index.html e a API em /api/gerar.
Uso:  python dev.py   ->  http://localhost:5000
(Na Vercel este arquivo não é usado.)"""
import os

from flask import send_from_directory

from api.gerar import app

PUBLIC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "public")


@app.get("/")
def _index():
    return send_from_directory(PUBLIC, "index.html")


if __name__ == "__main__":
    app.run(debug=True, port=5000)
