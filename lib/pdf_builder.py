"""Gera o PDF a partir do .docx já preenchido.

Na Vercel não há Word nem LibreOffice, então o PDF é montado com ReportLab
percorrendo o corpo do próprio .docx (parágrafos, negritos, quebras de linha,
tabelas de fotos e imagens) na mesma ordem e com o mesmo tamanho.
"""
import io
import os
from xml.sax.saxutils import escape

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image as RLImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONTES = os.path.join(ROOT, "assets", "fonts")

# Carlito = clone métrico do Calibri (fonte do modelo)
pdfmetrics.registerFont(TTFont("Carlito", os.path.join(FONTES, "Carlito-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Carlito-Bold", os.path.join(FONTES, "Carlito-Bold.ttf")))
pdfmetrics.registerFontFamily(
    "Carlito", normal="Carlito", bold="Carlito-Bold", italic="Carlito", boldItalic="Carlito-Bold"
)
_GLIFOS = pdfmetrics.getFont("Carlito").face.charToGlyph

EMU_POR_PT = 12700


def _so_glifos_suportados(texto: str) -> str:
    """Descarta caracteres que a fonte não tem (ex.: emojis) para não gerar quadradinhos."""
    saida = []
    for ch in texto:
        if ch == "\n":
            saida.append(ch)
        elif ch in "\t\u00a0":
            saida.append(" ")
        elif ord(ch) in _GLIFOS:
            saida.append(ch)
    return "".join(saida)


def _marcacao(par: DocxParagraph) -> str:
    partes = []
    for run in par.runs:
        texto = _so_glifos_suportados(run.text)
        if not texto:
            continue
        texto = escape(texto).replace("\n", "<br/>")
        partes.append(f"<b>{texto}</b>" if run.bold else texto)
    return "".join(partes)


def _imagens_do_paragrafo(doc: Document, par: DocxParagraph):
    saida = []
    for run in par.runs:
        for blip in run._r.xpath(".//a:blip"):
            rid = blip.get(qn("r:embed"))
            ext = run._r.xpath(".//wp:extent")[0]
            w = int(ext.get("cx")) / EMU_POR_PT
            h = int(ext.get("cy")) / EMU_POR_PT
            blob = doc.part.related_parts[rid].blob
            saida.append(RLImage(io.BytesIO(blob), width=w, height=h))
    return saida


def docx_para_pdf(docx_bytes: bytes, titulo: str = "Solicitação") -> bytes:
    doc = Document(io.BytesIO(docx_bytes))

    estilo = ParagraphStyle(
        "corpo", fontName="Carlito", fontSize=12, leading=15.5, spaceAfter=8, alignment=TA_LEFT
    )
    estilo_centro = ParagraphStyle("centro", parent=estilo, alignment=TA_CENTER)

    sec = doc.sections[0]
    saida = io.BytesIO()
    pdf = SimpleDocTemplate(
        saida,
        pagesize=A4,
        leftMargin=sec.left_margin / EMU_POR_PT,
        rightMargin=sec.right_margin / EMU_POR_PT,
        topMargin=sec.top_margin / EMU_POR_PT,
        bottomMargin=sec.bottom_margin / EMU_POR_PT,
        title=titulo,
        author="Conselho Popular",
    )

    flow = []
    for filho in doc.element.body.iterchildren():
        if filho.tag == qn("w:p"):
            par = DocxParagraph(filho, doc)
            centro = par.alignment == WD_ALIGN_PARAGRAPH.CENTER
            est = estilo_centro if centro else estilo
            if par.paragraph_format.keep_with_next:
                est = ParagraphStyle("kwn", parent=est, keepWithNext=1)
            imagens = _imagens_do_paragrafo(doc, par)
            texto = _marcacao(par)
            if texto.strip():
                flow.append(Paragraph(texto, est))
            elif not imagens:
                flow.append(Spacer(1, 15.5 + 8))
            for img in imagens:
                img.hAlign = "CENTER" if centro else "LEFT"
                flow.append(img)
                flow.append(Spacer(1, 8))

        elif filho.tag == qn("w:tbl"):
            tabela = DocxTable(filho, doc)
            dados, larguras = [], None
            for row in tabela.rows:
                linha = []
                for cel in row.cells:
                    conteudo = []
                    for p in cel.paragraphs:
                        conteudo.extend(_imagens_do_paragrafo(doc, p))
                    linha.append(conteudo or "")
                dados.append(linha)
                if larguras is None:
                    larguras = [(c.width or 0) / EMU_POR_PT for c in row.cells]
            t = Table(dados, colWidths=larguras, hAlign="CENTER")
            t.setStyle(
                TableStyle(
                    [
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 5.4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 5.4),
                        ("TOPPADDING", (0, 0), (-1, -1), 0),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                    ]
                )
            )
            flow.append(t)

    pdf.build(flow)
    return saida.getvalue()
