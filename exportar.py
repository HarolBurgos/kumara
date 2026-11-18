"""
exportar.py  -  Kumara
Genera la cotización en Word (.docx) y PDF a partir del resultado de
modelos.armar_cotizacion(). Reemplaza las macros del Excel.
"""
import io
import os
from datetime import datetime

from modelos import cop, cant, num, SI, NO

AZUL = "193657"
TURQUESA = "087E8B"
VERDE = "DDF5E5"
CLARO = "F4F7FA"

LOGO = os.path.join("images", "logo_kumara.png")   # opcional


# -----------------------------------------------------------------------------
#  Datos comunes a ambos formatos
# -----------------------------------------------------------------------------
def encabezado(evento, cot):
    horas = (f"{num(evento.get('h_inicio')):g}:00 a {num(evento.get('h_fin')):g}:00 "
             f"({cot['duracion']:g} horas)")
    return [
        ("Empresa", evento.get("empresa") or "-"),
        ("Contacto", evento.get("contacto") or "-"),
        ("Ciudad", evento.get("ciudad") or "-"),
        ("Fecha", evento.get("fecha") or "-"),
        ("Lugar", evento.get("lugar") or "-"),
        ("Dirección", evento.get("direccion") or "-"),
        ("Asistentes", cant(evento.get("asistentes"))),
        ("Horario", horas),
        ("Nivel de brigada", evento.get("nivel") or "-"),
    ]


def filas_pista(pista):
    """Solo lo que realmente se cotiza (cantidad > 0 y valor > 0), como la macro original."""
    filas = []
    for l in pista["lineas"]:
        if l["incluir"] and l["a_cotizar"] > 0 and l["valor"] > 0:
            nombre = l["nombre"] + (f"  ({l['obs']})" if l["obs"] else "")
            filas.append([nombre, cant(l["a_cotizar"]), cop(l["valor"]), cop(l["subtotal"])])
    return filas


def filas_alimentacion(cot):
    filas = []
    for f in cot["alimentacion"]["filas"]:
        if f["incluir"] and f["total"] > 0:
            filas.append([f["servicio"], f["menu"], cant(f["unidades"]),
                          cop(f["valor"]), cop(f["domicilio"]), cop(f["total"])])
    return filas


def filas_transporte(cot):
    return [[f["vehiculo"], cant(f["a_cotizar"]), cop(f["valor"]), cop(f["subtotal"])]
            for f in cot["transporte"]["filas"] if f["a_cotizar"] > 0 and f["valor"] > 0]


def filas_resumen(cot):
    r = cot["resumen"]
    filas = [["Capacitadores / profesionales", cop(r["profesionales"])],
             ["Otros recursos de las pistas", cop(r["otros_recursos"])],
             ["Alimentación", cop(r["alimentacion"])],
             ["Transporte", cop(r["transporte"])],
             ["SUBTOTAL", cop(r["subtotal"])]]
    if r["iva_incluido"]:
        filas.append([f"IVA ({r['iva_pct']:g}%)", cop(r["iva"])])
    filas.append(["TOTAL GENERAL", cop(r["total"])])
    return filas


def nombre_archivo(evento, extension):
    base = (evento.get("empresa") or "evento").strip().replace(" ", "_")
    base = "".join(c for c in base if c.isalnum() or c in "_-") or "evento"
    fecha = (evento.get("fecha") or datetime.now().strftime("%Y-%m-%d")).replace("/", "-")
    return f"Kumara_Cotizacion_{base}_{fecha}.{extension}"


# -----------------------------------------------------------------------------
#  WORD
# -----------------------------------------------------------------------------
def generar_word(evento, cot):
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor, Cm

    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2)
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10)

    def sombrear(celda, color):
        tcPr = celda._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), color)
        tcPr.append(shd)

    def titulo(texto, tam=13, color=AZUL, espacio=10):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(espacio)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(texto)
        r.bold = True
        r.font.size = Pt(tam)
        r.font.color.rgb = RGBColor.from_string(color)
        return p

    def tabla(encabezados, filas, anchos=None, derecha=(), negrita_ultima=False):
        t = doc.add_table(rows=1, cols=len(encabezados))
        t.style = "Table Grid"
        t.autofit = False
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(encabezados):
            c = t.rows[0].cells[i]
            c.text = ""
            run = c.paragraphs[0].add_run(h)
            run.bold = True
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(255, 255, 255)
            sombrear(c, TURQUESA)
        for n, fila in enumerate(filas):
            celdas = t.add_row().cells
            for i, v in enumerate(fila):
                celdas[i].text = ""
                run = celdas[i].paragraphs[0].add_run(str(v))
                run.font.size = Pt(9)
                if negrita_ultima and n == len(filas) - 1:
                    run.bold = True
                    sombrear(celdas[i], VERDE)
                if i in derecha:
                    celdas[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if anchos:
            for fila in t.rows:
                for i, a in enumerate(anchos):
                    fila.cells[i].width = Cm(a)
        for i in derecha:
            t.rows[0].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        return t

    # Encabezado
    if os.path.exists(LOGO):
        doc.add_picture(LOGO, height=Cm(1.6))
    titulo("KUMARA", tam=22, espacio=0)
    p = doc.add_paragraph("Cotización de evento de emergencias")
    p.runs[0].font.size = Pt(12)
    p.runs[0].font.color.rgb = RGBColor.from_string(TURQUESA)

    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    t.autofit = False
    for k, v in encabezado(evento, cot):
        celdas = t.add_row().cells
        celdas[0].text = ""
        r = celdas[0].paragraphs[0].add_run(k)
        r.bold = True
        r.font.size = Pt(9)
        sombrear(celdas[0], CLARO)
        celdas[1].text = ""
        celdas[1].paragraphs[0].add_run(str(v)).font.size = Pt(9)
        celdas[0].width, celdas[1].width = Cm(4), Cm(13)

    # Pistas activas
    titulo("1. Pistas de emergencia")
    activas = [p for p in cot["pistas"] if p["activa"]]
    if not activas:
        doc.add_paragraph("No se seleccionaron pistas.")
    for pista in activas:
        filas = filas_pista(pista)
        titulo(pista["nombre"], tam=11, color=TURQUESA, espacio=6)
        if filas:
            filas.append(["Subtotal pista", "", "", cop(pista["subtotal"])])
            tabla(["Recurso", "Cantidad", "Valor unitario", "Valor total"], filas,
                  anchos=[8.5, 2.2, 3.2, 3.2], derecha=(1, 2, 3), negrita_ultima=True)
        else:
            doc.add_paragraph("Sin recursos con valor para cotizar.")

    # Alimentación
    titulo("2. Alimentación")
    filas = filas_alimentacion(cot)
    if filas:
        filas.append(["Total alimentación", "", "", "", "", cop(cot["alimentacion"]["total"])])
        tabla(["Servicio", "Menú", "Cantidad", "Valor unit.", "Domicilio", "Total"], filas,
              anchos=[3.2, 4.3, 2, 2.6, 2.4, 2.6], derecha=(2, 3, 4, 5), negrita_ultima=True)
    else:
        doc.add_paragraph("No se incluye alimentación.")

    # Transporte
    titulo("3. Transporte")
    filas = filas_transporte(cot)
    if filas:
        filas.append(["Total transporte", "", "", cop(cot["transporte"]["total"])])
        tabla(["Vehículo", "Cantidad", "Valor unitario", "Subtotal"], filas,
              anchos=[6, 3, 4, 4], derecha=(1, 2, 3), negrita_ultima=True)
    else:
        doc.add_paragraph("No se incluye transporte.")

    # Resumen
    titulo("4. Resumen económico")
    tabla(["Concepto", "Valor"], filas_resumen(cot), anchos=[11, 6], derecha=(1,),
          negrita_ultima=True)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    r = p.add_run("Cotización generada con Kumara el " +
                  datetime.now().strftime("%d/%m/%Y %H:%M") +
                  ". Los valores están expresados en pesos colombianos (COP).")
    r.italic = True
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# -----------------------------------------------------------------------------
#  PDF
# -----------------------------------------------------------------------------
def generar_pdf(evento, cot):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (Image, Paragraph, SimpleDocTemplate, Spacer,
                                    Table, TableStyle)

    c_azul = colors.HexColor("#" + AZUL)
    c_turq = colors.HexColor("#" + TURQUESA)
    c_verde = colors.HexColor("#" + VERDE)
    c_claro = colors.HexColor("#" + CLARO)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=1.8*cm, bottomMargin=1.8*cm,
                            title="Cotización Kumara")
    h1 = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=20, leading=26,
                        textColor=c_azul, spaceAfter=2)
    sub = ParagraphStyle("sub", fontName="Helvetica", fontSize=11, textColor=c_turq,
                         spaceAfter=8)
    h2 = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=12, textColor=c_azul,
                        spaceBefore=10, spaceAfter=4)
    h3 = ParagraphStyle("h3", fontName="Helvetica-Bold", fontSize=10, textColor=c_turq,
                        spaceBefore=6, spaceAfter=3)
    normal = ParagraphStyle("n", fontName="Helvetica", fontSize=9)
    pie = ParagraphStyle("pie", fontName="Helvetica-Oblique", fontSize=7.5,
                         textColor=colors.grey, spaceBefore=12)
    celda = ParagraphStyle("c", fontName="Helvetica", fontSize=8.5, leading=10)

    def tabla(encabezados, filas, anchos, derecha=(), ultima_resaltada=False):
        datos = [encabezados] + [[Paragraph(str(v), celda) if i == 0 else str(v)
                                  for i, v in enumerate(f)] for f in filas]
        t = Table(datos, colWidths=[a * cm for a in anchos], repeatRows=1)
        estilo = [
            ("BACKGROUND", (0, 0), (-1, 0), c_turq),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#C9D3DC")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, c_claro]),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]
        for i in derecha:
            estilo.append(("ALIGN", (i, 0), (i, -1), "RIGHT"))
        if ultima_resaltada:
            n = len(datos) - 1
            estilo += [("BACKGROUND", (0, n), (-1, n), c_verde),
                       ("FONTNAME", (0, n), (-1, n), "Helvetica-Bold")]
        t.setStyle(TableStyle(estilo))
        return t

    s = []
    if os.path.exists(LOGO):
        s.append(Image(LOGO, height=1.6*cm, width=1.6*cm, kind="proportional"))
    s += [Paragraph("KUMARA", h1), Paragraph("Cotización de evento de emergencias", sub)]

    datos = [[Paragraph(f"<b>{k}</b>", celda), Paragraph(str(v), celda)]
             for k, v in encabezado(evento, cot)]
    t = Table(datos, colWidths=[4*cm, 13*cm])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#C9D3DC")),
                           ("BACKGROUND", (0, 0), (0, -1), c_claro),
                           ("TOPPADDING", (0, 0), (-1, -1), 3),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
    s.append(t)

    s.append(Paragraph("1. Pistas de emergencia", h2))
    activas = [p for p in cot["pistas"] if p["activa"]]
    if not activas:
        s.append(Paragraph("No se seleccionaron pistas.", normal))
    for pista in activas:
        s.append(Paragraph(pista["nombre"], h3))
        filas = filas_pista(pista)
        if filas:
            filas.append(["Subtotal pista", "", "", cop(pista["subtotal"])])
            s.append(tabla(["Recurso", "Cantidad", "Valor unitario", "Valor total"], filas,
                           [8.3, 2.2, 3.3, 3.2], derecha=(1, 2, 3), ultima_resaltada=True))
        else:
            s.append(Paragraph("Sin recursos con valor para cotizar.", normal))

    s.append(Paragraph("2. Alimentación", h2))
    filas = filas_alimentacion(cot)
    if filas:
        filas.append(["Total alimentación", "", "", "", "", cop(cot["alimentacion"]["total"])])
        s.append(tabla(["Servicio", "Menú", "Cant.", "Valor unit.", "Domicilio", "Total"], filas,
                       [3.2, 4.2, 1.8, 2.7, 2.4, 2.7], derecha=(2, 3, 4, 5),
                       ultima_resaltada=True))
    else:
        s.append(Paragraph("No se incluye alimentación.", normal))

    s.append(Paragraph("3. Transporte", h2))
    filas = filas_transporte(cot)
    if filas:
        filas.append(["Total transporte", "", "", cop(cot["transporte"]["total"])])
        s.append(tabla(["Vehículo", "Cantidad", "Valor unitario", "Subtotal"], filas,
                       [6, 3.3, 4, 3.7], derecha=(1, 2, 3), ultima_resaltada=True))
    else:
        s.append(Paragraph("No se incluye transporte.", normal))

    s.append(Paragraph("4. Resumen económico", h2))
    s.append(tabla(["Concepto", "Valor"], filas_resumen(cot), [11, 6], derecha=(1,),
                   ultima_resaltada=True))

    s.append(Paragraph("Cotización generada con Kumara el " +
                       datetime.now().strftime("%d/%m/%Y %H:%M") +
                       ". Valores en pesos colombianos (COP).", pie))
    doc.build(s)
    return buf.getvalue()
