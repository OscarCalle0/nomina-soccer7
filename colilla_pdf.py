"""
colilla_pdf.py — Generador de colilla de pago en PDF
Replica exactamente el formato original de GRANDA VARGAS SAS / Soccer 7
2 copias por página tamaño carta
"""

import io
import os
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm, cm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

LOGO_PATH = os.path.join(os.path.dirname(__file__), "logo.png")

# Empresa
EMPRESA = {
    "nombre": "GRANDA VARGAS SAS",
    "nit":    "NIT: 901436091-0",
    "dir":    "DIR: CRA 51 45 -45",
    "tel":    "TEL: 8474747",
}

MESES_ES = {
    1:"enero", 2:"febrero", 3:"marzo", 4:"abril", 5:"mayo", 6:"junio",
    7:"julio", 8:"agosto", 9:"septiembre", 10:"octubre", 11:"noviembre", 12:"diciembre"
}
DIAS_ES = {
    0:"lunes", 1:"martes", 2:"miércoles", 3:"jueves",
    4:"viernes", 5:"sábado", 6:"domingo"
}

def cop(n):
    """Formatea número como COP sin decimales"""
    if n is None or n == "" or n == 0:
        return ""
    try:
        return f"{round(float(n)):,}".replace(",", ".")
    except:
        return str(n)

def num_fmt(n):
    """Formatea número con hasta 2 decimales"""
    if n is None or n == "" or n == 0:
        return ""
    try:
        f = float(n)
        if f == int(f):
            return f"{int(f)}"
        return f"{f:.2f}"
    except:
        return str(n)


def dibujar_colilla(c: canvas.Canvas, y_offset: float, datos: dict, p_ini: datetime, p_fin: datetime):
    """
    Dibuja UNA colilla en el canvas.
    y_offset: coordenada Y de inicio (en puntos desde abajo)
    La colilla ocupa aproximadamente 350pt de alto (media carta aprox)
    """
    W = letter[0]  # 612 pt
    M  = 18        # margen izquierdo/derecho en pt
    CW = W - 2*M   # ancho útil

    # Colores exactos de la colilla original
    AZUL_HEADER = colors.HexColor("#1B3A5C")
    GRIS_FILA   = colors.HexColor("#D9D9D9")
    NEGRO       = colors.black
    BLANCO      = colors.white

    # ── BORDE EXTERIOR ────────────────────────────────────────────────────────
    c.setStrokeColor(colors.HexColor("#888888"))
    c.setLineWidth(0.5)
    c.rect(M, y_offset, CW, 330, stroke=1, fill=0)

    # ── LOGO ──────────────────────────────────────────────────────────────────
    logo_w = 45
    logo_h = 52
    logo_x = M + 4
    logo_y = y_offset + 330 - logo_h - 4
    if os.path.exists(LOGO_PATH):
        try:
            c.drawImage(LOGO_PATH, logo_x, logo_y, width=logo_w, height=logo_h,
                       preserveAspectRatio=True, mask='auto')
        except:
            pass

    # ── ENCABEZADO EMPRESA ────────────────────────────────────────────────────
    enc_x = M + logo_w + 10
    enc_y = y_offset + 330 - 14
    c.setFont("Helvetica-Bold", 11)
    c.setFillColor(NEGRO)
    c.drawCentredString(W/2, enc_y, EMPRESA["nombre"])

    c.setFont("Helvetica", 8.5)
    c.drawCentredString(W/2, enc_y - 12, EMPRESA["nit"])
    c.drawCentredString(W/2, enc_y - 22, EMPRESA["dir"])
    c.drawCentredString(W/2, enc_y - 32, EMPRESA["tel"])

    # Fecha en letra (derecha)
    dia_nombre = DIAS_ES[p_fin.weekday()]
    mes_nombre = MESES_ES[p_fin.month]
    fecha_str  = f"{dia_nombre}, {p_fin.day} de {mes_nombre} de {p_fin.year}"
    c.setFont("Helvetica", 8)
    c.drawRightString(W - M - 4, enc_y - 32, fecha_str)

    # ── LÍNEA SEPARADORA ─────────────────────────────────────────────────────
    sep_y = y_offset + 330 - 60
    c.setStrokeColor(colors.HexColor("#888888"))
    c.setLineWidth(0.4)
    c.line(M, sep_y, W - M, sep_y)

    # ── PERÍODO ───────────────────────────────────────────────────────────────
    per_y = sep_y - 12
    c.setFont("Helvetica", 8.5)
    c.setFillColor(NEGRO)
    c.drawString(M + 4, per_y, f"Período")
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(M + 38, per_y, f"{p_ini.month}")
    c.drawString(M + 55, per_y, f"{p_ini.year}")
    c.setFont("Helvetica", 8.5)
    c.drawString(M + 75, per_y,
                 f"Nómina Entre  {p_ini.strftime('%d/%m/%Y')}  y   {p_fin.strftime('%d/%m/%Y')}")

    # ── EMPLEADO / CARGO ──────────────────────────────────────────────────────
    emp_y = per_y - 13
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(M + 4, emp_y, "EMPLEADO:")
    c.setFont("Helvetica", 8.5)
    c.drawString(M + 52, emp_y, datos.get("cedula", ""))
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(M + 105, emp_y, datos.get("nombre_completo", ""))
    c.setFont("Helvetica", 8.5)
    c.drawRightString(W - M - 4, emp_y,
                      f"SALARIO MENSUAL:   {cop(datos.get('salario', 0))}")

    cargo_y = emp_y - 11
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(M + 4, cargo_y, "CARGO:")
    c.setFont("Helvetica", 8.5)
    c.drawString(M + 38, cargo_y, datos.get("cargo", ""))

    # ── LÍNEA SEPARADORA ─────────────────────────────────────────────────────
    c.setLineWidth(0.4)
    c.line(M, cargo_y - 5, W - M, cargo_y - 5)

    # ── TABLA DE CONCEPTOS ────────────────────────────────────────────────────
    tabla_y = cargo_y - 8

    # Cabecera de tabla
    cols_x = [M+4, M+52, M+280, M+340, M+400, M+480, M+545]
    # CODIGO | DESCRIPCION | DOC | CANT | DEVENGADO | DEDUCCION | SALDO
    hdrs = ["CODIGO", "DESCRIPCION", "DOC", "CANT", "DEVENGADO", "DEDUCCION", "SALDO"]
    col_widths = [48, 228, 50, 60, 80, 80, 50]

    c.setFont("Helvetica-Bold", 7.5)
    c.setFillColor(NEGRO)

    # Dibujar cabecera
    hdr_h = 12
    c.setFillColor(colors.HexColor("#F0F0F0"))
    c.rect(M, tabla_y - hdr_h, CW, hdr_h, stroke=0, fill=1)
    c.setFillColor(NEGRO)
    c.setLineWidth(0.3)
    c.rect(M, tabla_y - hdr_h, CW, hdr_h, stroke=1, fill=0)

    hdr_labels = ["CODIGO", "DESCRIPCION", "DOC", "CANT", "DEVENGADO", "DEDUCCION", "SALDO"]
    hdr_align  = ["center", "left", "center", "center", "center", "center", "center"]
    x_pos = M
    for i, (lbl, w, ha) in enumerate(zip(hdr_labels, col_widths, hdr_align)):
        c.setFont("Helvetica-Bold", 7)
        if ha == "center":
            c.drawCentredString(x_pos + w/2, tabla_y - hdr_h + 3, lbl)
        else:
            c.drawString(x_pos + 3, tabla_y - hdr_h + 3, lbl)
        x_pos += w

    # Líneas verticales cabecera
    x_pos = M
    for w in col_widths[:-1]:
        x_pos += w
        c.line(x_pos, tabla_y - hdr_h, x_pos, tabla_y)

    # ── FILAS DE CONCEPTOS ────────────────────────────────────────────────────
    row_h = 11
    conceptos = datos.get("conceptos", [])
    fila_y = tabla_y - hdr_h

    for i, row in enumerate(conceptos):
        fy = fila_y - row_h
        # Fondo gris para deducciones
        if row.get("tipo") == "deduccion" and row.get("deduccion"):
            c.setFillColor(colors.HexColor("#F5F5F5"))
            c.rect(M, fy, CW, row_h, stroke=0, fill=1)

        c.setFillColor(NEGRO)
        c.setFont("Helvetica", 7.5)

        vals = [
            str(row.get("codigo", "")),
            str(row.get("descripcion", "")),
            num_fmt(row.get("doc", "")),
            num_fmt(row.get("cant", "")),
            cop(row.get("devengado", "")),
            cop(row.get("deduccion", "")),
            cop(row.get("saldo", "")),
        ]
        aligns = ["center","left","center","right","right","right","right"]

        x_pos = M
        for j, (val, w, ha) in enumerate(zip(vals, col_widths, aligns)):
            if val and val != "0" and val != "":
                if ha == "center":
                    c.drawCentredString(x_pos + w/2, fy + 2.5, val)
                elif ha == "right":
                    c.drawRightString(x_pos + w - 3, fy + 2.5, val)
                else:
                    c.drawString(x_pos + 3, fy + 2.5, val)
            x_pos += w

        # Línea inferior fila
        c.setLineWidth(0.2)
        c.setStrokeColor(colors.HexColor("#CCCCCC"))
        c.line(M, fy, W - M, fy)
        # Líneas verticales
        x_pos = M
        for w in col_widths[:-1]:
            x_pos += w
            c.setStrokeColor(colors.HexColor("#CCCCCC"))
            c.line(x_pos, fy, x_pos, fila_y - hdr_h)

        fila_y = fy
        c.setStrokeColor(colors.HexColor("#888888"))

    # Borde tabla
    c.setLineWidth(0.4)
    c.rect(M, fila_y, CW, tabla_y - fila_y, stroke=1, fill=0)

    tot_y = fila_y

    # ── FILA TOTALES ──────────────────────────────────────────────────────────
    tot_row_y = tot_y - row_h
    c.setFillColor(colors.HexColor("#E8E8E8"))
    c.rect(M, tot_row_y, CW, row_h, stroke=0, fill=1)
    c.setFillColor(NEGRO)
    c.setFont("Helvetica-Bold", 7.5)

    # "TOTALES" centrado en col descripcion
    c.drawCentredString(M + col_widths[0] + col_widths[1]/2, tot_row_y + 2.5, "TOTALES")

    # Total devengado
    total_dev = datos.get("total_devengado", 0)
    total_ded = datos.get("total_deducciones", 0)
    x_dev = M + sum(col_widths[:4])
    x_ded = M + sum(col_widths[:5])
    c.drawRightString(x_dev + col_widths[4] - 3, tot_row_y + 2.5, cop(total_dev))
    c.drawRightString(x_ded + col_widths[5] - 3, tot_row_y + 2.5, cop(total_ded))

    c.setLineWidth(0.4)
    c.setStrokeColor(colors.HexColor("#888888"))
    c.rect(M, tot_row_y, CW, row_h, stroke=1, fill=0)

    # ── ESPACIO EN BLANCO (para firma) ────────────────────────────────────────
    blank_h = 16
    blank_y = tot_row_y - blank_h

    # ── NETO A PAGAR ─────────────────────────────────────────────────────────
    neto_h = 14
    neto_y = blank_y - neto_h
    c.setFillColor(NEGRO)
    c.setFont("Helvetica-Bold", 8.5)
    # Alineado a la derecha como en el original
    neto_label_x = M + sum(col_widths[:4]) + 10
    c.drawRightString(W - M - col_widths[6] - 3, neto_y + 3, "NETO A PAGAR")
    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(W - M - 4, neto_y + 3, cop(datos.get("neto_a_pagar", 0)))

    c.setLineWidth(0.4)
    c.line(M, neto_y, W - M, neto_y)
    c.line(M, neto_y + neto_h, W - M, neto_y + neto_h)

    # ── PIE: NOMBRE, CC, CUENTA ───────────────────────────────────────────────
    pie_y = neto_y - 12
    c.setFont("Helvetica-Bold", 8)
    c.drawString(M + 4, pie_y, datos.get("nombre_completo", ""))

    cc_y = pie_y - 11
    c.setFont("Helvetica-Bold", 8)
    c.drawString(M + 4, cc_y, "C.C")
    c.setFont("Helvetica", 8)
    c.drawString(M + 22, cc_y, datos.get("cedula", ""))

    cuenta_y = cc_y - 10
    c.setFont("Helvetica-Bold", 8)
    c.drawString(M + 4, cuenta_y, "Cuenta:")
    c.setFont("Helvetica", 8)
    banco_txt = datos.get("banco", "")
    cuenta_txt = datos.get("cuenta", "")
    c.drawString(M + 42, cuenta_y, banco_txt)
    if cuenta_txt:
        c.drawString(M + 42 + c.stringWidth(banco_txt, "Helvetica", 8) + 10,
                     cuenta_y, cuenta_txt)

    # Línea de firma sobre el nombre
    c.setLineWidth(0.5)
    c.line(M + 4, pie_y + 10, M + 180, pie_y + 10)


def generar_colilla_pdf(lista_datos: list, p_ini: datetime, p_fin: datetime) -> bytes:
    """
    Genera un PDF con todas las colillas — 2 por página (carta).
    lista_datos: lista de dicts con info de cada colaborador
    Retorna bytes del PDF.
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    W, H = letter  # 612 x 792

    for i, datos in enumerate(lista_datos):
        # Colilla superior: y_offset = 396 + algunos puntos
        # Colilla inferior: y_offset = 20
        for copia in range(2):  # 2 copias por persona por página
            if copia == 0:
                y_off = H/2 + 5   # colilla superior
            else:
                y_off = 15        # colilla inferior

            dibujar_colilla(c, y_off, datos, p_ini, p_fin)

        # Línea punteada de corte en el centro
        c.setDash(3, 3)
        c.setStrokeColor(colors.HexColor("#999999"))
        c.setLineWidth(0.4)
        c.line(10, H/2, W - 10, H/2)
        c.setDash()  # reset

        # Leyenda de corte
        c.setFont("Helvetica", 6)
        c.setFillColor(colors.HexColor("#999999"))
        c.drawCentredString(W/2, H/2 + 2, "✂ — — — — — — — — — — — — — — — — — Cortar aquí — — — — — — — — — — — — — — — — — ✂")

        c.showPage()  # Nueva página por persona

    c.save()
    buf.seek(0)
    return buf.read()


def calcular_conceptos_colilla(t: dict, col_data: dict) -> dict:
    """
    Construye el dict de datos para la colilla a partir de los resultados
    del motor de nómina.
    """
    sal = col_data.get("salario_mensual", 0)
    tipo = col_data.get("tipo", "empleado")
    dias_trab = t.get("dias_trab", 15)

    AUX_Q = 249095 / 2  # auxilio transporte quincena 2026
    PENSION_PCT = 0.04
    SALUD_PCT   = 0.04

    if tipo == "prestador":
        val_hora = col_data.get("valor_hora_prestador", 10000)
        tot_dev = t.get("val_total_prest", 0)
        conceptos = [
            {"codigo":"001","descripcion":"HORAS TRABAJADAS","cant":round(t.get("tot_h",0),2),"devengado":tot_dev,"tipo":"devengado"},
            {"codigo":"002","descripcion":"AUXILIO DE TRANSPORTE","cant":0,"devengado":0,"tipo":"devengado"},
            {"codigo":"003","descripcion":"AUXILIO  ","cant":0,"devengado":0,"tipo":"devengado"},
            {"codigo":"004","descripcion":"RECARGO NOCTURNO","cant":0,"devengado":0,"tipo":"devengado"},
            {"codigo":"121","descripcion":"DEDUCCION DE PRESTAMOS","deduccion":0,"tipo":"deduccion"},
            {"codigo":"123","descripcion":"DEDUCCION PENSION","deduccion":0,"tipo":"deduccion"},
            {"codigo":"127","descripcion":"DEDUCCION SALUD","deduccion":0,"tipo":"deduccion"},
        ]
        total_dev = tot_dev
        total_ded = 0
        neto = tot_dev
    else:
        sal_q     = sal / 2
        aux_transp = AUX_Q * dias_trab / 15
        dev_ext   = t.get("val_en", 0)
        dev_noct  = t.get("val_noct", 0)
        ibc       = sal_q + dev_ext + dev_noct
        pension   = ibc * PENSION_PCT
        salud     = ibc * SALUD_PCT

        # Novedades devengadas y deducidas
        nov_dev = t.get("nov_devengado", 0)
        nov_ded = t.get("nov_deduccion", 0)

        total_dev = sal_q + aux_transp + dev_ext + dev_noct + nov_dev
        total_ded = pension + salud + nov_ded
        neto      = total_dev - total_ded

        conceptos = [
            {"codigo":"001","descripcion":"SALARIO BASICO",
             "cant":110,"devengado":round(sal_q),"tipo":"devengado"},
            {"codigo":"002","descripcion":"AUXILIO DE TRANSPORTE",
             "cant":dias_trab,"devengado":round(aux_transp),"tipo":"devengado"},
            {"codigo":"003","descripcion":"AUXILIO  ",
             "cant":round(t.get("en_h",0),4),"devengado":round(dev_ext) if dev_ext else "","tipo":"devengado"},
            {"codigo":"004","descripcion":"RECARGO NOCTURNO",
             "cant":round(t.get("noct_h",0),2),"devengado":round(dev_noct) if dev_noct else "","tipo":"devengado"},
        ]

        # Agregar novedades devengadas
        cod = 5
        for nd in t.get("nov_detalle", []):
            if nd.get("devengado", 0) > 0:
                desc = nd.get("desc", "").split(" —")[0].upper()
                if nd.get("pct"): desc += f" {nd['pct']:.2f}%"
                conceptos.append({
                    "codigo": f"00{cod}", "descripcion": desc,
                    "cant": nd.get("dias", ""), "devengado": round(nd["devengado"]),
                    "tipo": "devengado"
                })
                cod += 1

        conceptos += [
            {"codigo":"121","descripcion":"DEDUCCION DE PRESTAMOS","deduccion":0,"tipo":"deduccion"},
            {"codigo":"123","descripcion":f"DEDUCCION PENSION PROTECCIÓN 4%","deduccion":round(pension),"tipo":"deduccion"},
            {"codigo":"127","descripcion":f"DEDUCCION SALUD 4% {col_data.get('eps','').upper()}","deduccion":round(salud),"tipo":"deduccion"},
        ]

        # Novedades deducidas
        cod_ded = 130
        for nd in t.get("nov_detalle", []):
            if nd.get("deduccion", 0) > 0:
                desc = nd.get("desc", "").split(" —")[0].upper()
                conceptos.append({
                    "codigo": str(cod_ded), "descripcion": f"DEDUCCION {desc}",
                    "cant": nd.get("dias", ""), "deduccion": round(nd["deduccion"]),
                    "tipo": "deduccion"
                })
                cod_ded += 1

    return {
        "cedula":          col_data.get("id", ""),
        "nombre_completo": col_data.get("nombre_completo", ""),
        "cargo":           col_data.get("cargo", ""),
        "salario":         sal,
        "banco":           col_data.get("banco", ""),
        "cuenta":          col_data.get("cuenta", ""),
        "eps":             col_data.get("eps", ""),
        "conceptos":       conceptos,
        "total_devengado": round(total_dev),
        "total_deducciones": round(total_ded),
        "neto_a_pagar":    round(neto),
    }
