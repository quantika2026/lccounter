"""Genera el flujo del proceso como gráfica editable (PowerPoint con formas y conectores nativos
y archivo draw.io) a partir de la hoja Actividades de la plantilla de levantamiento.

Uso: python flujo.py plantilla.xlsx salida.pptx salida.drawio
"""
import html
import sys

import openpyxl
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

XLSX, PPTX, DRAWIO = sys.argv[1:4]
# opcionales: --vista "PROCESO ACTUAL (AS-IS)" para el sello y el título
VISTA = ""
if "--vista" in sys.argv:
    VISTA = sys.argv[sys.argv.index("--vista") + 1]

# ------------------------------------------------------------------ datos
wb = openpyxl.load_workbook(XLSX, data_only=True)
roles = [wb["Roles"][f"A{r}"].value for r in range(5, 13) if wb["Roles"][f"A{r}"].value]
pasos = []
ws = wb["Actividades"]
for r in range(5, 45):
    if not ws[f"B{r}"].value:
        continue
    pasos.append({k: ws[f"{c}{r}"].value for k, c in zip(
        ["orden", "id", "etapa", "act", "rol", "tipo", "sig", "no", "nat", "estado"], "ABCDEFJKNQ")})
pasos.sort(key=lambda p: p["orden"])
por_id = {p["id"]: p for p in pasos}
hipotesis = any(p["estado"] != "Validado" for p in pasos)

NAVY = RGBColor(0x1F, 0x38, 0x64)
GRIS = RGBColor(0x59, 0x59, 0x59)
LINEA = RGBColor(0x5B, 0x6B, 0x86)
COLORES = {"Dato": ("D6E4F5", "2F5597"), "Criterio": ("FCE4D6", "C55A11"), "Mixto": ("E7E0F0", "7030A0")}
SIN = ("F2F2F2", "7F7F7F")


def rgb(h):
    return RGBColor.from_string(h)


def sin_estilo(sh):
    sh.shadow.inherit = False
    st = sh._element.find(qn("p:style"))
    if st is not None:
        st.getparent().remove(st)


# ------------------------------------------------------------------ PowerPoint
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
blank = prs.slide_layouts[6]


def texto(s, x, y, w, h, t, size=10, bold=False, color=GRIS, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, linea in enumerate(t.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = linea
        r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(size), bold, color, "Calibri"
    return tb


def punteada(sh, p):
    """Borde punteado para los pasos que aún son hipótesis."""
    if p.get("estado") != "Validado":
        ln = sh.line._get_or_add_ln()
        d = etree.SubElement(ln, qn("a:prstDash"))
        d.set("val", "dash")
    return sh


def forma(s, tipo, x, y, w, h, fill, borde, lineas, size=8.5, nombre=None):
    sh = s.shapes.add_shape(tipo, Inches(x), Inches(y), Inches(w), Inches(h))
    if nombre:
        sh.name = nombre
    sh.fill.solid()
    sh.fill.fore_color.rgb = rgb(fill)
    sh.line.color.rgb = rgb(borde)
    sh.line.width = Pt(1.25)
    sh.shadow.inherit = False
    st = sh._element.find(qn("p:style"))
    if st is not None:
        st.getparent().remove(st)
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, (t, b) in enumerate(lineas):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = t
        r.font.size, r.font.bold, r.font.name = Pt(size if not b else size + 0.5), b, "Calibri"
        r.font.color.rgb = RGBColor(0x1F, 0x1F, 0x1F)
    return sh


def flecha(s, a, b, ia, ib, color=LINEA, etiqueta=None, discontinua=False):
    c = s.shapes.add_connector(MSO_CONNECTOR.ELBOW, 0, 0, Inches(1), Inches(1))
    c.begin_connect(a, ia)
    c.end_connect(b, ib)
    c.line.color.rgb = color
    c.line.width = Pt(1.25)
    ln = c.line._get_or_add_ln()
    if discontinua:
        d = etree.SubElement(ln, qn("a:prstDash"))
        d.set("val", "dash")
    tail = etree.SubElement(ln, qn("a:tailEnd"))
    tail.set("type", "triangle")
    tail.set("w", "med")
    tail.set("len", "med")
    st = c._element.find(qn("p:style"))
    if st is not None:
        st.getparent().remove(st)
    return c


def encabezado(s, titulo, sub):
    texto(s, 0.35, 0.18, 10.5, 0.45, titulo, size=22, bold=True, color=NAVY, anchor=MSO_ANCHOR.TOP)
    texto(s, 0.35, 0.62, 10.5, 0.3, sub, size=11, color=GRIS, anchor=MSO_ANCHOR.TOP)
    if hipotesis or VISTA:
        st = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.55), Inches(0.22), Inches(2.45), Inches(0.36))
        st.fill.solid()
        st.fill.fore_color.rgb = rgb("FFF2CC")
        st.line.color.rgb = rgb("BF9000")
        sin_estilo(st)
        tf = st.text_frame
        tf.text = VISTA or "HIPÓTESIS POR VALIDAR"
        r = tf.paragraphs[0].runs[0]
        r.font.size, r.font.bold, r.font.color.rgb = Pt(10), True, rgb("7F6000")
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER


def leyenda(s, y=7.03):
    x = 0.35
    for t, tipo in (("Inicio / fin", MSO_SHAPE.FLOWCHART_TERMINATOR), ("Actividad", MSO_SHAPE.FLOWCHART_PROCESS),
                    ("Decisión", MSO_SHAPE.FLOWCHART_DECISION)):
        forma(s, tipo, x, y + 0.02, 0.38, 0.24, "FFFFFF", "595959", [])
        texto(s, x + 0.45, y, 1.1, 0.28, t, size=9)
        x += 1.45
    for nat, (f, b) in COLORES.items():
        forma(s, MSO_SHAPE.RECTANGLE, x, y + 0.04, 0.26, 0.2, f, b, [])
        texto(s, x + 0.33, y, 1.6, 0.28, {"Dato": "Dato (automatizable)", "Criterio": "Criterio del revisor",
                                           "Mixto": "Mixto"}[nat], size=9)
        x += 1.95 if nat != "Mixto" else 1.0
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x + 0.1), Inches(y + 0.14), Inches(x + 0.55), Inches(y + 0.14))
    c.line.color.rgb, c.line.width = LINEA, Pt(1.25)
    d = etree.SubElement(c.line._get_or_add_ln(), qn("a:prstDash"))
    d.set("val", "dash")
    texto(s, x + 0.62, y, 1.3, 0.28, "Retorno (No)", size=9)
    if hipotesis:
        h = forma(s, MSO_SHAPE.RECTANGLE, x + 1.95, y + 0.04, 0.26, 0.2, "FFFFFF", "595959", [])
        punteada(h, {"estado": "Hipótesis"})
        texto(s, x + 2.28, y, 1.4, 0.28, "Hipótesis por validar", size=9)


# --- Diapositiva 1: vista general por etapas
etapas = []
for p in pasos:
    if not etapas or etapas[-1]["nombre"] != p["etapa"]:
        etapas.append({"nombre": p["etapa"], "pasos": []})
    etapas[-1]["pasos"].append(p)
s = prs.slides.add_slide(blank)
encabezado(s, "Proceso de elaboración del informe de Revisoría Fiscal",
           "Vista general por etapas · de la solicitud de información al seguimiento de hallazgos"
           + (" · los pasos con línea punteada son hipótesis" if hipotesis and VISTA else ""))
n = len(etapas)
max_n = max(len(e["pasos"]) for e in etapas)
ancho = (13.333 - 0.7 - 0.12 * (n - 1)) / n
x = 0.35
for i, e in enumerate(etapas):
    num, nombre = (e["nombre"].split(" · ", 1) + [""])[:2] if " · " in e["nombre"] else ("", e["nombre"])
    ch = forma(s, MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON, x, 1.25, ancho + 0.1, 0.95, "1F3864", "1F3864",
               [(num, True), (nombre, False)], size=11, nombre=f"Etapa {num}")
    for par in ch.text_frame.paragraphs:
        for r in par.runs:
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    # actividades de la etapa (el alto se ajusta a la etapa con más actividades)
    yy = 2.4
    paso_y = min(1.1, 4.5 / max_n)
    for p in e["pasos"]:
        f, b = COLORES.get(p["nat"], SIN)
        tipo = {"Decisión": MSO_SHAPE.FLOWCHART_DECISION, "Inicio": MSO_SHAPE.FLOWCHART_TERMINATOR,
                "Fin": MSO_SHAPE.FLOWCHART_TERMINATOR}.get(p["tipo"], MSO_SHAPE.FLOWCHART_PROCESS)
        alto = paso_y - 0.22
        tam = 8.5 if paso_y >= 1.0 else 7
        punteada(forma(s, tipo, x + 0.05, yy, ancho - 0.1, alto, f, b, [(p["id"], True), (p["act"], False)], size=tam,
              nombre=p["id"]), p)
        texto(s, x + 0.05, yy + alto + 0.01, ancho - 0.1, 0.18, p["rol"], size=7, color=GRIS, align=PP_ALIGN.CENTER)
        yy += paso_y
    x += ancho + 0.12
leyenda(s)
s.notes_slide.notes_text_frame.text = ("Vista general. Las formas, textos y colores son editables. "
                                       "Fuente: hoja Actividades de la plantilla de levantamiento.")

# --- Diapositivas de carriles (swimlane), partidas por etapas para que quepan
grupos, actual = [], []
for e in etapas:
    if actual and sum(len(x["pasos"]) for x in actual) + len(e["pasos"]) > 7:
        grupos.append(actual)
        actual = []
    actual.append(e)
grupos.append(actual)

X0, LBL = 1.6, 1.25
Y_FASE, Y0, Y1 = 1.02, 1.3, 6.9
alto_carril = (Y1 - Y0) / len(roles)
for gi, grupo in enumerate(grupos):
    ps = [p for e in grupo for p in e["pasos"]]
    s = prs.slides.add_slide(blank)
    e0, e1 = grupo[0]['nombre'].split(' · ')[0], grupo[-1]['nombre'].split(' · ')[0]
    rango = f"etapa {e0}" if e0 == e1 else f"etapas {e0}–{e1}"
    encabezado(s, f"Flujo del proceso por responsable ({gi + 1}/{len(grupos)})",
               f"Diagrama de carriles · {rango} · pasos {ps[0]['orden']} a {ps[-1]['orden']}")
    ncol = len(ps) + (1 if gi < len(grupos) - 1 else 0) + (1 if gi > 0 else 0)
    cw = (13.333 - 0.3 - X0) / ncol
    off = 1 if gi > 0 else 0
    # carriles
    for i, rol in enumerate(roles):
        y = Y0 + i * alto_carril
        lane = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.35), Inches(y), Inches(13.333 - 0.65), Inches(alto_carril))
        lane.fill.solid()
        lane.fill.fore_color.rgb = rgb("FFFFFF" if i % 2 else "F7F9FC")
        lane.line.color.rgb = rgb("C9D1DE")
        lane.line.width = Pt(0.75)
        sin_estilo(lane)
        lane.name = f"Carril {rol}"
        lb = forma(s, MSO_SHAPE.RECTANGLE, 0.35, y, LBL, alto_carril, "DCE6F2", "C9D1DE", [(rol, True)], size=9,
                   nombre=f"Rol {rol}")
        for par in lb.text_frame.paragraphs:
            for r in par.runs:
                r.font.color.rgb = NAVY
    # franja de etapas
    col = off
    for e in grupo:
        k = len(e["pasos"])
        nom = e["nombre"] if k * cw >= 1.9 else e["nombre"].split(" · ")[0] + " · " + e["nombre"].split(" · ")[-1].split()[0]
        fz = forma(s, MSO_SHAPE.RECTANGLE, X0 + col * cw + 0.03, Y_FASE, k * cw - 0.06, 0.24, "1F3864", "1F3864",
                   [(nom, True)], size=8.5 if k * cw >= 1.9 else 7.5)
        for par in fz.text_frame.paragraphs:
            for r in par.runs:
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        col += k
    # formas
    shapes = {}
    for j, p in enumerate(ps):
        li = roles.index(p["rol"]) if p["rol"] in roles else 0
        cx = X0 + (j + off) * cw + cw / 2
        cy = Y0 + li * alto_carril + alto_carril / 2
        f, b = COLORES.get(p["nat"], SIN)
        if p["tipo"] == "Decisión":
            w, h, tipo, sz = min(cw - 0.1, 1.5), min(alto_carril - 0.06, 0.9), MSO_SHAPE.FLOWCHART_DECISION, 7
        else:
            w = min(cw - 0.2, 1.4)
            sz = 8 if w >= 1.25 else 7
            cpl = max(10, int(w * 138 / sz))
            lineas = 1 + -(-len(p["act"]) // cpl)
            if 0.12 + lineas * sz * 0.0205 > alto_carril - 0.04:     # texto largo: letra más pequeña
                sz = 6.5
                lineas = 1 + -(-len(p["act"]) // max(10, int(w * 138 / sz)))
            h = min(alto_carril - 0.04, max(0.62, 0.12 + lineas * sz * 0.0205))
            tipo = MSO_SHAPE.FLOWCHART_TERMINATOR if p["tipo"] in ("Inicio", "Fin") else MSO_SHAPE.FLOWCHART_PROCESS
        shapes[p["id"]] = punteada(forma(s, tipo, cx - w / 2, cy - h / 2, w, h, f, b, [(p["id"], True), (p["act"], False)],
                                         size=sz, nombre=p["id"]), p)
    # conectores fuera de página
    def conector_pagina(j, p_ref, entrada):
        li = roles.index(p_ref["rol"]) if p_ref["rol"] in roles else 0
        cx = X0 + j * cw + cw / 2
        cy = Y0 + li * alto_carril + alto_carril / 2
        t = f"Viene de\n{p_ref['id']}" if entrada else f"Sigue en\n{p_ref['id']}"
        sh = forma(s, MSO_SHAPE.FLOWCHART_OFFPAGE_CONNECTOR, cx - 0.4, cy - 0.33, 0.8, 0.66, "FFFFFF", "595959",
                   [(l, True) for l in t.split("\n")], size=7.5, nombre=t.replace("\n", " "))
        sh.text_frame.margin_left = sh.text_frame.margin_right = 0
        sh.text_frame.margin_bottom = Inches(0.12)
        return sh
    if gi > 0:
        prev = [p for e in grupos[gi - 1] for p in e["pasos"]]
        entra = conector_pagina(0, prev[-1], True)
        flecha(s, entra, shapes[ps[0]["id"]], 3, 1)
    # flechas
    for j, p in enumerate(ps):
        a = shapes[p["id"]]
        if p["sig"] and p["sig"] in shapes:
            b = shapes[p["sig"]]
            mismo = (roles.index(p["rol"]) if p["rol"] in roles else 0) == (roles.index(por_id[p["sig"]]["rol"]) if por_id[p["sig"]]["rol"] in roles else 0)
            # si entre origen y destino hay otra columna y cambian de carril, entrar por arriba o por abajo
            # para que la flecha no parezca llegar a la forma intermedia
            salto = (b.left - (a.left + a.width)) > cw * 914400 * 0.9
            if salto and abs(b.top - a.top) > 914400 * 0.3:
                flecha(s, a, b, 3, 2 if b.top < a.top else 0)
            else:
                flecha(s, a, b, 3, 1)
            if p["tipo"] == "Decisión":
                texto(s, a.left / 914400 + a.width / 914400 - 0.05, a.top / 914400 + a.height / 914400 / 2 - 0.24,
                      0.3, 0.2, "Sí", size=8, bold=True, color=rgb("2E7D32"))
        elif p["sig"] and gi < len(grupos) - 1:
            sale = conector_pagina(len(ps) + off, por_id[p["sig"]], False)
            flecha(s, a, sale, 3, 1)
        if p["no"] and p["no"] not in shapes:          # el retorno va a un paso de otra diapositiva
            texto(s, a.left / 914400, a.top / 914400 + a.height / 914400 + 0.02, a.width / 914400, 0.2,
                  f"No → vuelve a {p['no']}", size=8, bold=True, color=rgb("C00000"), align=PP_ALIGN.CENTER)
        if p["no"] and p["no"] in shapes:
            b = shapes[p["no"]]
            flecha(s, a, b, 2, 2, color=rgb("C00000"), discontinua=True)
            texto(s, a.left / 914400 + a.width / 914400 / 2 + 0.05, a.top / 914400 + a.height / 914400 - 0.02,
                  0.3, 0.2, "No", size=8, bold=True, color=rgb("C00000"))
    leyenda(s)
    s.notes_slide.notes_text_frame.text = (
        "Diagrama de carriles editable: mueva las formas y los conectores las siguen. Para agregar un paso, copie una "
        "forma existente y conéctela con Insertar > Formas > Conector de codo. Colores: azul = dato, naranja = criterio, "
        "morado = mixto.")
prs.save(PPTX)

# ------------------------------------------------------------------ draw.io
CW, LW, LH, TOP = 190, 150, 130, 40
celdas = ['<mxCell id="0"/>', '<mxCell id="1" parent="0"/>']
ancho_total = LW + CW * len(pasos) + 20
celdas.append(f'<mxCell id="pool" value="Proceso de elaboración del informe de Revisoría Fiscal{(" — " + VISTA) if VISTA else (" — HIPÓTESIS POR VALIDAR" if hipotesis else "")}" '
              f'style="text;html=1;fillColor=none;strokeColor=none;'
              f'fontSize=18;fontStyle=1;fontColor=#1F3864;align=left;verticalAlign=top;" vertex="1" parent="1">'
              f'<mxGeometry x="0" y="-60" width="{ancho_total}" height="40" as="geometry"/></mxCell>')
# franja de etapas
col = 0
for i, e in enumerate(etapas):
    k = len(e["pasos"])
    celdas.append(f'<mxCell id="fase{i}" value="{html.escape(e["nombre"])}" style="rounded=0;whiteSpace=wrap;html=1;'
                  f'fillColor=#1F3864;fontColor=#FFFFFF;strokeColor=#FFFFFF;fontStyle=1;fontSize=12;" vertex="1" parent="1">'
                  f'<mxGeometry x="{LW + col * CW}" y="0" width="{k * CW}" height="{TOP - 6}" as="geometry"/></mxCell>')
    col += k
for i, rol in enumerate(roles):
    celdas.append(f'<mxCell id="lane{i}" value="{html.escape(rol)}" style="swimlane;horizontal=0;html=1;startSize={LW};'
                  f'fillColor=#DCE6F2;swimlaneFillColor={"#F7F9FC" if i % 2 == 0 else "#FFFFFF"};strokeColor=#C9D1DE;'
                  f'fontColor=#1F3864;fontStyle=1;fontSize=12;whiteSpace=wrap;" vertex="1" parent="1">'
                  f'<mxGeometry x="0" y="{TOP + i * LH}" width="{ancho_total}" height="{LH}" as="geometry"/></mxCell>')
for j, p in enumerate(pasos):
    li = roles.index(p["rol"]) if p["rol"] in roles else 0
    f, b = COLORES.get(p["nat"], SIN)
    if p["tipo"] == "Decisión":
        w, h, forma_s = 160, 110, "rhombus;"
    elif p["tipo"] in ("Inicio", "Fin"):
        w, h, forma_s = 160, 76, "rounded=1;arcSize=50;"
    else:
        w, h, forma_s = 160, 76, "rounded=0;"
    x = LW + j * CW + (CW - w) / 2
    y = (LH - h) / 2
    val = f"<b>{html.escape(p['id'])}</b><br>{html.escape(p['act'])}"
    celdas.append(f'<mxCell id="{p["id"]}" value="{html.escape(val)}" style="{forma_s}{"dashed=1;" if p["estado"] != "Validado" else ""}whiteSpace=wrap;html=1;'
                  f'fillColor=#{f};strokeColor=#{b};strokeWidth=1.5;fontSize=11;" vertex="1" parent="lane{li}">'
                  f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/></mxCell>')
for p in pasos:
    if p["sig"] and p["sig"] in por_id:
        lbl = "Sí" if p["tipo"] == "Decisión" else ""
        celdas.append(f'<mxCell id="e_{p["id"]}_{p["sig"]}" value="{lbl}" style="edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;'
                      f'endArrow=block;endFill=1;strokeColor=#5B6B86;strokeWidth=1.5;fontColor=#2E7D32;fontStyle=1;" '
                      f'edge="1" parent="1" source="{p["id"]}" target="{p["sig"]}"><mxGeometry relative="1" as="geometry"/></mxCell>')
    if p["no"] and p["no"] in por_id:
        celdas.append(f'<mxCell id="e_{p["id"]}_{p["no"]}_no" value="No" style="edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;'
                      f'endArrow=block;endFill=1;dashed=1;strokeColor=#C00000;strokeWidth=1.5;fontColor=#C00000;fontStyle=1;'
                      f'exitX=0.5;exitY=1;exitDx=0;exitDy=0;entryX=0.5;entryY=1;entryDx=0;entryDy=0;" '
                      f'edge="1" parent="1" source="{p["id"]}" target="{p["no"]}"><mxGeometry relative="1" as="geometry"/></mxCell>')
# leyenda
ly = TOP + len(roles) * LH + 30
items = [("Dato (automatizable)", "D6E4F5", "2F5597"), ("Criterio del revisor", "FCE4D6", "C55A11"), ("Mixto", "E7E0F0", "7030A0")]
for i, (t, f, b) in enumerate(items):
    celdas.append(f'<mxCell id="ley{i}" value="{t}" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#{f};strokeColor=#{b};fontSize=11;" '
                  f'vertex="1" parent="1"><mxGeometry x="{LW + i * 190}" y="{ly}" width="170" height="34" as="geometry"/></mxCell>')
xml = ('<mxfile host="app.diagrams.net"><diagram id="flujo-revisoria" name="Flujo del proceso">'
       f'<mxGraphModel dx="1600" dy="900" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" '
       f'page="1" pageScale="1" pageWidth="{ancho_total + 40}" pageHeight="{ly + 120}" math="0" shadow="0"><root>'
       + "".join(celdas) + "</root></mxGraphModel></diagram></mxfile>")
etree.fromstring(xml.encode())          # valida que el XML esté bien formado
open(DRAWIO, "w", encoding="utf-8").write(xml)
print("ok", len(pasos), "pasos,", len(roles), "roles,", len(grupos), "diapositivas de carriles")
