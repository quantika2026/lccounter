"""Genera la sección financiera del informe de Revisoría Fiscal en PPTX.

Todas las tablas y gráficos son NATIVOS (editables en PowerPoint), no capturas.
Las notas del orador de cada diapositiva listan las celdas fuente del Excel.
"""
from __future__ import annotations

import math
import re
from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from . import narrativa as N
from .excel import Periodo

# --------------------------------------------------------------- identidad
NAVY = RGBColor(0x00, 0x20, 0x60)
AZUL = RGBColor(0x0B, 0x2E, 0x9C)
ROJO = RGBColor(0xE8, 0x49, 0x2F)
GRANATE = RGBColor(0x7B, 0x1C, 0x3E)
TEXTO = RGBColor(0x33, 0x41, 0x55)
GRIS = RGBColor(0x64, 0x74, 0x8B)
GRIS_CLARO = RGBColor(0xF3, 0xF5, 0xF9)
LINEA = RGBColor(0xD5, 0xDB, 0xE5)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)
NEG = RGBColor(0xC0, 0x1E, 0x1E)
FUENTE_T = "Aptos Display"
FUENTE = "Aptos"

ANCHO, ALTO = 10.0, 5.625
Y0, Y_MAX = 0.86, 5.08          # zona útil de contenido
PLANTILLA = Path(__file__).resolve().parent.parent / "plantilla"


def _entero(v):
    from decimal import ROUND_HALF_UP, Decimal
    return int(Decimal(str(abs(v))).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def fmt_pesos(v, parentesis=True):
    if v is None:
        return ""
    s = f"{_entero(v):,}".replace(",", ".")
    return (f"$ ({s})" if parentesis else f"-$ {s}") if v < -0.5 else f"$ {s}"


def fmt_pct(v, dec=1):
    return "" if v is None else f"{v * 100:.{dec}f}%".replace(".", ",")


# ----------------------------------------------------------------- primitivas

def _txt(slide, x, y, w, h, texto, size=10, bold=False, color=TEXTO, font=FUENTE, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, italic=False, nombre=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if nombre:
        tb.name = nombre
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.02)
    tf.margin_top = tf.margin_bottom = Inches(0.01)
    tf.vertical_anchor = anchor
    parrafos = texto if isinstance(texto, list) else [texto]
    for i, p_txt in enumerate(parrafos):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if i:
            p.space_before = Pt(5)
        r = p.add_run()
        r.text = p_txt
        f = r.font
        f.size, f.bold, f.italic, f.name = Pt(size), bold, italic, font
        f.color.rgb = color
    return tb


def _sin_sombra(shape):
    shape.shadow.inherit = False
    st = shape._element.find(qn("p:style"))
    if st is not None:
        ref = st.find(qn("a:effectRef"))
        if ref is not None:
            ref.set("idx", "0")


def _rect(slide, x, y, w, h, fill, forma=MSO_SHAPE.RECTANGLE, linea=None):
    s = slide.shapes.add_shape(forma, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if linea:
        s.line.color.rgb = linea
        s.line.width = Pt(0.75)
    else:
        s.line.fill.background()
    _sin_sombra(s)
    if forma == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = min(0.12, 0.09 / max(min(w, h), 0.1))
    return s


def _triangulo(slide, puntos, color):
    emu = [(int(Inches(px)), int(Inches(py))) for px, py in puntos]
    ff = slide.shapes.build_freeform(emu[0][0], emu[0][1], scale=1.0)
    ff.add_line_segments(emu[1:], close=True)
    s = ff.convert_to_shape()
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    _sin_sombra(s)
    return s


# secciones del modelo que usa cada diapositiva financiera (para marcar las que se leyeron con propuestas)
SECCIONES_DIAPOSITIVA = {"d_monetizacion": ("liquidez",), "d_liquidez": ("liquidez",), "d_cartera": ("cartera", "deudores"),
                         "d_concentracion": ("cartera",), "d_mayores_deudores": ("mayores_deudores",),
                         "d_anticipos": ("anticipos", "deudores_varios"), "d_cxp": ("cxp",),
                         "d_resultado": ("presupuesto", "patrimonio"), "d_excedentes": ("patrimonio", "reservas")}
ETIQUETA_PROPUESTA = "Leído con propuesta del agente, pendiente de aprobación del revisor"


class Informe:
    def __init__(self, m: dict, cfg: dict, redactor: N.Redactor | None = None, propuestas=None):
        self.m, self.cfg = m, cfg
        self.propuestas = set(propuestas or ())   # secciones leídas con propuestas del adaptador aún sin aprobar
        self.con_propuesta: list[int] = []        # diapositivas (1..n) que llevan la etiqueta
        self.red = redactor or N.Redactor(False)
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(ANCHO), Inches(ALTO)
        self.blank = self.prs.slide_layouts[6]
        self.n = 0
        self.p = Periodo.desde_texto(m["meta"]["periodo"])
        self.pie = (" ".join(x for x in ("Informe", str(cfg.get("numero_informe") or "").strip(), "de Revisoría Fiscal") if x) + " · "
                    + f"{cfg.get('copropiedad', '')} · "
                    f"{self.p.etiqueta}")

    # ------------------------------------------------------------ marcos
    def _nueva(self, titulo: str, subtitulo: str | None = None, fuentes: list[str] | None = None):
        s = self.prs.slides.add_slide(self.blank)
        self.n += 1
        # motivo de marca: esquina con los triángulos azul y rojo de la plantilla LC Counter
        _triangulo(s, [(0, ALTO - 0.55), (0.62, ALTO), (0, ALTO)], AZUL)
        _triangulo(s, [(0, ALTO - 0.95), (0.30, ALTO - 0.73), (0, ALTO - 0.60)], ROJO)
        s.shapes.add_picture(str(PLANTILLA / "logo.png"), Inches(8.98), Inches(0.12), Inches(0.80), Inches(0.565))
        _txt(s, 0.42, 0.20, 8.3, 0.36, titulo, size=15, bold=True, color=NAVY, font=FUENTE_T,
             anchor=MSO_ANCHOR.MIDDLE, nombre="Titulo")
        if subtitulo:
            _txt(s, 0.42, 0.53, 8.3, 0.26, subtitulo, size=9.5, color=GRIS, nombre="Subtitulo")
        _txt(s, 0.75, ALTO - 0.30, 7.5, 0.2, self.pie, size=7, color=GRIS)
        _txt(s, 9.0, ALTO - 0.30, 0.6, 0.2, str(self.n), size=7, color=GRIS, align=PP_ALIGN.RIGHT)
        if fuentes:
            s.notes_slide.notes_text_frame.text = (
                "Fuentes (Excel de papeles de trabajo):\n" + "\n".join(f"• {f}" for f in fuentes if f))
        return s

    def d_sin_datos(self, titulo: str, subtitulo: str, detalle: str, fuentes: list[str] | None = None):
        """Conserva el ítem del modelo cuando el papel de trabajo no trae evidencia."""
        s = self._nueva(titulo, subtitulo, fuentes)
        _rect(s, 0.9, 1.45, 8.2, 2.35, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE)
        _rect(s, 0.9, 1.63, 0.07, 1.99, ROJO)
        _txt(s, 1.25, 1.78, 7.45, 0.38, "Información pendiente en los papeles de trabajo",
             size=15, bold=True, color=NAVY, font=FUENTE_T, nombre="Pendiente")
        _txt(s, 1.25, 2.34, 7.45, 0.82, detalle, size=11, color=TEXTO,
             anchor=MSO_ANCHOR.MIDDLE)
        _txt(s, 1.25, 3.28, 7.45, 0.24, "El ítem se conserva para mantener la estructura estándar del informe.",
             size=8, color=GRIS)
        return s

    def seccion(self, titulo: str):
        s = self.prs.slides.add_slide(self.blank)
        self.n += 1
        s.shapes.add_picture(str(PLANTILLA / "fondo_seccion.jpg"), 0, 0, Inches(ANCHO), Inches(ALTO))
        s.shapes.add_picture(str(PLANTILLA / "logo.png"), Inches(8.93), Inches(0.125), Inches(0.883), Inches(0.625))
        lineas = max(1, math.ceil(len(titulo) / 34))
        _txt(s, 3.67, 1.38, 6.0, 0.42 * lineas + 0.2, titulo, size=22, bold=True, color=NAVY, font=FUENTE_T)
        sub = [self.cfg.get("copropiedad", ""), " ".join(x for x in ("Informe", str(self.cfg.get("numero_informe") or "").strip()) if x) + f" · {self.p.etiqueta}"]
        _txt(s, 3.67, 1.38 + 0.42 * lineas + 0.22, 6.0, 0.6, sub if lineas == 1 else " · ".join(sub),
             size=11 if lineas == 1 else 10, color=GRIS)
        return s

    # ------------------------------------------------------------ componentes
    def kpi(self, s, x, y, w, h, etiqueta, valor, color=NAVY, nota=None):
        _rect(s, x, y, w, h, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE)
        _txt(s, x + 0.12, y + 0.07, w - 0.24, 0.2, etiqueta.upper(), size=7.5, bold=True, color=GRIS)
        _txt(s, x + 0.12, y + 0.25, w - 0.24, 0.36, valor, size=16, bold=True, color=color, font=FUENTE_T)
        if nota:
            _txt(s, x + 0.12, y + h - 0.27, w - 0.24, 0.22, nota, size=7, color=GRIS)

    @staticmethod
    def alto_texto(texto, w, size, titulo=False):
        cpl = max(10, int(w * 128 / size))
        lineas = sum(math.ceil(max(1, len(par)) / cpl) for par in str(texto).split("\n"))
        return 0.18 + lineas * size * 0.0182 + (0.24 if titulo else 0)

    def nota(self, s, x, y, w, h, texto, titulo=None, size=8.5):
        h = h or self.alto_texto(texto, w - 0.35, size, bool(titulo))
        _rect(s, x, y, w, h, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE)
        _rect(s, x, y + 0.12, 0.05, h - 0.24, ROJO)
        partes = ([titulo] if titulo else []) + [texto]
        tb = _txt(s, x + 0.17, y + 0.08, w - 0.3, h - 0.16, partes, size=size, color=TEXTO)
        if titulo:
            r = tb.text_frame.paragraphs[0].runs[0]
            r.font.bold, r.font.color.rgb, r.font.size = True, NAVY, Pt(size + 0.5)
        return h

    def tabla(self, s, x, y, columnas, filas, size=8, alto_fila=0.21, total=None, alto_enc=0.27, fusionar=None):
        """columnas: [(titulo, ancho, 'l'|'r'|'c')]; filas: [[valor,...]] con str ya formateado
        o (str, {'color':..,'bold':..}); total: fila final resaltada."""
        if not filas and not total:                 # una tabla sin filas no se dibuja vacía: se dice qué falta
            w = sum(c[1] for c in columnas)
            _rect(s, x, y, w, 0.62, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE)
            _rect(s, x, y + 0.1, 0.05, 0.42, ROJO)
            _txt(s, x + 0.18, y + 0.08, w - 0.3, 0.46, "Información pendiente: los papeles de trabajo no traen registros "
                 "para este periodo.", size=9, color=TEXTO, anchor=MSO_ANCHOR.MIDDLE)
            return None
        todas = filas + ([total] if total else [])
        n_filas = len(todas) + 1
        w = sum(c[1] for c in columnas)
        alturas = [alto_enc] + [self._alto_fila(f, columnas, size, alto_fila) for f in todas]
        shp = s.shapes.add_table(n_filas, len(columnas), Inches(x), Inches(y), Inches(w), Inches(sum(alturas)))
        tbl = shp.table
        # sin estilo de banda predeterminado: aplicamos el nuestro
        tblPr = tbl._tbl.tblPr
        tblPr.set("bandRow", "0")
        tblPr.set("firstRow", "1")
        for j, (_, ancho, _) in enumerate(columnas):
            tbl.columns[j].width = Inches(ancho)
        for i, h in enumerate(alturas):
            tbl.rows[i].height = Inches(h)
        alin = {"l": PP_ALIGN.LEFT, "r": PP_ALIGN.RIGHT, "c": PP_ALIGN.CENTER}
        for j, (titulo, _, a) in enumerate(columnas):
            self._celda(tbl.cell(0, j), titulo, size - 0.5, True, BLANCO, NAVY, alin[a])
        for i, fila in enumerate(todas, start=1):
            es_total = total is not None and i == len(todas)
            fondo_fila = RGBColor(0xE6, 0xEB, 0xF5) if es_total else (GRIS_CLARO if i % 2 == 0 else BLANCO)
            for j, val in enumerate(fila):
                estilo = {}
                if isinstance(val, tuple):
                    val, estilo = val
                fondo = estilo.get("fondo", fondo_fila)
                color = estilo.get("color", NAVY if es_total else TEXTO)
                if isinstance(val, str) and val.startswith("$ (") and "color" not in estilo:
                    color = NEG
                self._celda(tbl.cell(i, j), val, estilo.get("size", size), estilo.get("bold", es_total), color,
                            fondo, alin[columnas[j][2]], seg=estilo.get("seg"))
        if fusionar is not None:                   # p. ej. una observación que aplica a varios inmuebles
            crudo = [(f[fusionar][0] if isinstance(f[fusionar], tuple) else f[fusionar]) for f in filas]
            i = 0
            while i < len(crudo):
                j = i
                while j + 1 < len(crudo) and crudo[j + 1] == crudo[i] and crudo[i] not in ("", "—"):
                    j += 1
                if j > i:
                    for k in range(i + 2, j + 2):          # el texto queda solo en la celda de origen
                        tf = tbl.cell(k, fusionar).text_frame
                        for par in list(tf.paragraphs)[1:]:
                            par._p.getparent().remove(par._p)
                        for run in tf.paragraphs[0].runs:
                            run.text = ""
                    tbl.cell(i + 1, fusionar).merge(tbl.cell(j + 1, fusionar))
                i = j + 1
        return shp, sum(alturas)

    @staticmethod
    def _alto_fila(fila, columnas, size, minimo):
        lineas = 1
        for val, (_, ancho, _) in zip(fila, columnas):
            t = val[0] if isinstance(val, tuple) else val
            cpl = max(6, int((ancho - 0.1) * 138 / size))   # caracteres por línea aprox.
            if isinstance(val, tuple) and any(x["b"] for x in (val[1].get("seg") or [])):
                cpl = max(6, int(cpl * 0.9))
            lineas = max(lineas, sum(math.ceil(max(1, len(par)) / cpl) for par in str(t).split("\n")))
        return max(minimo, 0.05 + lineas * size * 0.0172 + (0.025 if lineas > 1 else 0))

    @staticmethod
    def _celda(c, texto, size, bold, color, fondo, align, seg=None):
        c.fill.solid()
        c.fill.fore_color.rgb = fondo
        c.margin_left = c.margin_right = Inches(0.05)
        c.margin_top = c.margin_bottom = Inches(0.025)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = c.text_frame
        tf.word_wrap = True
        if seg:                                   # texto con énfasis tomado del Excel (negrilla / rojo)
            p = tf.paragraphs[0]
            p.alignment = align
            p.font.size = Pt(size)
            for x in seg:
                partes = x["t"].split("\n")
                for k, parte in enumerate(partes):
                    if k:
                        p = tf.add_paragraph()
                        p.alignment = align
                        p.font.size = Pt(size)
                    if not parte:
                        continue
                    r = p.add_run()
                    r.text = parte
                    r.font.size, r.font.name = Pt(size), FUENTE
                    r.font.bold = bold or x["b"]
                    r.font.color.rgb = NEG if x["rojo"] else color
            lineas = []
        else:
            lineas = str(texto).split("\n")
        for k, linea in enumerate(lineas):
            p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            p.alignment = align
            p.font.size = Pt(size)
            r = p.add_run()
            r.text = linea if linea else " "
            r.font.size, r.font.bold, r.font.name = Pt(size), bold, FUENTE
            epr = p._p.get_or_add_endParaRPr()
            epr.set("sz", str(int(size * 100)))
            r.font.color.rgb = color
        # bordes finos
        tcPr = c._tc.get_or_add_tcPr()
        for borde in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            ln = tcPr.find(qn(borde))
            if ln is None:
                ln = tcPr.makeelement(qn(borde), {})
                tcPr.append(ln)
            ln.set("w", "6350")
            for hijo in list(ln):
                ln.remove(hijo)
            sf = ln.makeelement(qn("a:solidFill"), {})
            clr = sf.makeelement(qn("a:srgbClr"), {"val": "D5DBE5"})
            sf.append(clr)
            ln.append(sf)

    @staticmethod
    def _tendencia(serie, color="C0504D"):
        """Línea de tendencia lineal nativa de PowerPoint (editable) en una serie."""
        ser = serie._element
        tl = ser.makeelement(qn("c:trendline"), {})
        nm = tl.makeelement(qn("c:name"), {})
        nm.text = "Tendencia"
        tl.append(nm)
        sp = tl.makeelement(qn("c:spPr"), {})
        ln = sp.makeelement(qn("a:ln"), {"w": "15875", "cap": "rnd"})
        sf = ln.makeelement(qn("a:solidFill"), {})
        sf.append(sf.makeelement(qn("a:srgbClr"), {"val": color}))
        ln.append(sf)
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "sysDash"}))
        sp.append(ln)
        tl.append(sp)
        tl.append(tl.makeelement(qn("c:trendlineType"), {"val": "linear"}))
        tl.append(tl.makeelement(qn("c:dispRSqr"), {"val": "0"}))
        tl.append(tl.makeelement(qn("c:dispEq"), {"val": "0"}))
        cat = ser.find(qn("c:cat"))
        if cat is not None:
            cat.addprevious(tl)
        else:
            ser.append(tl)

    def grafico_linea(self, s, x, y, w, h, categorias, series, titulo, formato='#,##0.0,,"M"', colores=None,
                      tendencia=None):
        cd = CategoryChartData()
        cd.categories = categorias
        for nombre, vals in series:
            cd.add_series(nombre, vals)
        gf = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(x), Inches(y), Inches(w), Inches(h), cd)
        ch = gf.chart
        ch.has_title = True
        ch.chart_title.text_frame.text = titulo
        tp = ch.chart_title.text_frame.paragraphs[0]
        tp.runs[0].font.size, tp.runs[0].font.bold = Pt(9), True
        tp.runs[0].font.color.rgb = NAVY
        ch.has_legend = len(series) > 1
        if ch.has_legend:
            ch.legend.position = XL_LEGEND_POSITION.TOP
            ch.legend.include_in_layout = True
            ch.legend.font.size = Pt(7.5)
        va = ch.value_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = LINEA
        va.tick_labels.font.size = Pt(7)
        va.tick_labels.font.color.rgb = GRIS
        va.tick_labels.number_format = formato
        va.tick_labels.number_format_is_linked = False
        va.format.line.fill.background()
        if len(series) > 1:                       # deja aire arriba para la leyenda
            tope = max(max(v) for _, v in series)
            paso = 10 ** (len(str(int(tope))) - 2) * 2
            va.maximum_scale = math.ceil(tope * 1.18 / paso) * paso
            va.minimum_scale = 0
        ca = ch.category_axis
        ca.tick_labels.font.size = Pt(7)
        ca.tick_labels.font.color.rgb = GRIS
        ca.format.line.color.rgb = LINEA
        from pptx.enum.chart import XL_TICK_LABEL_POSITION
        ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
        colores = colores or [NAVY, ROJO, GRANATE]
        for k, serie in enumerate(ch.plots[0].series):
            serie.format.line.color.rgb = colores[k % len(colores)]
            serie.format.line.width = Pt(2)
            serie.smooth = False
            serie.marker.size = 5
            serie.marker.format.fill.solid()
            serie.marker.format.fill.fore_color.rgb = colores[k % len(colores)]
            serie.marker.format.line.color.rgb = colores[k % len(colores)]
        if tendencia is not None:
            self._tendencia(ch.plots[0].series[tendencia])
        posiciones = [XL_LABEL_POSITION.BELOW, XL_LABEL_POSITION.ABOVE] if len(series) > 1 else [XL_LABEL_POSITION.ABOVE]
        for k, serie in enumerate(ch.plots[0].series):
            dl = serie.data_labels
            dl.show_value = True
            dl.number_format = formato
            dl.number_format_is_linked = False
            dl.font.size = Pt(7.5)
            dl.font.color.rgb = colores[k % len(colores)] if len(series) > 1 else TEXTO
            dl.position = posiciones[k % len(posiciones)]
            if len(series) == 1:                  # en una sola serie, los negativos van debajo
                for i, v in enumerate(series[0][1]):
                    if v < 0:
                        pdl = serie.points[i].data_label
                        pdl.position = XL_LABEL_POSITION.BELOW
                        pdl.font.size = Pt(6.5)
                        pdl.font.color.rgb = TEXTO
                        pdl.has_text_frame = False
        return ch

    def grafico_dona(self, s, x, y, w, h, categorias, valores, colores):
        cd = CategoryChartData()
        cd.categories = categorias
        cd.add_series("Cartera", valores)
        gf = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(x), Inches(y), Inches(w), Inches(h), cd)
        ch = gf.chart
        ch.has_title = False
        ch.has_legend = True
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(7)
        pl = ch.plots[0]
        for k, pt in enumerate(pl.series[0].points):
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = colores[k]
        pl.has_data_labels = True
        pl.data_labels.number_format = "0%"
        pl.data_labels.number_format_is_linked = False
        pl.data_labels.show_percentage = True
        pl.data_labels.show_value = False
        pl.data_labels.font.size = Pt(8)
        pl.data_labels.font.bold = True
        pl.data_labels.font.color.rgb = BLANCO
        # agujero de la dona
        hole = pl._element.find(qn("c:holeSize"))
        if hole is not None:
            hole.set("val", "55")
        return ch

    def grafico_barras(self, s, x, y, w, h, categorias, series, titulo, colores):
        cd = CategoryChartData()
        cd.categories = categorias
        for nombre, vals in series:
            cd.add_series(nombre, vals)
        gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), cd)
        ch = gf.chart
        ch.has_title = True
        ch.chart_title.text_frame.text = titulo
        r = ch.chart_title.text_frame.paragraphs[0].runs[0]
        r.font.size, r.font.bold, r.font.color.rgb = Pt(9), True, NAVY
        ch.has_legend = False
        va = ch.value_axis
        va.minimum_scale = 0
        va.visible = False
        va.has_major_gridlines = False
        ca = ch.category_axis
        ca.tick_labels.font.size = Pt(7)
        ca.tick_labels.font.color.rgb = GRIS
        ca.format.line.color.rgb = LINEA
        pl = ch.plots[0]
        pl.gap_width = 60
        pl.vary_by_categories = False
        ser = pl.series[0]
        for k, pt in enumerate(ser.points):
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = colores[k % len(colores)]
        pl.has_data_labels = True
        pl.data_labels.number_format = '#,##0.0,,"M"'
        pl.data_labels.number_format_is_linked = False
        pl.data_labels.font.size = Pt(7.5)
        pl.data_labels.font.bold = True
        pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        return ch

    # ------------------------------------------------------------ paginación
    def tabla_paginada(self, titulo, subtitulo, columnas, filas, fuentes, y=Y0, size=7.5, total=None,
                       alto_max=None, antes=None, fusionar=None):
        alto_max = alto_max or (Y_MAX - y - 0.08)
        # grupos: filas consecutivas con la misma celda fusionada no se separan entre páginas
        grupos = []
        for f in filas:
            clave = None
            if fusionar is not None:
                clave = f[fusionar][0] if isinstance(f[fusionar], tuple) else f[fusionar]
            if grupos and clave not in (None, "", "—") and grupos[-1]["clave"] == clave:
                sin_obs = list(f)
                sin_obs[fusionar] = ""
                grupos[-1]["filas"].append(f)
                grupos[-1]["alto"] += self._alto_fila(sin_obs, columnas, size, 0.21)
            else:
                grupos.append({"clave": clave, "filas": [f], "alto": self._alto_fila(f, columnas, size, 0.21)})
        alturas = [g["alto"] for g in grupos]
        util = alto_max - 0.27 - (0.25 if total else 0)
        n_pag = max(1, math.ceil(sum(alturas) / util))
        while True:                      # reparte los grupos de forma pareja entre las páginas
            objetivo = sum(alturas) / n_pag
            paginas, actual, acum = [], [], 0.0
            for g, h in zip(grupos, alturas):
                if actual and (acum + h > util or (acum + h / 2 > objetivo and len(paginas) < n_pag - 1)):
                    paginas.append(actual)
                    actual, acum = [], 0.0
                actual += g["filas"]
                acum += h
            paginas.append(actual)
            if len(paginas) <= n_pag:
                break
            n_pag += 1
        for k, pag in enumerate(paginas):
            t = titulo + (f" ({k + 1}/{len(paginas)})" if len(paginas) > 1 else "")
            s = self._nueva(t, subtitulo, fuentes)
            if antes and k == 0:
                antes(s)
            self.tabla(s, 0.42, y, columnas, pag, size=size, total=total if k == len(paginas) - 1 else None,
                       fusionar=fusionar)
        return len(paginas)

    # ================================================================ diapositivas
    def d_monetizacion(self):
        L, M = self.m["liquidez"], self.m["meta"]
        if not L.get("monetizacion"):                 # la copropiedad no lleva el análisis de monetización
            self.d_sin_datos(
                "Análisis de monetización del fondo de imprevistos",
                "Recursos en fiducia frente a la reserva registrada en el patrimonio",
                "El Excel no contiene el análisis de monetización. Complete la fiducia, la reserva y los rendimientos."
            )
            return
        s = self._nueva("Análisis de monetización del fondo de imprevistos",
                        "Recursos en fiducia frente a la reserva registrada en el patrimonio",
                        [r["fuente"] for r in L["monetizacion"]])
        filas = []
        for r in L["monetizacion"]:
            filas.append([r["concepto"], fmt_pesos(r["actual"]), fmt_pesos(r["referencia"])])
        total = filas.pop() if L["monetizacion"][-1]["es_total"] else None
        _, h_tab = self.tabla(s, 0.42, Y0 + 0.1, [("Concepto", 3.6, "l"), (M["periodo_corto"], 1.2, "r"),
                                                  (M["cierre_corto"], 1.2, "r")], filas, size=8.5, alto_fila=0.3,
                              total=total)
        v = L["mayor_valor_monetizado"]
        self.kpi(s, 6.75, Y0 + 0.1, 2.85, 0.95, "Mayor valor monetizado", fmt_pesos(v),
                 color=NAVY if v >= 0 else NEG, nota=f"{M['cierre_corto']}: {fmt_pesos(L['mayor_valor_monetizado_cierre'])}")
        fid = L["monetizacion"][0]["actual"]
        oblig = -sum(r["actual"] for r in L["monetizacion"][1:-1])
        self.grafico_barras(s, 6.6, Y0 + 1.15, 3.1, 2.3, ["Fiducia", "Reserva +\nrendimientos"],
                            [("Valor", [fid, oblig])], "Cobertura del fondo", [NAVY, ROJO])
        self.nota(s, 0.42, Y0 + 0.1 + h_tab + 0.3, 6.0, None,
                  self.red.pulir(N.texto_monetizacion(self.m), "monetización"), "Lectura")

    def d_liquidez(self):
        L, M = self.m["liquidez"], self.m["meta"]
        s = self._nueva("Liquidez de efectivo",
                        f"Disponible menos obligaciones de corto plazo · {M['periodo_corto']} vs. {M['cierre_corto']}",
                        [r["fuente"] for r in L["filas"]])
        filas, total = [], None
        for r in L["filas"]:
            fila = [r["concepto"], fmt_pesos(r["actual"]), fmt_pesos(r["referencia"])]
            if r["es_total"] and norm_(r["concepto"]).startswith("total equivalente"):
                total = fila
            elif r["es_total"]:
                filas.append([(x, {"bold": True, "color": NAVY}) for x in fila])
            else:
                filas.append(fila)
        self.tabla(s, 0.42, Y0, [("Concepto", 3.05, "l"), (M["periodo_corto"], 1.0, "r"), (M["cierre_corto"], 1.0, "r")],
                   filas, size=7.5, alto_fila=0.19, alto_enc=0.24, total=total)
        cats = [Periodo.desde_texto(x["periodo"]).corto for x in L["serie"]]
        self.grafico_linea(s, 5.65, Y0, 4.05, 2.75, cats, [("Equivalente al efectivo", [x["valor"] for x in L["serie"]])],
                           "Equivalente al efectivo positivo (millones)", tendencia=0)
        self.nota(s, 5.65, Y0 + 2.9, 4.05, None, self.red.pulir(N.texto_liquidez(self.m), "liquidez"), "Lectura",
                  size=8)

    def d_cartera(self):
        D, C, M = self.m["deudores"], self.m["cartera"], self.m["meta"]
        s = self._nueva(N.titulo_cartera(self.m), "Composición de los deudores de propietarios",
                        [r["fuente"] for r in D["filas"]] + [C["fuente_total"]])
        filas = [[r["concepto"], fmt_pesos(r["actual"]), fmt_pesos(r["referencia"])] for r in D["filas"][:-1]]
        tot = D["filas"][-1]
        self.tabla(s, 0.42, Y0 + 0.05, [("Detalle", 2.95, "l"), (M["periodo_corto"], 1.0, "r"), (M["cierre_corto"], 1.0, "r")],
                   filas, size=8.5, alto_fila=0.27, total=["Deudores copropietarios (neto)", fmt_pesos(tot["actual"]),
                                                          fmt_pesos(tot["referencia"])])
        cats = [Periodo.desde_texto(x["periodo"]).corto for x in D["serie"]]
        self.grafico_linea(s, 5.6, Y0, 4.1, 2.6, cats, [("Neto", [x["valor"] for x in D["serie"]]),
                                                        ("Bruto", [x["bruto"] for x in D["serie"]])],
                           "Evolución de la cartera (millones)", tendencia=0)
        self.nota(s, 0.42, 3.62, 9.28, None, self.red.pulir(N.texto_cartera(self.m), "cartera"), "Lectura", size=8.5)

    def d_concentracion(self):
        C, M = self.m["cartera"], self.m["meta"]
        s = self._nueva(N.titulo_concentracion(self.m), "Distribución de la cartera por rango de deuda",
                        [C["fuente_total"]])
        self.grafico_dona(s, 0.3, Y0, 3.4, 3.1, [r["rango"] for r in C["rangos"]], [r["valor"] for r in C["rangos"]],
                          [NAVY, AZUL, ROJO])
        filas = [[r["rango"], fmt_pesos(r["valor"]), str(r["inmuebles"]), fmt_pct(r["pct_cartera"]),
                  fmt_pct(r["pct_unidades"], 2)] for r in C["rangos"]]
        u = M.get("unidades")                       # ficha nueva sin unidades: la columna sale sin porcentaje
        total = [f"Total cartera {M['periodo_corto']}", fmt_pesos(C["total"]), str(C["inmuebles"]), "100,0%",
                 fmt_pct(C["inmuebles"] / u, 2) if u else "—"]
        self.tabla(s, 3.9, Y0 + 0.15, [("Rango", 2.0, "l"), ("Cifra", 1.15, "r"), ("Inmuebles", 0.75, "c"),
                                       ("% cartera", 0.8, "r"), (f"% de {u} uds." if u else "% de unidades", 1.1, "r")],
                   filas, size=8, alto_fila=0.32, total=total)
        nota = (self.cfg.get("notas") or {}).get("concentracion_cartera")
        if nota:
            self.nota(s, 3.9, 2.75, 5.8, None, " ".join(nota.split()), "Gestión de cobro", size=8)

    def d_mayores_deudores(self):
        X = self.m["mayores_deudores"]
        M = self.m["meta"]
        if not X.get("items"):
            self.d_sin_datos(
                "Detalle de mayores deudores", "Inmuebles con deuda mayor a $1 millón",
                "No se encontraron mayores deudores documentados. Complete saldos, gestión de cobro y observación de Revisoría Fiscal."
            )
            return
        filas = []
        for r in X["items"]:
            gest = r["gestion_abogado"] + (f"\n{r['detalle_gestion']}" if r["detalle_gestion"] else "")
            filas.append([(r["apto"], {"bold": True, "color": NAVY}), fmt_pesos(r["saldo_anterior"]),
                          fmt_pesos(r["saldo_actual"]), gest, r["proceso_manual"],
                          (r["observacion_rf"] or "—", {"color": TEXTO if r.get("seg_observacion") else (NEG if r["observacion_rf"] else GRIS),
                                                        "seg": r.get("seg_observacion")})])
        cols = [("Apto", 0.55, "c"), (f"Saldo {M['anterior_corto']}", 0.95, "r"), (f"Saldo {M['periodo_corto']}", 0.95, "r"),
                (X.get("encabezado_abogado") or "Gestión según informe del abogado", 2.25, "l"),
                ("Proceso según Manual de Cobro de Cartera", 1.0, "c"), ("Observación Revisoría Fiscal", 3.58, "l")]
        self.tabla_paginada("Detalle de mayores deudores", "Inmuebles con deuda mayor a $1 millón",
                            cols, filas, [r["fuente"] for r in X["items"]], size=8.5, fusionar=5)

    def d_anticipos(self):
        A, DV = self.m["anticipos"], self.m["deudores_varios"]
        fin = self.p.fin

        def meses(iso):
            try:
                return f"{(fin - date.fromisoformat(iso)).days // 30} m"
            except (TypeError, ValueError):
                return ""

        cols = [("Fecha", 0.75, "c"), ("Tercero", 1.5, "l"), ("Valor", 0.9, "r"), ("Antig.", 0.5, "c"),
                ("Detalle", 5.63, "l")]
        filas = []
        for it in A["items"]:
            ant = meses(it["fecha_iso"])
            viejo = ant and int(ant.split()[0]) >= 12
            filas.append([it["fecha"], it["tercero"], fmt_pesos(it["valor"]),
                          (ant, {"bold": viejo, "color": NEG if viejo else TEXTO}), (it["detalle"], {"seg": it.get("seg_detalle")})])
        total_a = ["", "Total anticipos", fmt_pesos(sum(i["valor"] for i in A["items"])), "", ""]
        if not A["items"] and not DV["items"]:
            self.d_sin_datos(
                "Anticipos a contratistas y deudores varios", "Cuentas 1330 y 1380",
                "El papel de trabajo no reporta partidas en anticipos ni deudores varios para el periodo."
            )
            self.d_sin_datos(
                "Opinión de la Revisoría Fiscal", "Anticipos y avances a contratistas — cuenta 1330",
                "No hay partidas ni una nota profesional para emitir la conclusión de esta sección."
            )
            return
        fuentes = [i["fuente"] for i in A["items"]] + [i["fuente"] for i in DV["items"]]
        s = self._nueva("Anticipos a contratistas y deudores varios",
                        "Anticipos y avances a contratistas — cuenta 1330", fuentes)
        # si anticipos y deudores varios no caben juntos en 7,5 pt, se bajan a 7 pt para mantenerlos en una diapositiva
        filas_dv0 = [[it["fecha"], it["tercero"], "", "", (it["detalle"], {"seg": it.get("seg_detalle")})] for it in DV["items"]]
        tam = 7.5
        def alto_total(t):
            return (sum(self._alto_fila(f, cols, t, 0.22) for f in filas + [total_a]) + 0.27 + 0.12 +
                    (0.25 + 0.27 + sum(self._alto_fila(f, cols, t, 0.22) for f in filas_dv0) + 0.22 if filas_dv0 else 0))
        if alto_total(7.5) > Y_MAX - Y0 + 0.1:
            tam = 7.0
        _, h = self.tabla(s, 0.42, Y0, cols, filas, size=tam, alto_fila=0.2, total=total_a)
        y = Y0 + h + 0.12
        filas_dv = [[it["fecha"], it["tercero"], fmt_pesos(it["valor"]), meses(it["fecha_iso"]),
                     (it["detalle"], {"seg": it.get("seg_detalle")})]
                    for it in DV["items"]]
        if filas_dv:
            if y + 0.25 + sum(self._alto_fila(f, cols, tam, 0.2) for f in filas_dv) + 0.27 > Y_MAX + 0.05:
                s = self._nueva("Anticipos a contratistas y deudores varios (cont.)", None, fuentes)
                y = Y0
            _txt(s, 0.42, y, 6, 0.22, "Deudores varios — cuenta 1380", size=9, bold=True, color=NAVY)
            self.tabla(s, 0.42, y + 0.25, cols, filas_dv, size=tam, alto_fila=0.2,
                       total=["", "Total deudores varios", fmt_pesos(sum(i["valor"] for i in DV["items"])), "", ""])
        op = (self.cfg.get("notas") or {}).get("opinion_anticipos")
        if op:
            s = self._nueva("Opinión de la Revisoría Fiscal", "Anticipos y avances a contratistas — cuenta 1330")
            parrafos = [" ".join(p.split()) for p in re.split(r"\n+", op.strip()) if p.strip()]
            h = min(Y_MAX - Y0 - 0.1, self.alto_texto("\n".join(parrafos), 8.7, 11) + 0.15 * len(parrafos))
            _rect(s, 0.42, Y0 + 0.1, 9.28, h, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE)
            _rect(s, 0.42, Y0 + 0.25, 0.06, h - 0.3, ROJO)   # acento de cita
            _txt(s, 0.7, Y0 + 0.22, 8.8, h - 0.2, parrafos, size=11, color=TEXTO)
        else:
            self.d_sin_datos(
                "Opinión de la Revisoría Fiscal", "Anticipos y avances a contratistas — cuenta 1330",
                "Complete la nota profesional sobre antigüedad, recuperabilidad y acciones requeridas."
            )

    def d_cxp(self):
        X = self.m["cxp"]
        if not X or (not X.get("items") and X.get("total") is None):   # hoja ausente: el ítem queda «pendiente»
            self.d_sin_datos(
                "Cuentas por pagar", f"Principales acreedores a {self.p.etiqueta}",
                "El Excel no contiene el detalle de cuentas por pagar del periodo."
            )
            return
        edades = {}
        for it in X["items"]:
            clave = "Mayor a 90 días" if "90" in it["edad"] and "ayor" in it["edad"] else "Hasta 90 días"
            edades[clave] = edades.get(clave, 0) + it["valor"]

        def antes(s):
            x = 0.42
            for k, (nombre, val) in enumerate(sorted(edades.items(), key=lambda t: t[0])):
                self.kpi(s, x, Y0 - 0.02, 2.2, 0.62, nombre, fmt_pesos(val),
                         color=NEG if "Mayor" in nombre else NAVY)
                x += 2.35
            self.kpi(s, x, Y0 - 0.02, 2.2, 0.62, "Acreedores", str(len(X["items"])))

        filas = [[it["edad"], it["tercero"], fmt_pesos(it["valor"]), fmt_pct(it["valor"] / X["total"]), (it["concepto"], {"seg": it.get("seg_concepto")})]
                 for it in X["items"]]
        cols = [("Edad", 1.05, "l"), ("Tercero", 2.2, "l"), ("Valor", 1.0, "r"), ("%", 0.55, "r"),
                ("Concepto", 4.48, "l")]
        self.tabla_paginada(f"CUENTAS POR PAGAR {fmt_pesos(X['total'])}", f"Principales acreedores a {self.p.etiqueta}",
                            cols, filas, [i["fuente"] for i in X["items"]], y=Y0 + 0.72, size=8.5,
                            total=["", "Total cuentas por pagar", fmt_pesos(X["total"]), "100,0%", ""], antes=antes)

    def d_resultado(self):
        P, PR = self.m.get("patrimonio"), self.m.get("presupuesto")
        if not PR:                                       # sin presupuesto legible: la diapositiva no tiene tabla
            raise KeyError("presupuesto")
        obs = (self.cfg.get("notas") or {}).get("presupuesto") or {}
        if P:
            titulo = (f"RESULTADO DEL EJERCICIO {self.p.anio}: EXCEDENTE DE {fmt_pesos(P['resultado'])}"
                      if P["resultado"] >= 0 else
                      f"RESULTADO DEL EJERCICIO {self.p.anio}: DÉFICIT DE {fmt_pesos(-P['resultado'])}")
        else:
            titulo = f"Ejecución presupuestal {self.p.anio}"
        s = self._nueva(titulo, f"Acumulado enero – {self.p.nombre_mes} {self.p.anio}",
                        ([P["fuente_resultado"]] if P else []) + [r["fuente"] for r in PR["rubros"]] + [PR["fuente_total"]])
        if P:
            self.kpi(s, 0.42, Y0 - 0.02, 3.0, 0.68, "Total ingresos", fmt_pesos(P["ingresos"]))
            self.kpi(s, 3.56, Y0 - 0.02, 3.0, 0.68, "Total gastos", fmt_pesos(P["gastos"]), color=NEG)
            self.kpi(s, 6.70, Y0 - 0.02, 3.0, 0.68, "Resultado del ejercicio", fmt_pesos(P["resultado"]),
                     color=NAVY if P["resultado"] >= 0 else NEG)
        else:
            _txt(s, 0.42, Y0, 9.2, 0.5, "Resultado del ejercicio pendiente: la hoja PATRIMONIO no se pudo leer "
                 "(ver hallazgos de validación).", size=9, color=GRIS)
        _txt(s, 0.42, Y0 + 0.75, 9, 0.22, "Rubros sobreejecutados (valores mayores a $1.000.000)", size=9,
             bold=True, color=NAVY)
        rubros = [r for r in PR["rubros"] if r["sobre_ejecutado"] > 1_000_000]
        def obs_de(r):
            t = obs.get(r["rubro"], r["observacion"]) or ""
            if r.get("soportes"):
                t = (t + " " if t else "") + "Ver soporte del auxiliar."
            return t
        filas = [[r["rubro"], fmt_pesos(r["presupuesto_anual"]), fmt_pesos(r["presupuesto_periodo"]),
                  fmt_pesos(r["ejecucion"]), (fmt_pesos(r["sobre_ejecutado"]), {"color": NEG, "bold": True}), obs_de(r)]
                 for r in rubros]
        mes = self.p.nombre_mes.capitalize()
        cols = [("Rubro", 1.95, "l"), (f"Presupuesto Ene–Dic {self.p.anio}", 1.05, "r"),
                (f"Presupuesto Ene–{mes[:3]}", 1.05, "r"), (f"Ejecución Ene–{mes[:3]}", 1.05, "r"),
                ("Sobre ejecutado", 1.05, "r"), ("Observaciones", 3.13, "l")]
        self.tabla(s, 0.42, Y0 + 1.0, cols, filas, size=7.5, alto_fila=0.21, alto_enc=0.3,
                   total=["Total rubros sobreejecutados", "", "", "", fmt_pesos(sum(r["sobre_ejecutado"] for r in rubros)), ""])
        self.d_soportes(rubros)

    def d_soportes(self, rubros):
        """Diapositivas con los extractos del auxiliar que la hoja PRESUPUESTO tiene pegados por rubro."""
        import io
        ALTO_UTIL, X_IMG, W_IMG = Y_MAX - Y0 - 0.05, 2.75, 6.95
        bloques = []
        for r in rubros:
            imgs = r.get("soportes") or []
            if not imgs:
                continue
            cols = 1 if len(imgs) == 1 else 2
            w = (W_IMG - 0.1 * (cols - 1)) / cols
            altos = [0.0] * cols
            pos = []
            for im in imgs:
                c = altos.index(min(altos))
                h = w * im["alto"] / max(im["ancho"], 1)
                pos.append((c, altos[c], h))
                altos[c] += h + 0.06
            bloques.append({"r": r, "imgs": imgs, "pos": pos, "w": w, "alto": max(altos), "cols": cols})
        if not bloques:
            return
        # un rubro con un solo extracto comparte diapositiva con otro; los de varios extractos van solos
        paginas, pendiente = [], []
        for bl in bloques:
            if len(bl["imgs"]) == 1 and bl["alto"] <= ALTO_UTIL * 0.8:
                pendiente.append(bl)
                if len(pendiente) == 2:
                    paginas.append(pendiente)
                    pendiente = []
            else:
                paginas.append([bl])
        if pendiente:
            paginas.append(pendiente)
        for k, pag in enumerate(paginas):
            t = "Soporte de rubros sobreejecutados" + (f" ({k + 1}/{len(paginas)})" if len(paginas) > 1 else "")
            s = self._nueva(t, "Extractos del auxiliar contable incluidos en el papel de trabajo",
                            [b["r"]["fuente"] for b in pag])
            y = Y0
            hueco = (ALTO_UTIL - 0.12 * (len(pag) - 1)) / len(pag)
            for b in pag:
                esc = min(1.0, hueco / b["alto"])
                alto = b["alto"] * esc
                _rect(s, 0.42, y, 2.2, alto, GRIS_CLARO)
                _rect(s, 0.42, y, 0.05, alto, ROJO)
                r = b["r"]
                _txt(s, 0.55, y + 0.06, 2.0, 0.4, r["rubro"], size=9, bold=True, color=NAVY)
                _txt(s, 0.55, y + 0.42, 2.0, 0.8,
                     [f"Presupuesto: {fmt_pesos(r['presupuesto_periodo'])}", f"Ejecución: {fmt_pesos(r['ejecucion'])}",
                      f"Sobre ejecutado: {fmt_pesos(r['sobre_ejecutado'])}"], size=7.5, color=TEXTO)
                for im, (c, yy, h) in zip(b["imgs"], b["pos"]):
                    s.shapes.add_picture(io.BytesIO(im["bytes"]), Inches(X_IMG + c * (b["w"] * esc + 0.1)),
                                         Inches(y + yy * esc), Inches(b["w"] * esc), Inches(h * esc))
                y += alto + 0.12

    def d_excedentes(self):
        P = self.m["patrimonio"]
        s = self._nueva("Excedentes acumulados", "Origen de los excedentes y activos que los respaldan",
                        [r["fuente"] for r in P["origen_excedentes"] + P["disponibilidad"]])
        _txt(s, 0.42, Y0, 4, 0.22, "Origen de los excedentes acumulados", size=9, bold=True, color=NAVY)
        filas = [[r["concepto"], fmt_pesos(r["actual"])] for r in P["origen_excedentes"] if not r["es_total"]]
        self.tabla(s, 0.42, Y0 + 0.26, [("Concepto", 2.85, "l"), ("Valor", 1.25, "r")], filas, size=9, alto_fila=0.32,
                   total=["Total excedentes acumulados", fmt_pesos(P["total_excedentes"])])
        if not P.get("disponibilidad"):               # sin análisis de respaldo: solo el origen y la lectura
            self.nota(s, 4.85, Y0, 4.8, None, self.red.pulir(N.texto_excedentes(self.m), "excedentes"), "Lectura", size=8)
            return
        _txt(s, 4.85, Y0, 4.8, 0.22, "Composición de la disponibilidad de excedentes", size=9, bold=True, color=NAVY)
        disp = sorted([d for d in P["disponibilidad"] if not d["es_total"]], key=lambda d: -d["actual"])
        filas = [[d["concepto"], fmt_pesos(d["actual"]), fmt_pct(d["pct"])] for d in disp]
        self.tabla(s, 4.85, Y0 + 0.26, [("Activos que respaldan", 2.8, "l"), ("Valor", 1.25, "r"), ("%", 0.8, "r")],
                   filas, size=9, alto_fila=0.3, total=["Disponibilidad de excedentes",
                                                         fmt_pesos(P["total_disponibilidad"]), "100,0%"])
        dif = P["total_disponibilidad"] - P["total_excedentes"]
        self.nota(s, 0.42, 2.55, 4.1, None, self.red.pulir(N.texto_excedentes(self.m), "excedentes"),
                  "Conclusión" if abs(dif) <= 1 else "Conclusión (con diferencia por conciliar)", size=8)

    # ================================================================ orquestación
    def marcar(self, desde: int, secciones):
        """Etiqueta las diapositivas desde+1..n si usan una sección leída con una propuesta del adaptador sin aprobar.
        Las diapositivas «pendiente» no la llevan: no muestran datos leídos."""
        if not self.propuestas or not set(secciones) & self.propuestas:
            return
        for i in range(desde, self.n):
            s = self.prs.slides[i]
            if any(sh.name in ("Pendiente", "EtiquetaPropuesta") for sh in s.shapes):
                continue
            _txt(s, 0.42, 0.01, 8.3, 0.2, "⚠ " + ETIQUETA_PROPUESTA, size=7.5, bold=True, color=ROJO,
                 nombre="EtiquetaPropuesta")
            self.con_propuesta.append(i + 1)

    def construir(self):
        self.seccion(f"Auditoría Financiera a {self.p.etiqueta}.")
        for nombre, secciones in SECCIONES_DIAPOSITIVA.items():
            antes = self.n
            getattr(self, nombre)()
            self.marcar(antes, secciones)
        return self

    def guardar(self, ruta):
        self.prs.save(ruta)
        return ruta


def norm_(t):
    from .excel import norm
    return norm(t)
