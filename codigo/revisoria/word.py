"""Documentos Word del flujo: informe de validación y opinión del revisor (antes HTML y Markdown).

Python arma el documento con los datos ya calculados; aquí no se calcula ni se redacta nada nuevo.
"""
from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

NAVY = RGBColor(0x00, 0x20, 0x60)
GRIS = RGBColor(0x64, 0x74, 0x8B)
COLOR_SEV = {"ALTA": RGBColor(0xB4, 0x23, 0x18), "MEDIA": RGBColor(0xB5, 0x47, 0x08),
             "BAJA": RGBColor(0x2E, 0x6B, 0xC6), "INFO": GRIS}
LOGO = Path(__file__).resolve().parent.parent / "app" / "static" / "logo.png"


def _base(titulo: str, subtitulo: str, aviso: str | None = None) -> Document:
    w = Document()
    sec = w.sections[0]
    sec.left_margin = sec.right_margin = Cm(2.2)
    st = w.styles["Normal"]
    st.font.name, st.font.size = "Arial", Pt(10.5)
    if LOGO.exists():
        sec.header.paragraphs[0].add_run().add_picture(str(LOGO), width=Cm(2.4))
    if aviso:
        p = sec.header.add_paragraph()
        r = p.add_run(aviso)
        r.font.size, r.font.color.rgb, r.bold = Pt(8.5), COLOR_SEV["ALTA"], True
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pie = sec.footer.paragraphs[0]
    pie.add_run("LC Counter Contadores y Auditores SAS · Desarrollado por Quantika IA SAS").font.size = Pt(8)
    h = w.add_paragraph()
    r = h.add_run(titulo)
    r.bold, r.font.size, r.font.color.rgb = True, Pt(18), NAVY
    s = w.add_paragraph()
    r = s.add_run(subtitulo)
    r.font.size, r.font.color.rgb = Pt(10.5), GRIS
    return w


def _negritas(par, texto: str, tam: Pt | None = None):
    """Escribe el texto respetando **negrita** y *cursiva* del Markdown."""
    for trozo in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)", texto):
        if not trozo:
            continue
        if trozo.startswith("**"):
            r = par.add_run(trozo[2:-2]); r.bold = True
        elif trozo.startswith("`"):
            r = par.add_run(trozo[1:-1]); r.font.name = "Consolas"
        elif trozo.startswith("*") and len(trozo) > 2:
            r = par.add_run(trozo[1:-1]); r.italic = True
        else:
            r = par.add_run(trozo)
        if tam:
            r.font.size = tam


def _sombra(celda, hexa: str):
    tc = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"), shd.set(qn("w:color"), "auto"), shd.set(qn("w:fill"), hexa)
    tc.append(shd)


def markdown_a_docx(texto: str, ruta: Path, titulo: str, subtitulo: str, aviso: str | None = None) -> Path:
    """Convierte el Markdown sencillo del flujo (títulos, listas, citas, tablas, negritas) a Word."""
    w = _base(titulo, subtitulo, aviso)
    lineas = texto.splitlines()
    i = 0
    while i < len(lineas):
        s = lineas[i].rstrip()
        i += 1
        if not s.strip():
            continue
        if i == 1 and s.startswith("# "):                     # el título ya va en la portada del documento
            continue
        if s.lstrip().startswith("|"):
            filas = [s]
            while i < len(lineas) and lineas[i].lstrip().startswith("|"):
                filas.append(lineas[i]); i += 1
            filas = [f for f in filas if not re.match(r"^\|?\s*:?-{2,}", f.strip())]
            celdas = [[c.strip() for c in f.strip().strip("|").split("|")] for f in filas]
            n = max(len(c) for c in celdas)
            t = w.add_table(rows=len(celdas), cols=n)
            t.style = "Table Grid"
            for r, fila in enumerate(celdas):
                for c, valor in enumerate(fila):
                    par = t.cell(r, c).paragraphs[0]
                    _negritas(par, valor, Pt(9))
                    if r == 0:
                        for run in par.runs:
                            run.bold = True
                        _sombra(t.cell(r, c), "E8EEF8")
            continue
        m = re.match(r"^(#{1,4})\s+(.*)", s)
        if m:
            p = w.add_paragraph()
            r = p.add_run(m.group(2).replace("**", ""))
            r.bold, r.font.color.rgb = True, NAVY
            r.font.size = Pt(14 if len(m.group(1)) <= 2 else 12)
            p.paragraph_format.space_before = Pt(10)
            continue
        m = re.match(r"^\s*([-*]|\d+\.)\s+(.*)", s)
        if m:
            _negritas(w.add_paragraph(style="List Bullet" if m.group(1) in "-*" else "List Number"), m.group(2))
            continue
        if s.startswith(">"):
            p = w.add_paragraph()
            _negritas(p, s.lstrip("> ").strip())
            p.paragraph_format.left_indent = Cm(0.8)
            for r in p.runs:
                r.italic, r.font.color.rgb = True, GRIS
            continue
        p = w.add_paragraph()
        _negritas(p, s.strip())
        p.paragraph_format.space_after = Pt(6)
    w.save(ruta)
    return Path(ruta)


def validacion_docx(hallazgos, meta: dict, cfg: dict, archivos: dict, ruta: Path) -> Path:
    """Informe de validación en Word: resumen por severidad y cada hallazgo con su regla, detalle y celdas."""
    w = _base("Informe de validación",
              f"{cfg.get('copropiedad', '')} · {meta.get('periodo_etiqueta', '')} · Excel «{archivos.get('excel', '')}»")
    c = Counter(h.severidad for h in hallazgos)
    t = w.add_table(rows=2, cols=4)
    t.style = "Table Grid"
    for j, s in enumerate(("ALTA", "MEDIA", "BAJA", "INFO")):
        r = t.cell(0, j).paragraphs[0].add_run(str(c.get(s, 0)))
        r.bold, r.font.size, r.font.color.rgb = True, Pt(20), COLOR_SEV[s]
        t.cell(1, j).paragraphs[0].add_run(s).font.size = Pt(9)
        for f in (0, 1):
            t.cell(f, j).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    w.add_paragraph()
    nota = w.add_paragraph()
    r = nota.add_run("Python calcula, la IA organiza y redacta, el profesional decide. Cada hallazgo indica la regla "
                     "que lo generó y las celdas del Excel donde revisarlo.")
    r.italic, r.font.color.rgb, r.font.size = True, GRIS, Pt(9)
    for origen, titulo in (("excel", "Hallazgos del Excel de papeles de trabajo"), ("pptx", "Revisión del informe generado")):
        grupo = [h for h in hallazgos if h.origen == origen]
        if not grupo:
            continue
        p = w.add_paragraph()
        r = p.add_run(titulo)
        r.bold, r.font.size, r.font.color.rgb = True, Pt(14), NAVY
        p.paragraph_format.space_before = Pt(12)
        for h in grupo:
            p = w.add_paragraph()
            r = p.add_run(f"{h.severidad}  ")
            r.bold, r.font.color.rgb = True, COLOR_SEV.get(h.severidad, GRIS)
            r = p.add_run(f"{h.regla} · {h.categoria} — {h.titulo}")
            r.bold = True
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.keep_with_next = True
            w.add_paragraph(h.detalle)
            if h.recomendacion:
                q = w.add_paragraph()
                q.add_run("Recomendación: ").bold = True
                q.add_run(h.recomendacion)
            if h.ubicacion:
                q = w.add_paragraph()
                r = q.add_run("Dónde revisar: " + ", ".join(h.ubicacion))
                r.font.size, r.font.color.rgb = Pt(9), GRIS
    w.save(ruta)
    return Path(ruta)
