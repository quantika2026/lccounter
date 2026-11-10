"""Conceptos del revisor fiscal para todos los clientes (o algunos), cuando haga falta: no es un informe mensual.

Cada concepto lleva:
  1. Texto explicativo detallado (borrador de IA que el revisor edita). Solo puede apoyarse en la OPINIÓN escrita por
     el revisor y en los fragmentos de norma que él eligió de la base normativa: nada de memoria.
  2. Infografía: «Lo que dice la norma» frente a «Opinión del revisor fiscal» (imagen que arma Python).
  3. La norma como anexo: el texto literal de los fragmentos elegidos, tal como están en la base normativa.
Se emite como un Word por copropiedad (con su nombre y su logo) más un ZIP con todos. La aplicación no envía nada:
el revisor descarga y comparte.
"""
from __future__ import annotations

import json
import re
import textwrap
import zipfile
from datetime import datetime
from pathlib import Path

from agente.comun import Espacio, ahora, guardar_json, leer_json

SISTEMA = """Eres especialista en redacción institucional de LC Counter Contadores y Auditores SAS (Revisoría Fiscal \
de propiedades horizontales en Colombia). Redactas el BORRADOR de un concepto del revisor fiscal dirigido a todas \
las copropiedades clientes.

Reglas:
- Tus únicas fuentes son la OPINIÓN DEL REVISOR (F0) y los fragmentos de norma (F1, F2, …). Usa solo los números \
de fuente que aparecen en FUENTES: si hay un solo fragmento, la norma es F1. No agregues normas, \
artículos, fechas, cifras ni opiniones que no estén ahí. Cita así: (F2, Ley 675 de 2001, artículo 35).
- La opinión y las recomendaciones son las del revisor: explícalas y ordénalas, no las cambies ni las amplíes.
- Si algo hace falta, deja un marcador entre corchetes, por ejemplo [fecha de entrada en vigencia].
- Estilo formal, claro y pedagógico para consejos de administración; tercera persona («esta Revisoría Fiscal»).
- No escribas el nombre de una copropiedad: el mismo texto va para todas (el encabezado lo pone el sistema)."""

ESQUEMA = {
    "type": "object",
    "properties": {
        "asunto": {"type": "string"},
        "texto": {"type": "string", "description": "Concepto completo en Markdown: contexto, qué dice la norma, "
                                                    "concepto del revisor, recomendaciones y conclusión."},
        "infografia": {"type": "object", "properties": {
            "titulo": {"type": "string"},
            "norma": {"type": "array", "items": {"type": "string"}, "description": "3 a 5 ideas cortas de la norma, cada una con su cita."},
            "opinion": {"type": "array", "items": {"type": "string"}, "description": "3 a 5 ideas cortas de la opinión del revisor."},
            "recomendacion": {"type": "string"}},
            "required": ["titulo", "norma", "opinion", "recomendacion"], "additionalProperties": False},
    },
    "required": ["asunto", "texto", "infografia"],
    "additionalProperties": False,
}


def carpeta(esp: Espacio) -> Path:
    return esp.raiz / "conceptos"


def _ruta(esp: Espacio, cid: str) -> Path:
    if not re.match(r"^\d{8}-\d{6}$", cid or ""):
        raise FileNotFoundError("Concepto inexistente.")
    return carpeta(esp) / cid


def listar(esp: Espacio) -> list[dict]:
    c = carpeta(esp)
    out = [leer_json(p / "concepto.json", {}) for p in sorted(c.glob("*"), reverse=True)] if c.exists() else []
    return [{k: x.get(k) for k in ("id", "tema", "asunto", "estado", "creado", "por", "emitido", "clientes")} for x in out if x]


def leer(esp: Espacio, cid: str) -> dict:
    d = leer_json(_ruta(esp, cid) / "concepto.json", None)
    if not d:
        raise FileNotFoundError("Concepto inexistente.")
    return d


def fuentes(opinion: str, fragmentos: list[dict]) -> str:
    partes = [f"[F0] OPINIÓN DEL REVISOR FISCAL\n{opinion.strip()}"]
    for i, f in enumerate(fragmentos, start=1):
        partes.append(f"[F{i}] Norma: {f['documento']}, {f['referencia']}\n{f['texto']}")
    return "\n\n".join(partes)


def crear(esp: Espacio, tema: str, opinion: str, fragmentos: list[dict], clientes: list[str], autor: str,
          redactar=None) -> dict:
    """Arma el borrador. `redactar(sistema, usuario, esquema)` permite probar sin proveedor de IA."""
    from . import proveedores
    if not tema.strip() or len(opinion.strip()) < 20:
        raise ValueError("Escriba el tema y la opinión del revisor (al menos un par de frases).")
    if not fragmentos:
        raise ValueError("Elija al menos un fragmento de norma de la base normativa.")
    ids = esp.ids_clientes()
    clientes = [c for c in (clientes or ids) if c in ids]
    if not clientes:
        raise ValueError("No hay clientes destinatarios.")
    usuario = f"TEMA: {tema.strip()}\n\nFUENTES:\n\n{fuentes(opinion, fragmentos)}"
    if redactar:
        r = redactar(SISTEMA, usuario, ESQUEMA)
    else:
        r, _ = proveedores.json_estructurado(esp.raiz, "comunicaciones", SISTEMA, usuario, ESQUEMA)
    cid = datetime.now().strftime("%Y%m%d-%H%M%S")
    d = {"id": cid, "tema": tema.strip(), "opinion": opinion.strip(), "fragmentos": fragmentos, "clientes": clientes,
         "asunto": r["asunto"], "texto": r["texto"], "infografia": r["infografia"], "estado": "borrador",
         "creado": ahora(), "por": autor, "emitido": None}
    d["citas_invalidas"] = citas_invalidas(d)
    _ruta(esp, cid).mkdir(parents=True, exist_ok=True)
    guardar_json(_ruta(esp, cid) / "concepto.json", d)
    return d


def citas_invalidas(d: dict) -> list[str]:
    """Citas a fuentes que no existen (p. ej. «F2» cuando solo hay F0 y F1): el revisor debe corregirlas."""
    validas = {f"F{i}" for i in range(len(d.get("fragmentos") or []) + 1)}
    texto = d.get("texto", "") + " " + json.dumps(d.get("infografia") or {}, ensure_ascii=False)
    return sorted(set(re.findall(r"\bF\d+\b", texto)) - validas)


def guardar(esp: Espacio, cid: str, cambios: dict, autor: str) -> dict:
    d = leer(esp, cid)
    if d["estado"] == "emitido":
        raise ValueError("El concepto ya se emitió: queda congelado. Cree uno nuevo si hace falta corregirlo.")
    for k in ("asunto", "texto", "infografia", "clientes"):
        if k in cambios and cambios[k] is not None:
            d[k] = cambios[k]
    d["editado"] = {"por": autor, "fecha": ahora()}
    d["citas_invalidas"] = citas_invalidas(d)
    guardar_json(_ruta(esp, cid) / "concepto.json", d)
    return d


# ------------------------------------------------------------------ infografía (la arma Python)

def infografia(d: dict, ruta: Path, logo: Path | None = None) -> Path:
    from PIL import Image, ImageDraw, ImageFont
    W, NAVY, VINO, NARANJA, GRIS, FONDO = 1600, (31, 58, 122), (142, 42, 78), (245, 160, 68), (71, 85, 105), (244, 246, 250)

    def fuente(tam, negrita=False):
        for nombre in (("Arial Bold.ttf", "Helvetica-Bold.ttc", "DejaVuSans-Bold.ttf") if negrita else
                       ("Arial.ttf", "Helvetica.ttc", "DejaVuSans.ttf")):
            for base in ("/System/Library/Fonts/Supplemental/", "/System/Library/Fonts/", "/Library/Fonts/", ""):
                try:
                    return ImageFont.truetype(base + nombre, tam)
                except OSError:
                    continue
        return ImageFont.load_default()

    ig = d["infografia"]
    f_tit, f_col, f_txt, f_pie = fuente(46, True), fuente(32, True), fuente(26), fuente(22)
    col_w = (W - 3 * 60) // 2
    envolver = lambda t: textwrap.wrap(t, width=46)
    def alto(items):
        return sum(len(envolver(x)) * 36 + 26 for x in items)
    cuerpo = max(alto(ig["norma"]), alto(ig["opinion"])) + 110
    rec = textwrap.wrap(ig.get("recomendacion", ""), width=100)
    H = 200 + cuerpo + 80 + len(rec) * 36 + 140
    im = Image.new("RGB", (W, H), FONDO)
    dr = ImageDraw.Draw(im)
    dr.rectangle([0, 0, W, 150], fill=NAVY)
    y = 40
    for linea in textwrap.wrap(ig["titulo"], width=48)[:2]:
        dr.text((60, y), linea, font=f_tit, fill="white"); y += 54
    if logo and logo.exists():
        lg = Image.open(logo).convert("RGBA"); lg.thumbnail((140, 110))
        caja = Image.new("RGBA", (lg.width + 20, lg.height + 20), "white"); caja.paste(lg, (10, 10), lg)
        im.paste(caja.convert("RGB"), (W - caja.width - 40, 20))
    for i, (titulo, items, color) in enumerate((("Lo que dice la norma", ig["norma"], NAVY),
                                                ("Opinión del revisor fiscal", ig["opinion"], VINO))):
        x0, y0 = 60 + i * (col_w + 60), 200
        dr.rounded_rectangle([x0, y0, x0 + col_w, y0 + cuerpo], radius=24, fill="white", outline=color, width=4)
        dr.rounded_rectangle([x0, y0, x0 + col_w, y0 + 80], radius=24, fill=color)
        dr.rectangle([x0, y0 + 50, x0 + col_w, y0 + 80], fill=color)
        dr.text((x0 + 30, y0 + 20), titulo, font=f_col, fill="white")
        y = y0 + 110
        for it in items:
            dr.ellipse([x0 + 30, y + 10, x0 + 46, y + 26], fill=NARANJA)
            for linea in envolver(it):
                dr.text((x0 + 62, y), linea, font=f_txt, fill=GRIS); y += 36
            y += 26
    y = 200 + cuerpo + 40
    dr.rounded_rectangle([60, y, W - 60, y + 60 + len(rec) * 36], radius=20, fill=(255, 244, 230), outline=NARANJA, width=3)
    dr.text((90, y + 14), "Recomendación", font=f_col, fill=(180, 83, 9))
    yy = y + 56
    for linea in rec:
        dr.text((90, yy), linea, font=f_txt, fill=GRIS); yy += 36
    dr.text((60, H - 60), "LC Counter Contadores y Auditores SAS · Revisoría Fiscal", font=f_pie, fill=GRIS)
    im.save(ruta)
    return ruta


# ------------------------------------------------------------------ emisión: un Word por copropiedad + ZIP

def emitir(esp: Espacio, cid: str, revisor: str) -> dict:
    from docx.shared import Cm, Pt
    from revisoria import word
    d = leer(esp, cid)
    if d["estado"] == "emitido":
        return d
    dir_c = _ruta(esp, cid)
    salida = dir_c / "emitidos"
    salida.mkdir(exist_ok=True)
    archivos = []
    for cli in d["clientes"]:
        try:
            ficha = esp.ficha(cli)
        except FileNotFoundError:
            continue
        logo = next((p for p in (esp.clientes / cli).glob("logo.*") if p.suffix in (".png", ".jpg", ".webp")), None)
        img = infografia(d, salida / f"infografia_{cli}.png",
                         logo or Path(__file__).resolve().parent / "static" / "logo.png")   # sin logo propio: el del despacho
        cop = ficha.get("copropiedad", cli)
        md = (f"**Señores Consejo de Administración**  \n{cop}\n\n**Asunto:** {d['asunto']}\n\n{d['texto']}")
        ruta = salida / f"Concepto_{re.sub(r'[^A-Za-z0-9]+', '_', cop).strip('_')}_{cid}.docx"
        word.markdown_a_docx(md, ruta, "Concepto del revisor fiscal", f"{cop} · {datetime.now():%d/%m/%Y}")
        from docx import Document
        doc = Document(ruta)
        doc.add_page_break()
        p = doc.add_paragraph(); r = p.add_run("Infografía"); r.bold, r.font.size = True, Pt(14)
        doc.add_picture(str(img), width=Cm(16.5))
        doc.add_page_break()
        p = doc.add_paragraph(); r = p.add_run("Anexo: texto de la norma citada"); r.bold, r.font.size = True, Pt(14)
        nota = doc.add_paragraph(); nota.add_run("Texto tomado literalmente de la base normativa del despacho "
                                                 "(versión indicada en cada documento).").italic = True
        for i, f in enumerate(d["fragmentos"], start=1):
            h = doc.add_paragraph(); h.add_run(f"F{i}. {f['documento']} — {f['referencia']}"
                                               + (f" ({f['version']})" if f.get("version") else "")).bold = True
            for parrafo in str(f["texto"]).split("\n"):
                if parrafo.strip():
                    doc.add_paragraph(parrafo.strip())
        doc.save(ruta)
        archivos.append(ruta)
    paquete = dir_c / f"Conceptos_{cid}.zip"
    with zipfile.ZipFile(paquete, "w", zipfile.ZIP_DEFLATED) as z:
        for a in archivos:
            z.write(a, a.name)
    d.update({"estado": "emitido", "emitido": {"por": revisor, "fecha": ahora(), "archivos": [a.name for a in archivos],
                                               "zip": paquete.name}})
    guardar_json(dir_c / "concepto.json", d)
    return d


def archivo(esp: Espacio, cid: str, nombre: str) -> Path:
    d = leer(esp, cid)
    permitidos = set((d.get("emitido") or {}).get("archivos", [])) | {(d.get("emitido") or {}).get("zip")}
    if nombre not in permitidos:
        raise FileNotFoundError("Archivo inexistente.")
    base = _ruta(esp, cid)
    return base / nombre if nombre.endswith(".zip") else base / "emitidos" / nombre
