"""Visor y editor de diapositivas del informe.

- Render: PPTX → PDF (LibreOffice) → PNG por diapositiva (PyMuPDF), en caché por huella del archivo.
- Textos editables: los mismos que compara la mejora continua (`agente.mejora._unidades`): párrafos de cuadros de
  texto (`s<k>.p<j>`) y celdas de tabla (`s<k>.t<r>.<c>`). Así, al aprobar con el PPTX editado, cada cambio se
  captura y se clasifica (texto, cifra o estructura).
- Edición: se escribe una COPIA (`Informe_Editado_...pptx`); el informe generado nunca se modifica.
"""
from __future__ import annotations

import hashlib
import re
import shutil
import threading
from pathlib import Path

from agente.comun import ahora, guardar_json, leer_json

_candado = threading.Lock()          # LibreOffice no admite conversiones simultáneas con el mismo perfil
RX_UNIDAD = re.compile(r"^s(\d+)\.(?:p(\d+)|t(\d+)\.(\d+))$")


def huella(ruta: Path) -> str:
    return hashlib.sha256(ruta.read_bytes()).hexdigest()[:16] if ruta.exists() else ""


def rutas(sal: Path, cid: str, periodo: str, generado: str) -> dict:
    return {"generado": sal / generado, "editado": sal / f"Informe_Editado_{cid}_{periodo}.pptx",
            "meta": sal / "informe_editado.json", "vista": sal / "vista"}


def estado_edicion(r: dict) -> dict | None:
    if not r["editado"].exists():
        return None
    m = leer_json(r["meta"], {})
    return {"guardado": m.get("fecha"), "por": m.get("por", ""), "cambios": m.get("cambios", 0),
            "desactualizado": m.get("base") != huella(r["generado"])}


def unidades(ruta: Path) -> list[dict]:
    from agente.mejora import _unidades
    datos = _unidades(ruta)
    for d in datos:
        d["u"] = {k: v for k, v in d["u"].items() if ".g" not in k}
    return datos


# ------------------------------------------------------------------ render

def render(pptx: Path, carpeta_vista: Path, escala: float = 1.6) -> tuple[str, int]:
    """Devuelve (huella, número de imágenes). Reutiliza el render si el archivo no cambió."""
    import pymupdf

    from revisoria.libreoffice import buscar, convertir
    h = huella(pptx)
    destino = carpeta_vista / h
    hecho = destino / "listo.json"
    if hecho.exists():
        return h, leer_json(hecho, {}).get("n", 0)
    if not buscar():
        raise RuntimeError("No se encontró LibreOffice: se necesita para mostrar las diapositivas.")
    with _candado:
        if hecho.exists():
            return h, leer_json(hecho, {}).get("n", 0)
        tmp = destino.with_name(h + "_tmp")
        shutil.rmtree(tmp, ignore_errors=True)
        tmp.mkdir(parents=True)
        pdf = convertir(pptx, "pdf", tmp, timeout=300)
        if not pdf:
            raise RuntimeError("LibreOffice no pudo convertir el informe a PDF.")
        doc = pymupdf.open(pdf)
        for i, pagina in enumerate(doc, 1):
            pagina.get_pixmap(matrix=pymupdf.Matrix(escala, escala)).save(tmp / f"{i:03d}.png")
        n = doc.page_count
        doc.close()
        pdf.unlink()
        guardar_json(tmp / "listo.json", {"n": n, "fecha": ahora()})
        shutil.rmtree(destino, ignore_errors=True)
        tmp.rename(destino)
    # conserva solo los dos renders más recientes (generado y editado)
    viejos = sorted((p for p in carpeta_vista.iterdir() if p.is_dir() and p.name != h), key=lambda p: p.stat().st_mtime)
    for p in viejos[:-1]:
        shutil.rmtree(p, ignore_errors=True)
    return h, n


# ------------------------------------------------------------------ edición

def _poner_texto(parrafo, texto: str):
    """Reemplaza el texto del párrafo conservando el formato del primer fragmento."""
    runs = parrafo.runs
    if not runs:
        parrafo.add_run().text = texto
        return
    runs[0].text = texto
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def aplicar(base: Path, destino: Path, cambios: dict[str, dict[str, str]]) -> int:
    """cambios = {"<n diapositiva>": {"s3.p0": "texto nuevo", "s5.t2.1": "celda"}}. Devuelve cuántos aplicó."""
    from pptx import Presentation
    prs = Presentation(str(base))
    diap = list(prs.slides)
    n = 0
    for num, porunidad in cambios.items():
        i = int(num) - 1
        if not 0 <= i < len(diap):
            raise ValueError(f"No existe la diapositiva {num}.")
        formas = list(diap[i].shapes)
        for clave, texto in porunidad.items():
            m = RX_UNIDAD.match(clave)
            if not m or int(m.group(1)) >= len(formas):
                raise ValueError(f"Elemento inválido «{clave}» en la diapositiva {num}.")
            forma = formas[int(m.group(1))]
            texto = str(texto).replace("\r", "")
            if m.group(2) is not None:
                if not forma.has_text_frame or int(m.group(2)) >= len(forma.text_frame.paragraphs):
                    raise ValueError(f"Elemento inválido «{clave}» en la diapositiva {num}.")
                _poner_texto(forma.text_frame.paragraphs[int(m.group(2))], texto)
            else:
                if not getattr(forma, "has_table", False) or not forma.has_table:
                    raise ValueError(f"Elemento inválido «{clave}» en la diapositiva {num}.")
                f, c = int(m.group(3)), int(m.group(4))
                tabla = forma.table
                if f >= len(tabla.rows) or c >= len(tabla.columns):
                    raise ValueError(f"Elemento inválido «{clave}» en la diapositiva {num}.")
                pars = tabla.cell(f, c).text_frame.paragraphs
                _poner_texto(pars[0], texto)
                for p in pars[1:]:                       # la celda queda con un solo párrafo
                    p._p.getparent().remove(p._p)
            n += 1
    tmp = destino.with_suffix(".tmp.pptx")
    prs.save(str(tmp))
    tmp.replace(destino)
    return n


def clasificar(antes: str, despues: str) -> str:
    """Igual que la mejora continua: si cambian los números es «cifra», si no es «texto»."""
    from collections import Counter

    from agente.mejora import RX_DIG
    return "cifra" if Counter(RX_DIG.findall(antes or "")) != Counter(RX_DIG.findall(despues or "")) else "texto"
