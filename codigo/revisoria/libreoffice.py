"""Ubicación de LibreOffice y conversión de archivos (PDF del informe aprobado, recálculo de Excel)."""
from __future__ import annotations

import os
import shutil
import signal
import subprocess
import sys
import tempfile
from pathlib import Path

# En macOS la instalación normal no deja «soffice» en el PATH
RUTAS_CONOCIDAS = [Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")]


def buscar() -> str | None:
    """Ruta del ejecutable de LibreOffice, o None si no está instalado."""
    encontrado = shutil.which("soffice") or shutil.which("libreoffice")
    if encontrado:
        return encontrado
    return next((str(r) for r in RUTAS_CONOCIDAS if r.exists()), None)


def convertir(ruta: Path, formato: str, carpeta: Path, timeout: int = 240) -> Path | None:
    """Convierte `ruta` a `formato` (pdf, xlsx…) dentro de `carpeta`. Devuelve el archivo creado o None."""
    soffice = buscar()
    if not soffice:
        return None
    ejecutar([soffice, "--headless", "--convert-to", formato, "--outdir", str(carpeta), str(ruta)], timeout)
    salida = carpeta / f"{ruta.stem}.{formato}"
    if not salida.exists():
        global ultimo_error
        ultimo_error = ultimo_error or f"LibreOffice no produjo {salida.name}."
        print(f"[libreoffice] {ruta.name} → {formato}: {ultimo_error}", file=sys.stderr, flush=True)
    return salida if salida.exists() else None


ultimo_error = ""      # motivo de la última conversión fallida (para mostrarlo en la app)


def ejecutar(args: list[str], timeout: int) -> bool:
    """Corre LibreOffice con un perfil propio y desechable, y lo mata entero si se pasa del tiempo.

    - Perfil propio: si otro LibreOffice quedó abierto (o colgado) con el perfil por defecto, una llamada nueva le
      pasa el trabajo y sale sin convertir nada; con perfil propio cada conversión es independiente.
    - Grupo de procesos: «soffice» es un lanzador que abre «soffice.bin»; matar solo el lanzador al vencer el tiempo
      dejaba vivo a soffice.bin, que seguía ocupando memoria y bloqueaba las conversiones siguientes.
    """
    global ultimo_error
    ultimo_error = ""
    perfil = Path(tempfile.mkdtemp(prefix="lo_perfil_"))
    cmd = [args[0], f"-env:UserInstallation={perfil.as_uri()}", "--norestore", "--nolockcheck", *args[1:]]
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, start_new_session=True)
    try:
        _, err = proc.communicate(timeout=timeout)
        if proc.returncode:
            ultimo_error = (err or b"").decode(errors="replace").strip()[-400:] or f"código de salida {proc.returncode}"
            if proc.returncode < 0:
                ultimo_error = (f"LibreOffice se cerró de golpe (señal {-proc.returncode}); en un servidor pequeño "
                                f"suele ser falta de memoria. {ultimo_error}")
        return proc.returncode == 0
    except subprocess.TimeoutExpired:
        try:
            os.killpg(proc.pid, signal.SIGKILL)
        except OSError:
            proc.kill()
        proc.communicate()
        ultimo_error = f"LibreOffice tardó más de {timeout} s y se detuvo."
        return False
    finally:
        shutil.rmtree(perfil, ignore_errors=True)


# ------------------------------------------------------------------ anclas de imágenes

_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
       "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
       "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
       "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
       "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def _ruta(base: str, destino: str) -> str:
    """Resuelve el Target de una relación respecto de la carpeta de `base` dentro del zip."""
    import posixpath
    if destino.startswith("/"):
        return destino.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base), destino))


def _rels(z, parte: str) -> dict[str, str]:
    import posixpath
    import xml.etree.ElementTree as ET
    ruta = posixpath.join(posixpath.dirname(parte), "_rels", posixpath.basename(parte) + ".rels")
    if ruta not in z.namelist():
        return {}
    return {r.get("Id"): _ruta(parte, r.get("Target")) for r in ET.fromstring(z.read(ruta)).findall("rel:Relationship", _NS)}


def _imagenes_por_hoja(z) -> dict[str, list[tuple[str, str, object]]]:
    """{hoja: [(parte del dibujo, huella de la imagen, elemento ancla)]} en el orden del dibujo."""
    import hashlib
    import xml.etree.ElementTree as ET
    libro = ET.fromstring(z.read("xl/workbook.xml"))
    rels_libro = _rels(z, "xl/workbook.xml")
    out, dibujos = {}, {}
    for h in libro.find("main:sheets", _NS):
        hoja = rels_libro.get(h.get(f"{{{_NS['r']}}}id"))
        if not hoja or hoja not in z.namelist():
            continue
        for parte in _rels(z, hoja).values():
            if "/drawings/" not in parte or parte not in z.namelist():
                continue
            raiz = dibujos.setdefault(parte, ET.fromstring(z.read(parte)))
            rels_d = _rels(z, parte)
            for ancla in list(raiz):
                blip = ancla.find(".//a:blip", _NS)
                img = rels_d.get(blip.get(f"{{{_NS['r']}}}embed")) if blip is not None else None
                if img and img in z.namelist():
                    out.setdefault(h.get("name"), []).append((parte, hashlib.sha256(z.read(img)).hexdigest(), ancla))
    return out, dibujos


def restaurar_anclas_imagenes(original: Path, recalculado: Path) -> int:
    """LibreOffice recalcula las posiciones de las imágenes con sus propias alturas de fila y al guardar puede
    anclarlas a otra fila (el soporte de un comprobante pasa a otro). Devuelve las imágenes de `recalculado` a la
    celda que tenían en `original` (se emparejan por hoja y contenido). Devuelve cuántas se corrigieron."""
    import xml.etree.ElementTree as ET
    import zipfile
    for pref, uri in {"xdr": _NS["xdr"], "a": _NS["a"], "r": _NS["r"]}.items():
        ET.register_namespace(pref, uri)
    with zipfile.ZipFile(original) as zo:
        antes, _ = _imagenes_por_hoja(zo)
    with zipfile.ZipFile(recalculado) as zr:
        despues, dibujos = _imagenes_por_hoja(zr)
        partes = {n: zr.read(n) for n in zr.namelist()}
    corregidas, tocados = 0, set()
    for hoja, ims in despues.items():
        pendientes = {}
        for _, huella, ancla in antes.get(hoja, []):
            pendientes.setdefault(huella, []).append(ancla.find("xdr:from", _NS))
        for parte, huella, ancla in ims:
            if not pendientes.get(huella):
                continue
            ref, actual = pendientes[huella].pop(0), ancla.find("xdr:from", _NS)
            if ref is None or actual is None:
                continue
            delta = {k: int(ref.find(f"xdr:{k}", _NS).text) - int(actual.find(f"xdr:{k}", _NS).text) for k in ("row", "col")}
            if not any(delta.values()):
                continue
            for k in ("col", "colOff", "row", "rowOff"):
                actual.find(f"xdr:{k}", _NS).text = ref.find(f"xdr:{k}", _NS).text
            hasta = ancla.find("xdr:to", _NS)
            if hasta is not None:                            # se conserva el tamaño: el extremo se mueve igual
                for k in ("row", "col"):
                    hasta.find(f"xdr:{k}", _NS).text = str(int(hasta.find(f"xdr:{k}", _NS).text) + delta[k])
            corregidas += 1
            tocados.add(parte)
    if corregidas:
        for parte in tocados:
            partes[parte] = ET.tostring(dibujos[parte], xml_declaration=True, encoding="UTF-8")
        tmp = recalculado.with_name(f"_{recalculado.name}")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
            for n, datos in partes.items():
                zw.writestr(n, datos)
        tmp.replace(recalculado)
    return corregidas


# ------------------------------------------------------------------ imágenes vectoriales (EMF/WMF)

VECTORIALES = (".emf", ".wmf")


def _recortar_blanco(png: bytes) -> bytes:
    """LibreOffice pone el EMF sobre una página entera: se recorta el margen blanco (si Pillow está disponible)."""
    try:
        import io

        from PIL import Image, ImageChops
    except ImportError:
        return png
    try:
        im = Image.open(io.BytesIO(png)).convert("RGB")
        caja = ImageChops.difference(im, Image.new("RGB", im.size, (255, 255, 255))).convert("L").point(
            lambda v: 255 if v > 12 else 0).getbbox()
        if not caja:
            return png
        m = 6
        caja = (max(caja[0] - m, 0), max(caja[1] - m, 0), min(caja[2] + m, im.width), min(caja[3] + m, im.height))
        buf = io.BytesIO()
        im.crop(caja).save(buf, "PNG")
        return buf.getvalue()
    except Exception:
        return png


def imagenes_vectoriales(ruta: Path, hoja: str, timeout: int = 240) -> list[dict]:
    """openpyxl descarta las imágenes EMF/WMF al abrir el libro. Se leen directo del zip, se convierten a PNG con
    LibreOffice y se devuelven con su celda de anclaje: [{fila, col, bytes, ancho, alto, origen}]. Sin LibreOffice
    o si la conversión falla, devuelve [] (el informe sale igual, sin esas imágenes)."""
    import io
    import tempfile
    import zipfile
    try:
        with zipfile.ZipFile(ruta) as z:
            por_hoja, _ = _imagenes_por_hoja_con_media(z)
            anclas = por_hoja.get(hoja, [])
            if not anclas or not buscar():
                return []
            with tempfile.TemporaryDirectory() as d:
                d = Path(d)
                origen = []
                for k, (media, ancla) in enumerate(anclas):
                    f = d / f"img{k}{Path(media).suffix.lower()}"
                    f.write_bytes(z.read(media))
                    origen.append((f, media, ancla))
                ejecutar([buscar(), "--headless", "--convert-to", "png", "--outdir", str(d)]
                         + [str(f) for f, _, _ in origen], timeout)
                out = []
                for f, media, ancla in origen:
                    png = f.with_suffix(".png")
                    if not png.exists():
                        continue
                    datos = _recortar_blanco(png.read_bytes())
                    desde = ancla.find("xdr:from", _NS)
                    ext = ancla.find(".//a:xfrm/a:ext", _NS)
                    ancho = int(ext.get("cx")) // 9525 if ext is not None else None
                    alto = int(ext.get("cy")) // 9525 if ext is not None else None
                    try:                                    # tras el recorte se respeta la proporción real
                        from PIL import Image
                        w, h = Image.open(io.BytesIO(datos)).size
                        ancho = ancho or w
                        alto = round(ancho * h / w)
                    except Exception:
                        pass
                    out.append({"fila": int(desde.find("xdr:row", _NS).text) + 1,
                                "col": int(desde.find("xdr:col", _NS).text) + 1,
                                "bytes": datos, "ancho": ancho, "alto": alto, "origen": media})
                return out
    except Exception:
        return []


def _imagenes_por_hoja_con_media(z) -> tuple[dict[str, list[tuple[str, object]]], None]:
    """{hoja: [(ruta de la imagen vectorial en el zip, ancla)]}; solo EMF/WMF con ancla de celda."""
    import xml.etree.ElementTree as ET
    libro = ET.fromstring(z.read("xl/workbook.xml"))
    rels_libro = _rels(z, "xl/workbook.xml")
    out = {}
    for h in libro.find("main:sheets", _NS):
        hoja = rels_libro.get(h.get(f"{{{_NS['r']}}}id"))
        if not hoja or hoja not in z.namelist():
            continue
        for parte in _rels(z, hoja).values():
            if "/drawings/" not in parte or parte not in z.namelist():
                continue
            rels_d = _rels(z, parte)
            for ancla in list(ET.fromstring(z.read(parte))):
                if ancla.find("xdr:from", _NS) is None:
                    continue
                blip = ancla.find(".//a:blip", _NS)
                img = rels_d.get(blip.get(f"{{{_NS['r']}}}embed")) if blip is not None else None
                if img and img.lower().endswith(VECTORIALES) and img in z.namelist():
                    out.setdefault(h.get("name"), []).append((img, ancla))
    return out, None
