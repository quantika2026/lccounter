"""Migración de un cliente a la plantilla estándar (opción C): se hace UNA vez por cliente.

Parte del Excel original (no de cero, para conservar fórmulas, formato e histórico):
  · hojas que el lector estándar ya entiende → se RENOMBRAN al nombre oficial (y se corrigen las fórmulas que las citan);
  · hojas con otra estructura (tienen receta) → se RECONSTRUYEN en la disposición estándar, con el origen de cada dato
    en un comentario; la hoja original queda como «ANEXO»;
  · secciones que el cliente no tiene → se crea la hoja estándar vacía (si la plantilla tiene escritor para ella);
  · hojas que sobran → quedan como «ANEXO - …» (no se pierde nada).
Después se recalcula con LibreOffice y se VERIFICA: el archivo migrado se lee con el lector estándar (sin mapa ni
recetas) y cada sección debe dar lo mismo que el original; también se comparan los hallazgos del validador.

El resultado es una PROPUESTA: la revisa y la aprueba el revisor fiscal. No reemplaza el Excel del cliente.
"""
from __future__ import annotations

import json
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl

from revisoria import documental, esquema, lectores, libreoffice, mapa as mapa_cliente, plantilla as PL
from revisoria.excel import Libro, Periodo, norm
from revisoria.extractor import extraer
from revisoria.validador import validar_excel

from .comun import guardar_json
from . import normalizador

IGNORAR = ("fuente", "_", "fila", "formula", "imagenes", "soportes", "seg_")
FORMATO = ("seg",)                   # segmentos de texto con formato: se comparan aparte (aviso, no error)
NORMALIZADOS = {"titulo"}            # títulos que la plantilla escribe en su formato estándar (se informan, no fallan)


def _limpio(x):
    if isinstance(x, dict):
        out = {k: _limpio(v) for k, v in x.items() if not k.startswith(IGNORAR)}
        if "imagenes" in x:
            out["#imagenes"] = len(x.get("imagenes") or [])
        if "soportes" in x:
            out["#soportes"] = len(x.get("soportes") or [])
        return out
    if isinstance(x, list):
        return [_limpio(v) for v in x]
    if isinstance(x, float):
        return round(x, 2)
    return x


def _diferencias(a, b, ruta="") -> list[str]:
    if isinstance(a, dict) and isinstance(b, dict):
        out = []
        for k in sorted(set(a) | set(b)):
            out += _diferencias(a.get(k), b.get(k), f"{ruta}/{k}")
        return out
    if isinstance(a, list) and isinstance(b, list):
        if len(a) != len(b):
            return [f"{ruta}: {len(a)} elementos en el original, {len(b)} en el migrado"]
        out = []
        for i, (x, y) in enumerate(zip(a, b)):
            out += _diferencias(x, y, f"{ruta}[{i}]")
        return out
    return [] if a == b else [f"{ruta}: original «{str(a)[:80]}» · migrado «{str(b)[:80]}»"]


def _j(x):
    return _limpio(json.loads(json.dumps(x, default=str)))


def _citar(titulo: str) -> str:
    return f"'{titulo}'" if re.search(r"[^A-Za-z0-9_]", titulo) else titulo


def _actualizar_formulas(wb, renombres: dict[str, str]):
    """openpyxl no corrige las referencias al renombrar hojas: se reemplazan en todas las fórmulas."""
    if not renombres:
        return
    pares = sorted(renombres.items(), key=lambda kv: -len(kv[0]))
    for ws in wb.worksheets:
        for fila in ws.iter_rows():
            for c in fila:
                v = c.value
                if isinstance(v, str) and v.startswith("=") and "!" in v:
                    for viejo, nuevo in pares:
                        v = v.replace(f"'{viejo}'!", f"{_citar(nuevo)}!")
                        if not re.search(r"[^A-Za-z0-9_]", viejo):
                            v = re.sub(rf"(?<![A-Za-z0-9_'.]){re.escape(viejo)}!", f"{_citar(nuevo)}!", v)
                    c.value = v


def _nombre_anexo(titulo: str, usados: set) -> str:
    base = f"ANEXO - {titulo}".strip()[:31]
    nombre, i = base, 2
    while nombre in usados:
        nombre = f"{base[:28]} {i}"
        i += 1
    usados.add(nombre)
    return nombre


def _reponer_vectoriales(wb, excel: Path, renombres: dict[str, str]) -> int:
    """openpyxl pierde las imágenes EMF/WMF al guardar: se reponen convertidas a PNG en su celda original."""
    import io
    from openpyxl.drawing.image import Image
    from openpyxl.utils import get_column_letter
    with zipfile.ZipFile(excel) as z:
        hojas = list(libreoffice._imagenes_por_hoja_con_media(z)[0])
    n = 0
    for original in hojas:
        titulo = renombres.get(original, original)
        if titulo not in wb.sheetnames:
            continue
        for v in libreoffice.imagenes_vectoriales(excel, original):
            im = Image(io.BytesIO(v["bytes"]))
            if v["ancho"] and v["alto"]:
                im.width, im.height = v["ancho"], v["alto"]
            wb[titulo].add_image(im, f"{get_column_letter(v['col'])}{v['fila']}")
            n += 1
    return n


def migrar(esp, cid: str, periodo: str, recetas_extra: dict | None = None, destino: Path | None = None, log=print) -> dict:
    p = Periodo.desde_texto(periodo)
    ficha = esp.ficha(cid)
    excel = normalizador.ubicar_excel(esp.entrada(periodo, cid), ficha)
    if not excel:
        raise ValueError(f"No hay Excel de {cid} para {periodo}.")
    mapa = json.loads(json.dumps(mapa_cliente.cargar(esp.clientes / cid) or mapa_cliente.detectar(Libro(str(excel)))))
    recetas = dict(mapa.get("lectores") or {})
    propuestas = sorted(set(recetas_extra or {}) - set(recetas))
    recetas.update(recetas_extra or {})
    mapa["lectores"] = recetas
    for sec, rec in recetas.items():                      # la hoja de una receta queda en el mapa
        mapa.setdefault("hojas", {})[PL.SECCION_HOJA[sec]] = rec["hoja"]
    meta = {k: ficha.get(k) for k in ("copropiedad", "nit", "unidades", "numero_informe") if ficha.get(k)}

    destino = Path(destino or esp.clientes / cid / "migracion" / f"{periodo}_{datetime.now():%Y%m%d_%H%M%S}")
    destino.mkdir(parents=True, exist_ok=True)
    shutil.copy2(excel, destino / f"original_{excel.name}")

    # 1. lectura del original (lector estándar + mapa + recetas): es la referencia que el migrado debe reproducir
    log("Leyendo el original…")
    lib_o = Libro(str(excel), mapa)
    mod_o, lib_o = extraer(str(excel), p, dict(meta), mapa, lib_o)
    doc_o = documental.extraer_documental(lib_o, p)
    orig = {**mod_o, **doc_o}
    h_o = validar_excel(mod_o, lib_o, p)

    # 2. plan por hoja lógica
    wb = openpyxl.load_workbook(excel, rich_text=True)
    titulos = set(wb.sheetnames)
    hojas = mapa.get("hojas") or {}
    secs_por_clave: dict[str, list[str]] = {}
    for sec, clave in PL.SECCION_HOJA.items():
        secs_por_clave.setdefault(clave, []).append(sec)
    plan, renombres, anexos, usados = [], {}, [], set()
    temporales = {}
    for i, t in enumerate(wb.sheetnames):                 # dos fases para que no choquen los nombres
        temporales[t] = f"__tmp{i}"
        wb[t].title = temporales[t]
    destino_titulo: dict[str, str] = {}
    reconstruir: list[str] = []
    for clave, oficial in PL.HOJAS.items():
        v = hojas.get(clave)
        actuales = [x[0] for x in v] if isinstance(v, list) else [v] if v else []
        actuales = [t for t in actuales if t in titulos]
        con_receta = [s for s in secs_por_clave.get(clave, []) if s in recetas]
        if con_receta:
            reconstruir.append(clave)
            for t in actuales:
                destino_titulo.setdefault(t, None)       # pasa a anexo
            plan.append({"hoja": oficial, "metodo": "reconstruida", "origen": actuales, "secciones": con_receta,
                         "recetas_propuestas": [s for s in con_receta if s in propuestas]})
        elif actuales:
            if clave == "AUDITORIA DOCUMENTAL" and len(actuales) > 1:
                bloques = {t: b for t, b in (v if isinstance(v, list) else [])}
                for t in actuales:
                    n = "AUD DOC CONTABLE" if bloques.get(t) == "Contable" else "AUD DOC ADMINISTRATIVA" \
                        if bloques.get(t) == "Administrativo" else _nombre_anexo(t, usados)
                    destino_titulo[t] = n
                plan.append({"hoja": " + ".join(destino_titulo[t] for t in actuales), "metodo": "renombrada",
                             "origen": actuales})
            else:
                destino_titulo[actuales[0]] = oficial
                plan.append({"hoja": oficial, "metodo": "renombrada" if actuales[0] != oficial else "igual",
                             "origen": actuales[:1]})
        elif clave in secs_por_clave and clave not in PL.SIN_HOJA_VACIA and any(
                s in PL.ESCRITORES or s in ("anticipos", "deudores_varios") for s in secs_por_clave[clave]):
            reconstruir.append(clave)
            plan.append({"hoja": oficial, "metodo": "vacía (el cliente no la tiene)", "origen": [],
                         "secciones": secs_por_clave[clave]})
        elif clave in PL.SIN_HOJA_VACIA:
            plan.append({"hoja": oficial, "metodo": "no se crea (el cliente no la tiene; el informe usa la evolución "
                                                    "de MAYORES DEUDORES)", "origen": []})
        else:
            plan.append({"hoja": oficial, "metodo": "no existe (sin hoja en la plantilla todavía)", "origen": []})
    usados |= {n for n in destino_titulo.values() if n} | set(PL.HOJAS.values())
    for t in wb.sheetnames:                              # títulos temporales → finales
        original = next(o for o, tmp in temporales.items() if tmp == t)
        nuevo = destino_titulo.get(original) or _nombre_anexo(original, usados)
        if not destino_titulo.get(original):
            anexos.append({"original": original, "anexo": nuevo})
        wb[t].title = nuevo
        if nuevo != original:
            renombres[original] = nuevo
    _actualizar_formulas(wb, renombres)

    # 3. reconstrucción de las secciones con receta y hojas vacías de la plantilla
    for clave in reconstruir:
        ws = wb.create_sheet(PL.HOJAS[clave])
        secs = secs_por_clave[clave]
        con_datos = orig.get                              # todas las secciones de la hoja, con o sin receta
        if clave == "ANTIC. E INGRESOS":
            PL.escribir_anticipos(ws, con_datos("anticipos"), con_datos("deudores_varios"), p)
        else:
            for s in secs:
                PL.ESCRITORES[s](ws, con_datos(s), p)
    # orden: hojas oficiales primero, anexos al final
    oficiales = [ws for ws in wb.worksheets if not ws.title.startswith("ANEXO - ")]
    wb._sheets = oficiales + [ws for ws in wb.worksheets if ws.title.startswith("ANEXO - ")]

    # 4. hoja de control
    wc = wb.create_sheet(PL.HOJA_CONTROL, 0)
    filas = [("Plantilla estándar LCCOUNTER", f"versión {PL.VERSION}"), ("Cliente", f"{ficha.get('copropiedad')} ({cid})"),
             ("Periodo de corte", p.etiqueta), ("Excel original", excel.name), ("Fecha de migración", datetime.now().isoformat(timespec="seconds")),
             ("Estado", "PROPUESTA — pendiente de revisión y aprobación del revisor fiscal"), ("", ""),
             ("Hoja de la plantilla", "Método · hoja(s) de origen")]
    filas += [(x["hoja"], f"{x['metodo']} · {', '.join(x['origen']) or '—'}") for x in plan]
    filas += [("", ""), ("Anexo", "Hoja original")] + [(a["anexo"], a["original"]) for a in anexos]
    for r, (a, b) in enumerate(filas, start=1):
        wc.cell(r, 1, a)
        wc.cell(r, 2, b)
        if r in (1, 8) or a == "Anexo":
            wc.cell(r, 1).font = PL.NEGRILLA
    wc.column_dimensions["A"].width, wc.column_dimensions["B"].width = 34, 90

    n_vect = _reponer_vectoriales(wb, excel, renombres)
    if n_vect:
        log(f"Imágenes EMF/WMF del original repuestas como PNG: {n_vect}")

    sin_recalcular = destino / f"_sin_recalcular_{cid}_{periodo}.xlsx"
    wb.save(sin_recalcular)
    log("Recalculando con LibreOffice…")
    tmp = destino / "_lo"
    tmp.mkdir(exist_ok=True)
    recalculado = libreoffice.convertir(sin_recalcular, "xlsx", tmp)
    migrado = destino / f"Plantilla_estandar_{cid}_{periodo}.xlsx"
    if recalculado:
        n_anclas = libreoffice.restaurar_anclas_imagenes(sin_recalcular, recalculado)
        if n_anclas:
            log(f"Imágenes devueltas a su celda original después de recalcular: {n_anclas}")
        shutil.move(str(recalculado), migrado)
        sin_recalcular.unlink()
    else:
        shutil.move(str(sin_recalcular), migrado)
    shutil.rmtree(tmp, ignore_errors=True)

    # 5. verificación: el migrado se lee con el lector ESTÁNDAR, sin mapa ni recetas
    log("Verificando con el lector estándar…")
    lib_m = Libro(str(migrado))
    errores = []
    try:
        mod_m, lib_m = extraer(str(migrado), p, dict(meta), None, lib_m)
        doc_m = documental.extraer_documental(lib_m, p)
        nuevo = {**mod_m, **doc_m}
        h_m = validar_excel(mod_m, lib_m, p)
    except Exception as e:                               # el migrado no se deja leer: la migración no sirve
        errores.append(f"El lector estándar no pudo leer el archivo migrado: {type(e).__name__}: {e}")
        nuevo, h_m = {}, []
    verificacion = []
    for s in [x.clave for x in esquema.SECCIONES]:
        if s not in orig or orig.get(s) is None:
            verificacion.append({"seccion": s, "resultado": "sin datos en el original", "diferencias": []})
            continue
        difs = _diferencias(_j(orig[s]), _j(nuevo.get(s)))
        formato = [d for d in difs if "/seg" in d.split(":")[0]]
        difs = [d for d in difs if d not in formato]
        normalizadas = [d for d in difs if d.split(":")[0].split("/")[-1] in NORMALIZADOS]
        limitaciones = [d for d in difs if s == "actas_consejo" and d.startswith("/todas")]
        reales = [d for d in difs if d not in normalizadas and d not in limitaciones]
        verificacion.append({"seccion": s, "resultado": "diferente" if reales else "idéntica (con normalizaciones)"
                             if normalizadas or limitaciones else "idéntica",
                             "diferencias": reales, "normalizaciones": normalizadas + limitaciones,
                             "formato": formato[:5] and [f"{len(formato)} diferencia(s) en negrillas/rojos/saltos de línea del texto"]})
    def clave_h(h, traducir=False):
        t = h.titulo
        if traducir:                                     # el hallazgo cita la hoja por su nombre: se traduce
            for viejo, nuevo in sorted(renombres.items(), key=lambda kv: -len(kv[0])):
                t = t.replace(f"'{viejo}'", f"'{nuevo}'").replace(f"«{viejo}»", f"«{nuevo}»")
        return (h.regla, t)
    h_orig, h_migr = {clave_h(h, True) for h in h_o}, {clave_h(h) for h in h_m}
    reporte = {
        "cliente": cid, "copropiedad": ficha.get("copropiedad"), "periodo": periodo, "plantilla": PL.VERSION,
        "excel_original": excel.name, "archivo_migrado": migrado.name, "recalculado_libreoffice": bool(recalculado),
        "fecha": datetime.now().isoformat(timespec="seconds"), "estado": "propuesta",
        "recetas_usadas": sorted(recetas), "recetas_propuestas_no_aprobadas": propuestas,
        "hojas": plan, "anexos": anexos, "verificacion": verificacion, "errores": errores,
        "hallazgos": {"iguales": len(h_orig & h_migr), "solo_original": sorted(map(list, h_orig - h_migr)),
                      "solo_migrado": sorted(map(list, h_migr - h_orig))},
    }
    reporte["aceptable"] = not errores and all(v["resultado"] != "diferente" for v in verificacion) \
        and not reporte["hallazgos"]["solo_original"] and not reporte["hallazgos"]["solo_migrado"]
    guardar_json(destino / "reporte_migracion.json", reporte)
    (destino / "reporte_migracion.md").write_text(_markdown(reporte), encoding="utf-8")
    return {**reporte, "carpeta": str(destino)}


def _markdown(r: dict) -> str:
    L = [f"# Migración a la plantilla estándar v{r['plantilla']} — {r['copropiedad']} ({r['cliente']})", "",
         f"Periodo de corte: **{r['periodo']}** · Excel original: `{r['excel_original']}` · Migrado: `{r['archivo_migrado']}`",
         f"Estado: **{'ACEPTABLE' if r['aceptable'] else 'CON DIFERENCIAS'}** — propuesta pendiente de aprobación del revisor fiscal.", ""]
    if r["recetas_propuestas_no_aprobadas"]:
        L += [f"⚠ Se usaron recetas **propuestas, aún no aprobadas**: {', '.join(r['recetas_propuestas_no_aprobadas'])}.", ""]
    L += ["## Hojas", "", "| Hoja de la plantilla | Método | Origen |", "|---|---|---|"]
    L += [f"| {h['hoja']} | {h['metodo']} | {', '.join(h['origen']) or '—'} |" for h in r["hojas"]]
    L += ["", "## Anexos (hojas del original que no forman parte de la plantilla)", ""]
    L += [f"- {a['anexo']} ← {a['original']}" for a in r["anexos"]] or ["- Ninguno"]
    L += ["", "## Verificación con el lector estándar", "", "| Sección | Resultado | Detalle |", "|---|---|---|"]
    for v in r["verificacion"]:
        det = "; ".join(v["diferencias"][:3]) or "; ".join((v.get("normalizaciones") or [])[:2] + (v.get("formato") or []))
        L.append(f"| {v['seccion']} | {v['resultado']} | {det.replace('|', '/')} |")
    hz = r["hallazgos"]
    L += ["", "## Hallazgos del validador", "", f"Iguales en ambos: {hz['iguales']}.",
          f"Solo en el original: {hz['solo_original'] or 'ninguno'}.", f"Solo en el migrado: {hz['solo_migrado'] or 'ninguno'}."]
    if r["errores"]:
        L += ["", "## Errores", ""] + [f"- {e}" for e in r["errores"]]
    return "\n".join(L) + "\n"
