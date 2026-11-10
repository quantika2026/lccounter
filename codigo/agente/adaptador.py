"""Agente ADAPTADOR: ajusta la lectura al Excel de cada cliente; el informe en PowerPoint no cambia.

El informe tiene siempre la misma estructura (config/estructura_informe.yaml). Lo que cambia entre clientes es cómo
organizan su Excel. El adaptador revisa TODAS las secciones del informe de una vez y, para cada una que no se pudo
leer con el mapa y las recetas vigentes, propone una receta de lectura (revisoria/lectores.py):

  1. ubica las hojas candidatas (nombre, encabezados y firma de la sección);
  2. intenta una receta AUTOMÁTICA (sin IA): busca la fila de encabezados y asigna cada campo estándar por las
     palabras de su columna; en tablas con una columna por mes usa la columna del periodo («{mes}»);
  3. si no basta y hay proveedor de IA configurado, le pide la receta al subagente estructurador (con ejemplos);
  4. PRUEBA cada receta con el lector genérico (lectores.probar): solo se proponen las que leen sin errores;
  5. simula el informe con todas las propuestas y dice cuántas secciones quedarían completas.

Python lee y calcula las cifras; la IA (si se usa) solo dice DÓNDE están. El adaptador NUNCA aprueba: las
propuestas quedan en salidas/<periodo>/<cliente>/adaptacion.json y las aprueba un revisor en la app
(Mapa del Excel → Adaptación al Excel). Lo que no tiene receta posible queda como «pendiente» en el informe, con
su motivo: la diapositiva se conserva para mantener la estructura estándar.
"""
from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

from revisoria import esquema, lectores
from revisoria.excel import Libro, Periodo, es_numero, norm

from .comun import guardar_json, leer_json

ARCHIVO = "adaptacion.json"
MAX_CANDIDATAS = 3

# palabras del TÍTULO de la hoja (o de su encabezado) que la delatan como candidata para cada sección
PALABRAS_HOJA = {
    "liquidez": ["efectivo", "liquidez", "disponible"],
    "cartera": ["cartera", "deudor", "moros", "propietario"],
    "presupuesto": ["presupuesto", "ppto", "ejecucion presupuestal", "sobre ejecutado", "sobreejecutado"],
    "cxp": ["cxp", "pagar", "proveedor", "acreedor"],
    "anticipos": ["anticipo", "antic", "avance"],
    "deudores_varios": ["deudores varios", "antic", "1380"],
    "retencion": ["impuesto", "retencion", "dian", "tribut"],
    "egresos": ["egreso", "comprobante"],
    "polizas": ["poliza", "contrato"],
    "actas_consejo": ["acta", "consejo"],
    "matriz_riesgo": ["matriz", "riesgo", "dictamen"],
    "matriz_manuales": ["manual", "matriz"],
}

# palabras con que suele empezar (o que contiene) el encabezado de cada campo estándar
CLAVES = {
    "cxp": {"tercero": ["tercero", "proveedor", "acreedor", "beneficiario", "nombre", "razon social", "concepto"],
            "valor": ["saldo", "valor", "monto", "importe"],
            "edad": ["edad", "antiguedad", "dias", "rango"],
            "concepto": ["detalle", "observ", "descrip", "concepto"]},
    "anticipos": {"tercero": ["tercero", "contratista", "proveedor", "nombre", "beneficiario"],
                  "valor": ["valor", "saldo", "monto"], "fecha": ["fecha"],
                  "detalle": ["detalle", "observ", "concepto", "descrip"]},
    "deudores_varios": {"tercero": ["tercero", "deudor", "contratista", "nombre", "proveedor", "beneficiario"],
                        "valor": ["valor", "saldo", "monto"], "fecha": ["fecha"],
                        "detalle": ["detalle", "observ", "concepto", "descrip"]},
    "retencion": {"periodo": ["periodo", "mes"], "valor": ["valor", "total", "monto"],
                  "presentado": ["presenta", "fecha de presenta", "fecha presenta"],
                  "pagado": ["pagado", "pago", "fecha de pago", "fecha pago"],
                  "observacion": ["observ", "detalle"]},
    "egresos": {"ce": ["ce", "c.e", "comprobante", "no.", "numero", "n°", "#"], "tercero": ["tercero", "beneficiario", "proveedor", "nombre"],
                "valor": ["valor", "monto", "total"], "fecha": ["fecha"],
                "descripcion": ["descrip", "concepto", "detalle"], "observacion": ["observ"]},
    "polizas": {"contrato": ["contrato", "proveedor", "contratista"], "verificacion": ["verific", "cumpl", "observ"],
                "n": ["n.", "no", "n°", "#", "item"]},
    "actas_consejo": {"acta": ["acta", "n. acta", "no acta", "numero"], "temas": ["tema", "asunto", "orden del dia"],
                      "fecha": ["fecha"], "firma_presidente": ["firma presidente", "presidente"],
                      "firma_administracion": ["firma secret", "firma admin", "secretari", "administra"]},
    "matriz_riesgo": {"objetivo": ["objetivo"], "descripcion": ["descrip", "opinion", "riesgo identificado", "hallazgo"],
                      "valoracion": ["valoriz", "valorac", "nivel de riesgo", "nivel riesgo"]},
    "matriz_manuales": {"objetivo": ["objetivo"], "detalle": ["detalle", "formaliz", "document", "descrip"],
                        "valoracion": ["valoriz", "valorac", "nivel de riesgo", "nivel riesgo"]},
    "presupuesto": {"rubro": ["rubro", "descrip", "concepto", "cuenta"],
                    "presupuesto_periodo": ["presupuesto ene", "presupuesto a", "presupuesto acum", "ppto acum",
                                            "presupuesto del periodo", "presupuesto periodo", "presupuesto"],
                    "ejecucion": ["ejecu", "real"],
                    "presupuesto_anual": ["presupuesto anual", "ppto anual", "anual", "diciembre"],
                    "diferencia": ["diferen", "variac", "valor sobre", "sobre ejecu", "sobreejecu"],
                    "observacion": ["observ"]},
}
# encabezados que NO sirven para un campo aunque traigan su palabra («Presupuesto enero-diciembre» no es el del periodo)
EXCLUIR = {("presupuesto", "presupuesto_periodo"): ["anual", "diciembre"]}
# títulos de bloque por mes: palabras que el título del bloque debe traer para ser de la sección
TITULO_BLOQUE = {"cxp": ["pagar", "proveedor", "cxp"], "anticipos": ["anticipo"], "deudores_varios": ["deudor"],
                 "egresos": ["egreso"]}


# ------------------------------------------------------------------ utilidades

def _tiene(texto: str, claves) -> bool:
    return any(texto == k or texto.startswith(k) or re.search(r"(^|[\s(/.,-])" + re.escape(k), texto) for k in claves)


def _encabezados(ws, r: int, max_col: int = 30) -> dict[int, str]:
    return {c: norm(ws.cell(r, c).value) for c in range(1, min(ws.max_column, max_col) + 1)
            if isinstance(ws.cell(r, c).value, str) and ws.cell(r, c).value.strip()}


def _con_datos(ws, r: int, c: int, n: int = 12) -> bool:
    return any(ws.cell(f, c).value not in (None, "") for f in range(r + 1, min(ws.max_row, r + n) + 1))


def _asignar(ws, r: int, seccion: str) -> dict:
    """{campo: letra} para la fila de encabezado r. Los obligatorios primero; cada columna se usa una vez."""
    from openpyxl.utils import get_column_letter
    enc = _encabezados(ws, r)
    campos = lectores.CAMPOS[seccion]["campos"]
    orden = sorted(campos, key=lambda c: (not campos[c][1],))
    usadas, out = set(), {}
    for campo in orden:
        tipo = campos[campo][0]
        for clave in CLAVES[seccion].get(campo, []):           # la clave más específica primero
            fuera = EXCLUIR.get((seccion, campo), [])
            col = next((c for c, t in enc.items() if c not in usadas and _tiene(t, [clave]) and not any(x in t for x in fuera)
                        and _con_datos(ws, r, c)), None)
            if col and tipo == "n":                             # un campo numérico debe tener números debajo
                nums = [ws.cell(f, col).value for f in range(r + 1, min(ws.max_row, r + 15) + 1)]
                if sum(es_numero(v) for v in nums) < 1:
                    col = None
            if col:
                out[campo] = get_column_letter(col)
                usadas.add(col)
                break
    return out


def _obligatorios(seccion: str, cols: dict | None = None) -> list[str]:
    ob = [c for c, (_, o, _) in lectores.CAMPOS[seccion]["campos"].items() if o]
    if seccion == "actas_consejo" and cols and "fecha" in cols:      # el acta se identifica por su fecha
        ob = [c for c in ob if c != "acta"]
    return ob


def _columna_mes(ws, r: int, p: Periodo) -> int | None:
    return lectores._col_mes(ws, r, p)


def _meses_en(ws, r: int) -> int:
    return sum(1 for c in range(1, min(ws.max_column, 40) + 1) if lectores.periodo_celda(ws.cell(r, c).value))


def _plantilla_titulo(texto: str, p: Periodo) -> str | None:
    """«CUENTAS POR PAGAR A AGOSTO 2026» → «cuentas por pagar a {mes} {anio}» (prefijo hasta el mes/año)."""
    t = norm(texto)
    for patron, repl in ((rf"{p.nombre_mes}\s+de\s+{p.anio}", "{mes} de {anio}"), (rf"{p.nombre_mes}\s+{p.anio}", "{mes} {anio}"),
                         (rf"{p.nombre_mes}\b", "{mes}")):
        m = re.search(patron, t)
        if m:
            otro_anio = re.match(r"\s*(de\s+)?(\d{4})", t[m.end():])
            if repl == "{mes}" and otro_anio and int(otro_anio.group(2)) != p.anio:
                return None                                  # «agosto 2025» no es el bloque de agosto 2026
            return t[:m.start()] + repl
    return None


# ------------------------------------------------------------------ receta automática (sin IA)

def _tabla(ws, seccion: str, p: Periodo) -> dict | None:
    """Receta de una tabla: con bloque por mes (título con el mes) o con una fila de encabezados fija."""
    base = {"seccion": seccion, "hoja": ws.title}
    oblig = _obligatorios(seccion)
    # 1) bloque por mes: un título con el mes del periodo y los encabezados en la fila siguiente
    if lectores.CAMPOS[seccion].get("por_mes") and seccion in TITULO_BLOQUE:
        bloques = []
        for fila in ws.iter_rows():
            for c in fila:
                if not (isinstance(c.value, str) and len(c.value) < 90):
                    continue
                t = norm(c.value)
                if not _tiene(t, TITULO_BLOQUE[seccion]) or t.startswith("total"):
                    continue
                plantilla = _plantilla_titulo(c.value, p)
                if plantilla:
                    cols = _asignar(ws, c.row + 1, seccion)
                    if all(o in cols for o in oblig):
                        bloques.append({**base, "titulo": plantilla, "columnas": cols})
        if bloques:                                         # el último bloque del mes es el vigente (como el lector)
            return bloques[-1]
    # 2) tabla única: la fila (entre las primeras 40) que más campos reconoce
    mejor, puntos = None, 0
    for r in range(1, min(ws.max_row, 40) + 1):
        cols = _asignar(ws, r, seccion)
        if not all(o in cols for o in _obligatorios(seccion, cols)):
            # tabla con una columna por mes (p. ej. CXP: tercero | may-26 | abr-26 …): el valor es la del periodo
            if "valor" not in oblig or _meses_en(ws, r) < 2:
                continue
            cols = {**_primer_texto(ws, r, seccion, {k: v for k, v in cols.items() if k != "valor"}), "valor": lectores.MES}
            if not all(o in cols for o in _obligatorios(seccion, cols)):
                continue
        n = len(cols)
        if n > puntos:
            mejor, puntos = {**base, "fila_encabezado": r, "columnas": cols}, n
    if mejor and seccion == "actas_consejo" and "acta" not in mejor["columnas"] and "fecha" in mejor["columnas"]:
        mejor["columnas"]["acta"] = mejor["columnas"].pop("fecha")   # «número y fecha en la misma celda»
    if mejor and CAMPOS_MATRIZ(seccion):
        mejor["niveles"] = _niveles(ws, mejor["fila_encabezado"], mejor["columnas"])
    return mejor


def CAMPOS_MATRIZ(seccion: str) -> bool:
    return bool(lectores.CAMPOS[seccion].get("matriz"))


def _primer_texto(ws, r: int, seccion: str, cols: dict) -> dict:
    """En una tabla por meses el campo de texto obligatorio (tercero) suele ser la primera columna con texto."""
    from openpyxl.utils import get_column_letter
    out = dict(cols)
    faltan = [o for o in _obligatorios(seccion) if o not in out and o != "valor"]
    if not faltan:
        return out
    usadas = {lectores._col(v) for v in out.values() if not isinstance(v, list) and not lectores.es_mes(v)}
    for c in range(1, min(ws.max_column, 10) + 1):
        if c in usadas or lectores.periodo_celda(ws.cell(r, c).value):
            continue
        textos = [ws.cell(f, c).value for f in range(r + 1, min(ws.max_row, r + 12) + 1)]
        if sum(isinstance(v, str) and v.strip() and not norm(v).startswith("total") for v in textos) >= 3:
            out[faltan[0]] = get_column_letter(c)
            break
    return out


def _niveles(ws, r: int, cols: dict) -> list[str]:
    """Columnas cuyos datos son niveles de cumplimiento (Bajo/Medio/Alto/ODM) y tienen una fecha encima."""
    from openpyxl.utils import get_column_letter
    usadas = {lectores._col(v) for v in cols.values() if not isinstance(v, list)}
    out = []
    for c in range(1, min(ws.max_column, 30) + 1):
        if c in usadas:
            continue
        vals = [norm(ws.cell(f, c).value) for f in range(r + 1, min(ws.max_row, r + 40) + 1) if ws.cell(f, c).value]
        if vals and sum(v in esquema.NIVELES_RIESGO for v in vals) >= 0.6 * len(vals) and lectores.periodo_celda(ws.cell(r - 1, c).value):
            out.append(get_column_letter(c))
    return out


def _liquidez(ws, p: Periodo) -> dict | None:
    """Fila con varios meses de encabezado y, debajo, una fila «total…» con cifra en el mes."""
    from openpyxl.utils import get_column_letter
    for r in range(1, min(ws.max_row, 60) + 1):
        if _meses_en(ws, r) < 2:
            continue
        col_mes = _columna_mes(ws, r, p)
        if not col_mes:
            continue
        c_lbl = next((c for c in range(1, col_mes) if isinstance(ws.cell(r, c).value, str) and ws.cell(r, c).value.strip()), None)
        if c_lbl is None:
            continue
        # la fila total que CIERRA el bloque: la última «total…» con cifra antes del siguiente encabezado de meses
        fin = next((f for f in range(r + 1, min(ws.max_row, r + 60) + 1) if _meses_en(ws, f) >= 2), min(ws.max_row, r + 60) + 1)
        totales = [f for f in range(r + 1, fin) if norm(ws.cell(f, c_lbl).value).startswith("total") and es_numero(ws.cell(f, col_mes).value)]
        f_tot = totales[-1] if totales else None
        if not f_tot:
            continue
        rec = {"seccion": "liquidez", "hoja": ws.title, "columnas": {"concepto": get_column_letter(c_lbl)},
               "encabezado": norm(ws.cell(r, c_lbl).value)[:40], "total": norm(ws.cell(f_tot, c_lbl).value)[:40]}
        f_mon = next((f for f in range(f_tot + 1, min(ws.max_row, f_tot + 30) + 1)
                      if isinstance(ws.cell(f, c_lbl).value, str) and _meses_en(ws, f) >= 2), None)
        if f_mon:
            mv = next((f for f in range(f_mon + 1, f_mon + 12)
                       if norm(ws.cell(f, c_lbl).value).startswith("mayor valor") and es_numero(ws.cell(f, col_mes).value)), None)
            if mv:
                rec["monetizacion"] = norm(ws.cell(f_mon, c_lbl).value)[:40]
                rec["mayor_valor"] = "mayor valor"
        return rec
    return None


def _cartera(ws, p: Periodo) -> dict | None:
    from openpyxl.utils import get_column_letter
    # a) listas por mes con título «cartera a <mes> <año>» y columnas apto/valor al lado del título
    for fila in ws.iter_rows(max_row=6):
        for c in fila:
            if isinstance(c.value, str) and "cartera" in norm(c.value) and not norm(c.value).startswith("total"):
                plantilla = _plantilla_titulo(c.value, p)
                if plantilla:
                    return {"seccion": "cartera", "hoja": ws.title, "titulo": plantilla, "columnas": {"apto": "+1", "valor": "+2"}}
    # b) tabla «inmueble | <mes> | <mes> …»: si hay varias, la que trae más meses (sirve para el comparativo)
    opciones = []
    for r in range(1, min(ws.max_row, 15) + 1):
        for c in range(1, min(ws.max_column, 30) + 1):
            t = norm(ws.cell(r, c).value)
            if t in ("apto", "apartamento", "inmueble", "unidad", "casa", "torre y apto", "apto.") or t.startswith(("inmueble", "apto")):
                if lectores._col_mes(ws, r, p, desde=c + 1, hasta=c + 8):
                    meses = sum(1 for k in range(c + 1, c + 9) if lectores.periodo_celda(ws.cell(r, k).value))
                    opciones.append((meses, -r, -c, r, c))
    if opciones:
        _, _, _, r, c = max(opciones)
        return {"seccion": "cartera", "hoja": ws.title, "fila_encabezado": r, "columnas": {"apto": get_column_letter(c)}}
    return None


def _presupuesto(ws, p: Periodo) -> dict | None:
    rec = _tabla(ws, "presupuesto", p)
    if not rec:
        return None
    c_r = lectores._col(rec["columnas"]["rubro"])
    f_tot = next((f for f in range(rec["fila_encabezado"] + 1, ws.max_row + 1) if norm(ws.cell(f, c_r).value).startswith("total")), None)
    if f_tot:
        rec["total"] = norm(ws.cell(f_tot, c_r).value)[:40]
    else:
        rec["columnas"].pop("diferencia", None)
    return rec


def receta_automatica(lib: Libro, seccion: str, hoja: str, p: Periodo) -> dict | None:
    """Receta propuesta sin IA para `seccion` en `hoja` (o None si la hoja no se deja leer así)."""
    ws = lib.valores[hoja]
    if seccion == "liquidez":
        rec = _liquidez(ws, p)
    elif seccion == "cartera":
        rec = _cartera(ws, p)
    elif seccion == "presupuesto":
        rec = _presupuesto(ws, p)
    elif seccion in CLAVES:
        rec = _tabla(ws, seccion, p)
    else:
        rec = None
    if rec and seccion == "presupuesto" and rec.get("total"):        # el signo de la diferencia: el que cuadra
        mejor = None
        for signo in (-1, 1):
            pr = lectores.probar(lib, {**rec, "signo_total": signo}, p)
            if pr["ok"] and not any("no coincide con el total" in a for a in pr["avisos"]):
                mejor = signo
                break
        rec["signo_total"] = mejor or -1
    return rec


# ------------------------------------------------------------------ hojas candidatas

def candidatas(lib: Libro, seccion: str, mapa: dict, ocupadas: set[str]) -> list[str]:
    """Hojas donde podría estar la sección, de la más a la menos probable. La del mapa va primero."""
    clave = next(s.hoja for s in esquema.SECCIONES if s.clave == seccion)
    actual = (mapa.get("hojas") or {}).get(clave)
    puntaje = {}
    for t in lib.valores.sheetnames:
        if t in ocupadas and t != actual:
            continue
        nt = norm(t)
        cab = esquema._texto_cabecera(lib.valores[t], filas=12, cols=20)
        s = 3 * sum(1 for w in PALABRAS_HOJA.get(seccion, []) if w in nt)
        s += 2 * sum(1 for f in esquema.FIRMAS.get(clave, ()) if esquema.cumple(f, cab))
        if s == 0:                       # ni el nombre ni la firma la delatan: una tabla cualquiera no es candidata
            continue
        s += sum(1 for ks in CLAVES.get(seccion, {}).values() if any(k in cab for k in ks if len(k) > 3))
        puntaje[t] = s
    orden = sorted(puntaje, key=lambda t: -puntaje[t])
    if isinstance(actual, str) and actual in lib.valores.sheetnames:
        orden = [actual] + [t for t in orden if t != actual]
    return orden[:MAX_CANDIDATAS]


# ------------------------------------------------------------------ correcciones del mapa

def ajustes_mapa(lib: Libro, mapa: dict, aprendidos: dict | None = None) -> dict:
    """Hojas del mapa que la detección actual ubica mejor: {clave: {"antes", "ahora", "motivo"}}.

    Pasa cuando el mapa se hizo con otra versión del Excel o del detector: dice «no tiene la hoja» aunque el Excel sí
    la trae, o asignó a una sección la hoja de otra. Solo se propone una hoja que cumple la firma de su sección; el
    mapa no se toca (lo aprueba el revisor). Las claves con varias hojas (auditoría documental) no se ajustan aquí."""
    from revisoria import mapa as mapa_cliente
    hojas = mapa.get("hojas") or {}
    nuevo = mapa_cliente.detectar(lib, aprendidos)["hojas"]
    dueno = {t: k for k, v in nuevo.items() for t in mapa_cliente._titulos(v)}   # a qué clave va cada hoja hoy
    out = {}
    for clave, ahora in nuevo.items():
        antes = hojas.get(clave)
        if clave not in hojas or ahora == antes or isinstance(ahora, list) or isinstance(antes, list) or not ahora:
            continue
        if esquema.revisar_formato(lib, {clave: ahora}):
            continue
        if antes is None:
            motivo = f"el mapa dice que no tiene esta hoja, pero el Excel trae «{ahora}» con el formato de la sección"
        elif antes not in lib.valores.sheetnames:
            motivo = f"la hoja «{antes}» ya no está en el Excel; «{ahora}» tiene el formato de la sección"
        elif dueno.get(antes, clave) != clave:
            motivo = f"«{antes}» es la hoja de «{dueno[antes]}»; la de esta sección es «{ahora}»"
        elif esquema.revisar_formato(lib, {clave: antes}):
            motivo = f"«{antes}» no tiene el formato de la sección; «{ahora}» sí"
        else:
            continue                                    # las dos sirven: se respeta lo que ya está en el mapa
        out[clave] = {"antes": antes, "ahora": ahora, "motivo": motivo}
    return out


# ------------------------------------------------------------------ el agente

def _estado(lib: Libro, mapa: dict, ficha: dict, p: Periodo) -> tuple[list[dict], dict]:
    from revisoria.documental import extraer_documental
    from revisoria.extractor import extraer
    lib.mapa = mapa
    meta = {k: ficha.get(k) for k in ("copropiedad", "nit", "unidades") if ficha.get(k)}
    modelo, _ = extraer(lib.ruta, p, meta, mapa, lib)
    completo = str(ficha.get("informe", "completo")).lower() == "completo"
    doc = extraer_documental(lib, p) if completo else {}
    return esquema.verificar(modelo, doc), {**(modelo.get("faltantes") or {}), **((doc or {}).get("faltantes") or {})}


def _propuesta_previa(sal: Path, seccion: str) -> dict | None:
    """Receta que ya dejó Claude Code (`probar-receta`) o el estructurador de la app para esta sección."""
    r = leer_json(sal / f"receta_propuesta_{seccion}.json", None)
    return r["receta"] if r and r.get("corresponde", True) and r.get("receta") else None


def adaptar(esp, cid: str, periodo: str, usar_ia: bool = False, llamar=None, log=None) -> dict:
    """Propone (sin aprobar) las recetas que faltan para leer el Excel del cliente. Guarda adaptacion.json."""
    from . import estructurador, normalizador
    from revisoria import mapa as mapa_cliente
    log = log or (lambda *_: None)
    p = Periodo.desde_texto(periodo)
    ficha = esp.ficha(cid)
    excel = normalizador.ubicar_excel(esp.entrada(periodo, cid), ficha)
    if not excel:
        raise ValueError(f"No hay Excel cargado para {cid} en {periodo}.")
    sal = esp.salida(periodo, cid)
    sal.mkdir(parents=True, exist_ok=True)
    mapa = mapa_cliente.cargar(esp.clientes / cid)
    lib = Libro(str(excel))
    if mapa is None:
        mapa = mapa_cliente.detectar(lib, mapa_cliente.aprendidos(esp.conocimiento))
    recetas = dict(mapa.get("lectores") or {})
    hoy, faltantes_hoy = _estado(lib, mapa, ficha, p)      # lo que sale hoy, con el mapa vigente
    estado_hoy = {c["seccion"]: c["estado"] for c in hoy}
    # 1) correcciones del mapa (hojas que el Excel sí trae); las recetas se buscan ya con el mapa corregido
    ajustes = ajustes_mapa(lib, mapa, mapa_cliente.aprendidos(esp.conocimiento))
    if ajustes:
        mapa = json.loads(json.dumps(mapa))
        for clave, a in ajustes.items():
            mapa["hojas"][clave] = a["ahora"]
    cobertura, faltantes = _estado(Libro(str(excel)), mapa, ficha, p) if ajustes else (hoy, faltantes_hoy)
    # secciones que se arreglan con el mapa corregido aunque su hoja sea otra (p. ej. «mayores deudores» se lee de la
    # hoja MAYORES DEUDORES): se prueba cada corrección sola para saber cuál las arregla
    causa = {c["seccion"]: c["hoja"] for c in cobertura if c["hoja"] in ajustes}
    mejoran = [c["seccion"] for c in cobertura if c["estado"] == "ok" and estado_hoy.get(c["seccion"]) != "ok"
               and c["seccion"] not in causa]
    if mejoran:
        original = mapa_cliente.cargar(esp.clientes / cid) or mapa
        for clave, a in ajustes.items():
            prueba = json.loads(json.dumps(original))
            prueba.setdefault("hojas", {})[clave] = a["ahora"]
            sola, _ = _estado(Libro(str(excel)), prueba, ficha, p)
            for c in sola:
                if c["seccion"] in mejoran and c["estado"] == "ok":
                    causa.setdefault(c["seccion"], clave)
    por_seccion = {c["seccion"]: c for c in cobertura}
    ocupadas = {t for c in cobertura if c["estado"] == "ok"
                for t in mapa_cliente._titulos((mapa.get("hojas") or {}).get(c["hoja"]))}
    ia = None
    if usar_ia:
        if llamar is not None:
            ia = llamar
        else:
            from app import proveedores
            if proveedores.disponible(esp.raiz, "estructurador")["ok"]:
                ia = True
    secciones, propuestas = [], {}
    for c in cobertura:
        sec = c["seccion"]
        fila = {"seccion": sec, "descripcion": c["descripcion"], "obligatoria": c["obligatoria"],
                "estado": estado_hoy.get(sec, c["estado"]),
                "detalle": faltantes.get(sec) or c["detalle"], "admite_receta": sec in lectores.CAMPOS,
                "receta_aprobada": sec in recetas, "propuesta": None, "origen": None, "prueba": None,
                "candidatas": [], "motivo": "", "ajuste_mapa": None}
        if sec in causa:
            fila["ajuste_mapa"] = {"clave": causa[sec], **ajustes[causa[sec]]}
        secciones.append(fila)
        if c["estado"] == "ok" and sec not in faltantes:
            if fila["ajuste_mapa"]:
                fila["motivo"] = "Se lee bien con la hoja corregida en el mapa."
            continue
        if sec not in lectores.CAMPOS:
            fila["motivo"] = ("Esta sección todavía no admite receta: se lee con el lector estándar"
                              + (f", que no la pudo leer ({fila['detalle'].rstrip('.')})" if fila["detalle"] else "")
                              + ". Mientras tanto sale como pendiente en el informe.")
            continue
        cands = candidatas(lib, sec, mapa, ocupadas)
        fila["candidatas"] = cands
        intentos = []
        previa = _propuesta_previa(sal, sec)
        if previa:
            intentos.append(("Claude", previa))
        for h in cands:
            try:
                rec = receta_automatica(lib, sec, h, p)
            except Exception as e:                              # una hoja rara no tumba al agente
                log(f"{sec} · {h}: {type(e).__name__} {e}")
                rec = None
            if rec:
                intentos.append(("automática", rec))
        elegido = None
        ultimo = None
        for origen, rec in intentos:
            try:
                pr = lectores.probar(lib, rec, p)
            except Exception as e:
                pr = {"ok": False, "errores": [f"{type(e).__name__}: {e}"], "avisos": [], "n": 0, "muestra": []}
            ultimo = (origen, rec, pr)
            if pr["ok"]:
                elegido = (origen, rec, pr)
                break
        if not elegido and ia and cands:
            try:
                r = estructurador.proponer(esp.raiz, lib, sec, cands[0], p, esp.conocimiento,
                                           llamar=None if ia is True else ia, log=log)
                if r.get("corresponde") and r.get("receta"):
                    ultimo = ("IA", r["receta"], r["prueba"])
                    if r["prueba"]["ok"]:
                        elegido = ultimo
            except Exception as e:
                fila["motivo"] = f"El estructurador con IA no respondió: {e}"
        if elegido:
            fila.update({"origen": elegido[0], "propuesta": elegido[1], "prueba": _resumen_prueba(elegido[2])})
            propuestas[sec] = elegido[1]
            ocupadas.add(elegido[1]["hoja"])
        else:
            if ultimo:
                fila["prueba"] = _resumen_prueba(ultimo[2])
            if fila["motivo"]:
                pass
            elif not cands:
                fila["motivo"] = "No se encontró en el Excel una hoja que traiga esta sección."
            elif ultimo and ultimo[2].get("errores"):
                err = ultimo[2]["errores"][0]
                fila["motivo"] = err if ultimo[1]["hoja"] in err else f"En la hoja «{ultimo[1]['hoja']}»: {err}"
            else:
                fila["motivo"] = ("Las hojas candidatas (" + ", ".join(f"«{h}»" for h in cands) + ") no se dejaron leer "
                                  "automáticamente" + ("" if ia else "; puede pedir la receta al estructurador con IA o "
                                                       "escribirla en «Otra estructura…»") + ".")
    # simulación: el informe con todas las propuestas aplicadas
    despues = {c["seccion"]: c["estado"] for c in cobertura} if ajustes else None
    if propuestas:
        sim = mapa_con_propuestas(mapa, {"secciones": [{"seccion": k, "propuesta": v} for k, v in propuestas.items()]})
        try:
            cob2, _ = _estado(Libro(str(excel)), sim, ficha, p)
            despues = {c["seccion"]: c["estado"] for c in cob2}
        except Exception as e:
            log(f"simulación: {type(e).__name__} {e}")
    for f in secciones:
        f["estado_con_propuestas"] = (despues or {}).get(f["seccion"], f["estado"])
    total = len(secciones)
    res = {"cliente": cid, "periodo": periodo, "excel": excel.name, "fecha": datetime.now().isoformat(timespec="seconds"),
           "ia": bool(ia), "secciones": secciones,
           "ajustes_mapa": ajustes,
           "resumen": {"secciones": total, "completas": sum(1 for f in secciones if f["estado"] == "ok"),
                       "completas_con_propuestas": sum(1 for f in secciones if f["estado_con_propuestas"] == "ok"),
                       "propuestas": len(propuestas), "ajustes_mapa": len(ajustes),
                       "sin_solucion": sum(1 for f in secciones if f["estado_con_propuestas"] != "ok")}}
    guardar_json(sal / ARCHIVO, res)
    return res


def mapa_con_propuestas(mapa: dict, ad: dict) -> dict:
    """Copia del mapa con las correcciones y recetas propuestas que aún no se aprueban (para el BORRADOR del informe).
    Solo en memoria: nunca se guarda; el mapa del cliente cambia únicamente con la aprobación del revisor."""
    sim = json.loads(json.dumps(mapa))
    for clave, a in (ad.get("ajustes_mapa") or {}).items():
        sim.setdefault("hojas", {})[clave] = a["ahora"]
    for f in ad.get("secciones", []):
        rec = f.get("propuesta")
        if rec and not f.get("receta_aprobada"):
            sim.setdefault("lectores", {})[f["seccion"]] = rec
            clave = next(s.hoja for s in esquema.SECCIONES if s.clave == f["seccion"])
            if not isinstance((sim.get("hojas") or {}).get(clave), list):
                sim.setdefault("hojas", {})[clave] = rec["hoja"]
    return sim


def secciones_con_propuesta(ad: dict) -> list[str]:
    """Secciones cuyos datos, en el borrador, vienen de una propuesta sin aprobar (receta o corrección del mapa)."""
    return [f["seccion"] for f in ad.get("secciones", [])
            if f.get("estado_con_propuestas") == "ok"
            and ((f.get("propuesta") and not f.get("receta_aprobada")) or (f.get("ajuste_mapa") and not f.get("mapa_aprobado")))]


def _resumen_prueba(pr: dict) -> dict:
    return json.loads(json.dumps({k: pr.get(k) for k in ("ok", "errores", "avisos", "n", "muestra", "encabezados")},
                                 default=str))


def aprobar(esp, cid: str, periodo: str, secciones: list[str], revisor: str) -> list[str]:
    """Lo ejecuta un REVISOR (la app lo exige): aprueba las correcciones del mapa y las recetas propuestas de las
    secciones elegidas (primero el mapa, porque las recetas se probaron con el mapa corregido)."""
    from app import mapeo
    from revisoria import mapa as mapa_cliente
    from . import estructurador, normalizador
    p = Periodo.desde_texto(periodo)
    excel = normalizador.ubicar_excel(esp.entrada(periodo, cid), esp.ficha(cid))
    datos = leer_json(esp.salida(periodo, cid) / ARCHIVO, None)
    if not datos:
        raise ValueError("No hay una adaptación propuesta para este cliente y periodo.")
    aprobadas = []
    cambios = {f["ajuste_mapa"]["clave"]: f["ajuste_mapa"]["ahora"] for f in datos["secciones"]
               if f["seccion"] in secciones and f.get("ajuste_mapa") and not f.get("mapa_aprobado")}
    if cambios:
        vigente = mapa_cliente.cargar(esp.clientes / cid) or mapa_cliente.detectar(Libro(str(excel)))
        hojas = {k: (mapa_cliente._titulos(v) if isinstance(v, list) else v) for k, v in vigente.get("hojas", {}).items()}
        mapeo.guardar(esp, cid, {**hojas, **cambios}, revisor, excel)
        for f in datos["secciones"]:
            if f.get("ajuste_mapa") and f["ajuste_mapa"]["clave"] in cambios:
                f["mapa_aprobado"], f["aprobada_por"] = True, revisor
                if f["seccion"] not in aprobadas:
                    aprobadas.append(f["seccion"])
    for f in datos["secciones"]:
        if f["seccion"] in secciones and f.get("propuesta"):
            estructurador.aprobar(esp, cid, f["propuesta"], p, excel, revisor)
            if f["seccion"] not in aprobadas:
                aprobadas.append(f["seccion"])
            f["receta_aprobada"], f["aprobada_por"] = True, revisor
    guardar_json(esp.salida(periodo, cid) / ARCHIVO, datos)
    return aprobadas
