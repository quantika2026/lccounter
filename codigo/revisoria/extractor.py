"""Extrae del Excel de papeles de trabajo el modelo de datos de la sección financiera.

Regla de oro: TODAS las cifras del informe salen de aquí, calculadas en Python.
El LLM (si se usa) solo redacta sobre cifras ya verificadas.
Cada cifra guarda su 'fuente' (hoja!celda) para trazabilidad.
"""
from __future__ import annotations

import re
from datetime import date, datetime

from openpyxl.utils import get_column_letter

from .excel import (Libro, Periodo, buscar_celdas, buscar_fila, columna_mes, coord, es_numero,
                    leer_tabla, meses_en_fila, norm, num, periodo_de)


def _f(ws, fila, col):
    return f"{ws.title.strip()}!{coord(fila, col)}" if fila and col else None


def _bloque_conceptos(ws, fila_ini: int, fila_fin: int, col_act: int, col_ref: int | None, col_lbl: int = 2):
    filas = []
    for r in range(fila_ini, fila_fin + 1):
        etiqueta = ws.cell(r, col_lbl).value
        if not etiqueta:
            continue
        filas.append({
            "concepto": str(etiqueta).strip(),
            "actual": num(ws.cell(r, col_act).value),
            "referencia": num(ws.cell(r, col_ref).value) if col_ref else None,
            "fuente": _f(ws, r, col_act),
            "es_total": norm(etiqueta).startswith("total"),
        })
    return filas


# ------------------------------------------------------------------ liquidez

def extraer_liquidez(lib: Libro, p: Periodo):
    ws = lib.hoja("CUADRO EFECTIVO")
    f_enc = buscar_fila(ws, "Análisis de liquidez")
    f_tot = buscar_fila(ws, "Total equivalente al efectivo", desde=f_enc)
    col = columna_mes(ws, f_enc, p, f_tot)
    col_cierre = columna_mes(ws, f_enc, p.cierre_anterior(), f_tot)
    filas = _bloque_conceptos(ws, f_enc + 1, f_tot, col, col_cierre)
    serie = []
    for per, c in meses_en_fila(ws, f_enc):
        v = ws.cell(f_tot, c).value
        if es_numero(v) and per.clave >= p.cierre_anterior().clave and per.clave <= p.clave:
            serie.append({"periodo": per.clave, "valor": float(v)})
    serie.sort(key=lambda x: x["periodo"])

    # Monetización del fondo de imprevistos (bloque opcional: no todas las copropiedades lo llevan)
    f_mon = buscar_fila(ws, "Análisis Monetización")
    f_mv = None
    for r in range(f_mon + 1, f_mon + 12) if f_mon else ():   # la fila con la fórmula/valor
        if norm(ws.cell(r, 2).value).startswith("mayor valor") and es_numero(ws.cell(r, col).value):
            f_mv = r
    c_mon = (columna_mes(ws, f_mon, p, f_mv) or col) if f_mv else None
    c_mon_cierre = (columna_mes(ws, f_mon, p.cierre_anterior(), f_mv) or col_cierre) if f_mv else None
    monet_ok = []
    for r in _bloque_conceptos(ws, f_mon + 1, f_mv, c_mon, c_mon_cierre) if f_mv else ():
        if norm(r["concepto"]).startswith("mayor valor"):
            if r["fuente"] and r["fuente"].endswith(str(f_mv)):
                r["es_total"] = True
                monet_ok.append(r)
            continue
        monet_ok.append(r)
    return {
        "filas": filas,
        "total": num(ws.cell(f_tot, col).value),
        "total_cierre": num(ws.cell(f_tot, col_cierre).value) if col_cierre else None,
        "fuente_total": _f(ws, f_tot, col),
        "serie": serie,
        "monetizacion": monet_ok,
        "mayor_valor_monetizado": num(ws.cell(f_mv, c_mon).value) if f_mv else None,
        "mayor_valor_monetizado_cierre": num(ws.cell(f_mv, c_mon_cierre).value) if f_mv and c_mon_cierre else None,
        "_hoja": ws.title, "_col": col,
    }


# ------------------------------------------------------------------ deudores

def _celda_que_empieza(ws, texto: str):
    """Primera celda (por filas) cuyo texto empieza con `texto`: la tabla puede no estar en la columna B."""
    objetivo = norm(texto)
    for fila in ws.iter_rows():
        for c in fila:
            if isinstance(c.value, str) and norm(c.value).startswith(objetivo):
                return c
    return None


def extraer_deudores(lib: Libro, p: Periodo):
    ws = lib.hoja("GRAFICA DEUDORES")
    c_b = _celda_que_empieza(ws, "(+) Deudores copropietarios")
    if c_b is None:
        raise KeyError(f"La hoja «{ws.title}» no tiene la fila «(+) Deudores copropietarios»")
    f_bruto, c_lbl = c_b.row, c_b.column
    letra = get_column_letter(c_lbl)
    # encabezado de meses: «Detalle» (modelo) o la fila más cercana arriba que trae meses a la derecha
    f_enc = buscar_fila(ws, "Detalle", columnas=(letra,), hasta=f_bruto) or next(
        (r for r in range(f_bruto - 1, 0, -1) if len(meses_en_fila(ws, r, desde_col=c_lbl + 1)) >= 2), None)
    f_neto = buscar_fila(ws, "Deudores copropietarios", columnas=(letra,), desde=f_bruto + 1, exacta=True) or \
        buscar_fila(ws, "Total deudores", columnas=(letra,), desde=f_bruto + 1)
    if not f_enc or not f_neto:
        raise KeyError(f"La hoja «{ws.title}» no trae el encabezado de meses o el total de deudores")
    col = columna_mes(ws, f_enc, p, f_neto, desde_col=c_lbl + 1)
    if not col:
        raise KeyError(f"La hoja «{ws.title}» no trae la columna de {p.etiqueta}")
    col_cierre = columna_mes(ws, f_enc, p.cierre_anterior(), f_neto, desde_col=c_lbl + 1)
    filas = _bloque_conceptos(ws, f_bruto, f_neto, col, col_cierre, c_lbl)
    filas[-1]["es_total"] = True
    serie = []
    for per, c in meses_en_fila(ws, f_enc, desde_col=c_lbl + 1, hasta_col=c_lbl + 30):
        if p.cierre_anterior().clave <= per.clave <= p.clave:
            c_ok = columna_mes(ws, f_enc, per, f_neto, desde_col=c_lbl + 1, hasta_col=c_lbl + 30)
            v = ws.cell(f_neto, c_ok).value
            if es_numero(v) and v:
                serie.append({"periodo": per.clave, "valor": float(v),
                              "bruto": num(ws.cell(f_bruto, c_ok).value)})
    serie.sort(key=lambda x: x["periodo"])
    # saldo bruto = todas las partidas «(+)» (p. ej. ordinarias + extraordinarias), no solo el primer renglón
    positivas = [r for r in range(f_bruto, f_neto) if str(ws.cell(r, c_lbl).value or "").strip().startswith("(+)")]
    bruto = sum(num(ws.cell(r, col).value) for r in positivas) if positivas else num(ws.cell(f_bruto, col).value)
    return {"filas": filas, "neto": num(ws.cell(f_neto, col).value),
            "bruto": bruto, "fuente_neto": _f(ws, f_neto, col),
            "serie": serie}


def extraer_cartera(lib: Libro, p: Periodo, unidades: int | None):
    ws = lib.hoja("Rangos de Cartera")

    def lista(per: Periodo):
        patron = rf"cartera a {per.nombre_mes}\s+{per.anio}"
        for c in buscar_celdas(ws, patron):
            if c.row > 3:      # el resumen también dice "TOTAL CARTERA A ..."
                continue
            col0 = c.column
            f_enc = c.row + 1
            items, total_hoja, fila_total = [], None, None
            for r in range(f_enc + 1, f_enc + 400):
                n0 = ws.cell(r, col0).value
                if norm(n0).startswith("total"):
                    total_hoja, fila_total = ws.cell(r, col0 + 2).value, r
                    break
                apto, val = ws.cell(r, col0 + 1).value, ws.cell(r, col0 + 2).value
                if apto is None and val is None:
                    break
                if es_numero(val):
                    items.append({"apto": str(apto).strip(), "valor": float(val), "fila": r,
                                  "fuente": f"{ws.title}!{coord(r, col0 + 2)}"})
            return {"items": items, "total_hoja": num(total_hoja) if total_hoja is not None else None,
                    "fuente_total": f"{ws.title}!{coord(fila_total, col0 + 2)}" if fila_total else None}
        # Segundo diseño: encabezado «APTO | NOMBRE | <fecha del mes>» sin título «CARTERA A …»
        for r in range(1, 7):
            col_apto = next((c for c in range(1, min(ws.max_column, 40) + 1) if norm(ws.cell(r, c).value) == "apto"), None)
            if not col_apto:
                continue
            col_val = columna_mes(ws, r, per, desde_col=col_apto + 1, hasta_col=col_apto + 8)
            if not col_val:
                return None
            items, total_hoja, fila_total = [], None, None
            for f in range(r + 1, r + 500):
                apto, val = ws.cell(f, col_apto).value, ws.cell(f, col_val).value
                if norm(apto).startswith("total") or norm(ws.cell(f, col_apto + 1).value).startswith("total"):
                    total_hoja, fila_total = val, f
                    break
                if apto is None and val is None:
                    break
                if es_numero(val):
                    items.append({"apto": str(apto).strip(), "valor": float(val), "fila": f,
                                  "fuente": f"{ws.title}!{coord(f, col_val)}"})
            return {"items": items, "total_hoja": num(total_hoja) if es_numero(total_hoja) else None,
                    "fuente_total": f"{ws.title}!{coord(fila_total, col_val)}" if fila_total else None}
        return None

    actual = lista(p)
    anterior = lista(p.anterior())
    if actual is None:
        raise ValueError(
            f"La hoja «{ws.title}» no contiene el detalle de cartera de {p.etiqueta}. "
            "Seleccione el periodo que realmente contiene el Excel y vuelva a cargarlo."
        )
    items = sorted(actual["items"], key=lambda x: -x["valor"])
    total = sum(i["valor"] for i in items)
    tramos = [
        ("Mayores a $3.000.000", lambda v: v > 3_000_000),
        ("Entre $1.000.001 y $3.000.000", lambda v: 1_000_000 < v <= 3_000_000),
        ("Hasta $1.000.000", lambda v: v <= 1_000_000),
    ]
    rangos = []
    for nombre, cond in tramos:
        sel = [i for i in items if cond(i["valor"])]
        s = sum(i["valor"] for i in sel)
        rangos.append({"rango": nombre, "valor": s, "inmuebles": len(sel),
                       "pct_cartera": s / total if total else 0,
                       "pct_unidades": len(sel) / unidades if unidades else None})

    # Resumen que trae la hoja (para contrastar en el validador)
    resumen_hoja = None
    f_tot = buscar_fila(ws, "TOTAL CARTERA A", columnas=("B",))
    if f_tot:
        resumen_hoja = {"total": num(ws.cell(f_tot, 3).value), "inmuebles": num(ws.cell(f_tot, 4).value),
                        "fuente": f"{ws.title}!C{f_tot}"}
    unidades_hoja = None
    for c in buscar_celdas(ws, r"^total todos los inmuebles"):
        for dc in range(1, 8):
            if es_numero(ws.cell(c.row, c.column + dc).value):
                unidades_hoja = int(ws.cell(c.row, c.column + dc).value)
                break
    return {
        "items": items, "total": total, "inmuebles": len(items), "rangos": rangos,
        "total_hoja_lista": actual["total_hoja"], "fuente_total": actual["fuente_total"],
        "resumen_hoja": resumen_hoja, "unidades_hoja": unidades_hoja,
        "anterior": {"total": sum(i["valor"] for i in anterior["items"]), "inmuebles": len(anterior["items"])}
        if anterior else None,
    }


def _mayores_por_evolucion(lib: Libro, p: Periodo, umbral: float):
    """Copropiedades sin hoja de seguimiento de cobro: saldos del mes y del anterior desde «MAYORES DEUDORES»
    (encabezado «# | APTO | <meses>»). Las columnas de gestión jurídica quedan vacías."""
    ws = lib.hoja("MAYORES DEUDORES")
    # encabezado «APTO» (o «Inmueble», «Apartamento», «Unidad»); si hay varias tablas, la que trae más meses
    opciones = []
    for r in range(1, 8):
        for c in range(1, min(ws.max_column, 30) + 1):
            if norm(ws.cell(r, c).value) in ("apto", "inmueble", "apartamento", "unidad"):
                opciones.append((len(meses_en_fila(ws, r, desde_col=c + 1, hasta_col=c + 12)), -r, -c))
    for _, menos_r, menos_c in sorted(opciones, reverse=True)[:1]:
        r, col_apto = -menos_r, -menos_c
        c_act = columna_mes(ws, r, p, desde_col=col_apto + 1, hasta_col=col_apto + 12)
        c_ant = columna_mes(ws, r, p.anterior(), desde_col=col_apto + 1, hasta_col=col_apto + 12)
        if not c_act:
            return {"items": [], "encabezado_abogado": "Gestión según informe del abogado", "sin_gestion": True}
        salida = []
        for f in range(r + 1, r + 300):
            apto, act = ws.cell(f, col_apto).value, ws.cell(f, c_act).value
            if apto is None and act is None:
                break
            if norm(apto).startswith("total") or not es_numero(act):
                continue
            if float(act) > umbral:
                salida.append({"apto": str(apto).strip(), "saldo_anterior": num(ws.cell(f, c_ant).value) if c_ant else 0.0,
                               "saldo_actual": float(act), "gestion_abogado": "", "detalle_gestion": "",
                               "proceso_manual": "", "observacion_rf": "", "fuente": f"{ws.title}!{coord(f, c_act)}",
                               "observacion_compartida": False, "seg_observacion": None, "seg_detalle": None})
        return {"items": salida, "encabezado_abogado": "Gestión según informe del abogado", "sin_gestion": True}
    return {"items": [], "encabezado_abogado": "Gestión según informe del abogado", "sin_gestion": True}


def detectar_hoja_seguimiento(lib: Libro) -> str | None:
    """Hoja de seguimiento de cobro («may deudores»), nombre exacto."""
    return next((w.title for w in reversed(lib.valores.worksheets) if norm(w.title) == norm("may deudores")), None)


def extraer_mayores_deudores(lib: Libro, umbral: float = 1_000_000, p: Periodo | None = None):
    titulo = lib.resolver("SEGUIMIENTO DEUDORES", lambda: detectar_hoja_seguimiento(lib))
    ws = lib.hoja_mapeada("SEGUIMIENTO DEUDORES", titulo) if titulo else None
    f_enc = buscar_fila(ws, "Apto", columnas=("A",)) if ws else None
    if not f_enc:
        if p is None:
            raise KeyError("No se encontró la hoja de seguimiento de mayores deudores («may deudores»).")
        return _mayores_por_evolucion(lib, p, umbral)
    cols = [f"c{i}" for i in range(8)]
    enc = [str(ws.cell(f_enc, 1 + i).value or "").strip() for i in range(8)]
    _, filas, _ = leer_tabla(ws, f_enc, 1, max_filas=200, columnas=cols)
    # una observación combinada (p. ej. G6:G7) aplica a todas las filas del rango
    combinadas = {}
    for rng in ws.merged_cells.ranges:
        v = ws.cell(rng.min_row, rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                combinadas[(r, c)] = (v, rng.max_row > rng.min_row and r > rng.min_row)
    salida = []
    for f in filas:
        vals = [f[c] for c in cols]
        compartida = False
        for i in range(len(vals)):
            if (f["_fila"], i + 1) in combinadas:
                vals[i], cont = combinadas[(f["_fila"], i + 1)]
                compartida = compartida or (cont and i == 6)
        if not es_numero(vals[2]) and not es_numero(vals[1]):
            continue
        reg = {
            "apto": str(vals[0]).strip(),
            "saldo_anterior": num(vals[1]),
            "saldo_actual": num(vals[2]),
            "gestion_abogado": (vals[3] or "").strip() if isinstance(vals[3], str) else "",
            "detalle_gestion": (vals[4] or "").strip() if isinstance(vals[4], str) else "",
            "proceso_manual": (vals[5] or "").strip() if len(vals) > 5 and isinstance(vals[5], str) else "",
            "observacion_rf": (vals[6] or "").strip() if len(vals) > 6 and isinstance(vals[6], str) else "",
            "fuente": f"{ws.title}!C{f['_fila']}",
            "observacion_compartida": compartida,
            "seg_observacion": lib.segmentos(ws, f"G{f['_fila']}"),
            "seg_detalle": lib.segmentos(ws, f"E{f['_fila']}"),
        }
        if reg["saldo_actual"] > umbral:
            salida.append(reg)
    encabezado_abogado = enc[3] if len(enc) > 3 and enc[3] else "Gestión según informe del abogado"
    return {"items": salida, "encabezado_abogado": encabezado_abogado}


# ------------------------------------------------------- bloques por título

def _bloque_titulado(ws, patron_titulo: str):
    """Último bloque cuyo título cumple el patrón. Devuelve items y total."""
    celdas = list(buscar_celdas(ws, patron_titulo))
    if not celdas:
        return None
    t = celdas[-1]
    f_enc = t.row + 1
    col0 = t.column
    enc = [norm(ws.cell(f_enc, col0 + i).value) for i in range(6)]
    # columnas de valor y detalle según el encabezado (no por posición fija)
    c_val = col0 + next((i for i, e in enumerate(enc) if e.startswith("valor")), 3)
    c_det = col0 + next((i for i, e in enumerate(enc) if e.startswith(("detalle", "concepto"))), 4)
    items, total, f_total = [], None, None
    for r in range(f_enc + 1, f_enc + 100):
        v0 = ws.cell(r, col0).value
        if norm(v0).startswith("total"):
            for dc in range(1, 5):
                if es_numero(ws.cell(r, col0 + dc).value):
                    total, f_total = float(ws.cell(r, col0 + dc).value), coord(r, col0 + dc)
                    break
            break
        fila = [ws.cell(r, col0 + i).value for i in range(6)]
        if all(x is None for x in fila):
            break
        items.append({"_fila": r, "_col_val": c_val, "_col_det": c_det,
                      **{enc[i] or f"c{i}": fila[i] for i in range(6)}})
    return {"titulo": str(t.value).strip(), "items": items, "total": total,
            "fuente_total": f"{ws.title}!{f_total}" if f_total else None, "fila_titulo": t.row}


def _fecha_txt(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    return str(v).strip() if v is not None else ""


def _fecha_parse(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        m = re.match(r"(\d{1,2}) de (\w+) de (\d{4})", norm(v))
        if m:
            from .excel import MESES
            try:
                return date(int(m.group(3)), MESES.index(m.group(2)) + 1, int(m.group(1)))
            except ValueError:
                return None
    return None


def _bloque_anual(ws, patron: str, p: Periodo):
    """Bloque «1330 ANTICIPOS Y AVANCES A CONTRATISTAS» con encabezado FECHA | CONTRATISTA | <mes> | <mes> | DETALLE:
    toma la columna del periodo y la deja como «valor»."""
    celdas = list(buscar_celdas(ws, patron))
    if not celdas:
        return None
    t = celdas[-1]
    f_enc, col0 = t.row + 1, t.column
    c_mes = columna_mes(ws, f_enc, p, desde_col=col0, hasta_col=col0 + 8)
    if not c_mes:
        return None
    enc = {c: norm(ws.cell(f_enc, c).value) for c in range(col0, col0 + 8)}
    c_det = next((c for c, e in enc.items() if e.startswith("detalle")), None)
    items, total, f_total = [], None, None
    for r in range(f_enc + 1, f_enc + 100):
        if norm(ws.cell(r, col0).value).startswith("total"):
            if es_numero(ws.cell(r, c_mes).value):
                total, f_total = float(ws.cell(r, c_mes).value), coord(r, c_mes)
            break
        fila = {"fecha": ws.cell(r, col0).value, "contratista": ws.cell(r, col0 + 1).value,
                "valor": ws.cell(r, c_mes).value, "detalle": ws.cell(r, c_det).value if c_det else None}
        if all(v is None for v in fila.values()):
            continue
        items.append({"_fila": r, "_col_val": c_mes, "_col_det": c_det, **fila})
    return {"titulo": str(t.value).strip(), "items": items, "total": total,
            "fuente_total": f"{ws.title}!{f_total}" if f_total else None, "fila_titulo": t.row}


def _bloque_del_mes(ws, base: str, p: Periodo):
    """Bloque «<base> A <MES> <AÑO>». Si el título no trae año («DEUDORES VARIOS A AGOSTO CTA 1380»), se acepta solo
    cuando es el último bloque de su tipo en la hoja y no menciona ningún otro año: así no se toma un mes de otro año."""
    mes = p.nombre_mes
    b = _bloque_titulado(ws, rf"^{base}\s+(?:al?\s+)?(?:\d{{1,2}}\s+de\s+)?{mes}\s+(?:de\s+)?{p.anio}\b")
    if b:
        return b
    todos = list(buscar_celdas(ws, rf"^{base}\b"))
    if not todos:
        return None
    ultimo = norm(todos[-1].value)
    # año vigente del papel de trabajo: el mayor año escrito en los títulos de la hoja (p. ej. «… ENERO 2026»)
    anios = [int(a) for c in buscar_celdas(ws, r"[a-z]") for a in re.findall(r"\b((?:19|20)\d{2})\b", norm(c.value))]
    if anios and max(anios) != p.anio:               # el bloque sin año pertenece al año vigente de la hoja, no a p
        return None
    if re.search(rf"^{base}\s+(?:al?\s+)?(?:\d{{1,2}}\s+de\s+)?{mes}\b", ultimo) and not re.search(r"\b(19|20)\d{2}\b", ultimo):
        b = _bloque_titulado(ws, rf"^{base}\s+(?:al?\s+)?(?:\d{{1,2}}\s+de\s+)?{mes}\b")
        if b and b["fila_titulo"] == todos[-1].row:
            b["anio_supuesto"] = p.anio            # trazabilidad: el año no estaba escrito en el título
            return b
    return None


def extraer_anticipos(lib: Libro, p: Periodo):
    ws = lib.hoja("ANTIC. E INGRESOS")
    ant = _bloque_del_mes(ws, "anticipos y avances a contratistas", p)
    dv = _bloque_del_mes(ws, "deudores varios", p)
    if ant is None and dv is None:                   # otro formato: un bloque anual con una columna por mes
        ant = _bloque_anual(ws, r"^(\d{4}\s+)?anticipos y avances a contratistas", p)
        dv = _bloque_anual(ws, r"^(\d{4}\s+)?deudores varios", p)

    def items(b, valor_key="valor"):
        out = []
        for it in (b or {}).get("items", []):
            v = it.get(valor_key)
            if not es_numero(v):
                continue
            fp = _fecha_parse(it.get("fecha"))
            out.append({"fecha": fp.strftime("%d/%m/%Y") if fp else _fecha_txt(it.get("fecha")), "fecha_iso": (_fecha_parse(it.get("fecha")) or "")
                        and _fecha_parse(it.get("fecha")).isoformat(),
                        "tercero": str(it.get("contratista") or "").strip(), "valor": float(v),
                        "detalle": str(it.get("detalle") or "").strip(),
                        "seg_detalle": lib.segmentos(ws, coord(it["_fila"], it.get("_col_det") or 5)),
                        "fuente": f"{ws.title}!{coord(it['_fila'], it.get('_col_val') or 4)}"})
        return out

    return {
        "anticipos": {"titulo": ant and ant["titulo"], "items": items(ant), "total_hoja": ant and ant["total"],
                      "fuente_total": ant and ant["fuente_total"]},
        "deudores_varios": {"titulo": dv and dv["titulo"], "items": items(dv), "total_hoja": dv and dv["total"],
                            "fuente_total": dv and dv["fuente_total"]},
    }


def extraer_cxp(lib: Libro, p: Periodo):
    ws = lib.hoja("CXP")
    # «CUENTAS POR PAGAR A AGOSTO 2026», «… AL 31 DE AGOSTO DE 2026», «… 31 DE AGOSTO DE 2026»
    b = _bloque_titulado(ws, rf"^cuentas por pagar\s+(al?\s+)?(\d{{1,2}}\s+de\s+)?{p.nombre_mes}\s+(de\s+)?{p.anio}")
    if not b:
        return None
    items = []
    for it in b["items"]:
        v = it.get("valor")
        if not es_numero(v):
            continue
        items.append({"edad": str(it.get("edad") or "").strip(), "tercero": str(it.get("tercero") or "").strip(),
                      "valor": float(v), "concepto": str(it.get("concepto") or "").strip(),
                      "seg_concepto": lib.segmentos(ws, f"E{it['_fila']}"),
                      "fuente": f"{ws.title}!D{it['_fila']}", "_fila": it["_fila"]})
    items.sort(key=lambda x: -x["valor"])
    return {"titulo": b["titulo"], "items": items, "total": sum(i["valor"] for i in items),
            "total_hoja": b["total"], "fuente_total": b["fuente_total"]}


# ---------------------------------------------------- resultado y patrimonio

def extraer_patrimonio(lib: Libro, p: Periodo):
    ws = lib.hoja("PATRIMONIO")
    f_enc = buscar_fila(ws, "RESULTADO DEL EJERCICIO")
    if not f_enc or not meses_en_fila(ws, f_enc):
        raise KeyError(f"La hoja «{ws.title}» no trae el encabezado «Resultado del ejercicio» con los meses en columnas; "
                       "sin los meses no se sabe qué columna es la del periodo")
    f_ing = buscar_fila(ws, "Total ingresos", desde=f_enc)
    f_gas = buscar_fila(ws, "Total gastos", desde=f_enc)
    f_res = buscar_fila(ws, "Resultado del ejercicio", desde=f_enc + 1)
    col = columna_mes(ws, f_enc, p, f_res)
    if not col or not f_res:
        raise KeyError(f"La hoja «{ws.title}» no trae la columna de {p.etiqueta} o la fila del resultado del ejercicio")
    col_ant = columna_mes(ws, f_enc, p.anterior(), f_res)

    f_exc = buscar_fila(ws, "Excedentes acumulados", desde=f_res + 1, exacta=True)
    f_tot_exc = buscar_fila(ws, "Total excedentes acumulados", desde=f_exc)
    c_exc = columna_mes(ws, f_exc, p, f_tot_exc) or col
    origen = _bloque_conceptos(ws, f_exc + 1, f_tot_exc, c_exc, None)

    # Bloque opcional: no todas las copropiedades analizan qué activos respaldan los excedentes
    f_disp = buscar_fila(ws, "Disponible de excedentes acumulados", desde=f_tot_exc)
    f_tot_disp = buscar_fila(ws, "Disponibilidad de excedentes", desde=f_disp + 1) if f_disp else None
    c_disp = (columna_mes(ws, f_disp, p, f_tot_disp) or col) if f_tot_disp else None
    disponibilidad = _bloque_conceptos(ws, f_disp + 1, f_tot_disp, c_disp, None) if f_tot_disp else []
    if disponibilidad:
        disponibilidad[-1]["es_total"] = True
    total_disp = num(ws.cell(f_tot_disp, c_disp).value) if f_tot_disp else None
    total_exc = num(ws.cell(f_tot_exc, c_exc).value)
    for d in disponibilidad:
        d["pct"] = d["actual"] / total_disp if total_disp and not d["es_total"] else None
        d["formula"] = lib.formula(ws, d["fuente"].split("!")[1]) if d["fuente"] else None
    return {
        "ingresos": num(ws.cell(f_ing, col).value), "gastos": num(ws.cell(f_gas, col).value),
        "resultado": num(ws.cell(f_res, col).value),
        "resultado_anterior": num(ws.cell(f_res, col_ant).value) if col_ant else None,
        "fuente_resultado": _f(ws, f_res, col),
        "origen_excedentes": origen, "total_excedentes": total_exc, "fuente_excedentes": _f(ws, f_tot_exc, c_exc),
        "disponibilidad": disponibilidad, "total_disponibilidad": total_disp,
        "fuente_disponibilidad": _f(ws, f_tot_disp, c_disp) if f_tot_disp else None,
    }


def extraer_presupuesto(lib: Libro):
    ws = lib.hoja("PRESUPUESTO")
    # «Rubro» exacto: el título «RUBROS SOBRE EJECUTADOS…» también empieza por «rubro» y no es el encabezado
    f_enc = buscar_fila(ws, "Rubro", columnas=("B",), exacta=True) or buscar_fila(ws, "Rubro", columnas=("B",))
    f_tot = buscar_fila(ws, "Total rubros", columnas=("B",), desde=f_enc)
    rubros = []
    for r in range(f_enc + 1, f_tot):
        nombre = ws.cell(r, 2).value
        if not nombre:
            continue
        ejec = num(ws.cell(r, 5).value)
        pres = num(ws.cell(r, 4).value)
        rubros.append({
            "rubro": str(nombre).strip(), "presupuesto_anual": num(ws.cell(r, 3).value),
            "presupuesto_periodo": pres, "ejecucion": ejec, "sobre_ejecutado": ejec - pres,
            "observacion": str(ws.cell(r, 7).value or "").strip(), "fuente": f"{ws.title}!E{r}", "_fila": r,
        })
    # imágenes de soporte (extractos del auxiliar) ancladas en la fila de cada rubro
    wf = lib.hoja_formulas(ws)
    por_fila = {}
    for im in getattr(wf, "_images", []):
        try:
            fila = im.anchor._from.row + 1
            por_fila.setdefault(fila, []).append({"bytes": im._data(), "ancho": im.width, "alto": im.height})
        except Exception:
            continue
    for r in rubros:
        r["soportes"] = por_fila.get(r["_fila"], [])
    rubros.sort(key=lambda x: -x["sobre_ejecutado"])
    encabezado_periodo = str(ws.cell(f_enc, 4).value or "").strip()
    return {"rubros": rubros, "total_calculado": sum(r["sobre_ejecutado"] for r in rubros),
            "total_hoja": -num(ws.cell(f_tot, 6).value), "fuente_total": f"{ws.title}!F{f_tot}",
            "formula_total": lib.formula(ws, f"F{f_tot}"), "encabezado_periodo": encabezado_periodo}


def extraer_reservas(lib: Libro, p: Periodo):
    ws = lib.hoja("CUADRO RESERVA")
    f_enc = buscar_fila(ws, "Reservas patrimonio")
    f_dif = buscar_fila(ws, "Diferencia contable de reserva")
    f_res = buscar_fila(ws, "Reserva fondo de imprevistos", desde=f_enc)
    col = columna_mes(ws, f_enc, p, f_res)
    col_ant = columna_mes(ws, f_enc, p.anterior(), f_res)
    if not col:                                  # cuadro de un solo corte: columna «Valor» en vez de meses
        col = next((c for c in range(2, ws.max_column + 1) if norm(ws.cell(f_enc, c).value) == "valor"), None)
    if not col:
        raise KeyError(f"La hoja «{ws.title}» no trae la columna de {p.etiqueta} ni una columna «Valor»")
    return {"reserva": num(ws.cell(f_res, col).value),
            "reserva_anterior": num(ws.cell(f_res, col_ant).value) if col_ant else None,
            "fuente_reserva": _f(ws, f_res, col),
            "diferencia": num(ws.cell(f_dif, col).value) if f_dif else 0,
            "fuente_diferencia": _f(ws, f_dif, col)}


def extraer_conciliaciones(lib: Libro, p: Periodo):
    ws = lib.hoja("CONCILIACIONES")
    cuentas = []
    for c in buscar_celdas(ws, r"^banco$"):
        r0, c0 = c.row, c.column
        nombre = ws.cell(r0 - 1, c0).value
        col = columna_mes(ws, r0, p, None, desde_col=c0 + 1, hasta_col=c0 + 16)
        if not col:
            continue
        datos = {"cuenta": str(nombre or "").strip(), "anexos": {}, "fuente": None}
        for r in range(r0 + 1, r0 + 12):
            et = norm(ws.cell(r, c0).value)
            v = ws.cell(r, col).value
            if et.startswith("anexo"):
                datos["anexos"][str(ws.cell(r, c0).value).strip()] = (str(v).strip().upper() if v else None)
            elif et.startswith("saldo en libros"):
                datos["libros"] = v
            elif et.startswith("saldo extracto"):
                datos["extracto"] = v
            elif et == "diferencia":
                datos["diferencia_hoja"] = v
                datos["formula_diferencia"] = lib.formula(ws, coord(r, col))
                datos["fuente"] = f"{ws.title}!{coord(r, col)}"
            elif et.startswith("pagos pendientes") or et.startswith("diferencia ("):
                datos.setdefault("partidas", []).append({"concepto": str(ws.cell(r, c0).value).strip(),
                                                         "valor": v})
        if datos.get("libros") is None and datos.get("extracto") is None:
            datos["sin_datos"] = True
        cuentas.append(datos)
    return cuentas


def extraer_retencion(lib: Libro):
    ws = lib.hoja("IMPUESTOS")
    f_enc = buscar_fila(ws, "PERIODO", columnas=("B",))
    filas = []
    for r in range(f_enc + 1, ws.max_row + 1):
        per = periodo_de(ws.cell(r, 2).value)
        if not per:
            continue
        filas.append({"periodo": per.clave, "periodo_corto": per.corto,
                      "presentado": _fecha_parse(ws.cell(r, 3).value),
                      "pagado": _fecha_parse(ws.cell(r, 4).value),
                      "valor": num(ws.cell(r, 5).value), "observacion": str(ws.cell(r, 6).value or "").strip(),
                      "fuente": f"{ws.title}!B{r}"})
    return filas


def extraer_caja_menor(lib: Libro):
    ws = lib.hoja("CAJA MENOR")
    fecha = None
    for c in buscar_celdas(ws, r"^fecha:?$"):
        for dc in range(1, 4):
            f = _fecha_parse(ws.cell(c.row, c.column + dc).value)
            if f:
                fecha = f
                break
        if fecha:
            break
    return {"fecha_arqueo": fecha.isoformat() if fecha else None}


# --------------------------------------------------------------- orquestador

def detectar_periodo(lib: Libro) -> Periodo:
    receta = ((lib.mapa or {}).get("lectores") or {}).get("liquidez")
    if receta:                                      # el cliente tiene su propia estructura de liquidez
        from .lectores import _col, _fila_etiqueta, periodo_celda
        ws = lib.hoja_mapeada("liquidez", receta["hoja"])
        c = _col(receta["columnas"]["concepto"])
        f = int(receta["fila_encabezado"]) if receta.get("fila_encabezado") else _fila_etiqueta(ws, receta["encabezado"], c)
        f_tot = _fila_etiqueta(ws, receta["total"], c, f or 1)
        meses = [periodo_celda(ws.cell(f, cc).value) for cc in range(1, ws.max_column + 1)] if f else []
        con_dato = [m for cc, m in zip(range(1, ws.max_column + 1), meses)
                    if m and f_tot and isinstance(ws.cell(f_tot, cc).value, (int, float))]
        if con_dato or any(meses):
            return max(con_dato or [m for m in meses if m], key=lambda x: x.clave)
    ws = lib.hoja("CUADRO EFECTIVO")
    f_enc = buscar_fila(ws, "Análisis de liquidez")
    meses = []
    for p, c in meses_en_fila(ws, f_enc):
        if any(es_numero(ws.cell(r, c).value) for r in range(f_enc + 1, min(ws.max_row, f_enc + 25) + 1)):
            meses.append(p)
    if not meses:
        raise ValueError("No se encontró un periodo con cifras bajo el encabezado de liquidez.")
    return max(meses, key=lambda x: x.clave)


def extraer(ruta_excel: str, periodo: Periodo | None = None, meta: dict | None = None,
            mapa: dict | None = None, libro: Libro | None = None) -> dict:
    lib = libro or Libro(ruta_excel, mapa)
    if libro is not None and mapa is not None:
        lib.mapa = mapa
    p = periodo or detectar_periodo(lib)
    meta = dict(meta or {})
    recetas = (lib.mapa or {}).get("lectores") or {}

    def _cartera(unidades):
        if "cartera" in recetas:                    # receta aprobada del cliente (revisoria/lectores.py)
            from .lectores import leer
            return leer(lib, recetas["cartera"], p, unidades)
        return extraer_cartera(lib, p, unidades)
    faltantes = {}

    def opcional(clave, fn, vacio=None):
        """Sección que no se pudo leer: si su hoja no está, o tiene otro formato que el lector estándar no entiende,
        sale vacía (None) y queda anotada en `faltantes`. No bloquea el informe: la validación la marca como ALTA
        y la diapositiva sale como pendiente. Una receta del cliente (revisoria/lectores.py) la puede leer después."""
        try:
            return fn()
        except KeyError as e:
            faltantes[clave] = str(e.args[0]) if e.args else str(e)    # str(KeyError) agrega comillas
        except Exception as e:                       # la hoja está pero su formato es otro
            faltantes[clave] = ("La hoja tiene un formato distinto al que espera el lector estándar; se puede leer con "
                                f"una receta de lectura del cliente (detalle técnico: {type(e).__name__}: {e})")
        return vacio

    cartera = opcional("cartera", lambda: _cartera(meta.get("unidades")))
    if cartera and not meta.get("unidades") and cartera.get("unidades_hoja"):
        meta["unidades"] = cartera["unidades_hoja"]
        cartera = opcional("cartera", lambda: _cartera(meta["unidades"]))

    VACIO_ANT = {"items": [], "titulo": "", "total_hoja": None, "fuente_total": None}

    def receta(sec, vacio=None):
        """Sección leída con la receta aprobada del cliente (revisoria/lectores.py), en vez del lector nativo."""
        from .lectores import RecetaInvalida, leer
        try:
            return leer(lib, recetas[sec], p)
        except RecetaInvalida as e:
            faltantes[sec] = f"La receta de lectura no se pudo aplicar: {e}"
            lib.desajustes.append(f"Receta de «{sec}»: {e}")
            return vacio

    def nativo_o_receta(sec, fn, vacio=None):
        if sec in recetas:
            return receta(sec, vacio)
        return opcional(sec, fn, vacio)

    ant_nativo = ({} if {"anticipos", "deudores_varios"} <= set(recetas) else
                  opcional("anticipos", lambda: extraer_anticipos(lib, p),
                           {"anticipos": dict(VACIO_ANT), "deudores_varios": dict(VACIO_ANT)}))
    modelo = {
        "meta": {**meta, "periodo": p.clave, "periodo_etiqueta": p.etiqueta, "periodo_corto": p.corto,
                 "cierre_corto": p.cierre_anterior().corto, "anterior_corto": p.anterior().corto,
                 "archivo_excel": __import__("os").path.basename(ruta_excel)},
        "liquidez": nativo_o_receta("liquidez", lambda: extraer_liquidez(lib, p)),
        "deudores": opcional("deudores", lambda: extraer_deudores(lib, p)),
        "cartera": cartera,
        "mayores_deudores": opcional("mayores_deudores", lambda: extraer_mayores_deudores(lib, p=p),
                                     {"items": [], "encabezado_abogado": ""}),
        "anticipos": receta("anticipos", dict(VACIO_ANT)) if "anticipos" in recetas else ant_nativo["anticipos"],
        "deudores_varios": (receta("deudores_varios", dict(VACIO_ANT)) if "deudores_varios" in recetas
                            else ant_nativo["deudores_varios"]),
        "cxp": nativo_o_receta("cxp", lambda: extraer_cxp(lib, p), {"titulo": "", "items": [], "total": None}),
        "patrimonio": opcional("patrimonio", lambda: extraer_patrimonio(lib, p)),
        "presupuesto": nativo_o_receta("presupuesto", lambda: extraer_presupuesto(lib)),
        "reservas": opcional("reservas", lambda: extraer_reservas(lib, p)),
        "conciliaciones": opcional("conciliaciones", lambda: extraer_conciliaciones(lib, p), []),
        "retencion": nativo_o_receta("retencion", lambda: extraer_retencion(lib), []),
        "caja_menor": opcional("caja_menor", lambda: extraer_caja_menor(lib), {"fecha_arqueo": None}),
    }
    if faltantes:
        modelo["faltantes"] = faltantes
    return modelo, lib
