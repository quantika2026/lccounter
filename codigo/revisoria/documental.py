"""Extracción de las secciones NO financieras del informe desde el Excel de papeles de trabajo.

Cubre lo que el informe modelo presenta en la Auditoría Regulatoria (egresos del mes, campos auditados
contables y administrativos, actas de consejo, decisiones de asamblea) y en la Auditoría de Control Interno
(contratos, verificación de pólizas frente al manual, matriz de riesgo, controles de actos administrativos).

Igual que el extractor financiero, busca por etiquetas y guarda la celda fuente de cada dato.
Las celdas que parecen notas internas del equipo (preguntas, comentarios a una persona) se separan en
`notas_internas` para que NUNCA lleguen al informe.
"""
from __future__ import annotations

import re
from datetime import date, datetime

from .excel import MESES, Libro, Periodo, norm

# Notas de trabajo que a veces quedan en el Excel y no deben salir en el informe
RX_NOTA_INTERNA = re.compile(r"(\?\s*$)|(^\s*(sandr|qu[eé] ha pasado|aqu[ií] ir[ií]a|pendiente revisar|ojo\b))", re.I)


def _txt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    return str(v).strip()


def _imagenes_por_fila(lib: Libro, ws) -> dict[int, list[dict]]:
    wf = lib.hoja_formulas(ws)
    out: dict[int, list[dict]] = {}
    for im in getattr(wf, "_images", []):
        try:
            fila = im.anchor._from.row + 1
            out.setdefault(fila, []).append({"bytes": im._data(), "ancho": im.width, "alto": im.height,
                                             "col": im.anchor._from.col + 1})
        except Exception:
            continue
    # EMF/WMF: openpyxl los descarta al abrir; se recuperan del zip y se convierten a PNG (una vez por hoja)
    cache = lib.__dict__.setdefault("_vectoriales", {})
    if ws.title not in cache:
        from . import libreoffice
        cache[ws.title] = libreoffice.imagenes_vectoriales(lib.ruta, ws.title)
    for v in cache[ws.title]:
        out.setdefault(v["fila"], []).append({k: v[k] for k in ("bytes", "ancho", "alto", "col")})
    return out


def _seg(lib: Libro, ws, r: int, c: int):
    try:
        return lib.segmentos(ws, f"{ws.cell(r, c).column_letter}{r}")
    except Exception:
        return None


# ---------------------------------------------------------------- egresos del mes

def extraer_egresos(lib: Libro, p: Periodo) -> dict:
    ws = lib.hoja("EGRESO")
    objetivos = {f"comprobantes de egreso {p.nombre_mes} {p.anio}", f"comprobantes de egresos {p.nombre_mes} {p.anio}"}
    ini = None
    for r in range(1, ws.max_row + 1):
        if norm(ws.cell(r, 2).value) in objetivos:
            ini = r
            break
    if ini is None:
        return {"items": [], "faltante": f"No se encontró el bloque «COMPROBANTES DE EGRESO {p.nombre_mes.upper()} {p.anio}»"}
    imgs = _imagenes_por_fila(lib, ws)
    items, total, fin = [], None, None
    for r in range(ini + 2, ws.max_row + 1):
        b = ws.cell(r, 2).value
        if isinstance(b, str) and norm(b).startswith("total egresos"):
            total, fin = ws.cell(r, 6).value, r
            etiqueta_total = b.strip()
            break
        if b in (None, "") and not ws.cell(r, 4).value:
            continue
        obs = _txt(ws.cell(r, 7).value)
        items.append({"ce": _txt(b), "fecha": _txt(ws.cell(r, 3).value), "tercero": _txt(ws.cell(r, 4).value),
                      "descripcion": _txt(ws.cell(r, 5).value), "valor": ws.cell(r, 6).value,
                      "observacion": obs, "seg_observacion": _seg(lib, ws, r, 7),
                      "correcto": norm(obs).startswith("correcto"), "fuente": f"{ws.title}!F{r}", "_fila": r,
                      "imagenes": []})
    # imágenes ancladas en el bloque: se asignan al comprobante de la fila más cercana hacia arriba
    if fin:
        for fila, ims in imgs.items():
            if ini <= fila <= fin and items:
                dueño = max((i for i in items if i["_fila"] <= fila), key=lambda i: i["_fila"], default=items[0])
                dueño["imagenes"] += ims
    suma = sum(i["valor"] or 0 for i in items if isinstance(i["valor"], (int, float)))
    out = {"titulo": _txt(ws.cell(ini, 2).value), "items": items, "total_hoja": total, "suma": suma,
           "fuente_total": f"{ws.title}!F{fin}" if fin else None,
           "etiqueta_total": etiqueta_total if fin else None,
           "con_observacion": sum(1 for i in items if not i["correcto"])}
    if fin and not norm(etiqueta_total).endswith(f"{p.nombre_mes} {p.anio}"):
        out["etiqueta_inconsistente"] = etiqueta_total
    return out


# ---------------------------------------------------------------- campos auditados

def _clasificar(t: str) -> str:
    n = norm(t)
    if n.startswith("(op)") or n.startswith("op:") or n.startswith("(op "):
        return "opinion"
    if n.startswith("hallazgo"):
        return "hallazgo"
    return "evidencia"


def _limpiar(t: str) -> str:
    return re.sub(r"^\s*\((ev|op)\)\s*:?\s*", "", t, flags=re.I).strip()


def detectar_hojas_auditoria(lib: Libro) -> list:
    """[[título, bloque]]: «AUD CONT Y ADMON» (una hoja con los dos bloques) o, si la copropiedad los separa, todas
    las hojas con el encabezado «Campo auditado» (p. ej. «REV DOC CONTABLE» y «REV DOC ADMON»)."""
    try:
        return [[lib.hoja("AUD CONT Y ADMON").title, None]]
    except KeyError:
        pass
    hojas = []
    for ws in lib.valores.worksheets:
        if any(norm(_txt(ws.cell(r, 2).value)).startswith("campo auditado") for r in range(1, 8)):
            titulo = norm(" ".join(_txt(ws.cell(r, 2).value) for r in range(1, 4)) + " " + ws.title)
            hojas.append([ws.title, "Contable" if "contab" in titulo else "Administrativo" if "admin" in titulo else None])
    return hojas


def _hojas_auditoria(lib: Libro) -> list:
    hojas = lib.resolver("AUDITORIA DOCUMENTAL", lambda: detectar_hojas_auditoria(lib))
    if not hojas:
        raise KeyError("No se encontró la hoja de auditoría documental («AUD CONT Y ADMON» o «Campo auditado»).")
    return [(lib.hoja_mapeada("AUDITORIA DOCUMENTAL", t), b) for t, b in hojas]


def extraer_auditoria_documental(lib: Libro) -> dict:
    """Auditoría documental: bloques Contable y Administrativo, un campo auditado por grupo de filas."""
    campos, notas = [], []
    for ws, bloque_hoja in _hojas_auditoria(lib):
        _leer_auditoria(lib, ws, bloque_hoja, campos, notas)
    for cpo in campos:
        cpo["nivel"] = {"MEDIO": "MEDIO", "ALTO": "ALTO", "ODM": "ODM", "BAJO": "BAJO"}.get(cpo["nivel"].strip(), cpo["nivel"])
    return {"campos": campos, "notas_internas": notas}


def _leer_auditoria(lib: Libro, ws, bloque_hoja, campos: list, notas: list):
    imgs = _imagenes_por_fila(lib, ws)
    bloque, actual = bloque_hoja, None
    for r in range(1, ws.max_row + 1):
        b = _txt(ws.cell(r, 2).value)
        nb = norm(b)
        if nb.startswith("auditoria documental"):
            bloque = "Contable" if "contable" in nb else "Administrativo" if "administ" in nb else (bloque_hoja or "Administrativo")
            actual = None
            continue
        if nb.startswith("campo auditado"):
            continue
        c, d, e = _txt(ws.cell(r, 3).value), _txt(ws.cell(r, 4).value), _txt(ws.cell(r, 5).value)
        f = _txt(ws.cell(r, 6).value)
        extras = [_txt(ws.cell(r, k).value) for k in range(7, 10)]
        if b:
            actual = {"bloque": bloque, "campo": b, "normatividad": [c] if c else [], "textos": [],
                      "nivel": e.upper() if e else "", "documentos": f if f not in ("", "N/A") else "",
                      "n_documentos": ws.cell(r, 7).value if isinstance(ws.cell(r, 7).value, (int, float)) else None,
                      "pct_incumplimiento": ws.cell(r, 8).value if isinstance(ws.cell(r, 8).value, (int, float)) else None,
                      "fila_inicio": r, "fila_fin": r, "imagenes": [], "fuente": f"{ws.title}!B{r}"}
            campos.append(actual)
        if actual is None:
            continue
        actual["fila_fin"] = r
        if c and not b:
            # una segunda norma (p. ej. impuesto de timbre) o un texto que quedó en la columna C
            (actual["normatividad"] if len(c) < 120 and not c.startswith("(") else actual["textos"]).append(
                c if len(c) < 120 and not c.startswith("(") else {"tipo": _clasificar(c), "texto": _limpiar(c),
                                                                  "fuente": f"{ws.title}!C{r}"})
        if d:
            actual["textos"].append({"tipo": _clasificar(d), "texto": _limpiar(d), "fuente": f"{ws.title}!D{r}",
                                     "seg": _seg(lib, ws, r, 4)})
        if e and not b and not actual["nivel"]:
            actual["nivel"] = e.upper()
        for k, x in zip((6, 7, 8, 9), [f] + extras):
            if x and isinstance(x, str) and RX_NOTA_INTERNA.search(x) and not x.replace(".", "").isdigit():
                notas.append({"celda": f"{ws.title}!{ws.cell(r, k).column_letter}{r}", "texto": x,
                              "campo": actual["campo"]})
        for fila, ims in imgs.items():
            if fila == r:
                actual["imagenes"] += ims


# ---------------------------------------------------------------- actas y asamblea

RX_ACTA = re.compile(r"^\s*(\d{1,3})\s+(.*)$", re.S)


def extraer_actas_consejo(lib: Libro, p: Periodo) -> dict:
    ws = lib.hoja("ACTAS DE CONSEJO")
    actas = []
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 3).value
        temas = _txt(ws.cell(r, 4).value)
        if a in (None, "") or norm(a).startswith("n. acta") or norm(a).startswith("relacion") or not temas:
            continue
        numero, fecha_txt, anio = parse_acta(a)
        actas.append({"numero": numero, "fecha": fecha_txt, "anio": anio, "temas": temas,
                      "firma_presidente": _txt(ws.cell(r, 5).value), "firma_administracion": _txt(ws.cell(r, 6).value),
                      "fuente": f"{ws.title}!C{r}", "_fila": r})
    return actas_del_anio(actas, p)


def parse_acta(a, fecha=None) -> tuple[str, str, int | None]:
    """(número, fecha en texto, año) de la celda «N. acta y fecha» (o de número y fecha en columnas separadas)."""
    if fecha not in (None, ""):
        numero = _txt(a)
        a = fecha
        if isinstance(a, (datetime, date)):
            return numero, a.strftime("%d/%m/%Y"), a.year
        _, fecha_txt, anio = parse_acta(_txt(a))
        return numero, fecha_txt, anio
    txt = _txt(a)
    m = RX_ACTA.match(txt)
    numero, fecha_txt = (m.group(1), re.sub(r"\s+", " ", m.group(2)).strip()) if m else ("", txt)
    if not m and txt.isdigit():
        numero, fecha_txt = txt, ""
    anio = None
    if isinstance(a, (datetime, date)):
        anio = a.year
    else:
        mm = re.search(r"(20\d{2})|/(\d{2})\s*$", fecha_txt)
        if mm:
            anio = int(mm.group(1)) if mm.group(1) else 2000 + int(mm.group(2))
        elif re.search(r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\s+(20\d{2})", norm(fecha_txt)):
            anio = int(re.search(r"(20\d{2})", fecha_txt).group(1))
    return numero, fecha_txt, anio


def actas_del_anio(actas: list[dict], p: Periodo) -> dict:
    del_anio = [x for x in actas if x["anio"] == p.anio]
    # las actas sin año reconocible que siguen a una del año se consideran del año (p. ej. «54» sin fecha)
    if del_anio:
        primera = del_anio[0]["_fila"]
        del_anio = [x for x in actas if x["_fila"] >= primera]
    sin_firma = [x for x in del_anio if not x["firma_presidente"] or "sin firma" in norm(x["firma_presidente"])]
    return {"actas": del_anio, "todas": len(actas), "sin_firma": sin_firma}


def extraer_asamblea(lib: Libro) -> dict:
    """Bloques «Decisiones …» de la hoja de asamblea, en la columna donde estén (B en un formato, F en otro).
    Cada fila: tema (decisión o pregunta), detalle (texto aprobado o resultado de la votación) y seguimiento."""
    ws = lib.hoja("ACTA ASAMBLEA")
    bloques, actual, col0 = [], None, None
    for r in range(1, ws.max_row + 1):
        titulo = next((c for c in range(1, min(ws.max_column, 15) + 1)
                       if re.match(r"(decisiones|desiciones)\b", norm(_txt(ws.cell(r, c).value)))
                       and not norm(_txt(ws.cell(r, c).value)).startswith(("decisiones importantes", "desiciones importantes"))), None)
        if titulo:
            col0 = titulo
            actual = {"titulo": _txt(ws.cell(r, col0).value), "filas": [], "fuente": f"{ws.title}!{ws.cell(r, col0).coordinate}"}
            bloques.append(actual)
            continue
        if actual is None:
            continue
        b, c, d = (_txt(ws.cell(r, col0 + k).value) for k in range(3))
        nb = norm(b)
        if nb in ("temas", "votacion") or nb.startswith(("decisiones importantes", "desiciones importantes")) \
                or norm(c).startswith("respuesta"):
            continue                                     # fila de encabezados del bloque
        if b or c or d:
            if not b and actual["filas"]:
                actual["filas"][-1]["detalle"] += ("\n" if actual["filas"][-1]["detalle"] else "") + c
                continue
            actual["filas"].append({"tema": b, "detalle": c, "seguimiento": d, "fuente": f"{ws.title}!{ws.cell(r, col0).coordinate}"})
    return {"bloques": bloques}


# ---------------------------------------------------------------- control interno

# columnas de la tabla de contratos, reconocidas por su encabezado (cada copropiedad las ordena distinto)
COLUMNAS_CONTRATO = {"proveedor": ("nombre proveedor", "proveedor", "tercero", "contratista"),
                     "vigencia": ("vigencia",), "valor": ("valor",), "total": ("total contrato", "total"),
                     "observaciones": ("polizas y garantias", "observaciones"), "verificacion": ("verificacion",)}


def _columnas_contrato(ws, f_enc: int, c_n: int) -> dict:
    enc = {c: norm(ws.cell(f_enc, c).value) for c in range(c_n + 1, min(ws.max_column, c_n + 15) + 1)}
    cols = {}
    for campo, alias in COLUMNAS_CONTRATO.items():
        for a in alias:
            c = next((c for c, e in enc.items() if e.startswith(a) and c not in cols.values()), None)
            if c:
                cols[campo] = c
                break
    # formato original (N. | proveedor | vigencia | valor | total | pólizas | verificación) si no hay encabezados
    por_defecto = dict(zip(COLUMNAS_CONTRATO, range(c_n + 1, c_n + 8)))
    return {k: cols.get(k, por_defecto[k]) for k in COLUMNAS_CONTRATO}


def extraer_contratos(lib: Libro) -> dict:
    ws = lib.hoja("CONTRATOS 2026")
    f_enc, c_n = next(((r, c) for r in range(1, 10) for c in (1, 2, 3)
                       if norm(ws.cell(r, c).value) in ("n.", "n", "no.", "no")), (2, 2))
    col = _columnas_contrato(ws, f_enc, c_n)
    contratos, actual = [], None
    for r in range(f_enc + 1, ws.max_row + 1):
        n, prov = ws.cell(r, c_n).value, _txt(ws.cell(r, col["proveedor"]).value)
        celdas = {k: (ws.cell(r, col[k]).value if k in ("valor", "total") and isinstance(ws.cell(r, col[k]).value, (int, float))
                      else _txt(ws.cell(r, col[k]).value)) for k in ("vigencia", "valor", "total", "observaciones", "verificacion")}
        segs = {k: _seg(lib, ws, r, col[k]) for k in ("observaciones", "valor", "vigencia")}
        if not prov and not any(celdas.values()):
            continue                                 # fila numerada pero sin contrato (tabla sin diligenciar)
        if n not in (None, "") or actual is None:
            actual = {"n": _txt(n), "partes": [], "fuente": f"{ws.title}!B{r}"}
            contratos.append(actual)
        # una anotación larga escrita en la columna VALOR pertenece a las observaciones (así se ve en el informe)
        if isinstance(celdas["valor"], str) and len(celdas["valor"]) > 120 and not celdas["observaciones"]:
            celdas["observaciones"], celdas["valor"] = celdas["valor"], ""
            segs["observaciones"], segs["valor"] = segs["valor"], None
        # cada fila (contrato u otrosí / anotación) se conserva como parte del contrato
        actual["partes"].append({"proveedor": prov, **celdas, "seg": segs, "_fila": r})
    contratos = [c for c in contratos if any(p["proveedor"] or p["valor"] or p["vigencia"] for p in c["partes"])]
    total = sum(1 for c in contratos)
    return {"contratos": contratos, "n": total}


def detectar_hoja_polizas(lib: Libro) -> str | None:
    # hoja aparte de verificación de pólizas («CONTRATOS 2026 (2)»); no se confunde con la hoja de contratos
    return next((w.title for w in lib.valores.worksheets
                 if ("(2)" in w.title and norm(w.title).startswith("contratos"))
                 or norm(w.title).startswith(("verificacion polizas", "verificacion de polizas"))), None)


def extraer_verificacion_polizas(lib: Libro) -> dict:
    titulo = lib.resolver("VERIFICACION POLIZAS", lambda: detectar_hoja_polizas(lib))
    if titulo is None:
        return {"filas": [], "no_cumplen": [], "matriz": {"titulo": "", "columnas": [], "filas": []}}
    ws = lib.hoja_mapeada("VERIFICACION POLIZAS", titulo)
    filas = []
    for r in range(3, ws.max_row + 1):
        n, prov, ver = ws.cell(r, 2).value, _txt(ws.cell(r, 3).value), _txt(ws.cell(r, 4).value)
        if prov:
            filas.append({"n": _txt(n), "contrato": prov, "verificacion": ver,
                          "cumple": norm(ver).startswith("cumple"), "fuente": f"{ws.title}!D{r}"})
    return {"filas": filas, "no_cumplen": [f for f in filas if not f["cumple"]], "matriz": matriz_lineamientos(ws)}


def matriz_lineamientos(ws) -> dict:
    """Matriz de lineamientos por tipo de contrato (bloque a la derecha de la verificación de pólizas)."""
    matriz = []
    enc = None
    for r in range(1, ws.max_row + 1):
        for c in range(5, ws.max_column + 1):
            if norm(ws.cell(r, c).value) == "tipo de contrato":
                enc = (r, c)
                break
        if enc:
            break
    if enc:
        r0, c0 = enc
        cols = [c for c in range(c0, c0 + 8) if ws.cell(r0, c).value]
        titulos = [_txt(ws.cell(r0, c).value) for c in cols]
        for r in range(r0 + 1, ws.max_row + 1):
            fila = [_txt(ws.cell(r, c).value) for c in cols]
            if any(fila):
                matriz.append(fila)
        titulo_matriz = next((_txt(ws.cell(r, c0).value) for r in range(1, r0) if ws.cell(r, c0).value), "")
    else:
        titulos, titulo_matriz = [], ""
    return {"titulo": titulo_matriz, "columnas": titulos, "filas": matriz}


def extraer_matriz_riesgo(lib: Libro) -> dict:
    ws = lib.hoja("MATRIZ DICTAMEN")
    enc = next((r for r in range(1, 10) if norm(ws.cell(r, 2).value).startswith("objetivo de control")), 3)
    periodos = []
    for c in (5, 6, 7):
        v = lib.hoja_formulas(ws).cell(enc - 1, c).value if False else ws.cell(enc - 1, c).value
        if isinstance(v, (datetime, date)):
            periodos.append((c, Periodo(v.year, v.month)))
    filas, actual = [], None
    for r in range(enc + 1, ws.max_row + 1):
        obj, desc = _txt(ws.cell(r, 2).value), _txt(ws.cell(r, 3).value)
        if obj:
            actual = {"objetivo": obj, "descripcion": [], "valoracion": _txt(ws.cell(r, 4).value),
                      "niveles": {p.corto: _txt(ws.cell(r, c).value).upper() for c, p in periodos},
                      "fuente": f"{ws.title}!B{r}"}
            filas.append(actual)
        if actual and desc:
            actual["descripcion"].append({"texto": desc, "seg": _seg(lib, ws, r, 3)})
        if actual and not obj:
            for c, p in periodos:
                v = _txt(ws.cell(r, c).value).upper()
                if v and not actual["niveles"].get(p.corto):
                    actual["niveles"][p.corto] = v
    return {"titulo": _txt(ws.cell(enc - 1, 2).value), "periodos": [p.corto for _, p in periodos], "filas": filas}


def extraer_matriz_manuales(lib: Libro) -> dict:
    try:
        ws = lib.hoja("matriz manuales")
    except KeyError:
        return {"filas": []}
    filas = []
    for r in range(4, ws.max_row + 1):
        obj = _txt(ws.cell(r, 2).value)
        if not obj and not ws.cell(r, 3).value:
            continue
        filas.append({"objetivo": obj, "detalle": " ".join(x for x in (_txt(ws.cell(r, 3).value), _txt(ws.cell(r, 4).value)) if x),
                      "valoracion": _txt(ws.cell(r, 5).value),
                      "niveles": [_txt(ws.cell(r, c).value) for c in (6, 7, 8)], "fuente": f"{ws.title}!B{r}"})
    return {"filas": filas}


def extraer_firmas_informes(lib: Libro, p: Periodo) -> dict:
    """Firma de estados financieros y entrega del informe de administración del periodo (no está en el informe modelo)."""
    try:
        ws = lib.hoja("FIRMAS E INF DE GESTION")
    except KeyError:
        return {}
    out = {}
    for r in range(3, ws.max_row + 1):
        v = ws.cell(r, 2).value
        if isinstance(v, (datetime, date)) and v.year == p.anio and v.month == p.mes:
            out = {"firma_contador": _txt(ws.cell(r, 3).value), "firma_rep_legal": _txt(ws.cell(r, 4).value),
                   "informe_gestion": _txt(ws.cell(r, 7).value), "fuente": f"{ws.title}!B{r}"}
    return out


def extraer_documental(lib: Libro, p: Periodo) -> dict:
    partes = {
        "egresos": lambda: extraer_egresos(lib, p),
        "auditoria": lambda: extraer_auditoria_documental(lib),
        "actas_consejo": lambda: extraer_actas_consejo(lib, p),
        "asamblea": lambda: extraer_asamblea(lib),
        "contratos": lambda: extraer_contratos(lib),
        "polizas": lambda: extraer_verificacion_polizas(lib),
        "matriz_riesgo": lambda: extraer_matriz_riesgo(lib),
        "matriz_manuales": lambda: extraer_matriz_manuales(lib),
        "firmas_informes": lambda: extraer_firmas_informes(lib, p),
    }
    recetas = (lib.mapa or {}).get("lectores") or {}
    out, errores = {}, {}
    for k, f in partes.items():
        if k in recetas:                            # sección con receta aprobada del cliente (revisoria/lectores.py)
            from .lectores import RecetaInvalida, leer
            try:
                out[k] = leer(lib, recetas[k], p)
            except RecetaInvalida as e:
                out[k] = None
                errores[k] = f"La receta de lectura no se pudo aplicar: {e}"
                lib.desajustes.append(f"Receta de «{k}»: {e}")
            continue
        try:
            out[k] = f()
        except KeyError as e:                       # hoja que el cliente no tiene: no bloquea
            out[k] = None
            errores[k] = str(e.args[0]) if e.args else str(e)
    out["faltantes"] = errores
    return out


# ---------------------------------------------------------------- validación de las secciones no financieras

CITAS_SOSPECHOSAS = [
    (r"ley\s+1561\s+de\s+2012", "La Ley 1561 de 2012 trata de titulación de la posesión, no de propiedad horizontal. "
                                 "Para libros de actas la norma es la Ley 675 de 2001."),
    (r"ley\s+1563\s+de\s+2012", "La Ley 1563 de 2012 es el estatuto de arbitraje. Para el libro de propietarios la norma es "
                                 "la Ley 675 de 2001 (art. 51)."),
    (r"ley\s+1523\s+de\s+(abril\s+de\s+)?2012", "La Ley 1523 de 2012 es de gestión del riesgo de desastres; para SG-SST "
                                                  "suele citarse la Ley 1562 de 2012. Verificar."),
]


def validar_documental(doc: dict, p: Periodo) -> list:
    from .validador import Hallazgo
    out = []
    A = doc.get("auditoria") or {}
    for n in A.get("notas_internas", []):
        out.append(Hallazgo("D01", "MEDIA", "Notas internas", f"Nota de trabajo en el Excel: «{n['texto'][:60]}»",
                            f"La celda {n['celda']} (campo «{n['campo']}») tiene una nota interna del equipo. "
                            "El generador no la lleva al informe, pero indica un punto sin cerrar.",
                            [n["celda"]], "Resolver la nota y borrarla del papel de trabajo."))
    sin_op = [c["campo"] for c in A.get("campos", []) if not any(t["tipo"] == "opinion" for t in c["textos"])
              and c["nivel"] in ("MEDIO", "BAJO", "ODM")]
    if sin_op:
        out.append(Hallazgo("D02", "MEDIA", "Opinión del revisor",
                            f"{len(sin_op)} campo(s) con nivel MEDIO/ODM/BAJO sin opinión o recomendación en el Excel",
                            "; ".join(sin_op) + ". En el informe modelo esa recomendación se escribía directamente en "
                            "PowerPoint; para generarla automáticamente debe quedar en el Excel con el prefijo «(Op):».",
                            [c["fuente"] for c in A.get("campos", []) if c["campo"] in sin_op],
                            "Escribir la recomendación en la hoja AUD CONT Y ADMON."))
    sin_nivel = [c["campo"] for c in A.get("campos", []) if c["nivel"] not in ("ALTO", "ODM", "MEDIO", "BAJO")]
    if sin_nivel:
        out.append(Hallazgo("D03", "MEDIA", "Semáforo", "Campos sin nivel de cumplimiento",
                            "; ".join(sin_nivel) + ".", [], "Asignar ALTO, ODM, MEDIO o BAJO."))
    for c in A.get("campos", []):
        todo = " ".join(c["normatividad"]) + " " + " ".join(t["texto"] for t in c["textos"])
        for patron, nota in CITAS_SOSPECHOSAS:
            if re.search(patron, norm(todo)):
                out.append(Hallazgo("D04", "MEDIA", "Normatividad", f"Verificar cita normativa en «{c['campo']}»", nota,
                                    [c["fuente"]], "Confirmar la norma antes de emitir."))
    # actas de consejo: fechas fuera de orden o sin firma
    act = (doc.get("actas_consejo") or {}).get("actas", [])
    prev = None
    for a in act:
        m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", a["fecha"] or "")
        if not m:
            continue
        d, mm, y = int(m.group(1)), int(m.group(2)), int(m.group(3)) % 100 + 2000
        try:
            f = date(y, mm, d)
        except ValueError:
            continue
        if f > p.fin:
            out.append(Hallazgo("D05", "BAJA", "Actas de consejo", f"Acta {a['numero']} con fecha posterior al periodo",
                                f"Fecha registrada {a['fecha']}; el informe es de {p.etiqueta}.", [a["fuente"]],
                                "Verificar la fecha del acta."))
        if prev and f < prev[1] and (prev[1] - f).days > 20:
            out.append(Hallazgo("D05", "BAJA", "Actas de consejo", f"Acta {a['numero']} con fecha anterior a la del acta {prev[0]}",
                                f"Acta {prev[0]}: {prev[1]:%d/%m/%Y}; acta {a['numero']}: {f:%d/%m/%Y}. Puede ser un error de digitación.",
                                [a["fuente"]], "Verificar la fecha en el libro de actas."))
        prev = (a["numero"], f)
    sf = (doc.get("actas_consejo") or {}).get("sin_firma", [])
    if sf:
        out.append(Hallazgo("D06", "MEDIA", "Actas de consejo", f"{len(sf)} acta(s) de {p.anio} sin firma del presidente",
                            ", ".join(a["numero"] for a in sf) + ".", [a["fuente"] for a in sf]))
    # egresos
    E = doc.get("egresos") or {}
    if E.get("items") and E.get("total_hoja") is not None and abs((E["total_hoja"] or 0) - E["suma"]) > 1:
        out.append(Hallazgo("D07", "ALTA", "Egresos", "El total de egresos no es la suma de los comprobantes",
                            f"Total {E['total_hoja']:,.0f} vs suma {E['suma']:,.0f}.".replace(",", "."), [E.get("fuente_total")]))
    if E.get("faltante"):
        out.append(Hallazgo("D07", "ALTA", "Egresos", "No está el bloque de egresos del mes", E["faltante"], ["EGRESO"]))
    # pólizas, matriz de riesgo
    P = doc.get("polizas") or {}
    if P.get("no_cumplen"):
        out.append(Hallazgo("D08", "MEDIA", "Contratos", f"{len(P['no_cumplen'])} contrato(s) no cumplen las pólizas del manual",
                            "; ".join(f["contrato"][:60] for f in P["no_cumplen"]) + ".", [f["fuente"] for f in P["no_cumplen"]]))
    M = doc.get("matriz_riesgo") or {}
    vacios = [f["objetivo"] for f in M.get("filas", []) if M.get("periodos") and not f["niveles"].get(M["periodos"][-1])]
    if vacios:
        out.append(Hallazgo("D09", "BAJA", "Matriz de riesgo", f"Sin calificación en {M['periodos'][-1]}",
                            "; ".join(vacios) + ".", [], "Calificar el objetivo de control o explicar por qué no aplica."))
    # información del Excel que el informe modelo no muestra
    F = doc.get("firmas_informes") or {}
    if F and ("sin" in norm(F.get("firma_contador")) or "pendiente" in norm(F.get("firma_contador"))
              or "no hay" in norm(F.get("informe_gestion")) or not F.get("informe_gestion")):
        out.append(Hallazgo("D10", "MEDIA", "Firmas e informes",
                            f"Estados financieros de {p.etiqueta}: contador «{F.get('firma_contador') or '—'}», "
                            f"representante legal «{F.get('firma_rep_legal') or '—'}»; informe de gestión: "
                            f"«{F.get('informe_gestion') or 'sin registro'}»",
                            "Esta información está en la hoja FIRMAS E INF DE GESTIÓN pero el informe modelo no la presenta.",
                            [F.get("fuente", "")], "Decidir si se incluye en el informe."))
    return out
