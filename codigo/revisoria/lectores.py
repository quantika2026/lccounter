"""Lectores declarativos: leer una sección del Excel cuando el cliente la organiza con OTRA estructura.

El mapa (revisoria/mapa.py) resuelve en qué HOJA está cada dato. Un lector resuelve CÓMO está dentro de la hoja:
fila de encabezado, qué columna es cada campo estándar, dónde empieza el bloque del mes. Es una receta en JSON
(sin código) que se guarda en el mapa del cliente (`mapa["lectores"][seccion]`), la propone el subagente
estructurador (agente/estructurador.py) y la aprueba el revisor.

Receta:
    {"seccion": "cxp", "hoja": "CXP",
     "titulo": "cuentas por pagar a {mes} {anio}",   # bloque del mes (hojas con un bloque por mes); o bien
     "fila_encabezado": 3,                             # tabla única con encabezado fijo
     "columnas": {"tercero": "C", "valor": "D", "concepto": ["E", "F"]},   # lista = se unen con espacio
     "niveles": ["E", "F"]}                            # solo matrices: columnas de nivel de cumplimiento

El resultado tiene EXACTAMENTE la forma que produce el lector nativo de Picasso (el informe modelo), así el
generador del PowerPoint no distingue de dónde vino. Las reglas de negocio (qué es «cumple», qué acta es del año,
qué egreso es «correcto») son las mismas del lector nativo: se importan, no se reescriben.
"""
from __future__ import annotations

import re
from datetime import date, datetime

from openpyxl.utils import column_index_from_string, get_column_letter

from .excel import Libro, Periodo, es_numero, norm, periodo_de

# ------------------------------------------------------------------ campos estándar por sección
# tipo: t texto · n número · f fecha · p periodo (mes) · obligatorio: el campo que define una fila
CAMPOS: dict[str, dict] = {
    "cxp": {"descripcion": "Cuentas por pagar del mes (un bloque por mes con título «CUENTAS POR PAGAR A <MES> <AÑO>»)",
            "por_mes": True, "campos": {
                "edad": ("t", False, "antigüedad de la cuenta (p. ej. «1 a 30 días»)"),
                "tercero": ("t", True, "proveedor o acreedor"),
                "valor": ("n", True, "saldo por pagar"),
                "concepto": ("t", False, "concepto u observación")}},
    "anticipos": {"descripcion": "Anticipos y avances a contratistas, cuenta 1330, del mes", "por_mes": True, "campos": {
        "fecha": ("f", False, "fecha del anticipo"), "tercero": ("t", True, "contratista"),
        "valor": ("n", True, "valor del anticipo"), "detalle": ("t", False, "detalle u observación")}},
    "deudores_varios": {"descripcion": "Deudores varios, cuenta 1380, del mes", "por_mes": True, "campos": {
        "fecha": ("f", False, "fecha"), "tercero": ("t", True, "deudor"),
        "valor": ("n", True, "saldo"), "detalle": ("t", False, "detalle u observación")}},
    "retencion": {"descripcion": "Retención en la fuente: una fila por periodo declarado", "por_mes": False, "campos": {
        "periodo": ("p", True, "mes declarado"), "presentado": ("f", False, "fecha de presentación"),
        "pagado": ("f", False, "fecha de pago"), "valor": ("n", True, "valor pagado"),
        "observacion": ("t", False, "observación")}},
    "egresos": {"descripcion": "Comprobantes de egreso del mes revisados", "por_mes": True, "campos": {
        "ce": ("t", True, "número del comprobante de egreso"), "fecha": ("f", False, "fecha"),
        "tercero": ("t", True, "beneficiario"), "descripcion": ("t", False, "descripción del pago"),
        "valor": ("n", True, "valor"), "observacion": ("t", False, "observación de la revisoría («Correcto» si está bien)")}},
    "polizas": {"descripcion": "Verificación de pólizas de cada contrato frente al manual de contratación (NO pagos de pólizas)",
                "por_mes": False, "campos": {
                    "n": ("t", False, "número"), "contrato": ("t", True, "contrato / proveedor"),
                    "verificacion": ("t", True, "resultado de la verificación («Cumple…» / «No cumple…»)")}},
    "matriz_riesgo": {"descripcion": "Matriz de riesgo / control interno: objetivo, descripción, valoración y niveles por periodo",
                      "por_mes": False, "matriz": True, "campos": {
                          "objetivo": ("t", True, "objetivo de control"),
                          "descripcion": ("t", True, "descripción del riesgo u opinión"),
                          "valoracion": ("t", True, "valoración del riesgo (Extremo, Alto, Moderado, Bajo)")}},
    "matriz_manuales": {"descripcion": "Matriz de manuales y actos administrativos", "por_mes": False, "matriz": True, "campos": {
        "objetivo": ("t", True, "objetivo de control"), "detalle": ("t", True, "detalle de la formalización/documentación"),
        "valoracion": ("t", True, "valoración del riesgo (Extremo, Alto, Moderado, Bajo)")}},
    "actas_consejo": {"descripcion": "Actas del consejo de administración del año", "por_mes": False, "campos": {
        "acta": ("t", True, "número del acta (o número y fecha en la misma celda)"),
        "fecha": ("f", False, "fecha del acta, si está en una columna aparte"),
        "temas": ("t", True, "temas tratados"), "firma_presidente": ("t", False, "firma del presidente"),
        "firma_administracion": ("t", False, "firma de la administración / secretario")}},
}
SECCIONES_MODELO = ("cxp", "anticipos", "deudores_varios", "retencion")       # van en extraer()
SECCIONES_DOC = ("egresos", "polizas", "matriz_riesgo", "matriz_manuales", "actas_consejo")  # en extraer_documental()


class RecetaInvalida(ValueError):
    """La receta no se puede ejecutar (columna o fila inexistente, campo desconocido…)."""


def _col(letra: str) -> int:
    try:
        return column_index_from_string(str(letra).strip().upper())
    except ValueError:
        raise RecetaInvalida(f"Columna inválida «{letra}».") from None


MES = "{mes}"


def es_mes(spec) -> bool:
    """La columna «{mes}» es la del encabezado con el mes del periodo (tablas con una columna por mes)."""
    return isinstance(spec, str) and spec.strip().lower() in ("{mes}", "mes")


def resolver_mes(ws, receta: dict, p: Periodo) -> dict:
    """Copia de la receta con cada columna «{mes}» reemplazada por la letra de la columna del mes en el encabezado.
    Si el encabezado no trae el mes, la receta no se puede aplicar ese mes (la sección queda pendiente)."""
    cols = receta.get("columnas") or {}
    if not any(es_mes(v) for v in cols.values() if not isinstance(v, list)):
        return receta
    if receta.get("titulo"):
        rx = _plantilla(receta["titulo"], p)
        t = [c for fila in ws.iter_rows() for c in fila if isinstance(c.value, str) and rx.search(norm(c.value))]
        f_enc = t[-1].row + int(receta.get("desplazamiento_encabezado", 1)) if t else None
    else:
        f_enc = int(receta["fila_encabezado"])
    c = _col_mes(ws, f_enc, p) if f_enc else None
    if not c:
        disponibles = sorted({x.clave for x in (periodo_celda(ws.cell(f_enc, k).value) for k in range(1, ws.max_column + 1)) if x}) if f_enc else []
        raise RecetaInvalida(f"La hoja «{ws.title}» no trae la columna de {p.etiqueta}"
                             + (f" (el último mes que trae es {disponibles[-1]})" if disponibles else "") + ".")
    nueva = dict(receta)
    nueva["columnas"] = {k: (get_column_letter(c) if not isinstance(v, list) and es_mes(v) else v) for k, v in cols.items()}
    return nueva


def _txt(v) -> str:
    """Mismo tratamiento de texto que el lector nativo (documental._txt), para que los resultados sean idénticos."""
    from .documental import _txt as txt_nativo
    return txt_nativo(v)


def _fecha(v) -> date | None:
    from .extractor import _fecha_parse
    if isinstance(v, str):
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return _fecha_parse(v)


def _plantilla(titulo: str, p: Periodo) -> re.Pattern:
    """«cuentas por pagar a {mes} {anio}» → regex sobre el texto normalizado (inicio de celda)."""
    t = norm(titulo)
    partes = re.split(r"(\{mes\}|\{anio\})", t)
    rx = "".join(p.nombre_mes if x == "{mes}" else str(p.anio) if x == "{anio}" else re.escape(x).replace(r"\ ", r"\s+")
                 for x in partes)
    return re.compile("^" + rx)


def validar_receta(receta: dict) -> dict:
    """Normaliza y valida la receta. Lanza RecetaInvalida."""
    if not isinstance(receta, dict):
        raise RecetaInvalida("La receta debe ser un objeto.")
    sec = receta.get("seccion")
    if sec not in CAMPOS:
        raise RecetaInvalida(f"Sección «{sec}» no admite lector declarativo. Disponibles: {', '.join(CAMPOS)}.")
    if not receta.get("hoja"):
        raise RecetaInvalida("Falta la hoja.")
    if sec in ESPECIALES:
        return _validar_especial(receta)
    if not receta.get("titulo") and not receta.get("fila_encabezado"):
        raise RecetaInvalida("Indique «titulo» (bloque del mes) o «fila_encabezado» (tabla única).")
    cols = receta.get("columnas") or {}
    campos = CAMPOS[sec]["campos"]
    desconocidos = set(cols) - set(campos)
    if desconocidos:
        raise RecetaInvalida(f"Campos desconocidos para {sec}: {', '.join(sorted(desconocidos))}.")
    faltan = [c for c, (_, ob, _) in campos.items() if ob and c not in cols and not (sec == "actas_consejo" and c == "acta" and "fecha" in cols)]
    if faltan:
        raise RecetaInvalida(f"Faltan columnas obligatorias: {', '.join(faltan)}.")
    for c, v in cols.items():
        for letra in (v if isinstance(v, list) else [v]):
            if not es_mes(letra):
                _col(letra)
    for letra in receta.get("niveles") or []:
        _col(letra)
    if receta.get("fila_encabezado") is not None:
        try:
            receta["fila_encabezado"] = int(receta["fila_encabezado"])
        except (TypeError, ValueError):
            raise RecetaInvalida("«fila_encabezado» debe ser un número de fila.") from None
    return receta


# ------------------------------------------------------------------ ejecución

def _filas(ws, receta: dict, p: Periodo):
    """(fila del título o None, fila de encabezado, filas de datos). Se detiene en una fila «TOTAL…» (que
    devuelve aparte) o, en un bloque del mes, al empezar otro bloque con el mismo título."""
    if receta.get("titulo"):
        rx = _plantilla(receta["titulo"], p)
        base = _plantilla(re.sub(r"\{mes\}.*$", "", receta["titulo"]).strip() or receta["titulo"], p)
        col_t = _col(receta.get("columna_titulo", "A")) if receta.get("columna_titulo") else None
        titulos = [c for fila in ws.iter_rows() for c in fila
                   if isinstance(c.value, str) and (col_t is None or c.column == col_t) and rx.search(norm(c.value))]
        if not titulos:
            return None, None, [], None
        t = titulos[-1]
        f_tit, f_enc = t.row, t.row + int(receta.get("desplazamiento_encabezado", 1))
    else:
        base, f_tit, f_enc = None, None, int(receta["fila_encabezado"])
    cols = sorted({_col(l) for v in receta["columnas"].values() for l in (v if isinstance(v, list) else [v])})
    filas, total = [], None
    fin = int(receta["fila_fin"]) if receta.get("fila_fin") else ws.max_row
    for r in range(f_enc + 1, fin + 1):
        textos = [norm(ws.cell(r, c).value) for c in cols if isinstance(ws.cell(r, c).value, str)]
        if any(t.startswith("total") for t in textos):
            nums = [(c, ws.cell(r, c).value) for c in range(min(cols), max(cols) + 1) if es_numero(ws.cell(r, c).value)]
            total = (r, nums[0][0], float(nums[0][1])) if nums else None
            break
        if base is not None and any(base.search(t) for t in textos) and r > f_enc:
            break                                              # empieza el bloque de otro mes
        if all(ws.cell(r, c).value in (None, "") for c in cols):
            continue
        filas.append(r)
    return f_tit, f_enc, filas, total


def _valor(ws, r: int, spec) -> object:
    if isinstance(spec, list):
        return " ".join(x for x in (_txt(ws.cell(r, _col(l)).value) for l in spec) if x)
    return ws.cell(r, _col(spec)).value


def leer(lib: Libro, receta: dict, p: Periodo, unidades: int | None = None) -> dict | list:
    """Ejecuta la receta y devuelve la sección con la forma estándar."""
    receta = validar_receta(dict(receta))
    sec = receta["seccion"]
    if sec in ESPECIALES:
        return leer_especial(lib, receta, p, unidades)
    try:
        ws = lib.hoja_mapeada(sec, receta["hoja"])
    except KeyError as e:
        raise RecetaInvalida(str(e).strip("'\"")) from None
    receta = resolver_mes(ws, receta, p)
    f_tit, f_enc, filas, total = _filas(ws, receta, p)
    C = receta["columnas"]
    hoja = ws.title

    def ref(campo, r):
        spec = C[campo]
        return f"{hoja}!{(spec[0] if isinstance(spec, list) else spec).upper()}{r}"

    def seg(campo, r):
        spec = C.get(campo)
        if not spec:
            return None
        return lib.segmentos(ws, f"{(spec[0] if isinstance(spec, list) else spec).upper()}{r}")

    def campo(r, nombre):
        return _valor(ws, r, C[nombre]) if nombre in C else None

    titulo = _txt(ws.cell(f_tit, _col(_letra_de(ws, f_tit))).value) if f_tit else ""

    if sec == "cxp":
        items = []
        for r in filas:
            v = campo(r, "valor")
            if not es_numero(v):
                continue
            items.append({"edad": _txt(campo(r, "edad")), "tercero": _txt(campo(r, "tercero")), "valor": float(v),
                          "concepto": _txt(campo(r, "concepto")), "seg_concepto": seg("concepto", r),
                          "fuente": ref("valor", r), "_fila": r})
        items.sort(key=lambda x: -x["valor"])
        return {"titulo": titulo, "items": items, "total": sum(i["valor"] for i in items),
                "total_hoja": total[2] if total else None,
                "fuente_total": f"{hoja}!{get_column_letter(total[1])}{total[0]}" if total else None}

    if sec in ("anticipos", "deudores_varios"):
        items = []
        for r in filas:
            v = campo(r, "valor")
            if not es_numero(v):
                continue
            f = _fecha(campo(r, "fecha"))
            items.append({"fecha": f.strftime("%d/%m/%Y") if f else _txt(campo(r, "fecha")),
                          "fecha_iso": f.isoformat() if f else "", "tercero": _txt(campo(r, "tercero")),
                          "valor": float(v), "detalle": _txt(campo(r, "detalle")), "seg_detalle": seg("detalle", r),
                          "fuente": ref("valor", r)})
        return {"titulo": titulo or None, "items": items, "total_hoja": total[2] if total else None,
                "fuente_total": f"{hoja}!{get_column_letter(total[1])}{total[0]}" if total else None}

    if sec == "retencion":
        out = []
        for r in filas:
            per = periodo_de(campo(r, "periodo")) or (lambda f: Periodo(f.year, f.month) if f else None)(_fecha(campo(r, "periodo")))
            if not per:
                continue
            v = campo(r, "valor")
            out.append({"periodo": per.clave, "periodo_corto": per.corto, "presentado": _fecha(campo(r, "presentado")),
                        "pagado": _fecha(campo(r, "pagado")), "valor": float(v) if es_numero(v) else 0.0,
                        "observacion": _txt(campo(r, "observacion")), "fuente": ref("periodo", r)})
        return out

    if sec == "egresos":
        from .documental import _imagenes_por_fila
        items = []
        for r in filas:
            obs = _txt(campo(r, "observacion"))
            v = campo(r, "valor")
            items.append({"ce": _txt(campo(r, "ce")), "fecha": _txt(campo(r, "fecha")), "tercero": _txt(campo(r, "tercero")),
                          "descripcion": _txt(campo(r, "descripcion")), "valor": v, "observacion": obs,
                          "seg_observacion": seg("observacion", r), "correcto": norm(obs).startswith("correcto"),
                          "fuente": ref("valor", r), "_fila": r, "imagenes": []})
        if total and items:
            for fila, ims in _imagenes_por_fila(lib, ws).items():
                if (f_tit or f_enc) <= fila <= total[0]:
                    dueño = max((i for i in items if i["_fila"] <= fila), key=lambda i: i["_fila"], default=items[0])
                    dueño["imagenes"] += ims
        suma = sum(i["valor"] or 0 for i in items if isinstance(i["valor"], (int, float)))
        etiqueta = _txt(ws.cell(total[0], _col(_letra_de(ws, total[0]))).value) if total else None
        out = {"titulo": titulo, "items": items, "total_hoja": total[2] if total else None, "suma": suma,
               "fuente_total": f"{hoja}!{get_column_letter(total[1])}{total[0]}" if total else None,
               "etiqueta_total": etiqueta, "con_observacion": sum(1 for i in items if not i["correcto"])}
        if total and not norm(etiqueta).endswith(f"{p.nombre_mes} {p.anio}"):
            out["etiqueta_inconsistente"] = etiqueta
        return out

    if sec == "polizas":
        from .documental import matriz_lineamientos
        filas_ = []
        for r in filas:
            contrato, ver = _txt(campo(r, "contrato")), _txt(campo(r, "verificacion"))
            if contrato:
                filas_.append({"n": _txt(campo(r, "n")), "contrato": contrato, "verificacion": ver,
                               "cumple": norm(ver).startswith("cumple"), "fuente": ref("verificacion", r)})
        return {"filas": filas_, "no_cumplen": [f for f in filas_ if not f["cumple"]], "matriz": matriz_lineamientos(ws)}

    if sec == "matriz_riesgo":
        niveles = [_col(l) for l in receta.get("niveles") or []]
        periodos = []
        for c in niveles:
            v = ws.cell(f_enc - 1, c).value
            if isinstance(v, (datetime, date)):
                periodos.append((c, Periodo(v.year, v.month).corto))
        out, actual = [], None
        for r in filas:
            obj, desc = _txt(campo(r, "objetivo")), _txt(campo(r, "descripcion"))
            if obj:
                actual = {"objetivo": obj, "descripcion": [], "valoracion": _txt(campo(r, "valoracion")),
                          "niveles": {pc: _txt(ws.cell(r, c).value).upper() for c, pc in periodos}, "fuente": ref("objetivo", r)}
                out.append(actual)
            if actual and desc:
                actual["descripcion"].append({"texto": desc, "seg": seg("descripcion", r)})
            if actual and not obj:
                for c, pc in periodos:
                    v = _txt(ws.cell(r, c).value).upper()
                    if v and not actual["niveles"].get(pc):
                        actual["niveles"][pc] = v
        return {"titulo": _txt(ws.cell(f_enc - 1, _col(_primera(C["objetivo"]))).value),
                "periodos": [pc for _, pc in periodos], "filas": out}

    if sec == "matriz_manuales":
        niveles = [_col(l) for l in receta.get("niveles") or []]
        out = []
        for r in filas:
            obj = _txt(campo(r, "objetivo"))
            det = _txt(campo(r, "detalle"))
            if not obj and not det:
                continue
            out.append({"objetivo": obj, "detalle": det, "valoracion": _txt(campo(r, "valoracion")),
                        "niveles": [_txt(ws.cell(r, c).value) for c in niveles], "fuente": ref("objetivo", r)})
        return {"filas": out}

    if sec == "actas_consejo":
        from .documental import actas_del_anio, parse_acta
        actas = []
        for r in filas:
            a, temas = campo(r, "acta"), _txt(campo(r, "temas"))
            if a in (None, "") or not temas or norm(a).startswith(("n. acta", "relacion", "no acta")):
                continue
            numero, fecha_txt, anio = parse_acta(a, campo(r, "fecha"))
            actas.append({"numero": numero, "fecha": fecha_txt, "anio": anio, "temas": temas,
                          "firma_presidente": _txt(campo(r, "firma_presidente")),
                          "firma_administracion": _txt(campo(r, "firma_administracion")),
                          "fuente": ref("acta", r), "_fila": r})
        return actas_del_anio(actas, p)

    raise RecetaInvalida(f"Sección «{sec}» sin lector.")


def _primera(spec) -> str:
    return spec[0] if isinstance(spec, list) else spec


def _letra_de(ws, r: int) -> str:
    """Columna de la primera celda con texto de la fila (donde suele estar el título o la etiqueta «TOTAL…»)."""
    for c in range(1, min(ws.max_column, 30) + 1):
        if isinstance(ws.cell(r, c).value, str) and ws.cell(r, c).value.strip():
            return get_column_letter(c)
    return "A"


# ------------------------------------------------------------------ prueba de una receta

def probar(lib: Libro, receta: dict, p: Periodo) -> dict:
    """Ejecuta la receta y revisa que lo leído tenga sentido. `errores` impiden aprobarla; `avisos` informan."""
    from .esquema import NIVELES_RIESGO
    try:
        res = leer(lib, receta, p)
    except RecetaInvalida as e:
        return {"ok": False, "errores": [str(e)], "avisos": [], "n": 0, "muestra": []}
    sec = receta["seccion"]
    if sec in ESPECIALES:
        return _probar_especial(sec, res, receta)
    items = (res if isinstance(res, list) else res.get("items") or res.get("filas") or res.get("actas") or [])
    errores, avisos = [], []
    if not items:
        errores.append("La receta no leyó ninguna fila.")
    campos = CAMPOS[sec]["campos"]
    ws = lib.hoja_mapeada(sec, receta["hoja"])
    con_mes = any(es_mes(v) for v in (receta.get("columnas") or {}).values() if not isinstance(v, list))
    receta = resolver_mes(ws, receta, p)
    cols = receta.get("columnas", {})
    _, f_enc, filas, total = _filas(ws, receta, p)
    for c, (tipo, ob, _) in campos.items():
        if c not in cols or not filas:
            continue
        vals = [_valor(ws, r, cols[c]) for r in filas]
        llenos = [v for v in vals if v not in (None, "")]
        if items and isinstance(items[0], dict) and c in items[0]:
            # obligatorio: se mide sobre lo LEÍDO (en una matriz la descripción sigue en filas de continuación)
            vals = [i.get(c) for i in items]
            llenos_ob = [v for v in vals if v not in (None, "", [])]
        else:
            llenos_ob = llenos
        if ob and len(llenos_ob) < 0.3 * len(vals):     # casi vacío: la columna es la equivocada
            errores.append(f"«{c}» (col. {cols[c]}) es obligatorio y está vacío en {len(vals) - len(llenos_ob)} de "
                           f"{len(vals)} filas: probablemente no es la columna correcta.")
        elif ob and len(llenos_ob) < 0.6 * len(vals):   # parcialmente vacío: dato faltante en el Excel del cliente
            avisos.append(f"«{c}» (col. {cols[c]}) viene vacío en {len(vals) - len(llenos_ob)} de {len(vals)} filas "
                          "del Excel: es información que falta en los papeles de trabajo.")
        if tipo == "n" and llenos and sum(es_numero(v) for v in llenos) < 0.8 * len(llenos):
            errores.append(f"«{c}» (col. {cols[c]}) debería ser numérico y no lo es en la mayoría de filas.")
        if tipo in ("f", "p") and llenos and sum(bool(_fecha(v) or periodo_de(v)) for v in llenos) < 0.7 * len(llenos):
            (errores if ob else avisos).append(f"«{c}» (col. {cols[c]}) debería ser fecha y no lo es en la mayoría de filas.")
        if tipo == "t" and c not in ("n", "ce", "acta") and llenos and sum(es_numero(v) for v in llenos) > 0.8 * len(llenos):
            avisos.append(f"«{c}» (col. {cols[c]}) trae números; ¿es la columna correcta?")
    if CAMPOS[sec].get("matriz"):
        vals = [norm(i.get("valoracion")) for i in items if i.get("valoracion")]
        if vals and sum(v in NIVELES_RIESGO for v in vals) / len(vals) < 0.6:
            errores.append("«valoracion» no trae niveles (Extremo/Alto/Moderado/Bajo) en la mayoría de filas.")
    if isinstance(res, dict) and res.get("total_hoja") is not None:
        suma = res.get("total") if sec == "cxp" else res.get("suma") if sec == "egresos" else sum(i["valor"] for i in items)
        if suma is not None and abs(suma - res["total_hoja"]) > 1:
            avisos.append(f"La suma leída ({suma:,.0f}) no coincide con el total de la hoja ({res['total_hoja']:,.0f}).")
    if CAMPOS[sec].get("por_mes") and not receta.get("titulo") and not con_mes:
        avisos.append("La sección es mensual pero la receta usa una fila fija: el próximo mes puede no coincidir.")
    enc = {c: _txt(ws.cell(f_enc, _col(_primera(v))).value) for c, v in cols.items()} if f_enc else {}
    return {"ok": not errores, "errores": errores, "avisos": avisos, "n": len(items), "encabezados": enc,
            "muestra": _muestra(items)}


def _probar_especial(sec: str, res: dict, receta: dict) -> dict:
    errores, avisos = [], []
    if sec == "liquidez":
        items = res["filas"]
        if not items:
            errores.append("No se leyó ningún concepto entre el encabezado y el total.")
        if not res.get("total"):
            errores.append("El total del mes quedó en cero o vacío: revise la fila total.")
        if res.get("total_cierre") is None:
            avisos.append("No se encontró la columna de diciembre del año anterior (comparativo de cierre).")
        if len(res.get("serie") or []) < 2:
            avisos.append("La serie mensual tiene menos de dos meses: el gráfico de evolución saldrá pobre.")
        muestra = [{"concepto": f["concepto"], "actual": f["actual"], "cierre": f["referencia"], "fuente": f["fuente"]} for f in items]
    elif sec == "cartera":
        items = res["items"]
        if not items:
            errores.append("No se leyó ningún inmueble con saldo.")
        if items and len({i["apto"] for i in items}) < 0.8 * len(items):
            errores.append("Muchos inmuebles repetidos: la columna de apartamento parece equivocada.")
        th = res.get("total_hoja_lista")
        if th is not None and abs(th - res["total"]) > 1:
            avisos.append(f"La suma de inmuebles ({res['total']:,.0f}) no coincide con el total de la lista ({th:,.0f}).")
        if res.get("anterior") is None:
            avisos.append("No se encontró la cartera del mes anterior (comparativo).")
        muestra = [{"apto": i["apto"], "valor": i["valor"], "fuente": i["fuente"]} for i in items]
    else:
        items = res["rubros"]
        if not items:
            errores.append("No se leyó ningún rubro entre el encabezado y la fila total.")
        if items and sum(1 for r in items if r["ejecucion"] or r["presupuesto_periodo"]) < 0.6 * len(items):
            errores.append("La mayoría de rubros no trae presupuesto ni ejecución numéricos: revise las columnas.")
        th = res.get("total_hoja")
        if th is not None and abs(th - res["total_calculado"]) > 1:
            msg = f"La suma de diferencias ({res['total_calculado']:,.0f}) no coincide con el total de la hoja ({th:,.0f})"
            avisos.append(msg + (": pruebe signo_total = " + ("1" if receta.get("signo_total", -1) == -1 else "-1") + "."
                                 if abs(-th - res["total_calculado"]) <= 1 else "."))
        muestra = [{"rubro": r["rubro"], "presupuesto": r["presupuesto_periodo"], "ejecucion": r["ejecucion"],
                    "sobre_ejecutado": r["sobre_ejecutado"], "fuente": r["fuente"]} for r in items]
    return {"ok": not errores, "errores": errores, "avisos": avisos, "n": len(items), "muestra": muestra[:5]}


def _muestra(items, n=5):
    def limpio(i):
        return {k: (v if not isinstance(v, (date, datetime)) else v.isoformat()) for k, v in i.items()
                if not k.startswith(("seg", "_", "imagenes")) and k != "descripcion"} | (
            {"descripcion": " / ".join(d["texto"] for d in i["descripcion"])[:200]} if isinstance(i.get("descripcion"), list) else {})
    return [limpio(i) for i in items[:n]]


# ================================================================== secciones con columnas por mes
# Liquidez, cartera y presupuesto no son una tabla de filas simples: la receta lleva, además de columnas, las
# ETIQUETAS de las filas clave (encabezado de meses, total) y la columna del mes se busca por su encabezado.

ESPECIALES = {
    "liquidez": {
        "descripcion": "Análisis de liquidez: conceptos (filas) por mes (columnas), con fila total; opcionalmente monetización del fondo",
        "por_mes": False, "campos": {"concepto": ("t", True, "columna con el nombre de cada concepto («(+) Saldo cuenta corriente…»)")},
        "parametros": {
            "encabezado": (True, "texto de la fila que trae los meses como encabezado (p. ej. «análisis de liquidez»)"),
            "total": (True, "texto de la fila total (p. ej. «total equivalente al efectivo»)"),
            "monetizacion": (False, "texto de la fila que abre el bloque de monetización del fondo (si existe)"),
            "mayor_valor": (False, "texto de la fila «mayor valor monetizado» dentro de ese bloque")}},
    "cartera": {
        "descripcion": "Cartera por inmueble del mes (apartamento y saldo), para rangos y concentración",
        "por_mes": True, "campos": {
            "apto": ("t", True, "columna del inmueble/apartamento; con «titulo» se puede dar relativa al título: +1"),
            "valor": ("n", False, "columna del saldo; con «titulo» relativa (+2). Sin titulo se toma la columna cuyo encabezado es el mes")},
        "parametros": {
            "resumen": (False, "texto de la fila del resumen que trae la hoja (p. ej. «total cartera a»)"),
            "unidades": (False, "texto de la celda con el total de inmuebles (p. ej. «total todos los inmuebles»)")}},
    "presupuesto": {
        "descripcion": "Rubros sobreejecutados: presupuesto anual, del periodo, ejecución, diferencia y observación",
        "por_mes": False, "campos": {
            "rubro": ("t", True, "nombre del rubro"), "presupuesto_anual": ("n", False, "presupuesto del año"),
            "presupuesto_periodo": ("n", True, "presupuesto acumulado al periodo"),
            "ejecucion": ("n", True, "ejecución acumulada al periodo"),
            "diferencia": ("n", False, "columna de la diferencia (para el total de la hoja)"),
            "observacion": ("t", False, "observación del revisor")},
        "parametros": {
            "total": (False, "texto de la fila total (p. ej. «total rubros»); sin ella se leen los rubros hasta la primera fila vacía"),
            "signo_total": (False, "-1 si la diferencia de la hoja es presupuesto − ejecución (negativa cuando hay sobreejecución); 1 si es al revés")}},
}
CAMPOS.update({k: {kk: vv for kk, vv in v.items() if kk != "parametros"} | {"parametros": v["parametros"]}
               for k, v in ESPECIALES.items()})
SECCIONES_MODELO = SECCIONES_MODELO + tuple(ESPECIALES)

_MESES_TXT = {m: i + 1 for i, m in enumerate(["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"])}


def periodo_celda(v) -> Periodo | None:
    """Mes de un encabezado: fecha de Excel o texto («ago-26», «agosto 2026», «08/2026», «AGOSTO DE 2026»)."""
    p = periodo_de(v)
    if p or not isinstance(v, str):
        return p
    t = norm(v)
    m = re.fullmatch(r"([a-z]{3})[a-z]*\.?(?:\s*de)?[\s\-/]*(\d{2}|\d{4})", t)
    if m and m.group(1) in _MESES_TXT:
        a = int(m.group(2))
        return Periodo(a + 2000 if a < 100 else a, _MESES_TXT[m.group(1)])
    m = re.fullmatch(r"(\d{1,2})[/\-](\d{4})", t)
    if m and 1 <= int(m.group(1)) <= 12:
        return Periodo(int(m.group(2)), int(m.group(1)))
    return None


def _col_mes(ws, fila: int, p: Periodo, fila_dato: int | None = None, desde: int = 1, hasta: int | None = None):
    """Como excel.columna_mes, pero reconoce también meses escritos como texto."""
    cands = [c for c in range(desde, (hasta or ws.max_column) + 1) if periodo_celda(ws.cell(fila, c).value) == p]
    if not cands:
        return None
    if fila_dato:
        for c in cands:
            v = ws.cell(fila_dato, c).value
            if es_numero(v) and v != 0:
                return c
        for c in cands:
            if es_numero(ws.cell(fila_dato, c).value):
                return c
    return cands[0]


def _fila_etiqueta(ws, etiqueta: str, col: int, desde: int = 1) -> int | None:
    objetivo = norm(etiqueta)
    for r in range(desde, ws.max_row + 1):
        if norm(ws.cell(r, col).value).startswith(objetivo):
            return r
    return None


def _col_rel(spec, base: int | None) -> int:
    """«D» absoluta, o «+2» relativa a la columna del título."""
    s = str(spec).strip()
    if s.startswith("+"):
        if base is None:
            raise RecetaInvalida(f"La columna «{s}» es relativa al título, pero la receta no tiene «titulo».")
        return base + int(s[1:])
    return _col(s)


def _validar_especial(receta: dict) -> dict:
    sec = receta["seccion"]
    d = ESPECIALES[sec]
    cols = receta.get("columnas") or {}
    desconocidos = set(cols) - set(d["campos"])
    if desconocidos:
        raise RecetaInvalida(f"Campos desconocidos para {sec}: {', '.join(sorted(desconocidos))}.")
    faltan = [c for c, (_, ob, _) in d["campos"].items() if ob and c not in cols]
    faltan += [k for k, (ob, _) in d["parametros"].items() if ob and not receta.get(k)]
    if sec == "cartera" and not receta.get("titulo") and not receta.get("fila_encabezado"):
        faltan.append("titulo o fila_encabezado")
    if faltan:
        raise RecetaInvalida(f"Faltan: {', '.join(faltan)}.")
    for v in cols.values():
        if not str(v).strip().startswith("+"):
            _col(v)
    return receta


def leer_especial(lib: Libro, receta: dict, p: Periodo, unidades: int | None = None):
    from .extractor import _bloque_conceptos, _f
    from .excel import coord, num
    sec = receta["seccion"]
    try:
        ws = lib.hoja_mapeada(sec, receta["hoja"])
    except KeyError as e:
        raise RecetaInvalida(str(e).strip("'\"")) from None
    C = receta.get("columnas") or {}

    if sec == "liquidez":
        c_lbl = _col(C["concepto"])
        f_enc = int(receta["fila_encabezado"]) if receta.get("fila_encabezado") else _fila_etiqueta(ws, receta["encabezado"], c_lbl)
        if not f_enc:
            raise RecetaInvalida(f"No se encontró la fila «{receta.get('encabezado')}» en la columna {C['concepto']}.")
        f_tot = _fila_etiqueta(ws, receta["total"], c_lbl, f_enc)
        if not f_tot:
            raise RecetaInvalida(f"No se encontró la fila total «{receta['total']}» debajo del encabezado.")
        col = _col_mes(ws, f_enc, p, f_tot)
        if not col:
            raise RecetaInvalida(f"La fila {f_enc} no trae una columna de {p.etiqueta}.")
        col_cierre = _col_mes(ws, f_enc, p.cierre_anterior(), f_tot)
        filas = _bloque_conceptos(ws, f_enc + 1, f_tot, col, col_cierre, c_lbl)
        serie, vistos = [], set()
        for c in range(1, ws.max_column + 1):
            per = periodo_celda(ws.cell(f_enc, c).value)
            if not per or per in vistos:
                continue
            vistos.add(per)
            v = ws.cell(f_tot, c).value
            if es_numero(v) and p.cierre_anterior().clave <= per.clave <= p.clave:
                serie.append({"periodo": per.clave, "valor": float(v)})
        serie.sort(key=lambda x: x["periodo"])
        f_mon = _fila_etiqueta(ws, receta["monetizacion"], c_lbl) if receta.get("monetizacion") else None
        etiqueta_mv = norm(receta.get("mayor_valor") or "mayor valor")
        f_mv = None
        for r in range(f_mon + 1, f_mon + 12) if f_mon else ():
            if norm(ws.cell(r, c_lbl).value).startswith(etiqueta_mv) and es_numero(ws.cell(r, col).value):
                f_mv = r
        c_mon = (_col_mes(ws, f_mon, p, f_mv) or col) if f_mv else None
        c_mon_cierre = (_col_mes(ws, f_mon, p.cierre_anterior(), f_mv) or col_cierre) if f_mv else None
        monet = []
        for r in _bloque_conceptos(ws, f_mon + 1, f_mv, c_mon, c_mon_cierre, c_lbl) if f_mv else ():
            if norm(r["concepto"]).startswith(etiqueta_mv):
                if r["fuente"] and r["fuente"].endswith(str(f_mv)):
                    r["es_total"] = True
                    monet.append(r)
                continue
            monet.append(r)
        return {"filas": filas, "total": num(ws.cell(f_tot, col).value),
                "total_cierre": num(ws.cell(f_tot, col_cierre).value) if col_cierre else None,
                "fuente_total": _f(ws, f_tot, col), "serie": serie, "monetizacion": monet,
                "mayor_valor_monetizado": num(ws.cell(f_mv, c_mon).value) if f_mv else None,
                "mayor_valor_monetizado_cierre": num(ws.cell(f_mv, c_mon_cierre).value) if f_mv and c_mon_cierre else None,
                "_hoja": ws.title, "_col": col}

    if sec == "cartera":
        def lista(per: Periodo):
            if receta.get("titulo"):
                rx = _plantilla(receta["titulo"], per)
                t = next((c for fila in ws.iter_rows() for c in fila if isinstance(c.value, str)
                          and rx.search(norm(c.value)) and not norm(c.value).startswith("total")), None)
                if t is None:
                    return None
                col0, f_enc = t.column, t.row + 1
                c_apto, c_val = _col_rel(C["apto"], col0), _col_rel(C.get("valor", "+2"), col0)
            else:
                f_enc = int(receta["fila_encabezado"])
                c_apto = _col(C["apto"])
                col0 = c_apto
                c_val = _col(C["valor"]) if C.get("valor") and per == p else _col_mes(ws, f_enc, per, desde=c_apto + 1, hasta=c_apto + 8)
                if not c_val:
                    return None
            items, total_hoja, fila_total = [], None, None
            for r in range(f_enc + 1, f_enc + 500):
                apto, val = ws.cell(r, c_apto).value, ws.cell(r, c_val).value
                if norm(ws.cell(r, col0).value).startswith("total") or norm(apto).startswith("total") \
                        or norm(ws.cell(r, c_apto + 1).value).startswith("total"):
                    total_hoja, fila_total = val, r
                    break
                if apto is None and val is None:
                    break
                if es_numero(val):
                    items.append({"apto": str(apto).strip(), "valor": float(val), "fila": r,
                                  "fuente": f"{ws.title}!{coord(r, c_val)}"})
            return {"items": items, "total_hoja": num(total_hoja) if es_numero(total_hoja) else None,
                    "fuente_total": f"{ws.title}!{coord(fila_total, c_val)}" if fila_total else None}

        actual, anterior = lista(p), lista(p.anterior())
        if actual is None:
            raise RecetaInvalida(f"La hoja «{ws.title}» no trae la cartera de {p.etiqueta} con esta receta.")
        items = sorted(actual["items"], key=lambda x: -x["valor"])
        total = sum(i["valor"] for i in items)
        tramos = [("Mayores a $3.000.000", lambda v: v > 3_000_000),
                  ("Entre $1.000.001 y $3.000.000", lambda v: 1_000_000 < v <= 3_000_000),
                  ("Hasta $1.000.000", lambda v: v <= 1_000_000)]
        rangos = []
        for nombre, cond in tramos:
            sel = [i for i in items if cond(i["valor"])]
            s = sum(i["valor"] for i in sel)
            rangos.append({"rango": nombre, "valor": s, "inmuebles": len(sel), "pct_cartera": s / total if total else 0,
                           "pct_unidades": len(sel) / unidades if unidades else None})
        resumen_hoja = None
        etq = norm(receta.get("resumen") or "total cartera a")
        for fila in ws.iter_rows():
            c = next((x for x in fila if isinstance(x.value, str) and norm(x.value).startswith(etq)), None)
            if c is not None:
                nums = [cc for cc in range(c.column + 1, c.column + 6) if es_numero(ws.cell(c.row, cc).value)]
                if nums:
                    resumen_hoja = {"total": num(ws.cell(c.row, nums[0]).value),
                                    "inmuebles": num(ws.cell(c.row, nums[1]).value) if len(nums) > 1 else 0.0,
                                    "fuente": f"{ws.title}!{coord(c.row, nums[0])}"}
                break
        unidades_hoja = None
        etu = norm(receta.get("unidades") or "total todos los inmuebles")
        for fila in ws.iter_rows():
            c = next((x for x in fila if isinstance(x.value, str) and norm(x.value).startswith(etu)), None)
            if c is not None:
                for dc in range(1, 8):
                    if es_numero(ws.cell(c.row, c.column + dc).value):
                        unidades_hoja = int(ws.cell(c.row, c.column + dc).value)
                        break
                break
        return {"items": items, "total": total, "inmuebles": len(items), "rangos": rangos,
                "total_hoja_lista": actual["total_hoja"], "fuente_total": actual["fuente_total"],
                "resumen_hoja": resumen_hoja, "unidades_hoja": unidades_hoja,
                "anterior": {"total": sum(i["valor"] for i in anterior["items"]), "inmuebles": len(anterior["items"])}
                if anterior else None,
                "_items_anterior": anterior}             # para reconstruir el comparativo al migrar a la plantilla

    if sec == "presupuesto":
        c_r = _col(C["rubro"])
        f_enc = int(receta["fila_encabezado"]) if receta.get("fila_encabezado") else next(
            (r for r in range(1, ws.max_row + 1) if norm(ws.cell(r, c_r).value) == "rubro"), None) or _fila_etiqueta(ws, "rubro", c_r)
        if not f_enc:
            raise RecetaInvalida("No se encontró la fila de encabezado («Rubro»): indique fila_encabezado.")
        if receta.get("total"):
            f_tot = _fila_etiqueta(ws, receta["total"], c_r, f_enc)
            if not f_tot:
                raise RecetaInvalida(f"No se encontró la fila total «{receta['total']}».")
        else:                                    # hoja sin fila total: hasta la primera fila sin rubro
            f_tot = next((r for r in range(f_enc + 1, ws.max_row + 2) if not ws.cell(r, c_r).value), ws.max_row + 1)
        con_total = bool(receta.get("total"))

        def n(campo, r):
            return num(ws.cell(r, _col(C[campo])).value) if C.get(campo) else 0.0
        rubros = []
        for r in range(f_enc + 1, f_tot):
            nombre = ws.cell(r, c_r).value
            if not nombre:
                continue
            ejec, pres = n("ejecucion", r), n("presupuesto_periodo", r)
            rubros.append({"rubro": str(nombre).strip(), "presupuesto_anual": n("presupuesto_anual", r),
                           "presupuesto_periodo": pres, "ejecucion": ejec, "sobre_ejecutado": ejec - pres,
                           "observacion": str(ws.cell(r, _col(C["observacion"])).value or "").strip() if C.get("observacion") else "",
                           "fuente": f"{ws.title}!{C['ejecucion'].upper()}{r}", "_fila": r})
        por_fila = {}
        for im in getattr(lib.hoja_formulas(ws), "_images", []):
            try:
                por_fila.setdefault(im.anchor._from.row + 1, []).append({"bytes": im._data(), "ancho": im.width, "alto": im.height})
            except Exception:
                continue
        for r in rubros:
            r["soportes"] = por_fila.get(r["_fila"], [])
        rubros.sort(key=lambda x: -x["sobre_ejecutado"])
        c_dif = C.get("diferencia") if con_total else None
        signo = float(receta.get("signo_total", -1))
        return {"rubros": rubros, "total_calculado": sum(r["sobre_ejecutado"] for r in rubros),
                "total_hoja": signo * num(ws.cell(f_tot, _col(c_dif)).value) if c_dif else None,
                "fuente_total": f"{ws.title}!{c_dif.upper()}{f_tot}" if c_dif else None,
                "formula_total": lib.formula(ws, f"{c_dif.upper()}{f_tot}") if c_dif else None,
                "encabezado_periodo": str(ws.cell(f_enc, _col(C["presupuesto_periodo"])).value or "").strip()}

    raise RecetaInvalida(f"Sección «{sec}» sin lector especial.")
