"""Validador de inconsistencias del Excel de papeles de trabajo.

Cada regla devuelve Hallazgos con severidad, ubicación exacta (hoja!celda) y
recomendación. El revisor fiscal decide; el validador solo señala.
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import asdict, dataclass, field
from datetime import date

from .excel import ERRORES_EXCEL, Libro, Periodo, coord, es_numero, norm

ORDEN = {"ALTA": 0, "MEDIA": 1, "BAJA": 2, "INFO": 3}
HOJAS_FINANCIERAS = {"cuadro efectivo", "grafica flujo de efectivo", "grafica deudores propietari",
                     "rangos de cartera", "may deudores", "antic. e ingresos x c", "cxp", "cuadro reserva",
                     "patrimonio", "presupuesto", "conciliaciones", "impuestos"}
_BASE_FINANCIERAS = ("cuadro efectivo", "grafica flujo de efectivo", "grafica deudores", "rangos de cartera",
                     "may deudores", "antic. e ingresos", "cxp", "cuadro reserva", "patrimonio", "presupuesto",
                     "conciliaciones", "impuestos")


def es_hoja_financiera(titulo: str) -> bool:
    """Reconoce la hoja con la misma tolerancia del resto del sistema («GRAFICA DE DEUDORES PROPIETARIO»,
    «GRAFICA DEUDORES» de la plantilla estándar…), no solo por el nombre exacto del informe modelo."""
    from .excel import buscar_hoja
    t = norm(titulo)
    return t in HOJAS_FINANCIERAS or any(buscar_hoja([titulo], b) for b in _BASE_FINANCIERAS) \
        and not t.startswith("anexo")


@dataclass
class Hallazgo:
    regla: str
    severidad: str
    categoria: str
    titulo: str
    detalle: str
    ubicacion: list[str] = field(default_factory=list)
    recomendacion: str = ""
    origen: str = "excel"           # excel | pptx

    def dict(self):
        return asdict(self)


def pesos(v: float) -> str:
    from decimal import ROUND_HALF_UP, Decimal
    s = f"{int(Decimal(str(abs(v))).quantize(Decimal('1'), rounding=ROUND_HALF_UP)):,}".replace(",", ".")
    return f"-${s}" if v < 0 else f"${s}"


# ------------------------------------------------------------ reglas libro

def _periodo_columna(ws, col: int):
    """Mes del encabezado de una columna (busca una fecha en las primeras filas)."""
    from .excel import periodo_de
    for r in range(1, 12):
        per = periodo_de(ws.cell(r, col).value)
        if per:
            return per
    return None


def r_errores_formula(lib: Libro, p: Periodo | None = None) -> list[Hallazgo]:
    out = []
    for ws in lib.valores.worksheets:
        celdas = [c for fila in ws.iter_rows() for c in fila
                  if isinstance(c.value, str) and c.value.strip() in ERRORES_EXCEL]
        if not celdas:
            continue
        fin = es_hoja_financiera(ws.title)
        del_periodo = [c for c in celdas if p and _periodo_columna(ws, c.column) == p]
        historicas = all(_periodo_columna(ws, c.column) not in (None, p) for c in celdas)
        if del_periodo:
            sev, nota = ("ALTA" if fin else "MEDIA"), f" {len(del_periodo)} están en la columna del periodo."
        elif historicas:
            sev, nota = "BAJA" if not fin else "MEDIA", " Todas están en columnas de meses anteriores."
        else:
            sev, nota = ("MEDIA" if fin else "BAJA"), ""
        out.append(Hallazgo(
            "V01", sev, "Integridad del Excel",
            f"{len(celdas)} celda(s) con error de fórmula en '{ws.title.strip()}'",
            "Las celdas devuelven #REF! u otro error." + nota +
            " Si alguna alimenta el informe, la cifra publicada puede estar incompleta.",
            [f"{ws.title.strip()}!{c.coordinate}" for c in celdas[:12]] + (["…"] if len(celdas) > 12 else []),
            "Reconstruir las referencias rotas o eliminar las columnas históricas que ya no se usan."))
    return out


def r_formulas_sin_resultado(lib: Libro) -> list[Hallazgo]:
    faltan = [(h, c, f) for h, c, f in lib.formulas_sin_resultado() if es_hoja_financiera(h)]
    if not faltan:
        return []
    return [Hallazgo("V21", "ALTA", "Integridad del Excel",
                      f"{len(faltan)} fórmula(s) sin resultado calculado",
                      "El archivo contiene fórmulas, pero no guardó su resultado. Esas celdas no pueden tratarse como cero.",
                      [f"{h}!{c}" for h, c, _ in faltan[:20]] + (["…"] if len(faltan) > 20 else []),
                      "Abra y recalcule el archivo en Excel o LibreOffice, guárdelo y vuelva a cargarlo.")]


RX_SUM = re.compile(r"^=\+?SUM\(\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\)$", re.I)


def r_rangos_suma_incompletos(lib: Libro) -> list[Hallazgo]:
    """=SUM(F4:F10) en F13 cuando F11 y F12 también tienen valores del mismo bloque."""
    out = []
    for wf in lib.formulas.worksheets:
        if not es_hoja_financiera(wf.title):
            continue
        wv = lib.valores[wf.title]
        for fila in wf.iter_rows():
            for c in fila:
                if not isinstance(c.value, str):
                    continue
                m = RX_SUM.match(c.value.replace(" ", ""))
                if not m or m.group(1) != m.group(3) or c.column_letter != m.group(1):
                    continue
                fin_rango, r = int(m.group(4)), c.row
                if r <= fin_rango + 1:
                    continue
                omitidas = []
                for rr in range(fin_rango + 1, r):
                    v = wv[f"{m.group(1)}{rr}"].value
                    etiqueta = wv[f"B{rr}"].value or wv[f"A{rr}"].value
                    if es_numero(v) and v != 0 and etiqueta:
                        omitidas.append((rr, v, str(etiqueta).strip()))
                if omitidas:
                    det = "; ".join(f"fila {rr} «{e[:40]}» {pesos(v)}" for rr, v, e in omitidas)
                    out.append(Hallazgo(
                        "V02", "ALTA", "Totales",
                        f"Total que no incluye todas las filas del bloque ({wf.title.strip()}!{c.coordinate})",
                        f"La fórmula {c.value} deja por fuera: {det}. "
                        f"Diferencia: {pesos(sum(v for _, v, _ in omitidas))}.",
                        [f"{wf.title.strip()}!{c.coordinate}"] +
                        [f"{wf.title.strip()}!{m.group(1)}{rr}" for rr, _, _ in omitidas],
                        "Ampliar el rango de la suma o convertir el bloque en una tabla de Excel."))
    return out


RX_CONST = re.compile(r"^=\+?[\d\s\.\+\-\*/\(\)]+$")


def r_valores_digitados(lib: Libro) -> list[Hallazgo]:
    out = []
    for wf in lib.formulas.worksheets:
        if not es_hoja_financiera(wf.title):
            continue
        celdas = []
        for fila in wf.iter_rows():
            for c in fila:
                if isinstance(c.value, str) and RX_CONST.match(c.value):
                    nums = re.findall(r"\d{5,}", c.value)
                    if len(nums) >= 2:
                        celdas.append(f"{wf.title.strip()}!{c.coordinate} {c.value[:60]}")
        if celdas:
            out.append(Hallazgo(
                "V03", "MEDIA", "Trazabilidad",
                f"{len(celdas)} fórmula(s) con cifras digitadas a mano en '{wf.title.strip()}'",
                "Son fórmulas que suman o restan valores tecleados, no referencias a celdas. "
                "Al cambiar de mes no se actualizan solas y es la causa típica de cifras arrastradas.",
                celdas[:8], "Reemplazar por referencias a las celdas de origen."))
    return out


# ----------------------------------------------------------- reglas modelo

def r_cruces(m: dict) -> list[Hallazgo]:
    """Cruces de la misma cifra entre hojas."""
    out = []
    # una sección que no se pudo leer llega como None: sus cruces no se hacen (la cobertura ya la marca)
    P = m.get("patrimonio") or {}
    disp = {norm(d["concepto"])[:20]: d for d in P.get("disponibilidad") or []}

    def dget(prefijo):
        for k, d in disp.items():
            if k.startswith(norm(prefijo)[:20]):
                return d
        return None

    L, D = m.get("liquidez") or {}, m.get("deudores") or {}
    A, DV = m.get("anticipos") or {}, m.get("deudores_varios") or {}
    cruces = [
        ("Disponible del equivalente", L.get("total"), L.get("fuente_total")),
        ("Cartera de propietarios", D.get("neto"), D.get("fuente_neto")),
        ("Anticipos a contratistas", A.get("total_hoja"), A.get("fuente_total")),
        ("Deudores Varios", DV.get("total_hoja"), DV.get("fuente_total")),
    ]
    for etiqueta, esperado, fuente in cruces:
        d = dget(etiqueta)
        if not d or esperado is None:
            continue
        dif = d["actual"] - esperado
        if abs(dif) > 1:
            extra = ""
            if d.get("formula") and RX_CONST.match(d["formula"]):
                extra = f" La celda usa cifras digitadas: {d['formula'][:70]}."
            out.append(Hallazgo(
                "V04", "ALTA", "Cruce entre hojas",
                f"«{d['concepto'][:45]}» no coincide con su hoja de origen",
                f"PATRIMONIO muestra {pesos(d['actual'])} y la hoja de origen {pesos(esperado)} "
                f"(diferencia {pesos(dif)}).{extra}",
                [d["fuente"], fuente], "Enlazar la celda a la hoja de origen."))

    dif = (P["total_disponibilidad"] - P["total_excedentes"]) if P.get("total_disponibilidad") is not None \
        and P.get("total_excedentes") is not None else 0
    if abs(dif) > 1:
        explicacion = ""
        difs_v04 = [h for h in out if h.regla == "V04"]
        for h in difs_v04:
            mm = re.search(r"diferencia (-?\$[\d\.]+)", h.detalle)
            if mm and mm.group(1).replace("-", "") == pesos(dif).replace("-", ""):
                explicacion = f" La diferencia es igual a la del cruce «{h.titulo[1:40]}…» (ver V04)."
        out.append(Hallazgo(
            "V05", "ALTA", "Excedentes",
            "La disponibilidad de excedentes no cuadra con los excedentes acumulados",
            f"Excedentes acumulados {pesos(P['total_excedentes'])} vs. activos que los respaldan "
            f"{pesos(P['total_disponibilidad'])} (diferencia {pesos(dif)}). La conclusión de que el 100% "
            f"está respaldado debe revisarse.{explicacion}",
            [P["fuente_excedentes"], P["fuente_disponibilidad"]],
            "Identificar la partida que explica la diferencia antes de emitir la conclusión."))

    C = m.get("cartera")
    if not C:
        return out
    if C["resumen_hoja"] and (abs(C["resumen_hoja"]["total"] - C["total"]) > 1
                              or int(C["resumen_hoja"]["inmuebles"]) != C["inmuebles"]):
        out.append(Hallazgo("V06", "ALTA", "Cartera", "El resumen de rangos no coincide con el detalle",
                            f"Resumen: {pesos(C['resumen_hoja']['total'])} / {int(C['resumen_hoja']['inmuebles'])} "
                            f"inmuebles. Detalle: {pesos(C['total'])} / {C['inmuebles']} inmuebles.",
                            [C["resumen_hoja"]["fuente"], C["fuente_total"]]))
    if D and abs(C["total"] - D["bruto"]) > 1:
        out.append(Hallazgo("V06", "ALTA", "Cartera", "El detalle de cartera no coincide con deudores copropietarios",
                            f"Detalle por inmueble {pesos(C['total'])} vs. saldo contable {pesos(m['deudores']['bruto'])}.",
                            [C["fuente_total"]]))
    longitudes = Counter(len(i["apto"]) for i in C["items"])
    if longitudes:
        tipica = longitudes.most_common(1)[0][0]
        raros = [i for i in C["items"] if len(i["apto"]) != tipica]
        if raros:
            out.append(Hallazgo("V07", "BAJA", "Cartera", "Códigos de inmueble con formato atípico",
                                "En el detalle aparecen códigos que no tienen el formato de los demás: " +
                                ", ".join(f"«{i['apto']}» ({pesos(i['valor'])})" for i in raros) + ".",
                                [i["fuente"] for i in raros], "Confirmar a qué inmueble corresponden."))
    return out


def r_reservas(m: dict) -> list[Hallazgo]:
    out = []
    R = m.get("reservas")
    if not R:
        return out
    if abs(R.get("diferencia") or 0) > 1:
        out.append(Hallazgo("V08", "ALTA", "Fondo de imprevistos",
                            "Diferencia contable en la reserva del fondo de imprevistos",
                            f"La hoja calcula una diferencia de {pesos(R['diferencia'])} entre la reserva contable "
                            "y el análisis de su aumento. Puede deberse a filas de análisis vacías para el mes.",
                            [R["fuente_diferencia"]],
                            "Completar el análisis de aumento del fondo o explicar la diferencia."))
    if R.get("reserva_anterior") and R["reserva"] < R["reserva_anterior"] - 1:
        out.append(Hallazgo("V09", "MEDIA", "Fondo de imprevistos",
                            "La reserva del fondo de imprevistos disminuyó frente al mes anterior",
                            f"Pasó de {pesos(R['reserva_anterior'])} a {pesos(R['reserva'])} "
                            f"({pesos(R['reserva'] - R['reserva_anterior'])}).",
                            [R["fuente_reserva"]],
                            "Verificar el soporte de la utilización del fondo y su aprobación."))
    return out


def r_conciliaciones(m: dict) -> list[Hallazgo]:
    out, signos = [], set()
    for c in m["conciliaciones"]:
        nombre = c["cuenta"][:70]
        if c.get("sin_datos"):
            out.append(Hallazgo("V10", "MEDIA", "Conciliaciones", f"Cuenta sin datos de conciliación: {nombre}",
                                "El bloque existe pero no tiene saldos ni anexos para el periodo.",
                                [c["fuente"]] if c["fuente"] else [],
                                "Confirmar si la cuenta sigue activa; si está cancelada, documentarlo."))
            continue
        if not (es_numero(c.get("extracto")) and es_numero(c.get("libros"))):
            falta = "saldo en libros" if not es_numero(c.get("libros")) else "saldo del extracto"
            out.append(Hallazgo("V10", "MEDIA", "Conciliaciones", f"Conciliación incompleta: {nombre}",
                                f"Para el periodo no hay {falta}; la diferencia no se puede calcular.",
                                [c["fuente"]] if c["fuente"] else [], "Completar la conciliación del mes."))
            continue
        f = (c.get("formula_diferencia") or "").replace("+", "")
        mm = re.match(r"=([A-Z]+)(\d+)-([A-Z]+)(\d+)", f)
        if mm:
            signos.add("extracto-libros" if int(mm.group(2)) > int(mm.group(4)) else "libros-extracto")
        dif = (c.get("extracto") or 0) - (c.get("libros") or 0)
        partidas = sum(p["valor"] for p in c.get("partidas", []) if es_numero(p["valor"]))
        if abs(dif) > 1 and abs(abs(dif) - abs(partidas)) > 1:
            out.append(Hallazgo("V11", "ALTA", "Conciliaciones", f"Diferencia no explicada en {nombre}",
                                f"Extracto {pesos(c['extracto'])} vs. libros {pesos(c['libros'])}; diferencia "
                                f"{pesos(dif)} y partidas conciliatorias por {pesos(partidas)}.", [c["fuente"]]))
        elif abs(dif) > 1:
            out.append(Hallazgo("V11", "INFO", "Conciliaciones", f"Partida conciliatoria en {nombre}",
                                f"Diferencia de {pesos(dif)} explicada como partidas pendientes. "
                                "Verificar que se regularicen el mes siguiente.", [c["fuente"]]))
        faltan = [k for k, v in c["anexos"].items() if v and norm(v).startswith("no")]
        if faltan:
            out.append(Hallazgo("V12", "MEDIA", "Conciliaciones", f"Anexos faltantes en {nombre}",
                                "Falta: " + ", ".join(faltan) + ".", [c["fuente"]]))
    if len(signos) > 1:
        out.append(Hallazgo("V12", "BAJA", "Conciliaciones", "Convención de signo distinta entre conciliaciones",
                            "Algunas diferencias se calculan como extracto − libros y otras como libros − extracto; "
                            "el signo de la diferencia no es comparable entre cuentas.", ["CONCILIACIONES"],
                            "Usar una sola convención (extracto − libros)."))
    return out


def r_retencion(m: dict, p: Periodo) -> list[Hallazgo]:
    out = []
    for r in m["retencion"]:
        if r["presentado"] and r["pagado"] and r["pagado"] > r["presentado"]:
            dias = (r["pagado"] - r["presentado"]).days
            if dias >= 3:
                out.append(Hallazgo("V13", "MEDIA", "Tributario",
                                    f"Retención {r['periodo_corto']}: pago {dias} días después de la presentación",
                                    f"Presentada el {r['presentado']:%d/%m/%Y} y pagada el {r['pagado']:%d/%m/%Y}. "
                                    f"La hoja la marca «{r['observacion']}».", [r["fuente"]],
                                    "Verificar si el pago se hizo dentro del vencimiento o si se liquidaron intereses."))
    esperado = p.anterior().clave
    if m["retencion"] and max(r["periodo"] for r in m["retencion"]) < esperado:
        out.append(Hallazgo("V14", "MEDIA", "Tributario", "Falta registrar la última declaración de retención",
                            f"El último periodo registrado es {max(r['periodo'] for r in m['retencion'])}; "
                            f"se esperaba al menos {esperado}.", ["IMPUESTOS"]))
    return out


def _antiguedad(fecha_iso: str, fin: date) -> int | None:
    try:
        return (fin - date.fromisoformat(fecha_iso)).days
    except (TypeError, ValueError):
        return None


def r_antiguedad(m: dict, p: Periodo) -> list[Hallazgo]:
    out = []
    for it in (m.get("anticipos") or {}).get("items") or []:
        d = _antiguedad(it["fecha_iso"], p.fin)
        if d and d > 365:
            out.append(Hallazgo("V15", "MEDIA", "Anticipos",
                                f"Anticipo con más de un año sin legalizar: {it['tercero']}",
                                f"{pesos(it['valor'])} desde el {it['fecha']} ({d // 30} meses).", [it["fuente"]],
                                "Gestionar legalización, devolución o reclasificación."))
    for it in (m.get("deudores_varios") or {}).get("items") or []:
        d = _antiguedad(it["fecha_iso"], p.fin)
        if d and d > 365:
            out.append(Hallazgo("V15", "MEDIA", "Deudores varios", f"Saldo antiguo en deudores varios: {it['tercero']}",
                                f"{pesos(it['valor'])} desde el {it['fecha']} ({d // 30} meses).", [it["fuente"]]))
    viejas = [i for i in (m["cxp"] or {}).get("items", []) if "90" in i["edad"] and "mayor" in norm(i["edad"])]
    if viejas:
        out.append(Hallazgo("V16", "MEDIA", "Cuentas por pagar", f"{len(viejas)} cuenta(s) por pagar con más de 90 días",
                            "; ".join(f"{i['tercero']} {pesos(i['valor'])}" for i in viejas) + ".",
                            [i["fuente"] for i in viejas], "Depurar o justificar los saldos antiguos."))
    f = (m.get("caja_menor") or {}).get("fecha_arqueo")
    d = _antiguedad(f, p.fin) if f else None
    if d and d > 60:
        out.append(Hallazgo("V17", "MEDIA", "Caja menor", "El último arqueo de caja menor está desactualizado",
                            f"Arqueo registrado el {date.fromisoformat(f):%d/%m/%Y}, {d} días antes del cierre del periodo.",
                            ["CAJA MENOR"], "Realizar arqueo en la visita del periodo."))
    return out


def r_presupuesto(m: dict) -> list[Hallazgo]:
    P = m.get("presupuesto")
    if P and P.get("total_hoja") is not None and abs(P["total_calculado"] - P["total_hoja"]) > 1:
        return [Hallazgo("V18", "ALTA", "Presupuesto", "El total de rubros sobreejecutados está subestimado",
                         f"La hoja reporta {pesos(P['total_hoja'])} ({P['formula_total']}); la suma de todos los "
                         f"rubros listados es {pesos(P['total_calculado'])}.", [P["fuente_total"]],
                         "El informe generado usa el total recalculado.")]
    return []


def r_totales_listas(m: dict) -> list[Hallazgo]:
    """V20: el total que trae la hoja debe ser la suma de las partidas que lista."""
    out = []
    for clave, nombre in (("cxp", "Cuentas por pagar"), ("anticipos", "Anticipos a contratistas"),
                          ("deudores_varios", "Deudores varios")):
        b = m.get(clave) or {}
        items, total = b.get("items") or [], b.get("total_hoja")
        if not items or total is None:
            continue
        suma = sum(i.get("valor") or 0 for i in items)
        if abs(suma - total) > 1:
            out.append(Hallazgo("V20", "ALTA", "Cruces", f"{nombre}: el total de la hoja no es la suma de sus partidas",
                                f"La hoja reporta {pesos(total)} y la suma de las {len(items)} partidas es {pesos(suma)} "
                                f"(diferencia {pesos(total - suma)}).", [b.get("fuente_total", "")],
                                "Revisar si el total está digitado o si falta una partida en la lista."))
    return out


def validar_excel(m: dict, lib: Libro, p: Periodo) -> list[Hallazgo]:
    hallazgos = []
    hallazgos += r_errores_formula(lib, p) + r_formulas_sin_resultado(lib) + r_rangos_suma_incompletos(lib) + r_valores_digitados(lib)
    hallazgos += r_cruces(m) + r_reservas(m) + r_conciliaciones(m) + r_presupuesto(m)
    hallazgos += r_retencion(m, p) + r_antiguedad(m, p) + r_totales_listas(m)
    # el V02 de PRESUPUESTO y el V18 dicen lo mismo: deja el más explícito
    if any(h.regla == "V18" for h in hallazgos):
        hallazgos = [h for h in hallazgos if not (h.regla == "V02" and "PRESUPUESTO" in h.titulo.upper())]
    return ordenar(hallazgos)


# dentro de cada severidad, primero lo que cambia cifras del informe
PRIORIDAD = ["P01", "P02", "V21", "V20", "V19", "V04", "V05", "V18", "V06", "V08", "V11", "V02", "P03", "P05", "V09", "V15", "V16",
             "V13", "V14", "V17", "V10", "V12", "V03", "V01", "V07", "P04", "P06"]


def ordenar(hs: list[Hallazgo]) -> list[Hallazgo]:
    pri = {r: i for i, r in enumerate(PRIORIDAD)}
    return sorted(hs, key=lambda h: (ORDEN[h.severidad], pri.get(h.regla, 99)))
