"""Plantilla estándar de herramientas de trabajo (v1) — la estructura del informe modelo (Picasso).

Opción C del proyecto: cada cliente se migra UNA vez a esta plantilla y desde entonces el auditor la diligencia cada
mes. Este módulo define los nombres oficiales de las hojas y sabe ESCRIBIR cada sección en la disposición que el
lector estándar (revisoria/extractor.py, documental.py) lee sin mapa ni recetas.

Cada escritor recibe la sección en su forma estándar (la que devuelven los lectores) y la escribe en una hoja nueva.
Con datos vacíos deja la hoja de la plantilla en blanco (títulos y encabezados). Cuando se migra, cada dato lleva un
comentario con su celda de origen en el Excel original (trazabilidad).
"""
from __future__ import annotations

from datetime import date, datetime
from io import BytesIO

from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill

from .excel import MESES, Periodo

VERSION = "1"
HOJA_CONTROL = "CONTROL PLANTILLA"

# clave lógica del mapa → nombre oficial de la hoja en la plantilla
HOJAS = {
    "CUADRO EFECTIVO": "CUADRO EFECTIVO", "GRAFICA DEUDORES": "GRAFICA DEUDORES", "Rangos de Cartera": "Rangos de Cartera",
    "MAYORES DEUDORES": "MAYORES DEUDORES", "SEGUIMIENTO DEUDORES": "may deudores", "ANTIC. E INGRESOS": "ANTIC. E INGRESOS",
    "CXP": "CXP", "CUADRO RESERVA": "CUADRO RESERVA", "PATRIMONIO": "PATRIMONIO", "PRESUPUESTO": "PRESUPUESTO",
    "CONCILIACIONES": "CONCILIACIONES", "IMPUESTOS": "IMPUESTOS", "CAJA MENOR": "CAJA MENOR", "EGRESO": "EGRESO",
    "AUDITORIA DOCUMENTAL": "AUD CONT Y ADMON", "ACTAS DE CONSEJO": "ACTAS DE CONSEJO", "ACTA ASAMBLEA": "ACTA ASAMBLEA",
    "CONTRATOS 2026": "CONTRATOS", "VERIFICACION POLIZAS": "VERIFICACION POLIZAS", "MATRIZ DICTAMEN": "MATRIZ DICTAMEN",
    "matriz manuales": "matriz manuales", "FIRMAS E INF DE GESTION": "FIRMAS E INF DE GESTION",
}
# hojas que NO se crean vacías al migrar: sin «may deudores» el lector toma los mayores deudores de la evolución
# mensual (MAYORES DEUDORES); una hoja de seguimiento vacía los haría desaparecer del informe
SIN_HOJA_VACIA = {"SEGUIMIENTO DEUDORES"}
# sección del informe → clave lógica de su hoja
SECCION_HOJA = {"liquidez": "CUADRO EFECTIVO", "cartera": "Rangos de Cartera", "cxp": "CXP", "anticipos": "ANTIC. E INGRESOS",
                "deudores_varios": "ANTIC. E INGRESOS", "retencion": "IMPUESTOS", "presupuesto": "PRESUPUESTO",
                "egresos": "EGRESO", "polizas": "VERIFICACION POLIZAS", "matriz_riesgo": "MATRIZ DICTAMEN",
                "matriz_manuales": "matriz manuales", "actas_consejo": "ACTAS DE CONSEJO",
                "mayores_deudores": "SEGUIMIENTO DEUDORES"}

AMARILLO = PatternFill("solid", fgColor="FFF2CC")       # celdas que diligencia el auditor
NEGRILLA = Font(bold=True)


def _mes(p: Periodo) -> str:
    return f"{MESES[p.mes - 1].upper()} {p.anio}"


def _fecha_dt(v):
    """«05/08/2026», «2026-08-05», date → datetime (para que Excel la guarde como fecha)."""
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, str) and v.strip():
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(v.strip()[:10], fmt)
            except ValueError:
                pass
    return v or None


def _periodo_dt(clave_o_corto) -> datetime | None:
    if not clave_o_corto:
        return None
    s = str(clave_o_corto)
    if len(s) == 7 and s[4] == "-":
        return datetime(int(s[:4]), int(s[5:]), 1)
    from .lectores import periodo_celda
    p = periodo_celda(s)
    return datetime(p.anio, p.mes, 1) if p else None


class Escritor:
    """Escribe en una hoja y deja el comentario de origen de cada dato."""

    def __init__(self, ws, trazar: bool = True):
        self.ws, self.trazar = ws, trazar

    def __call__(self, fila: int, col: int, valor, origen: str | None = None, negrilla: bool = False, amarillo: bool = False):
        c = self.ws.cell(fila, col)
        c.value = valor
        if negrilla:
            c.font = NEGRILLA
        if amarillo:
            c.fill = AMARILLO
        if origen and self.trazar and valor not in (None, ""):
            c.comment = Comment(f"Origen: {origen}", "LCCOUNTER migración")
        return c

    def texto_rico(self, fila: int, col: int, segmentos, texto: str, origen: str | None = None):
        """Conserva negrillas y rojos del texto original (se usan en las diapositivas)."""
        if segmentos:
            from openpyxl.cell.rich_text import CellRichText, TextBlock
            from openpyxl.cell.text import InlineFont
            partes = [TextBlock(InlineFont(b=s.get("b") or None, color="FFFF0000" if s.get("rojo") else None), s["t"])
                      for s in segmentos]
            return self(fila, col, CellRichText(*partes), origen)
        return self(fila, col, texto, origen)

    def imagenes(self, fila: int, col: int, imagenes):
        from openpyxl.drawing.image import Image
        for k, im in enumerate(imagenes or []):
            try:
                img = Image(BytesIO(im["bytes"]))
            except Exception:
                continue
            img.width, img.height = im.get("ancho") or img.width, im.get("alto") or img.height
            self.ws.add_image(img, f"{self.ws.cell(fila, col + k).column_letter}{fila}")


def _encabezado(e: Escritor, fila: int, col: int, titulos: list[str]):
    for i, t in enumerate(titulos):
        e(fila, col + i, t, negrilla=True)


# ------------------------------------------------------------------ escritores por sección (disposición de Picasso)

def escribir_cxp(ws, d: dict | None, p: Periodo, trazar=True):
    e, d = Escritor(ws, trazar), d or {}
    e(2, 2, f"CUENTAS POR PAGAR A {_mes(p)}", negrilla=True)
    _encabezado(e, 3, 2, ["EDAD", "TERCERO", "VALOR", "CONCEPTO"])
    r = 4
    for it in sorted(d.get("items") or [], key=lambda x: x.get("_fila") or 0):
        e(r, 2, it.get("edad"), amarillo=True)
        e(r, 3, it.get("tercero"), amarillo=True)
        e(r, 4, it.get("valor"), it.get("fuente"), amarillo=True)
        e.texto_rico(r, 5, it.get("seg_concepto"), it.get("concepto"))
        r += 1
    e(r, 2, "TOTAL", negrilla=True)
    e(r, 4, d.get("total_hoja"), d.get("fuente_total"), negrilla=True)


def _escribir_bloque_mes(e: Escritor, fila: int, titulo: str, etiqueta_total: str, d: dict | None):
    d = d or {}
    e(fila, 2, titulo, negrilla=True)
    _encabezado(e, fila + 1, 2, ["FECHA", "CONTRATISTA", "VALOR", "DETALLE"])
    r = fila + 2
    for it in d.get("items") or []:
        e(r, 2, _fecha_dt(it.get("fecha_iso") or it.get("fecha")), amarillo=True)
        e(r, 3, it.get("tercero"), amarillo=True)
        e(r, 4, it.get("valor"), it.get("fuente"), amarillo=True)
        e.texto_rico(r, 5, it.get("seg_detalle"), it.get("detalle"))
        r += 1
    e(r, 2, etiqueta_total, negrilla=True)
    e(r, 4, d.get("total_hoja"), d.get("fuente_total"), negrilla=True)
    return r + 2


def escribir_anticipos(ws, anticipos: dict | None, deudores_varios: dict | None, p: Periodo, trazar=True):
    e = Escritor(ws, trazar)
    r = _escribir_bloque_mes(e, 2, f"ANTICIPOS Y AVANCES A CONTRATISTAS A {_mes(p)} - CTA 1330", "TOTAL ANTICIPOS", anticipos)
    _escribir_bloque_mes(e, r, f"DEUDORES VARIOS A {_mes(p)} - CTA 1380", "TOTAL DEUDORES VARIOS", deudores_varios)


def escribir_retencion(ws, filas: list | None, p: Periodo, trazar=True):
    e = Escritor(ws, trazar)
    e(2, 2, "IMPUESTO RETENCION EN LA FUENTE", negrilla=True)
    _encabezado(e, 3, 2, ["PERIODO", "PRESENTADO", "PAGADO", "VALOR", "OBSERVACIÓN"])
    for r, f in enumerate(filas or [], start=4):
        e(r, 2, _periodo_dt(f.get("periodo")), f.get("fuente"), amarillo=True)
        e(r, 3, _fecha_dt(f.get("presentado")), amarillo=True)
        e(r, 4, _fecha_dt(f.get("pagado")), amarillo=True)
        e(r, 5, f.get("valor"), amarillo=True)
        e(r, 6, f.get("observacion"), amarillo=True)


def escribir_egresos(ws, d: dict | None, p: Periodo, trazar=True):
    e, d = Escritor(ws, trazar), d or {}
    e(2, 2, f"COMPROBANTES DE EGRESO {_mes(p)}", negrilla=True)
    _encabezado(e, 3, 2, ["CE", "FECHA", "TERCERO", "DESCRIPCIÓN", "VALOR", "OBSERVACIÓN"])
    r = 4
    for it in d.get("items") or []:
        e(r, 2, it.get("ce"), amarillo=True)
        e(r, 3, it.get("fecha"), amarillo=True)
        e(r, 4, it.get("tercero"), amarillo=True)
        e(r, 5, it.get("descripcion"), amarillo=True)
        e(r, 6, it.get("valor"), it.get("fuente"), amarillo=True)
        e.texto_rico(r, 7, it.get("seg_observacion"), it.get("observacion"))
        e.imagenes(r, 9, it.get("imagenes"))
        r += 1
    e(r, 2, d.get("etiqueta_total") or f"TOTAL EGRESOS {_mes(p)}", negrilla=True)
    e(r, 6, d.get("total_hoja"), d.get("fuente_total"), negrilla=True)


def escribir_polizas(ws, d: dict | None, p: Periodo, trazar=True):
    e, d = Escritor(ws, trazar), d or {}
    e(1, 2, "VERIFICACIÓN DE PÓLIZAS DE LOS CONTRATOS VIGENTES", negrilla=True)
    _encabezado(e, 2, 2, ["N.", "NOMBRE PROVEEDOR / CONTRATO", "VERIFICACIÓN DE POLIZAS DE ACUERDO AL MANUAL DE CONTRATACIÓN"])
    for r, f in enumerate(d.get("filas") or [], start=3):
        e(r, 2, f.get("n"), amarillo=True)
        e(r, 3, f.get("contrato"), amarillo=True)
        e(r, 4, f.get("verificacion"), f.get("fuente"), amarillo=True)
    m = d.get("matriz") or {}
    if m.get("columnas"):
        e(1, 6, m.get("titulo"), negrilla=True)
        _encabezado(e, 2, 6, ["Tipo de contrato"] + list(m["columnas"][1:]))
        for r, fila in enumerate(m.get("filas") or [], start=3):
            for c, v in enumerate(fila):
                e(r, 6 + c, v)


def escribir_matriz_riesgo(ws, d: dict | None, p: Periodo, trazar=True):
    e, d = Escritor(ws, trazar), d or {}
    e(2, 2, d.get("titulo") or "Matriz de riesgo, Ambiental, Mantenimiento y Salubridad", negrilla=True)
    periodos = d.get("periodos") or []
    for i, pc in enumerate(periodos[:3]):
        e(2, 5 + i, _periodo_dt(pc))
    _encabezado(e, 3, 2, ["Objetivo de Control", "DESCRIPCIÓN DEL RIESGO Y ACCIONES RECOMENDADAS", "Valorización del riesgo"]
                + ["Nivel de cumplimiento"] * max(len(periodos[:3]), 1))
    r = 4
    for f in d.get("filas") or []:
        e(r, 2, f.get("objetivo"), f.get("fuente"), amarillo=True)
        e(r, 4, f.get("valoracion"), amarillo=True)
        for i, pc in enumerate(periodos[:3]):
            e(r, 5 + i, (f.get("niveles") or {}).get(pc) or None, amarillo=True)
        descs = f.get("descripcion") or [{"texto": "", "seg": None}]
        for k, ds in enumerate(descs):
            e.texto_rico(r + k, 3, ds.get("seg"), ds.get("texto") or None)
        r += max(len(descs), 1)


def escribir_matriz_manuales(ws, d: dict | None, p: Periodo, trazar=True):
    e, d = Escritor(ws, trazar), d or {}
    e(2, 2, "CONTROLES INTERNOS: Actos administrativos y manuales", negrilla=True)
    _encabezado(e, 3, 2, ["Objetivo de Control", "Detalle la formalización y documentación", "", "Valorización del riesgo",
                          "Nivel de cumplimiento", "Nivel de cumplimiento", "Nivel de cumplimiento"])
    for r, f in enumerate(d.get("filas") or [], start=4):
        e(r, 2, f.get("objetivo"), f.get("fuente"), amarillo=True)
        e(r, 3, f.get("detalle"), amarillo=True)
        e(r, 5, f.get("valoracion"), amarillo=True)
        for i, v in enumerate((f.get("niveles") or [])[:3]):
            e(r, 6 + i, v or None, amarillo=True)


def escribir_actas(ws, d: dict | None, p: Periodo, trazar=True):
    e, d = Escritor(ws, trazar), d or {}
    e(2, 3, "Relación de actas de consejo de administración", negrilla=True)
    _encabezado(e, 3, 3, ["N. acta y fecha", "Temas", "Firma presidente", "Firma administración"])
    for r, a in enumerate(d.get("actas") or [], start=4):
        acta = " ".join(x for x in (a.get("numero"), a.get("fecha")) if x) or a.get("fecha") or ""
        e(r, 3, acta, a.get("fuente"), amarillo=True)
        e(r, 4, a.get("temas"), amarillo=True)
        e(r, 5, a.get("firma_presidente") or None, amarillo=True)
        e(r, 6, a.get("firma_administracion") or None, amarillo=True)


def escribir_presupuesto(ws, d: dict | None, p: Periodo, trazar=True):
    e, d = Escritor(ws, trazar), d or {}
    e(1, 2, "SOBREEJECUCIÓN PRESUPUESTAL", negrilla=True)   # no puede empezar por «Rubro»: es la fila de encabezado
    _encabezado(e, 2, 2, ["Rubro", f"Presupuesto Enero- Diciembre {p.anio}", d.get("encabezado_periodo")
                          or f"Presupuesto Enero- {MESES[p.mes - 1].capitalize()} {p.anio}",
                          f"Ejecución Enero-{MESES[p.mes - 1].capitalize()} {p.anio}", "Valor sobre ejecutado", "Observaciones"])
    r = 3
    for it in sorted(d.get("rubros") or [], key=lambda x: x.get("_fila") or 0):
        e(r, 2, it.get("rubro"), amarillo=True)
        e(r, 3, it.get("presupuesto_anual"), amarillo=True)
        e(r, 4, it.get("presupuesto_periodo"), amarillo=True)
        e(r, 5, it.get("ejecucion"), it.get("fuente"), amarillo=True)
        e(r, 6, (it.get("presupuesto_periodo") or 0) - (it.get("ejecucion") or 0))
        e(r, 7, it.get("observacion") or None, amarillo=True)
        e.imagenes(r, 9, it.get("soportes"))
        r += 1
    e(r, 2, "Total rubros sobre ejecutados", negrilla=True)
    th = d.get("total_hoja")
    e(r, 6, -th if th is not None else None, d.get("fuente_total"), negrilla=True)


def escribir_liquidez(ws, d: dict | None, p: Periodo, trazar=True):
    """Encabezado de meses (diciembre anterior → mes actual), conceptos con el mes y el cierre, total por mes."""
    e, d = Escritor(ws, trazar), d or {}
    meses = sorted({s["periodo"] for s in d.get("serie") or []} | {p.clave, p.cierre_anterior().clave}, reverse=True)
    e(2, 2, "Análisis de liquidez de efectivo", negrilla=True)
    col = {m: 3 + i for i, m in enumerate(meses)}
    for m, c in col.items():
        e(2, c, _periodo_dt(m), negrilla=True)
    r = 3
    filas = d.get("filas") or [{"concepto": "Total equivalente al efectivo", "es_total": True}]
    for f in filas:
        e(r, 2, f.get("concepto"), negrilla=bool(f.get("es_total")), amarillo=not f.get("es_total"))
        e(r, col[p.clave], f.get("actual"), f.get("fuente"))
        e(r, col[p.cierre_anterior().clave], f.get("referencia"))
        r += 1
    f_tot = r - 1
    if not norm_total(filas[-1].get("concepto")):
        e(r, 2, "Total equivalente al efectivo", negrilla=True)
        f_tot = r
        r += 1
    for s in d.get("serie") or []:
        e(f_tot, col[s["periodo"]], s["valor"])
    if d.get("monetizacion"):
        r += 1
        e(r, 2, "Análisis Monetización fondo de imprevistos", negrilla=True)
        for m, c in col.items():
            e(r, c, _periodo_dt(m), negrilla=True)
        for f in d["monetizacion"]:
            r += 1
            e(r, 2, f.get("concepto"), negrilla=bool(f.get("es_total")))
            e(r, col[p.clave], f.get("actual"), f.get("fuente"))
            e(r, col[p.cierre_anterior().clave], f.get("referencia"))


def norm_total(t) -> bool:
    return str(t or "").strip().lower().startswith("total")


def escribir_cartera(ws, d: dict | None, p: Periodo, trazar=True):
    """Resumen a la izquierda (B:E) y las listas por inmueble del mes y del mes anterior (con su título)."""
    e, d = Escritor(ws, trazar), d or {}
    e(1, 2, "Total todos los inmuebles", negrilla=True)
    e(1, 3, d.get("unidades_hoja"), amarillo=True)
    _encabezado(e, 2, 2, ["Detalle de rango deudores", "Cifra", "No inmuebles", "%Cartera"])
    r = 3
    for rg in d.get("rangos") or []:
        e(r, 2, rg["rango"]); e(r, 3, rg["valor"]); e(r, 4, rg["inmuebles"]); e(r, 5, rg["pct_cartera"])
        r += 1
    rh = d.get("resumen_hoja")
    if rh:
        e(r, 2, f"TOTAL CARTERA A {_mes(p)}", negrilla=True)
        e(r, 3, rh.get("total"), rh.get("fuente"))
        e(r, 4, rh.get("inmuebles"))

    def lista(col0, per, items, total, fuente_total):
        e(1, col0, f"CARTERA A {_mes(per)}", negrilla=True)
        _encabezado(e, 2, col0, ["No", "APTO", "VALOR"])
        f = 3
        for i, it in enumerate(sorted(items, key=lambda x: x.get("fila") or 0), start=1):
            e(f, col0, i)
            e(f, col0 + 1, it.get("apto"), amarillo=True)
            e(f, col0 + 2, it.get("valor"), it.get("fuente"), amarillo=True)
            f += 1
        e(f, col0, "TOTAL", negrilla=True)
        e(f, col0 + 2, total, fuente_total, negrilla=True)

    lista(8, p, d.get("items") or [], d.get("total_hoja_lista"), d.get("fuente_total"))
    ant = d.get("_items_anterior")
    if ant is not None:
        lista(12, p.anterior(), ant.get("items") or [], ant.get("total_hoja"), ant.get("fuente_total"))


def escribir_seguimiento_deudores(ws, d: dict | None, p: Periodo, trazar=True):
    """Hoja «may deudores»: un inmueble por fila con saldo anterior y actual, gestión del abogado y observación de la
    revisoría. Una observación que aplica a varios inmuebles seguidos va en una celda combinada (como en el original)."""
    e, d = Escritor(ws, trazar), d or {}
    _encabezado(e, 1, 1, ["Apto", f"Saldo {p.anterior().corto}", f"Saldo {p.corto}",
                          d.get("encabezado_abogado") or "Gestión según informe del abogado", "",
                          "Proceso según Manual de Cobro de Cartera", "Observación Revisoría Fiscal"])
    ws.merge_cells("D1:E1")
    r, inicio_obs = 2, None
    for it in d.get("items") or []:
        e(r, 1, it.get("apto"), amarillo=True)
        e(r, 2, it.get("saldo_anterior"), amarillo=True)
        e(r, 3, it.get("saldo_actual"), it.get("fuente"), amarillo=True)
        e(r, 4, it.get("gestion_abogado") or None, amarillo=True)
        e.texto_rico(r, 5, it.get("seg_detalle"), it.get("detalle_gestion") or None)
        e(r, 6, it.get("proceso_manual") or None, amarillo=True)
        if it.get("observacion_compartida") and inicio_obs is not None:
            ws.merge_cells(start_row=inicio_obs, start_column=7, end_row=r, end_column=7)
        else:
            e.texto_rico(r, 7, it.get("seg_observacion"), it.get("observacion_rf") or None)
            inicio_obs = r
        r += 1


ESCRITORES = {"cxp": escribir_cxp, "retencion": escribir_retencion, "egresos": escribir_egresos,
              "polizas": escribir_polizas, "matriz_riesgo": escribir_matriz_riesgo,
              "matriz_manuales": escribir_matriz_manuales, "actas_consejo": escribir_actas,
              "presupuesto": escribir_presupuesto, "liquidez": escribir_liquidez, "cartera": escribir_cartera,
              "mayores_deudores": escribir_seguimiento_deudores}
# anticipos y deudores varios comparten hoja: escribir_anticipos(ws, anticipos, deudores_varios, p)


# ================================================================== plantilla vacía (para los auditores)

def _meses_anio(p: Periodo) -> list[datetime]:
    """Columnas de meses como en el informe modelo: del mes más reciente al diciembre anterior (izquierda → derecha)."""
    meses = [datetime(p.anio, m, 1) for m in range(12, 0, -1)] + [datetime(p.anio - 1, 12, 1)]
    return meses


def _bloque_conceptos(e: Escritor, fila: int, titulo: str, meses, conceptos: list[str], total: str | None,
                      formula_total: bool = True) -> int:
    """Encabezado con meses, renglones de concepto (amarillos) y fila total con =SUMA por mes."""
    e(fila, 2, titulo, negrilla=True)
    for i, m in enumerate(meses):
        e(fila, 3 + i, m, negrilla=True).number_format = "mmm-yy"
    r0 = fila + 1
    for k, c in enumerate(conceptos):
        e(r0 + k, 2, c, amarillo=True)
        for i in range(len(meses)):
            e.ws.cell(r0 + k, 3 + i).fill = AMARILLO
    r = r0 + len(conceptos)
    if total:
        e(r, 2, total, negrilla=True)
        if formula_total:
            from openpyxl.utils import get_column_letter
            for i in range(len(meses)):
                L = get_column_letter(3 + i)
                e.ws.cell(r, 3 + i).value = f"=SUM({L}{r0}:{L}{r - 1})"
        r += 1
    return r


def plantilla_vacia(p: Periodo, ruta) -> str:
    """Genera la plantilla estándar vacía (versión VERSION) para el año de `p`. Devuelve la ruta."""
    import openpyxl
    from openpyxl.utils import get_column_letter
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    meses = _meses_anio(p)

    def hoja(clave):
        return wb.create_sheet(HOJAS[clave])

    wc = wb.create_sheet(HOJA_CONTROL)
    for r, (a, b) in enumerate([
            ("Plantilla estándar LCCOUNTER", f"versión {VERSION} · año {p.anio}"),
            ("Instrucciones", "Diligencie solo las celdas amarillas. No cambie nombres de hojas, títulos ni encabezados: "
                              "el sistema los usa para leer el archivo."),
            ("Bloques del mes", "En CXP, ANTIC. E INGRESOS y EGRESO cada mes es un bloque con su título "
                                "(«… A AGOSTO 2026»): copie el bloque y cambie el mes."),
            ("Información particular del cliente", "Si el cliente tiene información que no cabe aquí, póngala en una hoja "
                                                    "nueva cuyo nombre empiece por «ANEXO - »."),
            ("Hojas", ", ".join(HOJAS.values()))], start=1):
        wc.cell(r, 1, a).font = NEGRILLA
        wc.cell(r, 2, b)
    wc.column_dimensions["A"].width, wc.column_dimensions["B"].width = 34, 110

    e = Escritor(hoja("CUADRO EFECTIVO"), False)
    r = _bloque_conceptos(e, 2, "Análisis de liquidez de efectivo", meses,
                          ["(+) Saldo cuenta corriente (banco y número)", "(+) Saldo cuenta de ahorros (banco y número)",
                           "(+) Saldo fiducia (entidad y número)", "(-) Cuentas por pagar", "(-) Retención en la fuente e impuestos por pagar",
                           "(-) Anticipos y otros pasivos"], "Total equivalente al efectivo")
    r = _bloque_conceptos(e, r + 1, "Análisis Monetización fondo de imprevistos", meses,
                          ["(+) Recursos en fiducia del fondo de imprevistos", "(-) Reserva fondo de imprevistos registrada"],
                          "Mayor valor monetizado")
    e = Escritor(hoja("GRAFICA DEUDORES"), False)
    e(1, 2, "Análisis de deudores de propietarios", negrilla=True)
    _bloque_conceptos(e, 2, "Detalle", meses, ["(+) Deudores copropietarios", "(-) Consignaciones por identificar",
                                               "(-) Deterioro de cartera"], "Deudores copropietarios")
    escribir_cartera(hoja("Rangos de Cartera"), None, p, False)
    e = Escritor(hoja("MAYORES DEUDORES"), False)
    e(1, 1, f"MAYORES DEUDORES {p.anio}", negrilla=True)
    e(2, 1, "#", negrilla=True); e(2, 2, "APTO", negrilla=True)
    for i, m in enumerate(reversed(meses)):
        e(2, 3 + i, m, negrilla=True).number_format = "mmm-yy"
    escribir_seguimiento_deudores(hoja("SEGUIMIENTO DEUDORES"), None, p, False)
    escribir_anticipos(hoja("ANTIC. E INGRESOS"), None, None, p, False)
    escribir_cxp(hoja("CXP"), None, p, False)
    e = Escritor(hoja("CUADRO RESERVA"), False)
    e(1, 2, "RESERVAS DE PATRIMONIO", negrilla=True)
    r = _bloque_conceptos(e, 2, "Reservas patrimonio", meses, ["Reserva fondo de imprevistos"], "Total reservas patrimonio")
    e(r, 2, f"Fondo de imprevistos al 31 dic {p.anio - 1}", amarillo=True)
    e(r + 1, 2, "Diferencia contable de reserva", amarillo=True)
    e = Escritor(hoja("PATRIMONIO"), False)
    r = _bloque_conceptos(e, 2, "RESULTADO DEL EJERCICIO", meses, ["Total ingresos", "Total gastos"], f"Resultado del ejercicio {p.anio}")
    r = _bloque_conceptos(e, r + 1, "Excedentes acumulados", meses, ["Excedentes de ejercicios anteriores",
                                                                   "Resultado del ejercicio"], "Total excedentes acumulados")
    _bloque_conceptos(e, r + 1, "Disponible de excedentes acumulados", meses, ["(+) Disponible", "(-) Pasivos que lo comprometen"],
                      "Disponibilidad de excedentes")
    escribir_presupuesto(hoja("PRESUPUESTO"), None, p, False)
    e = Escritor(hoja("CONCILIACIONES"), False)
    e(3, 2, "Cuenta corriente BANCO ________ N.° ________", amarillo=True)
    e(4, 2, "BANCO", negrilla=True)
    for i, m in enumerate(meses):
        e(4, 3 + i, m, negrilla=True).number_format = "mmm-yy"
    for k, t in enumerate(["anexo extracto bancario", "anexo libro auxiliar contable", "anexo conciliación bancaria firmada",
                           "saldo en libros", "saldo extracto", "Diferencia", "Pagos pendientes por ingresar al banco"]):
        e(5 + k, 2, t, amarillo=t != "Diferencia")
    for i in range(len(meses)):
        L = get_column_letter(3 + i)
        e.ws.cell(10, 3 + i).value = f"=+{L}9-{L}8"
    escribir_retencion(hoja("IMPUESTOS"), None, p, False)
    e = Escritor(hoja("CAJA MENOR"), False)
    e(1, 2, "ARQUEO DE CAJA MENOR", negrilla=True)
    for k, t in enumerate(["Responsable:", "Fecha:", "Hora:", "Total fondo"]):
        e(2 + k, 2, t, negrilla=True); e.ws.cell(2 + k, 3).fill = AMARILLO
    escribir_egresos(hoja("EGRESO"), None, p, False)
    e = Escritor(hoja("AUDITORIA DOCUMENTAL"), False)
    r = 1
    for bloque in ("Contable", "Administrativa"):
        e(r, 2, f"Auditoria documental {bloque} de cumplimiento", negrilla=True)
        _encabezado(e, r + 1, 2, ["Campo auditado", "Normatividad", "Evidencia (Ev) / Opinión (Op)", "Nivel Cumplimiento", "Documento"])
        r += 12
    escribir_actas(hoja("ACTAS DE CONSEJO"), None, p, False)
    e = Escritor(hoja("ACTA ASAMBLEA"), False)
    e(1, 2, "DECISIONES RELEVANTES DE ASAMBLEA", negrilla=True)
    _encabezado(e, 2, 2, ["Temas", "Detalle de lo aprobado", "Seguimiento al cumplimiento"])
    e = Escritor(hoja("CONTRATOS 2026"), False)
    e(1, 2, "CONTRATOS VIGENTES", negrilla=True)
    _encabezado(e, 2, 2, ["N.", "NOMBRE PROVEEDOR", "VIGENCIA CONTRATO", "VALOR", "TOTAL CONTRATO",
                          "PÓLIZAS Y GARANTÍAS - OBSERVACIONES"])
    escribir_polizas(hoja("VERIFICACION POLIZAS"), None, p, False)
    escribir_matriz_riesgo(hoja("MATRIZ DICTAMEN"), {"periodos": [Periodo(p.anio - 1, 12).corto, Periodo(p.anio, 6).corto]}, p, False)
    escribir_matriz_manuales(hoja("matriz manuales"), None, p, False)
    e = Escritor(hoja("FIRMAS E INF DE GESTION"), False)
    e(1, 2, f"FIRMA ESTADOS FINANCIEROS {p.anio}", negrilla=True)
    _encabezado(e, 2, 2, ["Fecha", "Contador", "Rep. Legal", "Fecha", "", "Entrega informe de gestión"])
    for k, m in enumerate(range(1, 13)):
        e(3 + k, 2, datetime(p.anio, m, 1)).number_format = "mmm-yy"
        for c in (3, 4, 7):
            e.ws.cell(3 + k, c).fill = AMARILLO
    for ws in wb.worksheets:
        ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 10, 42)
    wb.save(ruta)
    return str(ruta)
