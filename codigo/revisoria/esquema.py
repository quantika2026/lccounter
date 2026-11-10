"""Esquema estándar de datos: el contrato fijo entre el Excel (que varía por cliente) y el informe (que no varía).

El generador del PowerPoint solo consume secciones de este esquema. Cada cliente llega a él a través de su mapa
(revisoria/mapa.py). `verificar` dice qué secciones faltan o vinieron vacías, sin inventar nada.
"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Seccion:
    clave: str
    descripcion: str
    hoja: str                   # nombre lógico de la hoja de origen (clave del mapa)
    obligatoria: bool           # sin ella el informe no es presentable
    claves: tuple = ()          # campos que la sección debe traer
    contenido: str | None = None  # lista que no debería venir vacía (None: la sección misma es la lista)
    documental: bool = False    # viene de extraer_documental (informe completo)


SECCIONES: list[Seccion] = [
    # ---- financiero
    Seccion("liquidez", "Análisis de liquidez y monetización", "CUADRO EFECTIVO", True,
            ("filas", "total", "serie"), "filas"),
    Seccion("deudores", "Deudores (bruto y neto)", "GRAFICA DEUDORES", True, ("bruto", "neto", "filas"), "filas"),
    Seccion("cartera", "Cartera por inmueble y rangos de edad", "Rangos de Cartera", True,
            ("items", "total", "rangos"), "items"),
    Seccion("mayores_deudores", "Mayores deudores y gestión de cobro", "SEGUIMIENTO DEUDORES", False, ("items",), "items"),
    Seccion("anticipos", "Anticipos (1330)", "ANTIC. E INGRESOS", False, ("items", "total_hoja"), "items"),
    Seccion("deudores_varios", "Deudores varios (1380)", "ANTIC. E INGRESOS", False, ("items", "total_hoja"), "items"),
    Seccion("cxp", "Cuentas por pagar", "CXP", True, ("items", "total"), "items"),
    Seccion("patrimonio", "Patrimonio, resultado y excedentes", "PATRIMONIO", True, ("resultado", "ingresos", "gastos")),
    Seccion("presupuesto", "Ejecución presupuestal", "PRESUPUESTO", True, ("rubros",), "rubros"),
    Seccion("reservas", "Fondo de imprevistos", "CUADRO RESERVA", False, ("reserva",)),
    Seccion("conciliaciones", "Conciliaciones bancarias", "CONCILIACIONES", True),
    Seccion("retencion", "Retención en la fuente", "IMPUESTOS", False),
    Seccion("caja_menor", "Arqueo de caja menor", "CAJA MENOR", False, ("fecha_arqueo",)),
    # ---- documental (informe completo)
    Seccion("egresos", "Revisión de egresos", "EGRESO", True, ("items",), "items", True),
    Seccion("auditoria", "Auditoría documental contable y administrativa", "AUDITORIA DOCUMENTAL", True,
            ("campos",), "campos", True),
    Seccion("actas_consejo", "Actas de consejo", "ACTAS DE CONSEJO", False, ("actas",), "actas", True),
    Seccion("asamblea", "Acta de asamblea", "ACTA ASAMBLEA", False, ("bloques",), "bloques", True),
    Seccion("contratos", "Contratos vigentes", "CONTRATOS 2026", False, ("contratos",), "contratos", True),
    Seccion("polizas", "Verificación de pólizas", "VERIFICACION POLIZAS", False, ("filas",), "filas", True),
    Seccion("matriz_riesgo", "Matriz de riesgo / dictamen", "MATRIZ DICTAMEN", False, ("filas",), "filas", True),
    Seccion("matriz_manuales", "Matriz de manuales", "matriz manuales", False, ("filas",), "filas", True),
    Seccion("firmas_informes", "Firmas e informe de gestión", "FIRMAS E INF DE GESTION", False, (), None, True),
]

# hojas que se ubican por nombre (lib.hoja); las demás claves del mapa se resuelven con un detector propio
HOJAS_POR_NOMBRE = sorted({s.hoja for s in SECCIONES} - {"SEGUIMIENTO DEUDORES", "AUDITORIA DOCUMENTAL",
                                                          "VERIFICACION POLIZAS"} | {"MAYORES DEUDORES"})


def verificar(modelo: dict, doc: dict | None = None) -> list[dict]:
    """Estado de cada sección del esquema: ok, vacia (vino sin filas), incompleta (faltan campos) o falta."""
    doc = doc or {}
    salida = []
    for s in SECCIONES:
        if s.documental and not doc:
            continue
        v = (doc if s.documental else modelo).get(s.clave)
        falt = (doc if s.documental else modelo).get("faltantes") or {}
        if v is None or v == {} or s.clave in falt:
            estado, detalle = "falta", falt.get(s.clave, "no se encontró en el Excel")
        elif faltan := [k for k in s.claves if not isinstance(v, dict) or k not in v]:
            estado, detalle = "incompleta", "faltan: " + ", ".join(faltan)
        elif isinstance(c := (v.get(s.contenido) if s.contenido else v), (list, dict)) and not c:
            estado, detalle = "vacia", "la tabla vino sin filas"
        else:
            estado, detalle = "ok", ""
        salida.append({"seccion": s.clave, "descripcion": s.descripcion, "hoja": s.hoja,
                       "obligatoria": s.obligatoria, "estado": estado, "detalle": detalle})
    return salida


# ---------------------------------------------------------------- firma de formato
# Textos que la hoja debe tener en sus primeras filas para que el lector la entienda. Un nombre de hoja parecido
# no basta: «POLIZA» (pagos de pólizas) no es la «verificación de pólizas» que el informe necesita.
# «a|b» acepta cualquiera de las variantes (cada despacho titula sus columnas a su manera).
FIRMAS = {
    "CUADRO EFECTIVO": ("analisis de liquidez",), "GRAFICA DEUDORES": ("deudores",), "Rangos de Cartera": ("rango",),
    "MAYORES DEUDORES": ("apto|inmueble",), "ANTIC. E INGRESOS": ("anticipo",), "CXP": ("cuentas por pagar",),
    "CUADRO RESERVA": ("reserva",), "PATRIMONIO": ("resultado del ejercicio|resultado ejercicio",), "PRESUPUESTO": ("rubro",),
    "CONCILIACIONES": ("extracto",), "IMPUESTOS": ("retencion",), "CAJA MENOR": ("arqueo",), "EGRESO": ("egreso",),
    "AUDITORIA DOCUMENTAL": ("campo auditado",), "ACTAS DE CONSEJO": ("acta", "consejo"), "ACTA ASAMBLEA": ("asamblea",),
    "CONTRATOS 2026": ("contrato|vencimiento",), "VERIFICACION POLIZAS": ("verificacion", "poliza"),
    "MATRIZ DICTAMEN": ("objetivo de control", "valorizacion del riesgo|nivel riesgo|nivel de riesgo"),
    "matriz manuales": ("objetivo de control", "valorizacion del riesgo|nivel riesgo|nivel de riesgo"),
    "FIRMAS E INF DE GESTION": ("contador",), "SEGUIMIENTO DEUDORES": ("apto", "gestion"),
}
NIVELES_RIESGO = ("extremo", "alto", "moderado", "medio", "bajo", "odm")


def _texto_cabecera(ws, filas: int = 8, cols: int = 16) -> str:
    from .excel import norm
    return " ".join(norm(v) for fila in ws.iter_rows(min_row=1, max_row=filas, max_col=cols, values_only=True)
                    for v in fila if isinstance(v, str))


def cumple(firma: str, texto: str) -> bool:
    """¿El texto trae alguna de las variantes de la firma («a|b»)?"""
    return any(v in texto for v in firma.split("|"))


def revisar_formato(lib, hojas: dict, doc: dict | None = None, lectores: dict | None = None) -> list[dict]:
    """Por cada hoja del mapa: ¿tiene la firma de su sección? y, para las matrices, ¿las columnas están en su lugar?
    Las claves cuyas secciones se leen con una receta aprobada (revisoria/lectores.py) no exigen la firma de
    encabezados del informe modelo: la receta ya describe su estructura. Devuelve solo los problemas."""
    from .excel import norm
    problemas = []
    con_receta = {s.hoja for s in SECCIONES if s.clave in (lectores or {})}
    secciones_por_clave = {}
    for s in SECCIONES:
        secciones_por_clave.setdefault(s.hoja, []).append(s.clave)
    for clave, v in hojas.items():
        if clave in con_receta and all(sc in (lectores or {}) for sc in secciones_por_clave.get(clave, [])):
            continue
        titulos = [x[0] for x in v] if isinstance(v, list) else [v] if v else []
        for t in titulos:
            if t not in lib.valores.sheetnames:
                continue                                  # eso lo reporta el mapa como desajuste
            cab = _texto_cabecera(lib.valores[t])
            faltan = [f.replace("|", "» o «") for f in FIRMAS.get(clave, ()) if not cumple(f, cab)]
            if faltan:
                problemas.append({"clave": clave, "hoja": t, "detalle": f"La hoja «{t}» no parece ser «{clave}»: "
                                  f"no tiene en sus encabezados {', '.join('«' + f + '»' for f in faltan)}."})
    for sec, clave in (("matriz_riesgo", "MATRIZ DICTAMEN"), ("matriz_manuales", "matriz manuales")):
        filas = ((doc or {}).get(sec) or {}).get("filas") or []
        vals = [norm(f.get("valoracion")) for f in filas if f.get("valoracion")]
        if vals and sum(v in NIVELES_RIESGO for v in vals) / len(vals) < 0.6:
            problemas.append({"clave": clave, "hoja": hojas.get(clave), "detalle":
                              f"En «{hojas.get(clave)}» la columna «Valorización del riesgo» no trae niveles "
                              "(Extremo/Alto/Moderado/Bajo) en la mayoría de filas: las columnas parecen corridas."})
    return problemas
