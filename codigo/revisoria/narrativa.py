"""Redacción de textos del informe.

Por defecto los textos se generan con plantillas determinísticas a partir de las
cifras ya calculadas. Opcionalmente (--llm) se pueden pulir con Claude, pero con un
CANDADO: si el texto devuelto contiene una cifra que no estaba en el original, se
descarta y se conserva la versión determinística.
"""
from __future__ import annotations

import os
import re

from .excel import Periodo

MODELO = os.environ.get("REVISORIA_MODELO", "claude-sonnet-5-5")
RX_NUM = re.compile(r"\d[\d\.,]*\d|\d")


def pesos(v: float, signo_parentesis: bool = False) -> str:
    from decimal import ROUND_HALF_UP, Decimal
    s = f"{int(Decimal(str(abs(v))).quantize(Decimal('1'), rounding=ROUND_HALF_UP)):,}".replace(",", ".")
    if v < 0:
        return f"(${s})" if signo_parentesis else f"-${s}"
    return f"${s}"


def millones(v: float) -> str:
    return f"${v / 1e6:,.1f} M".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(v: float, dec: int = 1) -> str:
    return f"{v * 100:.{dec}f}%".replace(".", ",")


def _cifras(texto: str) -> list[str]:
    """Conserva signo, separadores, porcentaje y cantidad de apariciones."""
    patron = re.compile(
        r"(?<!\w)(?:-?\s*\$\s*|\$\s*-?\s*|-?\s*)"
        r"(?:\(\s*)?\d+(?:\.\d{3})*(?:,\d+)?(?:\s*\))?\s*%?"
    )
    return [re.sub(r"\s+", "", n) for n in patron.findall(texto)]


class Redactor:
    def __init__(self, usar_llm: bool = False):
        self.usar_llm = usar_llm and bool(os.environ.get("ANTHROPIC_API_KEY"))
        self.cliente = None
        self.registro: list[dict] = []
        if self.usar_llm:
            try:
                import anthropic
                self.cliente = anthropic.Anthropic()
            except ImportError:
                self.usar_llm = False

    def pulir(self, texto: str, contexto: str = "") -> str:
        """Mejora redacción sin tocar cifras. Si algo falla, devuelve el texto original."""
        if not self.usar_llm or not texto.strip():
            return texto
        prompt = (
            "Eres revisor fiscal de propiedad horizontal en Colombia. Mejora la redacción del siguiente "
            "texto para un informe formal dirigido al Consejo de Administración. Reglas estrictas: no agregues, "
            "quites ni cambies ninguna cifra, fecha, porcentaje o nombre; no agregues opiniones nuevas; máximo "
            "el mismo largo. Devuelve solo el texto.\n\n"
            f"Contexto: {contexto}\n\nTexto:\n{texto}")
        try:
            r = self.cliente.messages.create(model=MODELO, max_tokens=600,
                                             messages=[{"role": "user", "content": prompt}])
            nuevo = "".join(b.text for b in r.content if getattr(b, "type", "") == "text").strip()
        except Exception as e:                     # sin red, sin cuota, etc.
            self.registro.append({"contexto": contexto, "estado": f"error: {e}"})
            return texto
        if not nuevo or _cifras(nuevo) != _cifras(texto):
            self.registro.append({"contexto": contexto, "estado": "descartado: cambió cifras"})
            return texto
        self.registro.append({"contexto": contexto, "estado": "pulido"})
        return nuevo


# ---------------------------------------------------- textos determinísticos

def texto_liquidez(m: dict) -> str:
    L = m["liquidez"]
    serie = L["serie"]
    minimo = min(serie, key=lambda x: x["valor"]) if serie else None
    t = (f"El equivalente al efectivo positivo cierra {m['meta']['periodo_etiqueta'].lower()} en "
         f"{pesos(L['total'])}, frente a {pesos(L['total_cierre'])} al cierre de "
         f"{m['meta']['cierre_corto']}.")
    if minimo and minimo["valor"] < 0:
        t += (f" El punto más bajo del año fue {pesos(minimo['valor'])} en {Periodo.desde_texto(minimo['periodo']).corto}, "
              "cuando las obligaciones superaron el disponible.")
    return t


def texto_monetizacion(m: dict) -> str:
    L = m["liquidez"]
    v = L["mayor_valor_monetizado"]
    if v is None:
        return ""
    if v >= 0:
        return (f"Los recursos en la fiducia del fondo de imprevistos cubren la reserva y sus rendimientos "
                f"registrados en el patrimonio, con un mayor valor monetizado de {pesos(v)}.")
    return (f"Los recursos en la fiducia NO cubren la reserva y rendimientos registrados en el patrimonio: "
            f"faltan {pesos(abs(v))} por monetizar.")


def titulo_cartera(m: dict) -> str:
    C = m["cartera"]
    return f"CARTERA ALCANZA {pesos(C['total'])} CON {C['inmuebles']} INMUEBLES MOROSOS"


def rango_principal(C: dict) -> dict:
    """El tramo de deuda más alto que tiene inmuebles: si nadie debe más de $3.000.000, el título habla del
    siguiente tramo (no de «0 inmuebles concentran el 0%»)."""
    return next((r for r in C["rangos"] if r["inmuebles"]), C["rangos"][0])


def titulo_concentracion(m: dict) -> str:
    r = rango_principal(m["cartera"])
    return f"{r['inmuebles']} INMUEBLES CONCENTRAN EL {round(r['pct_cartera'] * 100)}% DE LA CARTERA TOTAL"


def texto_cartera(m: dict) -> str:
    C, D = m["cartera"], m["deudores"]
    ant = C.get("anterior") or {}
    u = m["meta"].get("unidades")
    t = (f"La cartera bruta de copropietarios es {pesos(C['total'])} en {C['inmuebles']} inmuebles"
         + (f" ({pct(C['inmuebles'] / u)} de las {u} unidades)" if u else "") + ". "
         f"Neta de consignaciones por identificar, deterioro y cuentas correctoras queda en {pesos(D['neto'])}.")
    if ant:
        var = C["total"] - ant["total"]
        t += (f" Frente al mes anterior la cartera {'disminuyó' if var < 0 else 'aumentó'} {pesos(abs(var))} "
              f"y el número de inmuebles morosos pasó de {ant['inmuebles']} a {C['inmuebles']}.")
    return t


def texto_cxp(m: dict) -> str:
    X = m["cxp"]
    top = X["items"][0]
    return (f"Las cuentas por pagar suman {pesos(X['total'])} con {len(X['items'])} acreedores. "
            f"El principal es {top['tercero']} con {pesos(top['valor'])} ({pct(top['valor'] / X['total'])}).")


def texto_excedentes(m: dict) -> str:
    P = m["patrimonio"]
    if P.get("total_disponibilidad") is None or not P.get("disponibilidad"):
        return (f"Los excedentes acumulados suman {pesos(P['total_excedentes'])}. La hoja no trae el análisis de los "
                "activos que los respaldan.")
    dif = P["total_disponibilidad"] - P["total_excedentes"]
    principal = max((d for d in P["disponibilidad"] if not d["es_total"]), key=lambda d: d["actual"])
    t = (f"Los excedentes acumulados suman {pesos(P['total_excedentes'])}. Los activos que los respaldan "
         f"suman {pesos(P['total_disponibilidad'])}, concentrados principalmente en "
         f"«{principal['concepto'].lower()}» ({pct(principal['pct'])}).")
    if abs(dif) > 1:
        t += (f" Existe una diferencia de {pesos(abs(dif))} entre ambos valores, que se encuentra en "
              "revisión y debe conciliarse antes de concluir sobre el respaldo total.")
    else:
        t += " El 100% de los excedentes se encuentra respaldado."
    return t
