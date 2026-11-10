"""Agente redactor: borrador de hallazgos, nivel de riesgo y recomendación a partir de la biblioteca.

La biblioteca (conocimiento/biblioteca_opiniones.yaml) guarda los textos que el revisor fiscal ya usa
(p. ej. su opinión estándar cuando el flujo de efectivo es negativo). Las CONDICIONES están en código y se
calculan con las cifras del Excel; los TEXTOS están en la biblioteca y los mejora el revisor.

El resultado es un BORRADOR: el revisor decide qué entra al informe. Las cifras se insertan desde el modelo
y, si se pule la redacción con IA, el candado de cifras descarta cualquier texto que cambie un número.
"""
from __future__ import annotations

from revisoria.excel import Periodo
from revisoria.narrativa import Redactor, pct, pesos

from .comun import Espacio, leer_yaml

BIBLIOTECA_BASE = {
    "flujo_negativo": {
        "titulo": "Flujo de efectivo negativo en el año",
        "riesgo": "ALTO",
        "texto": ("Durante el año el equivalente al efectivo llegó a {minimo} en {mes_minimo}, lo que indica que las "
                  "obligaciones superaron el disponible. La Revisoría Fiscal recomienda a la administración un flujo "
                  "de caja proyectado mensual y priorizar la recuperación de cartera."),
    },
    "monetizacion_faltante": {
        "titulo": "Fondo de imprevistos sin monetizar en su totalidad",
        "riesgo": "MEDIO",
        "texto": ("Los recursos en fiducia no cubren la reserva registrada en el patrimonio; faltan {faltante} por "
                  "monetizar. Se recomienda trasladar el valor a la fiducia del fondo de imprevistos."),
    },
    "cartera_aumenta": {
        "titulo": "Aumento de la cartera de copropietarios",
        "riesgo": "MEDIO",
        "texto": ("La cartera aumentó {variacion} frente al mes anterior y se ubica en {cartera}. Se recomienda "
                  "solicitar a la abogada el informe de gestión de cobro y revisar los acuerdos de pago."),
    },
    "concentracion_cartera": {
        "titulo": "Concentración de la cartera en pocos inmuebles",
        "riesgo": "MEDIO",
        "texto": ("{inmuebles} inmuebles concentran el {pct_concentracion} de la cartera. Se recomienda priorizar el "
                  "cobro jurídico de estos inmuebles y hacer seguimiento mensual a su estado."),
    },
    "presupuesto_sobreejecutado": {
        "titulo": "Rubros del presupuesto sobreejecutados",
        "riesgo": "MEDIO",
        "texto": ("{n_rubros} rubro(s) superan lo presupuestado por un total de {sobre}. Se recomienda que el Consejo "
                  "apruebe los traslados presupuestales y que la administración soporte cada mayor valor."),
    },
    "excedentes_sin_respaldo": {
        "titulo": "Diferencia entre excedentes y su respaldo",
        "riesgo": "MEDIO",
        "texto": ("Existe una diferencia de {dif_excedentes} entre los excedentes acumulados y los activos que los "
                  "respaldan. Se recomienda conciliarla antes de concluir sobre el respaldo total."),
    },
    "anticipos_antiguos": {
        "titulo": "Anticipos pendientes de legalizar",
        "riesgo": "MEDIO",
        "texto": ("Hay {n_anticipos} anticipo(s) a contratistas por {anticipos} pendientes de legalizar. Se recomienda "
                  "solicitar las actas de avance o la devolución de los recursos."),
    },
}


def biblioteca(esp: Espacio) -> dict:
    b = dict(BIBLIOTECA_BASE)
    b.update(leer_yaml(esp.conocimiento / "biblioteca_opiniones.yaml", {}) or {})
    return b


def condiciones(m: dict) -> dict[str, dict]:
    """Qué condiciones se cumplen este mes y con qué cifras (todas salen del modelo)."""
    c = {}
    L = m.get("liquidez") or {}
    serie = L.get("serie") or []
    if serie:
        mn = min(serie, key=lambda x: x["valor"])
        if mn["valor"] < 0:
            c["flujo_negativo"] = {"minimo": pesos(mn["valor"]), "mes_minimo": Periodo.desde_texto(mn["periodo"]).etiqueta}
    if (L.get("mayor_valor_monetizado") or 0) < 0:
        c["monetizacion_faltante"] = {"faltante": pesos(abs(L["mayor_valor_monetizado"]))}
    C = m.get("cartera") or {}
    ant = C.get("anterior") or {}
    if ant and C.get("total", 0) > ant.get("total", 0):
        c["cartera_aumenta"] = {"variacion": pesos(C["total"] - ant["total"]), "cartera": pesos(C["total"])}
    if C.get("rangos") and C["rangos"][0]["pct_cartera"] >= 0.6:
        r = C["rangos"][0]
        c["concentracion_cartera"] = {"inmuebles": r["inmuebles"], "pct_concentracion": pct(r["pct_cartera"], 0)}
    sobre = [r for r in (m.get("presupuesto") or {}).get("rubros", []) if (r.get("sobre_ejecutado") or 0) > 0]
    if sobre:
        c["presupuesto_sobreejecutado"] = {"n_rubros": len(sobre), "sobre": pesos(sum(r["sobre_ejecutado"] for r in sobre))}
    P = m.get("patrimonio") or {}
    if P and P.get("total_disponibilidad") is not None and abs(P["total_disponibilidad"] - P.get("total_excedentes", 0)) > 1:
        c["excedentes_sin_respaldo"] = {"dif_excedentes": pesos(abs(P["total_disponibilidad"] - P["total_excedentes"]))}
    A = (m.get("anticipos") or {}).get("items") or []
    if A:
        c["anticipos_antiguos"] = {"n_anticipos": len(A), "anticipos": pesos(sum(a["valor"] for a in A))}
    return c


def borrador(esp: Espacio, m: dict, ficha: dict, hallazgos_excel: list, usar_llm: bool = False) -> tuple[str, list[dict]]:
    bib = biblioteca(esp)
    red = Redactor(usar_llm=usar_llm)
    items = []
    for clave, datos in condiciones(m).items():
        e = bib.get(clave)
        if not e:
            continue
        try:
            texto = e["texto"].format(**datos)
        except (KeyError, IndexError, ValueError) as ex:    # plantilla del revisor con un marcador inválido
            texto = BIBLIOTECA_BASE[clave]["texto"].format(**datos) if clave in BIBLIOTECA_BASE else ""
            texto += f"\n\n_(Se usó el texto base: la plantilla de la biblioteca tiene un marcador inválido {ex}.)_"
        texto = red.pulir(texto, contexto=e["titulo"])
        items.append({"clave": clave, "titulo": e["titulo"], "riesgo": e.get("riesgo", "MEDIO"), "texto": texto,
                      "cifras": datos})
    meta = m["meta"]
    lineas = [f"# Borrador de hallazgos y opinión — {ficha.get('copropiedad', '')}",
              f"Periodo {meta['periodo_etiqueta']} · Informe {ficha.get('numero_informe', '')}", "",
              "> BORRADOR generado por el agente redactor. Las cifras salen del Excel; la opinión, el nivel de riesgo "
              "y la recomendación finales los decide el revisor fiscal.", ""]
    for i, it in enumerate(items, 1):
        lineas += [f"## {i}. {it['titulo']}", f"**Nivel de riesgo propuesto:** {it['riesgo']}", "", it["texto"], ""]
    altas = [h for h in hallazgos_excel if h.severidad in ("ALTA",)]
    if altas:
        lineas += ["## Hallazgos del Excel que el revisor debe resolver", ""]
        lineas += [f"- **{h.regla}** {h.titulo}: {h.detalle}" for h in altas]
        lineas.append("")
    if red.registro:
        lineas += ["---", "Redacción con IA: " + ", ".join(r["estado"] for r in red.registro)]
    return "\n".join(lineas), items
