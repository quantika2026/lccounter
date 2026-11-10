"""Agente re-auditor: revisa de forma independiente el informe YA GENERADO.

No reutiliza el camino del generador: abre el PowerPoint terminado, lee lo que un lector vería y lo
contrasta contra el Excel. Corre mientras el profesional valida el informe anterior de la cola.

Reglas:
  R01  Afirmaciones de títulos vs Excel (cartera, concentración, cuentas por pagar, resultado)
  R02  Cifras visibles sin origen rastreable en el Excel o en las notas del revisor
  R03  Cifras del mes anterior o periodos fuera de rango
  R04  Secciones esperadas que no aparecen en el informe
  R05  Diapositivas que requieren opinión del revisor y no la tienen
  R06  Hallazgos ALTA del Excel que el revisor debe decidir antes de aprobar
  R07  Hallazgos abiertos del mes anterior (seguimiento)
"""
from __future__ import annotations

import re

from revisoria.auditor_pptx import _montos, _valores_columna, auditar, leer_diapositivas
from revisoria.excel import Periodo, es_numero, norm
from revisoria.narrativa import pesos
from revisoria.validador import Hallazgo, ordenar

SECCIONES = [("liquidez", "Liquidez", lambda m: m.get("liquidez")),
             ("cartera", "Cartera", lambda m: (m.get("cartera") or {}).get("total")),
             ("cuentas por pagar", "Cuentas por pagar", lambda m: m.get("cxp")),
             ("resultado del ejercicio", "Resultado del ejercicio", lambda m: m.get("patrimonio")),
             ("excedentes", "Excedentes acumulados", lambda m: m.get("patrimonio")),
             ("anticipos", "Anticipos", lambda m: (m.get("anticipos") or {}).get("items"))]


def _numeros(o, out: set, minimo=1000):
    if isinstance(o, dict):
        for v in o.values():
            _numeros(v, out, minimo)
    elif isinstance(o, list):
        for v in o:
            _numeros(v, out, minimo)
    elif isinstance(o, str):
        for v in _montos(o) + [int(re.sub(r"\D", "", x)) for x in re.findall(r"\d{1,3}(?:\.\d{3})+", o)]:
            out.add(v)
    elif es_numero(o) and not isinstance(o, bool) and abs(o) >= minimo:
        out.add(round(abs(o)))


def _sumas_listas(o, out: set):
    """Totales y subtotales que el generador calcula sobre las listas (total de la tabla, agrupación por edad)."""
    if isinstance(o, dict):
        for v in o.values():
            _sumas_listas(v, out)
    elif isinstance(o, list) and o and all(isinstance(x, dict) for x in o):
        for campo in {k for x in o for k, v in x.items() if es_numero(v) and not isinstance(v, bool)}:
            out.add(round(abs(sum(x.get(campo) or 0 for x in o if es_numero(x.get(campo))))))
            # subtotal de lo que supera el umbral de presentación ($1.000.000)
            out.add(round(abs(sum(x[campo] for x in o if es_numero(x.get(campo)) and x[campo] > 1_000_000))))
            for clave in {k for x in o for k, v in x.items() if isinstance(v, str)}:
                grupos = {}
                for x in o:
                    g = str(x.get(clave, ""))
                    grupos.setdefault(g, 0)
                    grupos[g] += x.get(campo) or 0 if es_numero(x.get(campo)) else 0
                if 1 < len(grupos) <= 8:
                    out |= {round(abs(v)) for v in grupos.values()}
        if any("edad" in x for x in o):                      # «Mayor a 90 días» vs «Hasta 90 días»
            viejo = sum(x.get("valor", 0) for x in o if "90" in str(x.get("edad")) and "ayor" in str(x.get("edad")))
            out |= {round(viejo), round(sum(x.get("valor", 0) for x in o) - viejo)}
        for x in o:
            _sumas_listas(x, out)


def rastreables(m: dict, ficha: dict, lib, p: Periodo, extra=None) -> set[int]:
    """Cifras que tienen origen: modelo, columnas del mes, notas del revisor y diferencias entre totales."""
    base: set[int] = set()
    _numeros(m, base)
    _numeros(ficha.get("notas") or {}, base)
    if extra:                                            # secciones regulatoria y de control interno
        _numeros(extra, base)
        _sumas_listas(extra, base)
    base |= _valores_columna(lib, p)
    _sumas_listas(m, base)
    totales = sorted({v for v in base if v >= 100_000})
    for a in totales:                                   # variaciones mes a mes y diferencias de cuadre
        for b in totales:
            if a > b:
                base.add(a - b)
    # umbrales de los rangos de cartera ("Mayores a $3.000.000", "$1.000.001 a $3.000.000")
    for r in (m.get("cartera") or {}).get("rangos", []):
        _numeros(r.get("rango", ""), base, 0)
    return base


def reauditar(ruta_pptx, m: dict, lib, p: Periodo, ficha: dict, hallazgos_excel: list,
              abiertos_anteriores: list | None = None, extra=None) -> list[Hallazgo]:
    out: list[Hallazgo] = []
    # R01 / R03: reutiliza las comprobaciones de lectura del PPTX (títulos, mes anterior, periodos)
    for h in auditar(str(ruta_pptx), m, lib, p):
        if h.regla in ("P01", "P02", "P04"):
            h.regla = {"P01": "R01", "P02": "R03", "P04": "R03"}[h.regla]
            h.categoria = "Re-auditoría"
            out.append(h)

    diap = leer_diapositivas(str(ruta_pptx))
    # R02 cifras sin origen
    ok = rastreables(m, ficha, lib, p, extra)
    for d in diap:
        texto = "\n".join(d["visible"])
        cifras_grafico = {round(abs(v)) for g in d.get("graficos", []) for v in g.get("valores", []) if abs(v) >= 1000}
        sin = sorted({v for v in set(_montos(texto)) | cifras_grafico
                      if v >= 1000 and not ({v - 1, v, v + 1} & ok)})
        if sin:
            out.append(Hallazgo("R02", "ALTA", "Re-auditoría", f"Diapositiva {d['n']}: cifras sin origen rastreable",
                                "No se encontró en el Excel ni en las notas del revisor: "
                                + ", ".join(pesos(v) for v in sin[:8]) + ".",
                                [f"Diapositiva {d['n']}"], "Verificar de dónde sale la cifra antes de aprobar."))

    # R04 secciones esperadas
    titulos = norm(" ".join((d["visible"] or [""])[0] for d in diap))
    for clave, nombre, hay in SECCIONES:
        if hay(m) and clave not in titulos:
            out.append(Hallazgo("R04", "MEDIA", "Re-auditoría", f"No aparece la sección «{nombre}»",
                                "El Excel tiene datos para esta sección pero el informe no la muestra.", [],
                                "Revisar la generación o confirmar que se omite a propósito."))

    # R05 opiniones pendientes
    notas = ficha.get("notas") or {}
    if (m.get("anticipos") or {}).get("items") and not str(notas.get("opinion_anticipos") or "").strip():
        out.append(Hallazgo("R05", "INFO", "Re-auditoría", "Sin opinión del revisor sobre anticipos",
                            "Hay anticipos pendientes de legalizar y la ficha no trae opinión para este mes.", [],
                            "Use el borrador de opinión como punto de partida si aplica."))
    if not str(notas.get("concentracion_cartera") or "").strip():
        out.append(Hallazgo("R05", "INFO", "Re-auditoría", "Sin nota sobre la gestión de cobro de cartera",
                            "La diapositiva de concentración de cartera sale sin comentario del revisor.", []))

    # R06 hallazgos ALTA del Excel
    altas = [h for h in hallazgos_excel if h.severidad == "ALTA"]
    if altas:
        out.append(Hallazgo("R06", "ALTA", "Re-auditoría", f"{len(altas)} hallazgo(s) ALTA del Excel requieren decisión",
                            "; ".join(f"{h.regla} {h.titulo}" for h in altas[:6]) + ".", [],
                            "Corregir el Excel y volver a correr, o aprobar con justificación expresa."))

    # R07 seguimiento a hallazgos abiertos del mes anterior
    if abiertos_anteriores:
        out.append(Hallazgo("R07", "INFO", "Seguimiento", f"{len(abiertos_anteriores)} hallazgo(s) abiertos del mes anterior",
                            "; ".join(h.get("titulo", "") for h in abiertos_anteriores[:6]) + ".", [],
                            "Confirmar en la visita si se cerraron e incluir su estado en el informe."))

    for h in out:
        h.origen = "reauditoria"
    return ordenar(out)
