"""Audita un informe PPTX ya elaborado contra el Excel del periodo.

Detecta: cifras de títulos que no coinciden, cifras de meses anteriores,
texto residual oculto detrás de imágenes, periodos mal digitados, citas
normativas dudosas y diapositivas construidas solo con capturas.
"""
from __future__ import annotations

import re

from pptx import Presentation
from pptx.util import Emu

from .excel import MESES_CORTOS, Libro, Periodo, buscar_fila, columna_mes, es_numero, norm
from .validador import Hallazgo, pesos

RX_PESOS = re.compile(r"\$\s?\(?\s?(\d{1,3}(?:[\.,]\d{3})+)")

# Citas que en informes de PH suelen estar mal referenciadas. Se marcan para VERIFICAR, no como error.
CITAS = [
    (r"ley 1561 de 2012", "La Ley 1561 de 2012 trata la titulación de la posesión; las actas de asamblea y consejo "
                          "en PH se regulan en la Ley 675 de 2001 (art. 47)."),
    (r"ley 1563 de 2012", "La Ley 1563 de 2012 es el Estatuto de Arbitraje; el libro de registro de propietarios es "
                          "función del administrador según la Ley 675 de 2001 (art. 51, num. 2)."),
    (r"decreto 2649 de 1993", "El Decreto 2649 de 1993 fue reemplazado por el marco técnico NIIF "
                              "(DUR 2420 de 2015) para la preparación de información financiera."),
]


def _area(sh):
    return max(sh.width or 0, 0) * max(sh.height or 0, 0)


def _cubierto(texto, imagen, umbral=0.85):
    x1, y1 = max(texto.left, imagen.left), max(texto.top, imagen.top)
    x2 = min(texto.left + texto.width, imagen.left + imagen.width)
    y2 = min(texto.top + texto.height, imagen.top + imagen.height)
    if x2 <= x1 or y2 <= y1 or not _area(texto):
        return False
    return (x2 - x1) * (y2 - y1) / _area(texto) >= umbral


def _textos(shape):
    if shape.has_text_frame:
        return shape.text_frame.text
    if getattr(shape, "has_table", False) and shape.has_table:
        return "\n".join(c.text for r in shape.table.rows for c in r.cells)
    return ""


def leer_diapositivas(ruta: str):
    prs = Presentation(ruta)
    ancho, alto = prs.slide_width, prs.slide_height
    salida = []
    for i, s in enumerate(prs.slides, 1):
        shapes = [sh for sh in s.shapes if sh.width is not None]
        imagenes = [(k, sh) for k, sh in enumerate(shapes) if sh.shape_type == 13]
        visibles, ocultos, graficos = [], [], []
        for k, sh in enumerate(shapes):
            if getattr(sh, "has_chart", False) and sh.has_chart:
                titulo = ""
                if sh.chart.has_title:
                    titulo = sh.chart.chart_title.text_frame.text.strip()
                    if titulo:
                        visibles.append(titulo)
                for ns, serie in enumerate(sh.chart.series):
                    valores = []
                    try:
                        valores = [float(v) for v in serie.values if v is not None]
                    except (TypeError, ValueError, AttributeError):
                        pass
                    graficos.append({"serie": ns, "nombre": str(getattr(serie, "name", "") or ""),
                                      "valores": valores})
            t = _textos(sh).strip()
            if not t:
                continue
            # cubierto por una imagen que está por encima (z-order posterior) y que no es fondo de toda la diapositiva
            tapa = any(j > k and _cubierto(sh, im) and _area(im) < 0.97 * ancho * alto for j, im in imagenes)
            (ocultos if tapa else visibles).append(t)
        area_img = sum(_area(im) for _, im in imagenes if _area(im) < 0.97 * ancho * alto
                       and im.width > ancho * 0.15)
        tiene_nativo = any(getattr(sh, "has_table", False) and sh.has_table or getattr(sh, "has_chart", False)
                           and sh.has_chart for sh in shapes)
        salida.append({"n": i, "visible": visibles, "oculto": ocultos, "graficos": graficos,
                       "titulo": visibles[0] if visibles else "",
                       "captura": area_img > 0.25 * ancho * alto and not tiene_nativo})
    return salida


def _montos(texto: str) -> list[int]:
    texto = re.sub(r"(?i)(mayores|superiores|menores) a \$\s?[\d\.,]+", "", texto)   # umbrales, no cifras
    return [int(re.sub(r"[\.,]", "", m)) for m in RX_PESOS.findall(texto)]


def _valores_columna(lib: Libro, per: Periodo) -> set[int]:
    """Todas las cifras del mes 'per' en las hojas con columnas mensuales."""
    vals = set()
    bloques = [("CUADRO EFECTIVO", "Análisis de liquidez"), ("CUADRO EFECTIVO", "Análisis Monetización"),
               ("PATRIMONIO", "RESULTADO DEL EJERCICIO"), ("PATRIMONIO", "Excedentes acumulados"),
               ("PATRIMONIO", "Disponible de excedentes"), ("GRAFICA DEUDORES", "Detalle"),
               ("CUADRO RESERVA", "Reservas patrimonio")]
    for hoja, etiqueta in bloques:
        try:
            ws = lib.hoja(hoja)
        except KeyError:                       # hoja opcional que el cliente no tiene este mes
            continue
        f = buscar_fila(ws, etiqueta)
        if not f:
            continue
        c = columna_mes(ws, f, per, f + 1)
        if not c:
            continue
        for r in range(f + 1, f + 25):
            v = ws.cell(r, c).value
            if es_numero(v) and abs(v) >= 100_000:
                vals.add(round(abs(v)))
    return vals


def _valores_modelo(m: dict) -> set[int]:
    vals = set()

    def rec(o):
        if isinstance(o, dict):
            for v in o.values():
                rec(v)
        elif isinstance(o, list):
            for v in o:
                rec(v)
        elif es_numero(o) and abs(o) >= 100_000:
            vals.add(round(abs(o)))
    rec(m)
    return vals


def auditar(ruta_pptx: str, m: dict, lib: Libro, p: Periodo) -> list[Hallazgo]:
    diap = leer_diapositivas(ruta_pptx)
    actuales = _valores_modelo(m) | _valores_columna(lib, p)
    anteriores = _valores_columna(lib, p.anterior()) - actuales
    if (m.get("cartera") or {}).get("anterior"):
        anteriores.add(round(m["cartera"]["anterior"]["total"]))
    out = []

    for d in diap:
        visible = "\n".join(d["visible"])
        tn = norm(visible)
        # 1. Afirmaciones de los títulos
        mm = re.search(r"cartera alcanza \$\s?([\d\.]+) con (\d+) inmuebles", tn)
        if mm and m.get("cartera"):
            v, n = int(mm.group(1).replace(".", "")), int(mm.group(2))
            C = m["cartera"]
            if abs(v - C["total"]) > 1 or n != C["inmuebles"]:
                prev = C.get("anterior") or {}
                pista = (" Las cifras del título coinciden con el mes anterior."
                         if prev and abs(v - prev.get("total", 0)) <= 1 and n == prev.get("inmuebles") else "")
                out.append(Hallazgo("P01", "ALTA", "Informe PPTX", f"Diapositiva {d['n']}: el título de cartera no coincide",
                                    f"El título dice {pesos(v)} y {n} inmuebles; el Excel del periodo da "
                                    f"{pesos(C['total'])} y {C['inmuebles']} inmuebles.{pista}",
                                    [f"Diapositiva {d['n']}", C["fuente_total"]], origen="pptx"))
        mm = re.search(r"(\d+) inmuebles concentran el (\d+)\s?%", tn)
        if mm and (m.get("cartera") or {}).get("rangos"):
            from .narrativa import rango_principal
            r0 = rango_principal(m["cartera"])
            if int(mm.group(1)) != r0["inmuebles"] or abs(int(mm.group(2)) - round(r0["pct_cartera"] * 100)) > 1:
                out.append(Hallazgo("P01", "ALTA", "Informe PPTX", f"Diapositiva {d['n']}: concentración de cartera",
                                    f"El título dice {mm.group(1)} inmuebles / {mm.group(2)}%; el Excel da "
                                    f"{r0['inmuebles']} / {r0['pct_cartera']:.0%}.", [f"Diapositiva {d['n']}"], origen="pptx"))
        mm = re.search(r"cuentas por pagar \$\s?([\d\.]+)", tn)
        if mm and m.get("cxp") and m["cxp"].get("total") is not None:
            v = int(mm.group(1).replace(".", ""))
            if abs(v - m["cxp"]["total"]) > 1:
                out.append(Hallazgo("P01", "ALTA", "Informe PPTX", f"Diapositiva {d['n']}: total de cuentas por pagar",
                                    f"El título dice {pesos(v)}; el Excel da {pesos(m['cxp']['total'])}.",
                                    [f"Diapositiva {d['n']}"], origen="pptx"))
        mm = re.search(r"genera excedente de \$\s?([\d\.]+)", tn)
        if mm and m.get("patrimonio"):
            v = int(mm.group(1).replace(".", ""))
            if abs(v - round(m["patrimonio"]["resultado"])) > 1:
                out.append(Hallazgo("P01", "ALTA", "Informe PPTX", f"Diapositiva {d['n']}: resultado del ejercicio",
                                    f"El título dice {pesos(v)}; el Excel da {pesos(m['patrimonio']['resultado'])}.",
                                    [f"Diapositiva {d['n']}"], origen="pptx"))

        # 2. Cifras visibles que solo existen en el mes anterior
        viejas = sorted({v for v in _montos(visible) if v in anteriores})
        if viejas:
            out.append(Hallazgo("P02", "ALTA", "Informe PPTX", f"Diapositiva {d['n']}: cifras del mes anterior",
                                "Estas cifras visibles coinciden con la columna del mes anterior y no con la del periodo: "
                                + ", ".join(pesos(v) for v in viejas) + ".", [f"Diapositiva {d['n']}"],
                                "Actualizar con las cifras del periodo.", origen="pptx"))

        # 3. Texto oculto detrás de imágenes
        oculto = "\n".join(d["oculto"])
        montos_ocultos = _montos(oculto)
        if montos_ocultos:
            de_otro_mes = sorted({v for v in montos_ocultos if v in anteriores or v not in actuales})
            sev = "MEDIA" if de_otro_mes else "BAJA"
            det = (f"Hay {len(d['oculto'])} cuadro(s) de texto tapados por una imagen, con {len(montos_ocultos)} cifra(s)."
                   + (f" Algunas no corresponden al periodo: {', '.join(pesos(v) for v in de_otro_mes[:6])}." if de_otro_mes else ""))
            out.append(Hallazgo("P03", sev, "Informe PPTX", f"Diapositiva {d['n']}: texto residual oculto",
                                det + " Al copiar, buscar o exportar el informe, ese texto sigue presente.",
                                [f"Diapositiva {d['n']}"], "Eliminar los cuadros de texto que quedaron debajo.", origen="pptx"))

        # 4. Periodos mal digitados (ej. «Jul-27» en un informe de 2026)
        todo = visible
        for mes, anio in re.findall(r"\b(" + "|".join(MESES_CORTOS) + r")-(\d{2})\b", norm(todo)):
            if int(anio) > p.anio % 100:
                out.append(Hallazgo("P04", "BAJA", "Informe PPTX", f"Diapositiva {d['n']}: periodo «{mes.capitalize()}-{anio}»",
                                    f"El informe es de {p.etiqueta} y aparece un periodo posterior.",
                                    [f"Diapositiva {d['n']}"], "Corregir el periodo.", origen="pptx"))

        # 5. Citas normativas para verificar
        for patron, nota in CITAS:
            if re.search(patron, tn):
                out.append(Hallazgo("P05", "MEDIA", "Normatividad", f"Diapositiva {d['n']}: verificar cita normativa",
                                    nota, [f"Diapositiva {d['n']}"], "Confirmar la norma aplicable antes de emitir.",
                                    origen="pptx"))

    capturas = [d["n"] for d in diap if d["captura"]]
    if capturas:
        out.append(Hallazgo("P06", "INFO", "Informe PPTX",
                            f"{len(capturas)} de {len(diap)} diapositivas presentan los datos como imágenes",
                            "Las tablas pegadas como captura no se actualizan con el Excel y no se pueden verificar "
                            "automáticamente. Diapositivas: " + ", ".join(map(str, capturas)) + ".",
                            [], "El generador produce tablas y gráficos nativos.", origen="pptx"))
    # agrupa el texto oculto de todas las diapositivas en un solo hallazgo
    ocultos = [h for h in out if h.regla == "P03"]
    if len(ocultos) > 1:
        out = [h for h in out if h.regla != "P03"]
        sev = "MEDIA" if any(h.severidad == "MEDIA" for h in ocultos) else "BAJA"
        nums = [h.ubicacion[0].split()[-1] for h in ocultos]
        out.append(Hallazgo("P03", sev, "Informe PPTX", f"Texto residual oculto en {len(ocultos)} diapositivas ({', '.join(nums)})",
                            " ".join(f"Diap. {n}: {h.detalle.split(' Al copiar')[0]}" for n, h in zip(nums, ocultos))
                            + " Al copiar, buscar o exportar el informe, ese texto sigue presente.",
                            [u for h in ocultos for u in h.ubicacion],
                            "Eliminar los cuadros de texto que quedaron debajo de las imágenes.", origen="pptx"))
    # deduplica citas repetidas en varias diapositivas
    vistos, final = set(), []
    for h in out:
        clave = (h.regla, h.detalle) if h.regla == "P05" else id(h)
        if clave in vistos:
            for f in final:
                if f.regla == "P05" and f.detalle == h.detalle:
                    f.ubicacion += h.ubicacion
                    f.titulo = "Diapositivas " + ", ".join(u.split()[-1] for u in f.ubicacion) + ": verificar cita normativa"
            continue
        vistos.add(clave)
        final.append(h)
    return final
