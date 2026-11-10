"""Pruebas de regresión del MVP con el caso Picasso agosto 2026.

Las cifras esperadas se leen del Excel por coordenada (camino independiente del
extractor, que busca por etiquetas) y se comparan con lo que quedó en el PPTX.

Ejecutar:  python -m pytest -q tests
Requiere:  variable REVISORIA_DATOS con la carpeta que contiene el Excel y el PPTX
           (por defecto ../lccounter-datos).
"""
import os
import re
import sys
from pathlib import Path

import openpyxl
import pytest
from pptx import Presentation

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

DATOS = Path(os.environ.get("REVISORIA_DATOS", RAIZ.parent / "lccounter-datos"))
EXCEL = DATOS / "EXCEL PICASSO AGOSTO 2026 .xlsx"
PPTX = DATOS / "Inf 006-2026 PICASSO AGOSTO 2026.pptx"

pytestmark = pytest.mark.skipif(not EXCEL.exists(), reason=f"No está el Excel de prueba en {DATOS}")


@pytest.fixture(scope="module")
def corrida(tmp_path_factory):
    from generar_informe import main
    salida = tmp_path_factory.mktemp("salida")
    args = ["--excel", str(EXCEL), "--config", str(RAIZ / "config/picasso.yaml"), "--salida", str(salida)]
    if PPTX.exists():
        args += ["--revisar-pptx", str(PPTX)]
    assert main(args) == 0
    deck = next(salida.glob("Seccion_Financiera_*.pptx"))
    return salida, Presentation(deck)


@pytest.fixture(scope="module")
def xl():
    return openpyxl.load_workbook(EXCEL, data_only=True)


def _todo_el_texto(prs):
    partes = []
    for s in prs.slides:
        for sh in s.shapes:
            if sh.has_text_frame:
                partes.append(sh.text_frame.text)
            if getattr(sh, "has_table", False) and sh.has_table:
                partes += [c.text for r in sh.table.rows for c in r.cells]
    return "\n".join(partes)


def _p(v):
    return f"{abs(round(v)):,.0f}".replace(",", ".")


def test_cifras_clave_estan_en_el_informe(corrida, xl):
    _, prs = corrida
    texto = _todo_el_texto(prs)
    esperadas = {
        "Total equivalente al efectivo": xl["CUADRO EFECTIVO"]["C24"].value,
        "Resultado del ejercicio": xl["PATRIMONIO"]["C5"].value,
        "Total ingresos": xl["PATRIMONIO"]["C3"].value,
        "Total gastos": xl["PATRIMONIO"]["C4"].value,
        "Total excedentes": xl["PATRIMONIO"]["C10"].value,
        "Disponibilidad de excedentes": xl["PATRIMONIO"]["C20"].value,
        "Cuentas por pagar": xl["CXP "]["D95"].value,
        "Cartera total": xl["Rangos de Cartera"]["C7"].value,
        "Deudores netos": xl["GRAFICA DEUDORES PROPIETARI"]["E11"].value,
        "Mayor valor monetizado": xl["CUADRO EFECTIVO"]["C31"].value,
        "Anticipos 1330": xl["ANTIC. E INGRESOS X C"]["D167"].value,
    }
    faltan = {k: _p(v) for k, v in esperadas.items() if _p(v) not in texto}
    assert not faltan, f"Cifras que no aparecen en el PPTX: {faltan}"


def test_titulos_calculados(corrida, xl):
    _, prs = corrida
    texto = _todo_el_texto(prs)
    assert f"CARTERA ALCANZA $ {_p(xl['Rangos de Cartera']['C7'].value)} CON 41 INMUEBLES" in texto.replace("$1", "$ 1")
    assert "7 INMUEBLES CONCENTRAN EL 85%" in texto


def test_total_presupuesto_recalculado(corrida, xl):
    _, prs = corrida
    ws = xl["PRESUPUESTO "]
    total = sum((ws[f"E{r}"].value or 0) - (ws[f"D{r}"].value or 0) for r in range(4, 13)
                if ws[f"B{r}"].value and (ws[f"E{r}"].value or 0) - (ws[f"D{r}"].value or 0) > 1_000_000)
    assert _p(total) in _todo_el_texto(prs)


def test_no_quedan_cifras_del_mes_anterior(corrida):
    _, prs = corrida
    texto = _todo_el_texto(prs)
    for vieja in ("170.547.727", "416.391.692", "8.909.820", "3.666.667"):
        assert vieja not in texto


def test_tablas_y_graficos_nativos(corrida):
    _, prs = corrida
    tablas = sum(1 for s in prs.slides for sh in s.shapes if getattr(sh, "has_table", False) and sh.has_table)
    graficos = sum(1 for s in prs.slides for sh in s.shapes if getattr(sh, "has_chart", False) and sh.has_chart)
    assert tablas >= 10 and graficos >= 4


def test_notas_con_fuentes(corrida):
    _, prs = corrida
    con_fuente = [s for s in prs.slides if s.has_notes_slide and "Fuentes" in s.notes_slide.notes_text_frame.text]
    assert len(con_fuente) >= 8


def test_hallazgos_esperados(corrida):
    import json
    salida, _ = corrida
    hs = json.loads(next(salida.glob("hallazgos_*.json")).read_text(encoding="utf-8"))
    reglas = {h["regla"] for h in hs}
    for r in ("V04", "V05", "V18", "V08", "V15"):
        assert r in reglas, f"Falta el hallazgo {r}"
    if PPTX.exists():
        assert {"P01", "P02", "P03"} <= reglas
        assert any("48 inmuebles" in h["detalle"] for h in hs if h["regla"] == "P01")


def test_candado_llm_descarta_cifras_nuevas():
    from revisoria.narrativa import Redactor

    class Falso:
        class messages:
            @staticmethod
            def create(**kw):
                class B:
                    type, text = "text", "La cartera es de $999.999.999."
                class R:
                    content = [B()]
                return R()

    r = Redactor(False)
    r.usar_llm, r.cliente = True, Falso()
    original = "La cartera es de $168.959.656."
    assert r.pulir(original) == original
    assert r.registro[-1]["estado"].startswith("descartado")


def test_observaciones_en_celdas_combinadas(corrida):
    """En 'may deudores' G6:G7, G8:G9 y G12:G14 están combinadas: la observación aplica a todo el rango."""
    _, prs = corrida
    texto = _todo_el_texto(prs)
    assert texto.count("El Manual contempla el traslado al abogado") == 1      # una sola celda fusionada
    merged = 0
    for s in prs.slides:
        for sh in s.shapes:
            if getattr(sh, "has_table", False) and sh.has_table:
                merged += sum(1 for r in sh.table.rows for c in r.cells if c.is_merge_origin)
    assert merged >= 3


def test_soportes_del_auxiliar(corrida):
    _, prs = corrida
    soportes = [s for s in prs.slides for sh in s.shapes
                if sh.has_text_frame and sh.text_frame.text.startswith("Soporte de rubros")]
    imagenes = sum(1 for s in soportes for sh in s.shapes if sh.shape_type == 13 and sh.width > 2_000_000)
    assert soportes and imagenes >= 10


def test_redondeo_como_excel(corrida):
    _, prs = corrida
    assert "12.632.881" in _todo_el_texto(prs)          # 12.632.880,5 se redondea hacia arriba, como Excel


def test_enfasis_del_excel(corrida):
    """Las frases en negrilla o rojo del Excel conservan el énfasis en el PPTX."""
    _, prs = corrida
    negrilla_roja = 0
    for s in prs.slides:
        for sh in s.shapes:
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        for p in c.text_frame.paragraphs:
                            for r in p.runs:
                                if "RADICACIÓN DE DEMANDA" in r.text and r.font.bold:
                                    negrilla_roja += 1
    assert negrilla_roja >= 1


def test_linea_de_tendencia(corrida):
    _, prs = corrida
    n = sum(1 for s in prs.slides for sh in s.shapes if getattr(sh, "has_chart", False) and sh.has_chart
            and sh.chart._chartSpace.xpath(".//c:trendline"))
    assert n >= 2
