"""Lectores declarativos y subagente estructurador.

El «entrenamiento» del estructurador son ejemplos aprobados; el primero es el informe modelo. Estas pruebas exigen que
las recetas del informe modelo reproduzcan EXACTAMENTE al lector nativo, y que el ciclo proponer → probar → corregir
funcione (con un modelo simulado, sin red)."""
import json
import os
import sys
from pathlib import Path

import openpyxl
import pytest

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))
DATOS = Path(os.environ.get("REVISORIA_DATOS", RAIZ.parent / "lccounter-datos"))
EXCEL = DATOS / "EXCEL PICASSO AGOSTO 2026 .xlsx"

from revisoria import lectores  # noqa: E402
from revisoria.excel import Libro, Periodo  # noqa: E402

P = Periodo(2026, 8)
RECETAS = json.loads((RAIZ / "config" / "lectores_modelo.json").read_text(encoding="utf-8"))["recetas"]


def _j(x):
    def quitar(v):                              # dato auxiliar del lector por receta (para migrar la cartera)
        if isinstance(v, dict):
            return {k: quitar(w) for k, w in v.items() if k != "_items_anterior"}
        return [quitar(w) for w in v] if isinstance(v, list) else v
    return quitar(json.loads(json.dumps(x, default=str, sort_keys=True)))


@pytest.mark.skipif(not EXCEL.exists(), reason=f"No está el Excel de prueba en {DATOS}")
@pytest.mark.parametrize("receta", RECETAS, ids=[r["seccion"] for r in RECETAS])
def test_recetas_del_modelo_reproducen_al_lector_nativo(receta):
    from revisoria.documental import extraer_documental
    from revisoria.extractor import extraer
    modelo, lib = extraer(str(EXCEL), P, {})
    nativo = {**modelo, **extraer_documental(lib, P)}[receta["seccion"]]
    generico = lectores.leer(Libro(str(EXCEL)), {k: v for k, v in receta.items() if k != "vistazo"}, P,
                             modelo["meta"].get("unidades"))
    assert _j(generico) == _j(nativo)


@pytest.mark.skipif(not EXCEL.exists(), reason=f"No está el Excel de prueba en {DATOS}")
def test_con_todas_las_recetas_el_modelo_completo_no_cambia():
    from revisoria.extractor import extraer
    rec = {r["seccion"]: {k: v for k, v in r.items() if k != "vistazo"} for r in RECETAS}
    a, _ = extraer(str(EXCEL), P, {})
    b, _ = extraer(str(EXCEL), P, {}, {"lectores": rec})
    assert _j(a) == _j(b)


def _otra_estructura(ruta):
    """Cuentas por pagar con otra organización: dos bloques mensuales, título distinto, columnas en otro orden."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Proveedores"
    r = 1
    for mes, valores in (("JULIO", [100, 200]), ("AGOSTO", [300, 400, 500])):
        ws.cell(r, 1, f"SALDOS PROVEEDORES AL 31 DE {mes} DE 2026")
        for c, t in enumerate(["PROVEEDOR", "CONCEPTO", "DÍAS", "SALDO"], start=1):
            ws.cell(r + 1, c, t)
        for i, v in enumerate(valores):
            ws.cell(r + 2 + i, 1, f"Proveedor {mes} {i}")
            ws.cell(r + 2 + i, 2, "Servicio")
            ws.cell(r + 2 + i, 3, "1 a 30 días")
            ws.cell(r + 2 + i, 4, v)
        ws.cell(r + 2 + len(valores), 1, "TOTAL")
        ws.cell(r + 2 + len(valores), 4, sum(valores))
        r += len(valores) + 5
    wb.save(ruta)
    return Libro(str(ruta))


def test_receta_lee_el_bloque_del_mes_con_otra_estructura(tmp_path):
    lib = _otra_estructura(tmp_path / "x.xlsx")
    receta = {"seccion": "cxp", "hoja": "Proveedores", "titulo": "saldos proveedores al 31 de {mes} de {anio}",
              "columnas": {"tercero": "A", "concepto": "B", "edad": "C", "valor": "D"}}
    r = lectores.leer(lib, receta, P)
    assert r["total"] == 1200 and r["total_hoja"] == 1200 and len(r["items"]) == 3
    assert r["items"][0]["tercero"] == "Proveedor AGOSTO 2" and r["items"][0]["fuente"] == "Proveedores!D12"
    assert lectores.probar(lib, receta, P)["ok"]
    malo = dict(receta, columnas={"tercero": "A", "valor": "B"})        # valor en una columna de texto
    assert not lectores.probar(lib, malo, P)["ok"]


def test_receta_invalida_se_explica():
    with pytest.raises(lectores.RecetaInvalida):
        lectores.validar_receta({"seccion": "cxp", "hoja": "X", "fila_encabezado": 2, "columnas": {"tercero": "A"}})
    with pytest.raises(lectores.RecetaInvalida):
        lectores.validar_receta({"seccion": "liquidez", "hoja": "X", "fila_encabezado": 1, "columnas": {}})


def test_estructurador_se_corrige_con_los_errores_de_la_prueba(tmp_path):
    """Primer intento con la columna equivocada; el sistema le devuelve el error y el segundo intento la corrige."""
    from agente import estructurador
    lib = _otra_estructura(tmp_path / "x.xlsx")
    vistos = []

    def modelo_simulado(sistema, usuario, esquema):
        ctx = json.loads(usuario)
        vistos.append(ctx)
        valor = "B" if "errores_al_probarla" not in ctx else "D"
        return ({"corresponde": True, "hoja": "Proveedores", "titulo": "saldos proveedores al 31 de {mes} de {anio}",
                 "fila_encabezado": 0, "niveles": [], "motivo": "Encabezados PROVEEDOR, CONCEPTO, DÍAS, SALDO.",
                 "columnas": [{"campo": "tercero", "columnas": ["A"]}, {"campo": "valor", "columnas": [valor]},
                              {"campo": "concepto", "columnas": ["B"]}, {"campo": "edad", "columnas": ["C"]}]}, "simulado")

    r = estructurador.proponer(tmp_path, lib, "cxp", "Proveedores", P, None, llamar=modelo_simulado)
    assert r["prueba"]["ok"] and len(r["intentos"]) == 2 and r["receta"]["columnas"]["valor"] == "D"
    assert vistos[0]["ejemplos"] and vistos[0]["ejemplos"][0]["receta"]["seccion"] == "cxp"   # aprende del modelo
    assert any("SALDOS PROVEEDORES AL 31 DE AGOSTO" in f for f in vistos[0]["vistazo_de_la_hoja"])


def test_estructurador_respeta_cuando_la_hoja_no_corresponde(tmp_path):
    from agente import estructurador
    lib = _otra_estructura(tmp_path / "x.xlsx")
    r = estructurador.proponer(tmp_path, lib, "polizas", "Proveedores", P, None, llamar=lambda *_: (
        {"corresponde": False, "hoja": "", "titulo": "", "fila_encabezado": 0, "columnas": [], "niveles": [],
         "motivo": "Son saldos de proveedores, no verificación de pólizas."}, "simulado"))
    assert r["corresponde"] is False and "receta" not in r


def test_receta_aprobada_se_usa_y_se_aprende(tmp_path):
    from agente import estructurador
    from agente.comun import Espacio
    from revisoria import mapa as mapa_cliente
    esp = Espacio(tmp_path / "esp")
    (esp.clientes / "9001234567").mkdir(parents=True)
    x = tmp_path / "x.xlsx"
    _otra_estructura(x)
    receta = {"seccion": "cxp", "hoja": "Proveedores", "titulo": "saldos proveedores al 31 de {mes} de {anio}",
              "columnas": {"tercero": "A", "concepto": "B", "edad": "C", "valor": "D"}}
    estructurador.aprobar(esp, "9001234567", receta, P, x, "Revisora")
    m = mapa_cliente.cargar(esp.clientes / "9001234567")
    assert m["lectores"]["cxp"]["hoja"] == "Proveedores" and m["hojas"]["CXP"] == "Proveedores"
    ej = estructurador.ejemplos(esp.conocimiento, "cxp")
    assert ej[0]["receta"]["hoja"] == "Proveedores" and ej[0]["vistazo"]          # el aprobado va primero


def _liquidez_otra(ruta):
    """Liquidez con otra organización: conceptos en la columna A, meses escritos como texto, otro rótulo."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Disponible"
    ws["A2"] = "DISPONIBLE Y FLUJO"
    for c, m in enumerate(["dic-25", "jun-26", "jul-26", "ago-26"], start=2):
        ws.cell(2, c, m)
    filas = [("Banco corriente", [10, 20, 30, 40]), ("Banco ahorros", [1, 2, 3, 4]), ("TOTAL DISPONIBLE", [11, 22, 33, 44])]
    for r, (n, vals) in enumerate(filas, start=3):
        ws.cell(r, 1, n)
        for c, v in enumerate(vals, start=2):
            ws.cell(r, c, v)
    ws["A8"] = "PRESUPUESTO 2026"
    for c, t in enumerate(["CONCEPTO", "PPTO AÑO", "PPTO A AGOSTO", "EJECUTADO", "DIFERENCIA", "NOTA"], start=1):
        ws.cell(9, c, t)
    for r, (n, pa, pp, ej) in enumerate([("Energía", 120, 80, 100), ("Aseo", 60, 40, 30)], start=10):
        for c, v in enumerate([n, pa, pp, ej, pp - ej, "ok"], start=1):
            ws.cell(r, c, v)
    ws.cell(12, 1, "TOTAL"); ws.cell(12, 5, -10)
    wb.save(ruta)
    return Libro(str(ruta))


def test_liquidez_y_presupuesto_con_otra_estructura(tmp_path):
    lib = _liquidez_otra(tmp_path / "l.xlsx")
    liq = {"seccion": "liquidez", "hoja": "Disponible", "encabezado": "disponible y flujo", "total": "total disponible",
           "columnas": {"concepto": "A"}}
    r = lectores.leer(lib, liq, P)
    assert r["total"] == 44 and r["total_cierre"] == 11 and r["fuente_total"] == "Disponible!E5"
    assert [s["periodo"] for s in r["serie"]] == ["2025-12", "2026-06", "2026-07", "2026-08"]
    assert lectores.probar(lib, liq, P)["ok"]
    pre = {"seccion": "presupuesto", "hoja": "Disponible", "fila_encabezado": 9, "total": "total",
           "columnas": {"rubro": "A", "presupuesto_anual": "B", "presupuesto_periodo": "C", "ejecucion": "D",
                        "diferencia": "E", "observacion": "F"}}
    r = lectores.leer(lib, pre, P)
    assert r["rubros"][0]["rubro"] == "Energía" and r["rubros"][0]["sobre_ejecutado"] == 20
    assert r["total_hoja"] == 10 and r["total_calculado"] == 10 and lectores.probar(lib, pre, P)["ok"]
    malo = dict(pre, signo_total=1)
    assert any("signo_total" in a for a in lectores.probar(lib, malo, P)["avisos"])


def test_cartera_con_columna_por_mes(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cartera"
    for c, t in enumerate(["APTO", "PROPIETARIO", "jul-26", "ago-26"], start=1):
        ws.cell(1, c, t)
    for r, (a, j, g) in enumerate([("101", 500, 600), ("102", 4_000_000, 4_100_000), ("103", 0, 0)], start=2):
        ws.cell(r, 1, a); ws.cell(r, 2, "x"); ws.cell(r, 3, j); ws.cell(r, 4, g)
    ws.cell(5, 1, "TOTAL"); ws.cell(5, 4, 4_100_600)
    ruta = tmp_path / "c.xlsx"
    wb.save(ruta)
    rec = {"seccion": "cartera", "hoja": "Cartera", "fila_encabezado": 1, "columnas": {"apto": "A"}}
    r = lectores.leer(Libro(str(ruta)), rec, P, 10)
    assert r["total"] == 4_100_600 and r["items"][0]["apto"] == "102" and r["total_hoja_lista"] == 4_100_600
    assert r["anterior"]["total"] == 4_000_500 and r["rangos"][0]["inmuebles"] == 1
    assert lectores.probar(Libro(str(ruta)), rec, P)["ok"]


def test_plantilla_vacia_se_lee_con_el_lector_estandar(tmp_path):
    from revisoria import documental, plantilla
    from revisoria.extractor import extraer
    ruta = tmp_path / "plantilla.xlsx"
    plantilla.plantilla_vacia(P, ruta)
    lib = Libro(str(ruta))
    assert set(plantilla.HOJAS.values()) <= set(lib.valores.sheetnames)
    m, lib = extraer(str(ruta), P, {})
    d = documental.extraer_documental(lib, P)
    assert m["cxp"]["items"] == [] and m["cartera"]["items"] == [] and not d["faltantes"]


@pytest.mark.skipif(not EXCEL.exists(), reason=f"No está el Excel de prueba en {DATOS}")
def test_escritores_de_la_plantilla_ida_y_vuelta(tmp_path):
    """Cada sección de Picasso escrita en la plantilla se vuelve a leer igual con el lector estándar."""
    from revisoria import documental as D, plantilla as PL
    from revisoria.extractor import extraer, extraer_cxp, extraer_retencion, extraer_presupuesto, extraer_cartera
    from agente.migrador import _diferencias, _j
    m, lib = extraer(str(EXCEL), P, {})
    doc = D.extraer_documental(lib, P)
    rec = {r["seccion"]: r for r in RECETAS}
    cart = lectores.leer(Libro(str(EXCEL)), rec["cartera"], P, m["meta"]["unidades"])
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    hoja = lambda s: wb.create_sheet(PL.HOJAS[PL.SECCION_HOJA[s]])
    PL.escribir_cxp(hoja("cxp"), m["cxp"], P); PL.escribir_retencion(hoja("retencion"), m["retencion"], P)
    PL.escribir_presupuesto(hoja("presupuesto"), m["presupuesto"], P); PL.escribir_cartera(hoja("cartera"), cart, P)
    PL.escribir_polizas(hoja("polizas"), doc["polizas"], P); PL.escribir_matriz_riesgo(hoja("matriz_riesgo"), doc["matriz_riesgo"], P)
    PL.escribir_matriz_manuales(hoja("matriz_manuales"), doc["matriz_manuales"], P)
    wb.save(tmp_path / "rt.xlsx")
    L = Libro(str(tmp_path / "rt.xlsx"))
    nuevo = {"cxp": extraer_cxp(L, P), "retencion": extraer_retencion(L), "presupuesto": extraer_presupuesto(L),
             "cartera": extraer_cartera(L, P, m["meta"]["unidades"]), "polizas": D.extraer_verificacion_polizas(L),
             "matriz_riesgo": D.extraer_matriz_riesgo(L), "matriz_manuales": D.extraer_matriz_manuales(L)}
    orig = {**m, **doc}
    for s, v in nuevo.items():
        assert _diferencias(_j(orig[s]), _j(v)) == [], s


@pytest.mark.skipif(not EXCEL.exists(), reason=f"No está el Excel de prueba en {DATOS}")
def test_migrar_picasso_es_aceptable(tmp_path):
    """El informe modelo migrado a la plantilla (identidad) debe verificarse sin diferencias."""
    import shutil
    from agente.comun import Espacio
    from agente import migrador
    from revisoria import libreoffice
    if not libreoffice.buscar():
        pytest.skip("LibreOffice no está instalado")
    esp = Espacio(tmp_path / "esp")
    (esp.clientes / "9014175393").mkdir(parents=True)
    (esp.clientes / "9014175393" / "ficha.yaml").write_text("copropiedad: Conjunto Residencial Picasso PH\nunidades: 408\n", encoding="utf-8")
    esp.entrada("2026-08", "9014175393").mkdir(parents=True)
    shutil.copy2(EXCEL, esp.entrada("2026-08", "9014175393") / "Excel_9014175393_2026-08.xlsx")
    r = migrador.migrar(esp, "9014175393", "2026-08", log=lambda *_: None)
    assert r["aceptable"], [v for v in r["verificacion"] if v["resultado"] == "diferente"] or r["hallazgos"]


@pytest.mark.skipif(not EXCEL.exists(), reason=f"No está el Excel de prueba en {DATOS}")
def test_escritor_de_seguimiento_de_deudores_ida_y_vuelta(tmp_path):
    """«may deudores» de Picasso escrita en la plantilla se lee igual, con observaciones combinadas y su formato."""
    from revisoria import plantilla as PL
    from revisoria.extractor import extraer, extraer_mayores_deudores
    from agente.migrador import _diferencias, _j
    m, _ = extraer(str(EXCEL), P, {})
    orig = m["mayores_deudores"]
    assert any(i["observacion_compartida"] for i in orig["items"])          # el caso difícil está presente
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    PL.escribir_seguimiento_deudores(wb.create_sheet(PL.HOJAS["SEGUIMIENTO DEUDORES"]), orig, P)
    wb.save(tmp_path / "md.xlsx")
    nuevo = extraer_mayores_deudores(Libro(str(tmp_path / "md.xlsx")), p=P)
    assert _diferencias(_j(orig), _j(nuevo)) == []
    assert [i["observacion_compartida"] for i in nuevo["items"]] == [i["observacion_compartida"] for i in orig["items"]]
    assert [bool(i["seg_observacion"]) for i in nuevo["items"]] == [bool(i["seg_observacion"]) for i in orig["items"]]


def test_migracion_no_crea_seguimiento_vacio():
    """Sin hoja de seguimiento, el lector usa la evolución de MAYORES DEUDORES: la migración no debe crearla vacía."""
    from revisoria import plantilla as PL
    assert "SEGUIMIENTO DEUDORES" in PL.SIN_HOJA_VACIA
