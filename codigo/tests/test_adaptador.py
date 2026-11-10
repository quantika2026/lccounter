"""Agente adaptador: la lectura se ajusta al Excel de cada cliente y el informe conserva siempre su estructura."""
import json
import os
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest

from revisoria import lectores
from revisoria.excel import Libro, Periodo

RAIZ = Path(__file__).resolve().parent.parent
EXCEL = Path(os.environ.get("REVISORIA_DATOS", RAIZ.parent.parent / "05-Datos-prueba-agosto-2026")) / "EXCEL PICASSO AGOSTO 2026 .xlsx"
P = Periodo(2026, 8)


def _libro_cliente(ruta: Path, meses_cxp=(5, 4, 3)) -> Path:
    """Un Excel «de otro despacho»: CXP con una columna por mes, presupuesto con otros encabezados, liquidez y actas."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CUADRO CXP"
    ws["A6"] = "Cuentas por pagar"
    ws["A7"] = "CONCEPTO"
    for i, m in enumerate(meses_cxp):
        ws.cell(7, 2 + i, datetime(2026, m, 1))
    ws.cell(7, 2 + len(meses_cxp), "Detalle")
    for r, (t, v) in enumerate([("Proveedor uno", 100), ("Proveedor dos", 250), ("Proveedor tres", 50)], start=8):
        ws.cell(r, 1, t)
        for i in range(len(meses_cxp)):
            ws.cell(r, 2 + i, v)
        ws.cell(r, 2 + len(meses_cxp), "saldo inicial")
    ws["A11"] = "total por tercero"
    ws["B11"] = 400

    ws = wb.create_sheet("PRESUPUESTO 2025")
    ws["B2"] = "RUBROS SOBRE EJECUTADOS AÑO 2026"
    for c, t in enumerate(["Descripción", "Presupuesto Enero a Diciembre 2026", "Presupuesto Enero-agosto 2026",
                           "Ejecutado Enero-agosto 2026", "Valor sobre ejecutado agosto", "Observaciones"], start=2):
        ws.cell(3, c, t)
    for r, (n, a, pp, e) in enumerate([("Reparaciones", 1200, 800, 1500), ("Aseo", 600, 400, 450)], start=4):
        ws.cell(r, 2, n), ws.cell(r, 3, a), ws.cell(r, 4, pp), ws.cell(r, 5, e), ws.cell(r, 6, pp - e)

    ws = wb.create_sheet("FLUJO DISPONIBLE")
    ws["B4"] = "Disponible por mes"
    for i, m in enumerate([8, 7, 6]):
        ws.cell(4, 3 + i, datetime(2026, m, 1))
    ws.cell(4, 6, datetime(2025, 12, 1))
    for r, (t, v) in enumerate([("(+) Banco cuenta corriente", 1000), ("(-) Cuentas por pagar", -300)], start=5):
        ws.cell(r, 2, t)
        for i in range(4):
            ws.cell(r, 3 + i, v)
    ws["B7"] = "Total disponible"
    for i in range(4):
        ws.cell(7, 3 + i, 700)

    ws = wb.create_sheet("ACTAS CONSEJO ")
    for c, t in enumerate(["Fecha", "Temas", "Firma Presidente", "Firma Secretaria"], start=3):
        ws.cell(4, c, t)
    ws["C5"], ws["D5"], ws["E5"], ws["F5"] = "feb /02/2026", "presupuesto 2026", "Presidente", "Secretaria"
    ws["C6"], ws["D6"], ws["E6"], ws["F6"] = "abril 22/2026", "proposiciones", "Presidente", "Secretaria"
    wb.save(ruta)
    return ruta


def test_tabla_con_columna_por_mes_usa_la_del_periodo(tmp_path):
    from agente.adaptador import receta_automatica
    lib = Libro(str(_libro_cliente(tmp_path / "x.xlsx", meses_cxp=(8, 7, 6))))
    rec = receta_automatica(lib, "cxp", "CUADRO CXP", P)
    assert rec["columnas"]["valor"] == lectores.MES and rec["columnas"]["tercero"] == "A"
    pr = lectores.probar(lib, rec, P)
    assert pr["ok"], pr["errores"]
    assert lectores.leer(lib, rec, P)["total"] == 400


def test_mes_ausente_queda_pendiente_con_motivo_claro(tmp_path):
    from agente.adaptador import receta_automatica
    lib = Libro(str(_libro_cliente(tmp_path / "x.xlsx")))            # el CXP llega solo hasta mayo
    pr = lectores.probar(lib, receta_automatica(lib, "cxp", "CUADRO CXP", P), P)
    assert not pr["ok"] and "no trae la columna de Agosto 2026" in pr["errores"][0] and "2026-05" in pr["errores"][0]


def test_receta_automatica_de_presupuesto_liquidez_y_actas(tmp_path):
    from agente.adaptador import receta_automatica
    lib = Libro(str(_libro_cliente(tmp_path / "x.xlsx")))
    pres = receta_automatica(lib, "presupuesto", "PRESUPUESTO 2025", P)
    assert pres["columnas"]["presupuesto_periodo"] == "D" and pres["columnas"]["presupuesto_anual"] == "C"
    assert pres["columnas"]["ejecucion"] == "E" and lectores.probar(lib, pres, P)["ok"]
    liq = receta_automatica(lib, "liquidez", "FLUJO DISPONIBLE", P)
    leido = lectores.leer(lib, liq, P)
    assert leido["total"] == 700 and leido["total_cierre"] == 700 and len(leido["filas"]) == 3
    actas = receta_automatica(lib, "actas_consejo", "ACTAS CONSEJO ", P)
    assert actas["columnas"]["acta"] == "C" and lectores.probar(lib, actas, P)["n"] == 2


def test_recetas_automaticas_reproducen_el_lector_del_informe_modelo():
    """En el Excel modelo, la receta que arma el adaptador lee lo mismo que el lector estándar."""
    from agente.adaptador import receta_automatica
    from revisoria.extractor import extraer
    if not EXCEL.exists():
        pytest.skip("sin datos de prueba")
    lib = Libro(str(EXCEL))
    m, _ = extraer(str(EXCEL), P, {"unidades": 408}, None, lib)
    hojas = {"cxp": "CXP ", "anticipos": "ANTIC. E INGRESOS X C", "deudores_varios": "ANTIC. E INGRESOS X C",
             "liquidez": "CUADRO EFECTIVO", "cartera": "Rangos de Cartera", "presupuesto": "PRESUPUESTO "}
    leido = {s: lectores.leer(lib, receta_automatica(lib, s, h, P), P, 408) for s, h in hojas.items()}
    assert leido["cxp"]["total"] == m["cxp"]["total"]
    assert leido["anticipos"]["total_hoja"] == m["anticipos"]["total_hoja"]
    assert leido["deudores_varios"]["total_hoja"] == m["deudores_varios"]["total_hoja"]
    assert leido["liquidez"]["total"] == m["liquidez"]["total"]
    assert leido["liquidez"]["mayor_valor_monetizado"] == m["liquidez"]["mayor_valor_monetizado"]
    assert leido["cartera"]["total"] == m["cartera"]["total"]
    assert [r["rubro"] for r in leido["presupuesto"]["rubros"]] == [r["rubro"] for r in m["presupuesto"]["rubros"]]


def _espacio(tmp_path, excel: Path, cid="9001112223"):
    from agente.comun import Espacio
    esp = Espacio(tmp_path / "esp")
    (esp.clientes / cid).mkdir(parents=True)
    (esp.clientes / cid / "ficha.yaml").write_text("copropiedad: Conjunto Residencial Picasso PH\nnit: 900.111.222-3\n",
                                                  encoding="utf-8")
    esp.entrada("2026-08", cid).mkdir(parents=True)
    excel.replace(esp.entrada("2026-08", cid) / "excel.xlsx")
    return esp, cid


def _cambiar_liquidez(origen: Path, destino: Path):
    """Excel modelo con la hoja de liquidez renombrada y titulada de otra forma: el lector estándar ya no la entiende."""
    wb = openpyxl.load_workbook(origen, data_only=True)
    out = openpyxl.Workbook()
    out.remove(out.active)
    for ws in wb.worksheets:
        n = out.create_sheet("FLUJO DISPONIBLE" if ws.title == "CUADRO EFECTIVO" else ws.title)
        for fila in ws.iter_rows():
            for c in fila:
                if c.value is not None:
                    v = c.value
                    if ws.title == "CUADRO EFECTIVO" and isinstance(v, str):
                        v = v.replace("Análisis de liquidez", "Disponible por mes").replace("ANALISIS DE LIQUIDEZ", "DISPONIBLE POR MES")
                    n.cell(c.row, c.column, v)
    out.save(destino)


def test_adaptar_propone_aprueba_y_el_informe_lee_la_seccion(tmp_path):
    from agente import adaptador, orquestador
    from agente.comun import PENDIENTE
    if not EXCEL.exists():
        pytest.skip("sin datos de prueba")
    x = tmp_path / "otro.xlsx"
    _cambiar_liquidez(EXCEL, x)
    esp, cid = _espacio(tmp_path, x)
    est = orquestador.procesar_cliente(str(esp.raiz), "2026-08", cid)        # el lote corre el adaptador solo
    assert est["estado"] == PENDIENTE
    sal = esp.salida("2026-08", cid)
    ad = json.loads((sal / adaptador.ARCHIVO).read_text(encoding="utf-8"))
    liq = next(f for f in ad["secciones"] if f["seccion"] == "liquidez")
    assert liq["estado"] != "ok" and liq["propuesta"]["hoja"] == "FLUJO DISPONIBLE" and liq["estado_con_propuestas"] == "ok"
    hs = json.loads((sal / "hallazgos.json").read_text(encoding="utf-8"))
    assert any(h["regla"] == "A01" for h in hs)
    # el adaptador no aprueba: el mapa no tiene la receta hasta que el revisor la aprueba
    mapa = json.loads((esp.clientes / cid / "mapa.json").read_text(encoding="utf-8"))
    assert "liquidez" not in (mapa.get("lectores") or {})
    assert adaptador.aprobar(esp, cid, "2026-08", ["liquidez"], "Revisor de prueba") == ["liquidez"]
    est = orquestador.procesar_cliente(str(esp.raiz), "2026-08", cid)
    modelo = json.loads((sal / "modelo.json").read_text(encoding="utf-8"))
    assert modelo["liquidez"]["total"] > 0 and "liquidez" not in (modelo.get("faltantes") or {})


def test_cli_adaptar_no_aprueba(tmp_path, capsys):
    from agente.__main__ import main
    esp, cid = _espacio(tmp_path, _libro_cliente(tmp_path / "x.xlsx", meses_cxp=(8, 7)))
    (esp.clientes / cid / "ficha.yaml").write_text("copropiedad: Edificio Prueba PH\nnit: 900.111.222-3\ninforme: financiero\n",
                                                  encoding="utf-8")
    main(["--raiz", str(esp.raiz), "adaptar", cid, "--periodo", "2026-08"])
    salida = capsys.readouterr().out
    assert "receta automática" in salida and "Las aprueba un revisor" in salida
    mapa = esp.clientes / cid / "mapa.json"
    assert not mapa.exists() or "lectores" not in json.loads(mapa.read_text(encoding="utf-8"))


def test_informe_conserva_la_estructura_aunque_falten_datos_o_falle_una_diapositiva():
    """Mismo orden y mismos ítems para cualquier cliente: lo que falta o falla sale como pendiente, en su lugar."""
    from revisoria.informe_completo import InformeCompleto
    modelo = {"meta": {"periodo": "2026-08", "periodo_etiqueta": "Agosto 2026"}, "faltantes": {"cartera": "no está"}}
    inf = InformeCompleto(modelo, {"copropiedad": "Edificio Prueba PH"}, {})

    def falla():
        raise RuntimeError("dato con forma inesperada")
    inf.d_portada = falla
    inf.construir()
    tipos = [x["tipo"] for x in inf.est["estructura"]]
    assert [x["tipo"] for x in inf.plan] == tipos                          # cada ítem, en el orden de la estructura
    assert all(x["diapositivas"] >= 1 for x in inf.plan)                   # ninguno desaparece
    assert inf.errores and "portada" in inf.errores[0].lower()
    assert inf.n == len(inf.prs.slides)


def _libro_mapa_viejo(ruta: Path) -> Path:
    """Excel con las hojas MAYORES DEUDORES y PATRIMONIO, y una matriz de riesgo; el mapa guardado no las ubica."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MAYORES DEUDORES"
    ws["A1"] = "DEUDA DE PROPIETARIOS"
    ws["A2"], ws["B2"], ws["C2"] = "Inmueble", datetime(2026, 8, 1), datetime(2026, 7, 1)
    for r, (apto, v) in enumerate([("201", 2421886), ("306", 1398290)], start=3):
        ws.cell(r, 1, apto), ws.cell(r, 2, v), ws.cell(r, 3, v)
    ws = wb.create_sheet("PATRIMONIO")
    ws["B3"], ws["C3"] = "CONCEPTO", datetime(2026, 8, 1)
    ws["B6"], ws["C6"] = "RESULTADO EJERCICIO", -4696874
    ws = wb.create_sheet("matriz riengo")
    ws["A1"] = "MATRIZ DE SEGUIMIENTO"
    ws["A2"], ws["B2"], ws["C2"] = "Objetivo de Control", "DESCRIPCIÓN DEL CONTROL", "NIVEL RIESGO"
    ws["A3"], ws["B3"], ws["C3"] = "Equipo motobombas", "Se evidencia la revisión", "ALTO"
    wb.save(ruta)
    return ruta


def test_ajustes_mapa_propone_las_hojas_que_el_mapa_no_ubicaba(tmp_path):
    from agente.adaptador import ajustes_mapa
    lib = Libro(str(_libro_mapa_viejo(tmp_path / "c.xlsx")))
    mapa = {"hojas": {"MAYORES DEUDORES": None, "PATRIMONIO": None, "MATRIZ DICTAMEN": None,
                      "CONTRATOS 2026": "matriz riengo"}}
    aj = ajustes_mapa(lib, mapa)
    assert aj["MAYORES DEUDORES"]["ahora"] == "MAYORES DEUDORES" and aj["MAYORES DEUDORES"]["antes"] is None
    assert aj["PATRIMONIO"]["ahora"] == "PATRIMONIO"
    assert aj["MATRIZ DICTAMEN"]["ahora"] == "matriz riengo"
    assert mapa["hojas"]["PATRIMONIO"] is None                        # el mapa no se toca: lo aprueba el revisor
    # una hoja que ya está bien ubicada no se propone de nuevo
    assert "PATRIMONIO" not in ajustes_mapa(lib, {"hojas": {"PATRIMONIO": "PATRIMONIO"}})


def test_hoja_ausente_en_el_mapa_da_mensaje_sin_comillas_sueltas(tmp_path):
    from revisoria.extractor import extraer
    ruta = _libro_mapa_viejo(tmp_path / "c.xlsx")
    modelo, _ = extraer(str(ruta), P, {}, {"hojas": {"PATRIMONIO": None}})
    assert "Según el mapa del cliente, no tiene la hoja «PATRIMONIO»" in modelo["faltantes"].values()


def test_adaptar_propone_corregir_el_mapa_y_el_revisor_lo_aprueba(tmp_path):
    from agente import adaptador
    if not EXCEL.exists():
        pytest.skip("sin datos de prueba")
    import shutil
    from revisoria import mapa as mapa_cliente
    copia = tmp_path / "copia.xlsx"
    shutil.copy2(EXCEL, copia)                                 # _espacio MUEVE el archivo: nunca el de los datos de prueba
    esp, cid = _espacio(tmp_path, copia)
    lib = Libro(str(esp.entrada("2026-08", cid) / "excel.xlsx"))
    viejo = mapa_cliente.detectar(lib)
    viejo["hojas"]["PATRIMONIO"] = None                       # mapa hecho con otra versión: «no tiene» la hoja
    mapa_cliente.guardar(esp.clientes / cid, viejo)
    ad = adaptador.adaptar(esp, cid, "2026-08")
    pat = next(f for f in ad["secciones"] if f["seccion"] == "patrimonio")
    assert pat["estado"] != "ok" and pat["estado_con_propuestas"] == "ok"
    assert pat["ajuste_mapa"]["clave"] == "PATRIMONIO" and pat["ajuste_mapa"]["ahora"]
    assert mapa_cliente.cargar(esp.clientes / cid)["hojas"]["PATRIMONIO"] is None    # proponer no aprueba
    assert adaptador.aprobar(esp, cid, "2026-08", ["patrimonio"], "Revisor de prueba") == ["patrimonio"]
    nuevo = mapa_cliente.cargar(esp.clientes / cid)
    assert nuevo["hojas"]["PATRIMONIO"] == pat["ajuste_mapa"]["ahora"] and nuevo["aprobado_por"] == "Revisor de prueba"


def test_borrador_usa_propuestas_con_etiqueta_y_bloquea_la_aprobacion(tmp_path):
    from pptx import Presentation
    from agente import adaptador, orquestador
    from agente.comun import PENDIENTE
    if not EXCEL.exists():
        pytest.skip("sin datos de prueba")
    x = tmp_path / "otro.xlsx"
    _cambiar_liquidez(EXCEL, x)                               # copia modificada: los datos de prueba no se tocan
    esp, cid = _espacio(tmp_path, x)
    est = orquestador.procesar_cliente(str(esp.raiz), "2026-08", cid)
    assert est["estado"] == PENDIENTE and "liquidez" in est["propuestas_sin_aprobar"]
    sal = esp.salida("2026-08", cid)
    modelo = json.loads((sal / "modelo.json").read_text(encoding="utf-8"))
    assert modelo["liquidez"]["total"] > 0                    # el borrador ya leyó la sección con la receta propuesta
    assert "liquidez" not in (json.loads((esp.clientes / cid / "mapa.json").read_text(encoding="utf-8")).get("lectores") or {})
    pr = Presentation(sal / est["archivos"]["pptx"])
    marcadas = [s for s in pr.slides if any(sh.name == "EtiquetaPropuesta" for sh in s.shapes)]
    assert marcadas and est["resumen"]["diapositivas_con_propuesta"] == len(marcadas)
    with pytest.raises(SystemExit, match="Adaptación al Excel"):
        orquestador.aprobar(esp, "2026-08", cid, "Revisor de prueba", "justificación")
    # aprobadas las propuestas y vuelto a validar, la etiqueta desaparece y ya se puede decidir
    adaptador.aprobar(esp, cid, "2026-08", list(est["propuestas_sin_aprobar"]), "Revisor de prueba")
    est = orquestador.procesar_cliente(str(esp.raiz), "2026-08", cid)
    assert est["propuestas_sin_aprobar"] == []
    pr = Presentation(sal / est["archivos"]["pptx"])
    assert not any(sh.name == "EtiquetaPropuesta" for s in pr.slides for sh in s.shapes)


def test_informe_sin_propuestas_no_lleva_etiquetas():
    from pptx import Presentation
    from revisoria.documental import extraer_documental
    from revisoria.extractor import extraer
    from revisoria.informe_completo import InformeCompleto
    from revisoria.narrativa import Redactor
    if not EXCEL.exists():
        pytest.skip("sin datos de prueba")
    modelo, lib = extraer(str(EXCEL), P, {"copropiedad": "Picasso"})
    inf = InformeCompleto(modelo, {"copropiedad": "Picasso"}, extraer_documental(lib, P), Redactor(False)).construir()
    assert inf.con_propuesta == []
    assert not any(sh.name == "EtiquetaPropuesta" for s in inf.prs.slides for sh in s.shapes)
