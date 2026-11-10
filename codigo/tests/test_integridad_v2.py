"""Pruebas independientes de los controles de integridad de la versión mejorada."""
from pathlib import Path

import openpyxl

from agente.comun import Espacio, guardar_json, mostrar_nit, normalizar_nit
from agente.integridad import manifiesto, misma_revision, verificar_identidad
from revisoria.excel import Libro, Periodo
from revisoria.extractor import detectar_periodo, extraer_anticipos
from revisoria.narrativa import Redactor
from revisoria.validador import r_formulas_sin_resultado


class _Respuesta:
    def __init__(self, texto):
        self.content = [type("Bloque", (), {"type": "text", "text": texto})()]


class _Cliente:
    def __init__(self, texto):
        self.texto = texto
        self.messages = self

    def create(self, **_):
        return _Respuesta(self.texto)


def _redactar(original, nuevo):
    r = Redactor(False)
    r.usar_llm, r.cliente = True, _Cliente(nuevo)
    return r.pulir(original), r


def test_candado_rechaza_signo_porcentaje_omision_y_permutacion():
    casos = [
        ("La variación fue -$1.234.", "La variación fue $1.234."),
        ("La cartera vencida es 25,0%.", "La cartera vencida es 250%."),
        ("Cartera $1.234 y deterioro $500.", "Cartera $1.234."),
        ("Cartera $1.234 y deterioro $500.", "Cartera $500 y deterioro $1.234."),
    ]
    for original, nuevo in casos:
        salida, r = _redactar(original, nuevo)
        assert salida == original
        assert r.registro[-1]["estado"].startswith("descartado")


def test_formula_sin_cache_no_se_trata_como_cero(tmp_path):
    ruta = tmp_path / "formula.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PATRIMONIO"
    ws["A1"], ws["A2"] = 100, "=SUM(A1:A1)"
    wb.save(ruta)
    hallazgos = r_formulas_sin_resultado(Libro(str(ruta)))
    assert hallazgos and hallazgos[0].regla == "V21"
    assert "PATRIMONIO!A2" in hallazgos[0].ubicacion


def test_periodo_automatico_ignora_mes_vacio(tmp_path):
    ruta = tmp_path / "periodo.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CUADRO EFECTIVO"
    ws["A1"] = "Análisis de liquidez"
    ws["B1"], ws["C1"] = __import__("datetime").date(2026, 8, 1), __import__("datetime").date(2026, 12, 1)
    ws["B2"] = 100
    wb.save(ruta)
    assert detectar_periodo(Libro(str(ruta))).clave == "2026-08"


def test_anticipos_exige_mes_y_anio_y_conserva_fuente(tmp_path):
    ruta = tmp_path / "anticipos.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ANTIC. E INGRESOS X C"
    for fila, anio, valor in ((1, 2026, 100), (6, 2025, 999)):
        ws.cell(fila, 1, f"ANTICIPOS Y AVANCES A CONTRATISTAS A AGOSTO {anio}")
        for c, v in enumerate(("FECHA", "CONTRATISTA", "OTRO", "VALOR", "DETALLE", ""), 1):
            ws.cell(fila + 1, c, v)
        ws.cell(fila + 2, 1, "01 de agosto de 2026")
        ws.cell(fila + 2, 2, "Proveedor")
        ws.cell(fila + 2, 4, valor)
        ws.cell(fila + 3, 1, "TOTAL")
        ws.cell(fila + 3, 4, valor)
    wb.save(ruta)
    resultado = extraer_anticipos(Libro(str(ruta)), Periodo(2026, 8))["anticipos"]
    assert resultado["items"][0]["valor"] == 100
    assert resultado["items"][0]["fuente"].endswith("!D3")


def test_identidad_y_revision_de_fuentes(tmp_path):
    ruta = tmp_path / "identidad.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CONTRATOS"
    ws["A2"] = "MATRIZ DE CONTRATOS CONJUNTO ALDEA DE SAN MARCO PH 2026"
    wb.save(ruta)
    libro = Libro(str(ruta))
    problemas = verificar_identidad(libro, {"copropiedad": "Conjunto Picasso PH", "unidades": 408},
                                     {"meta": {"periodo": "2026-08"}, "cartera": {"unidades_hoja": 249}}, "2026-08")
    assert len(problemas) == 2

    esp = Espacio(tmp_path / "espacio")
    (esp.clientes / "picasso").mkdir(parents=True)
    (esp.clientes / "picasso" / "ficha.yaml").write_text("copropiedad: Picasso PH\n", encoding="utf-8")
    entrada = esp.entrada("2026-08", "picasso")
    entrada.mkdir(parents=True)
    (entrada / "Excel.xlsx").write_bytes(b"uno")
    primero = manifiesto(esp, "2026-08", "picasso")
    assert misma_revision(esp, "2026-08", "picasso", primero)
    (entrada / "Excel.xlsx").write_bytes(b"dos")
    assert not misma_revision(esp, "2026-08", "picasso", primero)


def test_json_atomico_y_json_danado_visible(tmp_path):
    ruta = tmp_path / "estado.json"
    guardar_json(ruta, {"ok": True})
    assert not list(tmp_path.glob("*.tmp"))
    ruta.write_text("{", encoding="utf-8")
    from agente.comun import leer_json
    try:
        leer_json(ruta, {})
    except ValueError as e:
        assert "dañado" in str(e)
    else:
        raise AssertionError("Un JSON dañado no debe parecer un expediente nuevo")


def test_nit_es_el_identificador_estable_del_cliente():
    assert normalizar_nit("900.123.456-7") == "9001234567"
    assert normalizar_nit("900 123 456 7") == "9001234567"
    assert mostrar_nit("9001234567") == "900.123.456-7"
    try:
        normalizar_nit("torres_parque")
    except ValueError as e:
        assert "dígito de verificación" in str(e)
    else:
        raise AssertionError("Un nombre no puede volver a usarse como identificador de cliente")


def test_identidad_no_confunde_apartamento_o_resolucion_con_nit(tmp_path):
    ruta = tmp_path / "identidad_sin_nit.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MAYORES DEUDORES"
    ws["C2"], ws["C3"] = "APTO", "901444281"
    ws["D4"] = "Resolución No. 000184 2024"
    wb.save(ruta)
    problemas = verificar_identidad(
        Libro(str(ruta)), {"nit": "900.123.456-7"},
        {"meta": {"periodo": "2026-08"}, "cartera": {}}, "2026-08"
    )
    assert problemas == []


def test_identidad_compara_nit_solo_si_esta_declarado(tmp_path):
    ruta = tmp_path / "identidad_con_nit.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["B1"] = "NIT", "901.444.281-7"
    wb.save(ruta)
    problemas = verificar_identidad(
        Libro(str(ruta)), {"nit": "900.123.456-7"},
        {"meta": {"periodo": "2026-08"}, "cartera": {}}, "2026-08"
    )
    assert len(problemas) == 1
    assert "9014442817" in problemas[0]
    assert "Sheet!B1" in problemas[0]


def test_modelo_picasso_conserva_items_aunque_falten_datos():
    from revisoria.generador_pptx import Informe
    from revisoria.informe_completo import InformeCompleto

    modelo = {
        "meta": {"periodo": "2026-08"}, "liquidez": {},
        "mayores_deudores": {"items": []}, "anticipos": {"items": []},
        "deudores_varios": {"items": []}, "cxp": None,
    }
    financieros = []
    base = Informe(modelo, {})
    base.d_sin_datos = lambda titulo, *_args, **_kwargs: financieros.append(titulo)
    base.d_monetizacion()
    base.d_mayores_deudores()
    base.d_anticipos()
    base.d_cxp()
    assert {
        "Análisis de monetización del fondo de imprevistos", "Detalle de mayores deudores",
        "Anticipos a contratistas y deudores varios", "Opinión de la Revisoría Fiscal", "Cuentas por pagar",
    } <= set(financieros)

    regulatorios = []
    completo = InformeCompleto(modelo, {"copropiedad": "Prueba PH"}, {})
    completo.d_sin_datos = lambda titulo, *_args, **_kwargs: regulatorios.append(titulo)
    completo.d_egresos()
    completo.d_campo({"campo": "Factura electrónica", "titulo": "Facturación Electrónica"})
    completo.d_actas_consejo({"campo": "Libro actas de consejo", "titulo": "Actas de Consejo"})
    completo.d_asamblea({"bloque": "ordinaria"})
    completo.d_contratos({"titulo": "Contratos Vigentes"})
    completo.d_polizas_manual({"titulo": "Verificación de Pólizas"})
    completo.d_matriz_riesgo({})
    assert {
        "Comprobantes de Egreso Agosto 2026", "Facturación Electrónica", "Actas de Consejo",
        "Decisiones de la Asamblea", "Contratos Vigentes", "Verificación de Pólizas",
        "Matriz de lineamientos por tipo de contrato",
        "Matriz de riesgo, Ambiental, Mantenimiento y Salubridad",
    } <= set(regulatorios)


def _hoja_deudores(ruta, titulos):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ANTIC. E INGRESOS X C"
    r = 1
    for t, total in titulos:
        ws.cell(r, 1, t); ws.cell(r + 1, 1, "FECHA"); ws.cell(r + 1, 2, "CONTRATISTA"); ws.cell(r + 1, 4, "VALOR")
        ws.cell(r + 2, 1, "2026-08-01"); ws.cell(r + 2, 2, "Proveedor"); ws.cell(r + 2, 4, total)
        ws.cell(r + 3, 1, "TOTAL DEUDORES VARIOS"); ws.cell(r + 3, 4, total)
        r += 6
    wb.save(ruta)
    return Libro(str(ruta)).hoja("ANTIC. E INGRESOS")


def test_bloque_sin_anio_solo_si_es_el_ultimo_del_anio_vigente(tmp_path):
    """«DEUDORES VARIOS A AGOSTO CTA 1380» (sin año) se acepta para agosto del año vigente de la hoja, no para otro."""
    from revisoria.extractor import _bloque_del_mes
    ws = _hoja_deudores(tmp_path / "a.xlsx", [("ANTICIPOS Y AVANCES A CONTRATISTAS ENERO 2026", 1),
                                              ("DEUDORES VARIOS A JULIO CTA 1380", 2),
                                              ("DEUDORES VARIOS A AGOSTO CTA 1380", 5299410)])
    b = _bloque_del_mes(ws, "deudores varios", Periodo(2026, 8))
    assert b and b["total"] == 5299410 and b["anio_supuesto"] == 2026
    assert _bloque_del_mes(ws, "deudores varios", Periodo(2025, 8)) is None      # otro año: no
    assert _bloque_del_mes(ws, "deudores varios", Periodo(2026, 7)) is None      # julio no es el último bloque
    ws2 = _hoja_deudores(tmp_path / "b.xlsx", [("DEUDORES VARIOS A AGOSTO 2025", 7), ("DEUDORES VARIOS A AGOSTO 2026", 8)])
    assert _bloque_del_mes(ws2, "deudores varios", Periodo(2025, 8))["total"] == 7  # con año manda el año


def test_dos_excel_en_la_entrada_bloquean(tmp_path):
    from agente.normalizador import ExcelAmbiguo, ubicar_excel
    (tmp_path / "Excel_1_2026-08.xlsx").write_bytes(b"x")
    assert ubicar_excel(tmp_path, {}).name == "Excel_1_2026-08.xlsx"
    (tmp_path / "Excel_1_2026-08 2.xlsx").write_bytes(b"x")
    try:
        ubicar_excel(tmp_path, {})
    except ExcelAmbiguo as e:
        assert "2 archivos" in str(e)
    else:
        raise AssertionError("Con dos Excel el sistema no debe elegir uno por su cuenta")


def test_cambiar_nit_provisional_mueve_todo_y_respeta_aprobados(tmp_path):
    import yaml
    from agente.integridad import cambiar_nit
    esp = Espacio(tmp_path)
    viejo, real = "9001234567", "901.417.539-3"
    (esp.clientes / viejo).mkdir(parents=True)
    (esp.clientes / viejo / "ficha.yaml").write_text(
        f"id: '{viejo}'\ncopropiedad: Aldea\nnit: 900.123.456-7\nnit_provisional: true\n", encoding="utf-8")
    esp.entrada("2026-08", viejo).mkdir(parents=True)
    (esp.entrada("2026-08", viejo) / "Excel.xlsx").write_bytes(b"x")
    guardar_json(esp.salida("2026-08", viejo) / "estado.json", {"cliente": viejo, "estado": "PENDIENTE_VALIDACION", "bitacora": []})
    nuevo = cambiar_nit(esp, viejo, real, "Revisora")
    assert nuevo == "9014175393" and not (esp.clientes / viejo).exists()
    f = yaml.safe_load((esp.clientes / nuevo / "ficha.yaml").read_text(encoding="utf-8"))
    assert f["nit"] == "901.417.539-3" and f["nit_provisional"] is False and f["nit_anterior"] == "900.123.456-7"
    assert (esp.entrada("2026-08", nuevo) / "Excel.xlsx").exists()
    from agente.comun import leer_json
    assert leer_json(esp.salida("2026-08", nuevo) / "estado.json")["cliente"] == nuevo
    guardar_json(esp.salida("2026-08", nuevo) / "estado.json", {"cliente": nuevo, "estado": "APROBADO", "bitacora": []})
    try:
        cambiar_nit(esp, nuevo, "900.000.000-1", "Revisora")
    except ValueError as e:
        assert "aprobados" in str(e)
    else:
        raise AssertionError("No se reescribe el NIT de un cliente con informes aprobados")
