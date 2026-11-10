"""Controles del flujo con Claude: decisiones del revisor bloqueadas y aprendizaje solo de la redacción del revisor."""
import json
import subprocess
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent
HOOK = RAIZ / ".claude/hooks/proteger_decisiones_revisor.py"
import os  # noqa: E402
EXCEL = Path(os.environ.get("REVISORIA_DATOS", RAIZ.parent.parent / "05-Datos-prueba-agosto-2026")) / "EXCEL PICASSO AGOSTO 2026 .xlsx"


def _hook(entrada: dict) -> int:
    return subprocess.run([sys.executable, str(HOOK)], input=json.dumps(entrada), text=True,
                          capture_output=True).returncode


@pytest.mark.parametrize("cmd,bloquea", [
    (".venv/bin/python -m agente --raiz ../espacio-trabajo aprobar 9014175393 --periodo 2026-08 --revisor \"A\"", True),
    (".venv/bin/python -m agente --raiz ../espacio-trabajo aprobar-receta 1 --periodo 2026-08 --seccion cxp --revisor A", True),
    (".venv/bin/python -m agente --raiz ../espacio-trabajo devolver 1 --periodo 2026-08 --revisor A --motivo x", True),
    ("python -c \"from agente import orquestador; orquestador.aprobar(e, p, c, r)\"", True),
    ("curl -X POST http://127.0.0.1:8790/api/aprobar", True),
    ("CLAUDECODE= python -m agente --raiz x estado --periodo 2026-08", True),
    ("sed -i s/A/B/ ../espacio-trabajo/salidas/2026-08/1/estado.json", True),
    (".venv/bin/python -m agente --raiz ../espacio-trabajo correr --periodo 2026-08 --clientes 1", False),
    (".venv/bin/python -m agente --raiz ../espacio-trabajo probar-receta 1 --periodo 2026-08 --receta /tmp/r.json", False),
    ("cat ../espacio-trabajo/salidas/2026-08/1/estado.json", False),
    ("grep -n aprobar agente/orquestador.py", False),
])
def test_hook_bash(cmd, bloquea):
    assert (_hook({"tool_name": "Bash", "tool_input": {"command": cmd}}) == 2) is bloquea


def test_hook_archivos():
    assert _hook({"tool_name": "Edit", "tool_input": {"file_path": "/p/espacio-trabajo/clientes/1/mapa.json"}}) == 2
    assert _hook({"tool_name": "Write", "tool_input": {"file_path": "/p/espacio-trabajo/salidas/2026-08/1/Borrador_propuesta_claude.md"}}) == 0


def test_cli_rechaza_decisiones_dentro_de_claude(tmp_path, monkeypatch):
    from agente.__main__ import main
    monkeypatch.setenv("CLAUDECODE", "1")
    with pytest.raises(SystemExit) as e:
        main(["--raiz", str(tmp_path), "aprobar", "1", "--periodo", "2026-08", "--revisor", "A"])
    assert "revisor fiscal" in str(e.value)


def test_cli_exige_raiz(monkeypatch):
    from agente.__main__ import main
    monkeypatch.delenv("REVISORIA_RAIZ", raising=False)
    with pytest.raises(SystemExit):
        main(["estado", "--periodo", "2026-08"])


def test_redaccion_de_claude_no_entra_al_aprendizaje():
    from agente import mejora
    items = [{"clave": "cartera", "titulo": "Cartera", "texto": "Texto base 10.", "cifras": {}},
             {"clave": "cxp", "titulo": "Cuentas por pagar", "texto": "Base cxp.", "cifras": {}}]
    propuesta = "## 1. Cartera\nTexto de Claude 10.\n\n## 2. Cuentas por pagar\nPropuesta cxp de Claude.\n"
    final = "## 1. Cartera\nTexto de Claude 10.\n\n## 2. Cuentas por pagar\nRedacción propia del revisor.\n"
    cambios = mejora.diferencias_borrador(items, final)
    mejora.marcar_origen_claude(cambios, items, propuesta)
    origen = {c["clave"]: c.get("origen") for c in cambios}
    assert origen == {"cartera": "claude", "cxp": None}


# ------------------------------------------------------------------ Excel con otra estructura

def _variante(origen, destino, renombrar, quitar=()):
    """Copia con valores (sin fórmulas) del Excel modelo, con hojas renombradas o quitadas."""
    import openpyxl
    wb = openpyxl.load_workbook(origen, data_only=True)
    out = openpyxl.Workbook()
    out.remove(out.active)
    for ws in wb.worksheets:
        if ws.title in quitar:
            continue
        n = out.create_sheet(renombrar.get(ws.title, ws.title))
        for fila in ws.iter_rows():
            for c in fila:
                if c.value is not None:
                    n.cell(c.row, c.column, c.value)
    out.save(destino)


RENOMBRES = {"CXP ": "PROVEEDORES Y CUENTAS", "PRESUPUESTO ": "EJECUCION PPTO", "EGRESO": "COMPROBANTES MES",
             "ANTIC. E INGRESOS X C": "ANTICIPOS RECIBIDOS"}


def test_hojas_renombradas_se_detectan_por_contenido(tmp_path):
    from revisoria import mapa
    from revisoria.excel import Libro
    if not Path(EXCEL).exists():
        pytest.skip("sin datos de prueba")
    x = tmp_path / "otro_cliente.xlsx"
    _variante(EXCEL, x, RENOMBRES)
    hojas = mapa.detectar(Libro(str(x)))["hojas"]
    assert hojas["CXP"] == "PROVEEDORES Y CUENTAS"
    assert hojas["PRESUPUESTO"] == "EJECUCION PPTO"
    assert hojas["EGRESO"] == "COMPROBANTES MES"                 # antes calzaba por error con «FIRMAS E INF…»
    assert hojas["ANTIC. E INGRESOS"] == "ANTICIPOS RECIBIDOS"


def test_firma_ambigua_se_desempata_por_el_titulo_de_la_hoja():
    """Si dos hojas libres mencionan «cuentas por pagar», gana la que abre con ese título (no una matriz que lo
    nombra en un renglón). Sin título que desempate, no se adivina."""
    import openpyxl
    from revisoria import mapa
    from revisoria.excel import Libro
    import tempfile
    wb = openpyxl.Workbook()
    wb.active.title = "Cuentas x pagar"
    wb["Cuentas x pagar"]["B2"] = "CUENTAS POR PAGAR A AGOSTO 2026"
    wb.create_sheet("Matriz procedimientos")["B2"] = "Procedimiento"
    wb["Matriz procedimientos"]["C3"] = "Revisar cuentas por pagar"
    with tempfile.TemporaryDirectory() as d:
        ruta = Path(d) / "x.xlsx"
        wb.save(ruta)
        assert mapa._por_contenido(Libro(str(ruta)), {}, wb.sheetnames).get("CXP") == "Cuentas x pagar"
        wb["Matriz procedimientos"]["B1"] = "Cuentas por pagar: procedimientos"
        wb.save(ruta)
        assert mapa._por_contenido(Libro(str(ruta)), {}, wb.sheetnames).get("CXP") is None


def test_cliente_sin_hoja_de_cxp_no_se_bloquea(tmp_path):
    from agente import orquestador
    from agente.comun import PENDIENTE, Espacio
    if not Path(EXCEL).exists():
        pytest.skip("sin datos de prueba")
    esp = Espacio(tmp_path / "esp")
    (esp.clientes / "9001112223").mkdir(parents=True)
    (esp.clientes / "9001112223" / "ficha.yaml").write_text(
        "copropiedad: Conjunto Residencial Picasso PH\nnit: 900.111.222-3\nunidades: 408\n", encoding="utf-8")
    esp.entrada("2026-08", "9001112223").mkdir(parents=True)
    _variante(EXCEL, esp.entrada("2026-08", "9001112223") / "excel.xlsx", RENOMBRES, quitar=("CXP ",))
    est = orquestador.procesar_cliente(str(esp.raiz), "2026-08", "9001112223")
    assert est["estado"] == PENDIENTE, est["bitacora"][-1]
    hs = json.loads((esp.salida("2026-08", "9001112223") / "hallazgos.json").read_text(encoding="utf-8"))
    assert any(h["regla"] == "M02" and h["severidad"] == "ALTA" and "Cuentas por pagar" in h["titulo"] for h in hs)


def test_cliente_sin_hojas_financieras_sale_con_pendientes(tmp_path):
    """Un Excel sin «Rangos de Cartera» ni «PATRIMONIO» (como Arboleda Park) no bloquea: el informe sale con esas
    diapositivas como pendientes y la validación las marca ALTA."""
    from pptx import Presentation
    from agente import orquestador
    from agente.comun import PENDIENTE, Espacio
    if not Path(EXCEL).exists():
        pytest.skip("sin datos de prueba")
    esp = Espacio(tmp_path / "esp")
    (esp.clientes / "9001112223").mkdir(parents=True)
    (esp.clientes / "9001112223" / "ficha.yaml").write_text(
        "copropiedad: Conjunto Residencial Picasso PH\nnit: 900.111.222-3\n", encoding="utf-8")
    esp.entrada("2026-08", "9001112223").mkdir(parents=True)
    _variante(EXCEL, esp.entrada("2026-08", "9001112223") / "excel.xlsx", {}, quitar=("Rangos de Cartera", "PATRIMONIO"))
    est = orquestador.procesar_cliente(str(esp.raiz), "2026-08", "9001112223")
    assert est["estado"] == PENDIENTE, est["bitacora"][-1]
    sal = esp.salida("2026-08", "9001112223")
    hs = json.loads((sal / "hallazgos.json").read_text(encoding="utf-8"))
    titulos = " ".join(h["titulo"] for h in hs if h["regla"] == "M02" and h["severidad"] == "ALTA")
    assert "Cartera" in titulos and "Patrimonio" in titulos
    deck = Presentation(str(next(sal.glob("Informe_Revisoria_*.pptx"))))
    texto = [" ".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame) for s in deck.slides]
    assert sum("Información pendiente en los papeles de trabajo" in t for t in texto) >= 3   # cartera, concentración, excedentes
    assert not any("None Aptos" in t for t in texto[:1])                                    # ficha sin unidades


def test_concentracion_usa_el_tramo_con_inmuebles():
    from revisoria.narrativa import titulo_concentracion
    rangos = [{"rango": "> 3M", "valor": 0, "inmuebles": 0, "pct_cartera": 0},
              {"rango": "1M-3M", "valor": 5_808_261, "inmuebles": 3, "pct_cartera": 1.0},
              {"rango": "<= 1M", "valor": 0, "inmuebles": 0, "pct_cartera": 0}]
    assert titulo_concentracion({"cartera": {"rangos": rangos}}).startswith("3 INMUEBLES CONCENTRAN EL 100%")


def test_firma_acepta_variantes_de_encabezado():
    from revisoria.esquema import FIRMAS, cumple
    assert cumple(FIRMAS["PATRIMONIO"][0], "concepto total ingresos acumulado resultado ejercicio")
    assert cumple(FIRMAS["MAYORES DEUDORES"][0], "deuda de propietarios inmueble cta extra")
    assert not cumple(FIRMAS["PATRIMONIO"][0], "rubros sobre ejecutados")


# ------------------------------------------------------------------ memoria y RSI

def _esp_con_cliente(tmp_path):
    from agente.comun import Espacio
    esp = Espacio(tmp_path / "esp")
    (esp.clientes / "9001112223").mkdir(parents=True)
    (esp.clientes / "9001112223" / "ficha.yaml").write_text("copropiedad: Prueba PH\n", encoding="utf-8")
    return esp


def test_memoria_persiste_entre_sesiones(tmp_path):
    from agente import memoria
    from agente.comun import Espacio
    esp = _esp_con_cliente(tmp_path)
    n = memoria.anotar(esp, "9001112223", "pendiente", "Pedir actas de consejo", "2026-08")
    assert memoria.anotar(esp, "9001112223", "pendiente", "Pedir  actas de consejo", "2026-08")["id"] == n["id"]  # sin duplicar
    memoria.anotar(esp, "flujo", "leccion", "La matriz viene en otra hoja")
    otra_sesion = Espacio(esp.raiz)                                    # se lee de disco, no de la sesión
    ctx = memoria.contexto(otra_sesion, "9001112223", "2026-09")
    assert [x["texto"] for x in ctx["notas_abiertas"]] == ["Pedir actas de consejo"]
    assert ctx["lecciones_del_flujo"][0]["texto"] == "La matriz viene en otra hoja"
    memoria.resolver(otra_sesion, "9001112223", n["id"], "Llegaron")
    assert memoria.contexto(esp, "9001112223", "2026-09")["notas_abiertas"] == []
    with pytest.raises(ValueError):
        memoria.anotar(esp, "no_existe", "pendiente", "x")


def test_rsi_no_aprende_del_texto_de_claude(tmp_path):
    from agente import mejora
    from app import rsi
    esp = _esp_con_cliente(tmp_path)
    mejora.registrar(esp, "2026-08", "9001112223", [
        {"tipo": "opinion", "clave": "cartera", "accion": "editada", "despues": "a", "plantilla": "Texto del revisor"},
        {"tipo": "opinion", "clave": "cartera", "accion": "editada", "despues": "b", "plantilla": "Texto de Claude",
         "origen": "claude"}])
    assert [c["objetivo"] for c in rsi.casos(esp)["cartera"]] == ["Texto del revisor"]


def test_rsi_aprobar_es_del_revisor(tmp_path, monkeypatch):
    from agente.__main__ import main
    monkeypatch.setenv("CLAUDECODE", "1")
    with pytest.raises(SystemExit) as e:
        main(["--raiz", str(tmp_path), "rsi-aprobar", "RSI-x", "--revisor", "A"])
    assert "revisor fiscal" in str(e.value)
    assert _hook({"tool_name": "Bash", "tool_input": {"command": "curl -X PUT localhost:8790/api/configuracion/rsi"}}) == 2


def test_receta_probada_queda_lista_para_el_revisor(tmp_path, monkeypatch):
    """La receta que escribe Claude y pasa «probar-receta» queda como propuesta que «aprobar-receta» puede aprobar."""
    import shutil
    from agente.__main__ import main
    from agente.comun import Espacio
    if not Path(EXCEL).exists():
        pytest.skip("sin datos de prueba")
    esp = Espacio(tmp_path / "esp")
    (esp.clientes / "9001112223").mkdir(parents=True)
    (esp.clientes / "9001112223" / "ficha.yaml").write_text("copropiedad: Conjunto Residencial Picasso PH\n", encoding="utf-8")
    esp.entrada("2026-08", "9001112223").mkdir(parents=True)
    shutil.copy2(EXCEL, esp.entrada("2026-08", "9001112223") / "excel.xlsx")
    modelo = json.loads((RAIZ / "config/lectores_modelo.json").read_text(encoding="utf-8"))
    receta = {k: v for k, v in next(r for r in modelo["recetas"] if r["seccion"] == "cxp").items() if k != "vistazo"}
    (tmp_path / "r.json").write_text(json.dumps(receta), encoding="utf-8")
    monkeypatch.setenv("CLAUDECODE", "1")
    main(["--raiz", str(esp.raiz), "probar-receta", "9001112223", "--periodo", "2026-08", "--receta", str(tmp_path / "r.json")])
    assert (esp.salida("2026-08", "9001112223") / "receta_propuesta_cxp.json").exists()
    monkeypatch.delenv("CLAUDECODE")
    main(["--raiz", str(esp.raiz), "aprobar-receta", "9001112223", "--periodo", "2026-08", "--seccion", "cxp", "--revisor", "R"])
    assert "cxp" in json.loads((esp.clientes / "9001112223" / "mapa.json").read_text(encoding="utf-8"))["lectores"]
