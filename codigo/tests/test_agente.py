"""Pruebas del agente por lote (usa el espacio de demostración construido desde el Excel de Picasso)."""
import os
import subprocess
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))
DATOS = Path(os.environ.get("REVISORIA_DATOS", RAIZ.parent / "lccounter-datos"))
EXCEL = DATOS / "EXCEL PICASSO AGOSTO 2026 .xlsx"
pytestmark = pytest.mark.skipif(not EXCEL.exists(), reason=f"No está el Excel de prueba en {DATOS}")


@pytest.fixture(scope="module")
def esp(tmp_path_factory):
    from agente import orquestador
    from agente.comun import Espacio
    d = tmp_path_factory.mktemp("espacio")
    subprocess.run([sys.executable, str(RAIZ / "demo/preparar_demo.py"), "--excel", str(EXCEL), "--destino", str(d)],
                   check=True, capture_output=True)
    e = Espacio(d)
    orquestador.correr(e, "2026-08", paralelo=2, log=lambda *_: None)
    return e


def test_lote_estados(esp):
    est = {c: esp.estado("2026-08", c)["estado"] for c in esp.ids_clientes()}
    assert est == {"picasso": "PENDIENTE_VALIDACION", "mirador": "PENDIENTE_VALIDACION",
                   "arboleda": "PENDIENTE_VALIDACION", "sanrafael": "BLOQUEADO"}


def test_cuadre_y_error_sembrado(esp):
    from agente.comun import leer_json
    cuadre = leer_json(esp.salida("2026-08", "picasso") / "cuadre_balance.json")
    assert cuadre and all(c["ok"] for c in cuadre)                          # perfil genérico, cuadra
    assert all(c["ok"] for c in leer_json(esp.salida("2026-08", "mirador") / "cuadre_balance.json"))  # otro software
    reglas = {h["regla"] for h in leer_json(esp.salida("2026-08", "arboleda") / "hallazgos.json")}
    assert {"V20", "N03"} <= reglas                                         # total digitado + formato sin perfil
    assert (esp.conocimiento / "propuestas" / "mapeo_contable_y_arboleda.json").exists()


def test_reauditor_sin_cifras_huerfanas(esp):
    from agente.comun import leer_json
    for c in ("picasso", "mirador", "arboleda"):
        assert not [h for h in leer_json(esp.salida("2026-08", c) / "reauditoria.json") if h["regla"] == "R02"], c


def test_aprobacion_requiere_justificacion_y_alimenta_mejora(esp):
    from agente import mejora, orquestador
    with pytest.raises(SystemExit):
        orquestador.aprobar(esp, "2026-08", "picasso", "Revisora")
    sal = esp.salida("2026-08", "picasso")
    subprocess.run([sys.executable, str(RAIZ / "demo/simular_revisor.py"), str(sal)], check=True, capture_output=True)
    est = orquestador.aprobar(esp, "2026-08", "picasso", "Revisora", "justificado",
                              str(sal / "Informe_editado.pptx"), str(sal / "Borrador_editado.md"))
    assert est["estado"] == "APROBADO" and est["aprobacion"]["correcciones"] == 2
    assert (esp.conocimiento / "golden" / "2026-08_picasso" / "excel.xlsx").exists()
    props = {p["id"]: p for p in mejora.proponer(esp)}
    p = props["BIB-concentracion_cartera-1"]
    assert "{inmuebles}" in p["plantilla"] and "{pct_concentracion}" in p["plantilla"]
    r = mejora.aprobar_propuesta(esp, p["id"], "Revisora")
    assert r["ok"] and mejora.regresion(esp)["ok"]
    with pytest.raises(SystemExit):                       # un informe aprobado no se vuelve a aprobar
        orquestador.aprobar(esp, "2026-08", "picasso", "Revisora", "x")


def test_informe_completo_sigue_al_modelo(esp):
    """El informe generado trae las secciones del modelo 006-2026 y nada del Excel queda sin lugar."""
    from pptx import Presentation
    from agente.comun import leer_json
    sal = esp.salida("2026-08", "picasso")
    deck = Presentation(str(next(sal.glob("Informe_Revisoria_*.pptx"))))
    titulos = " | ".join(next((sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()), "")
                         for s in deck.slides).lower()
    for esperado in ("metodología", "comprobantes de egreso agosto 2026", "facturación electrónica", "conciliaciones bancarias",
                     "retención en la fuente", "libros actas de consejo", "decisiones asamblea ordinaria",
                     "auditoría financiera", "cartera alcanza", "excedentes acumulados", "contratos vigentes",
                     "verificación de pólizas", "matriz de riesgo", "solicitud de papeles", "gracias"):
        assert esperado in titulos, esperado
    reglas = {h["regla"] for h in leer_json(sal / "hallazgos.json")}
    assert {"D01", "D02", "D04", "D08", "D10"} <= reglas
    assert not [h for h in leer_json(sal / "hallazgos.json") if h["regla"] == "G01" and "no se pudo" in h["detalle"].lower()]


def test_imagen_emf_del_excel_se_recupera():
    """openpyxl descarta los EMF: la tabla de retención en la fuente (AUD CONT Y ADMON!C17) se recupera como PNG."""
    from revisoria import documental, libreoffice
    from revisoria.excel import Libro
    if not libreoffice.buscar():
        pytest.skip("LibreOffice no está instalado")
    vect = libreoffice.imagenes_vectoriales(EXCEL, "AUD CONT Y ADMON")
    assert [(v["fila"], v["col"]) for v in vect] == [(17, 3)] and vect[0]["bytes"].startswith(b"\x89PNG")
    campos = documental.extraer_auditoria_documental(Libro(str(EXCEL)))["campos"]
    assert any(c["fila_inicio"] <= 17 <= c["fila_fin"] and vect[0]["bytes"] in [i["bytes"] for i in c["imagenes"]]
               for c in campos)


# ---------------------------------------------------------------- correcciones v0.4.1 (no requieren el Excel)

def test_plantillar_no_toca_digitos_sueltos():
    from agente.mejora import plantillar
    t, sueltas = plantillar("8 inmuebles concentran el 87% de la cartera desde 2018.", {"inmuebles": 8, "pct": "87%"})
    assert t == "{inmuebles} inmuebles concentran el {pct} de la cartera desde 2018."
    assert sueltas == ["2018."] or sueltas == ["2018"]


def test_revertir_version_inexistente_no_borra_biblioteca(tmp_path):
    from agente import mejora
    from agente.comun import Espacio
    e = Espacio(tmp_path)
    bib = e.conocimiento / "biblioteca_opiniones.yaml"
    bib.parent.mkdir(parents=True)
    bib.write_text("flujo_negativo: {texto: x}\n", encoding="utf-8")
    with pytest.raises(SystemExit):
        mejora.revertir(e, "bib-999")
    assert bib.exists()


def test_lote_sigue_con_cliente_sin_ficha(tmp_path):
    from agente import orquestador
    from agente.comun import Espacio
    r = orquestador.correr(Espacio(tmp_path), "2026-08", ["no_existe"], paralelo=1, log=lambda *_: None)
    assert r[0]["estado"] == "BLOQUEADO"


def test_redactor_tolera_marcador_invalido(tmp_path):
    from agente import redaccion
    from agente.comun import Espacio
    e = Espacio(tmp_path)
    (e.conocimiento).mkdir(parents=True)
    (e.conocimiento / "biblioteca_opiniones.yaml").write_text(
        "anticipos_antiguos: {titulo: Anticipos, texto: 'Hay {marcador_malo}'}\n", encoding="utf-8")
    m = {"meta": {"periodo_etiqueta": "agosto 2026"}, "anticipos": {"items": [{"valor": 1000}]}}
    texto, items = redaccion.borrador(e, m, {}, [])
    assert items and "marcador inválido" in items[0]["texto"]


# ---------------------------------------------------------------- mejora continua v0.5 (no requieren el Excel)

def test_version_vigente_respeta_reversiones(tmp_path):
    from agente import mejora
    from agente.comun import Espacio, guardar_json
    e = Espacio(tmp_path)
    assert mejora.version_vigente(e) == "base"
    guardar_json(e.conocimiento / "versiones.json", [{"version": "bib-001"}, {"version": "bib-002"},
                                                     {"version": "revertir-bib-002"}])
    assert mejora.version_vigente(e) == "bib-001"
    guardar_json(e.conocimiento / "versiones.json", [{"version": "bib-001"}, {"version": "revertir-bib-001"}])
    assert mejora.version_vigente(e) == "base"


def test_metrica_por_version(tmp_path):
    from agente import mejora
    from agente.comun import Espacio, guardar_json
    e = Espacio(tmp_path)
    for cid, ver in (("a", "base"), ("b", "bib-001")):
        guardar_json(e.salida("2026-08", cid) / "estado.json",
                     {"cliente": cid, "periodo": "2026-08", "estado": "APROBADO", "bitacora": [],
                      "aprobacion": {"revisor": "R", "version_biblioteca": ver}})
    mejora.registrar(e, "2026-08", "a", [{"tipo": "opinion"}, {"tipo": "texto"}])
    porv = {v["version"]: v for v in mejora.metricas(e)["por_version"]}
    assert porv["base"]["correcciones_por_informe"] == 2 and porv["base"]["opiniones_editadas_por_informe"] == 1
    assert porv["bib-001"]["correcciones_por_informe"] == 0


def test_devoluciones_repetidas_generan_borrador_de_regla(tmp_path):
    from agente import mejora
    from agente.comun import Espacio
    e = Espacio(tmp_path)
    mejora.registrar(e, "2026-08", "a", [{"tipo": "devolucion", "motivo": "Total de CXP digitado; corregir la fórmula"}])
    mejora.registrar(e, "2026-09", "b", [{"tipo": "devolucion", "motivo": "El total de CXP está digitado"}])
    mejora.registrar(e, "2026-09", "c", [{"tipo": "devolucion", "motivo": "Falta el acta de consejo de julio"}])
    regs = [p for p in mejora.proponer(e) if p["tipo"] == "regla_borrador"]
    assert len(regs) == 1 and "a 2026-08" in regs[0]["evidencia"] and "b 2026-09" in regs[0]["evidencia"]
    assert (e.raiz / regs[0]["archivo"]).exists()


def test_auditoria_mensual_alimenta_mejora(tmp_path):
    from agente import auditoria, mejora
    from agente.comun import Espacio, guardar_json
    e = Espacio(tmp_path)
    guardar_json(e.conocimiento / "auditorias" / "muestra_2026-08.json",
                 {"periodo": "2026-08", "muestra": [{"cliente": "a", "copropiedad": "A", "aprobado_por": "Revisora A"}]})
    with pytest.raises(SystemExit):                                        # quien aprobó no se audita a sí mismo
        auditoria.registrar_conclusion(e, "2026-08", "a", "Revisora A", "menores", "x")
    with pytest.raises(SystemExit):                                        # con observaciones hay que describirlas
        auditoria.registrar_conclusion(e, "2026-08", "a", "Revisor B", "reproceso")
    auditoria.registrar_conclusion(e, "2026-08", "a", "Revisor B", "reproceso", "Cartera del título no cuadra")
    assert any(p["id"] == "AUD-a-2026-08" for p in mejora.proponer(e))
    assert mejora.metricas(e)["auditoria_mensual"]["requiere_reproceso"] == 1


# ---------------------------------------------------------------- RSI y proveedores (no requieren el Excel ni claves)

def test_rsi_restricciones_y_puntaje():
    from app import rsi
    obj = ["{inmuebles} inmuebles concentran el {pct_concentracion} de la cartera. En opinión de la Revisoría Fiscal, debe informarse al Consejo."]
    assert rsi.problemas(obj[0], "concentracion_cartera", obj) == []
    assert any("cifras escritas a mano" in p for p in rsi.problemas("{inmuebles} inmuebles; plazo 2025 para el cobro de la cartera.", "concentracion_cartera", obj))
    assert any("marcadores no permitidos" in p for p in rsi.problemas("Texto con {inventado} suficiente para la prueba de longitud.", "concentracion_cartera", obj))
    assert rsi.puntaje(obj[0], obj) == 1.0 and rsi.puntaje("otra cosa distinta", obj) < 0.3


def test_proveedores_claves_enmascaradas(tmp_path):
    from app import proveedores
    proveedores.guardar_proveedor(tmp_path, "local", {"nombre": "L", "tipo": "openai", "base_url": "http://localhost:11434/v1",
                                                      "api_key": "sk-secreta-123456", "modelos": ["m"]})
    pub = proveedores.publico(tmp_path)
    assert "sk-secreta" not in str(pub) and pub["proveedores"]["local"]["api_key"].endswith("3456")
    with pytest.raises(ValueError):
        proveedores.guardar_proveedor(tmp_path, "x", {"nombre": "x", "tipo": "openai", "base_url": "http://externo.com/v1"})
    proveedores.asignar(tmp_path, "asistente", "local", "m")
    with pytest.raises(ValueError):
        proveedores.eliminar_proveedor(tmp_path, "local")              # está en uso


def test_comunicaciones_controles_de_citas_y_cifras():
    from app.comunicaciones import cifras_sin_respaldo, citas_invalidas
    F = [{"id": f"F{i}", "texto": "Total $22.063.905"} for i in range(1, 8)]
    assert citas_invalidas("(=SUM(F4:F10)) PRESUPUESTO !F13, celda F12, (F5) y (F10, inventada)", F) == ["F10"]
    assert cifras_sin_respaldo("2.5.1 Presupuesto: $22.063.905 y $9.999.999", F) == ["9.999.999"]


def test_mapa_del_cliente(esp):
    """El orquestador deja el mapa propuesto; con o sin mapa el modelo es el mismo; un mapa viejo avisa."""
    import copy
    import json
    from revisoria import esquema, mapa
    from revisoria.excel import Libro, Periodo
    from revisoria.extractor import extraer
    m = mapa.cargar(esp.clientes / "picasso")
    assert m and m["hojas"]["SEGUIMIENTO DEUDORES"] == "may deudores" and m["hojas"]["CXP"] == "CXP "
    p = Periodo(2026, 8)
    con, _ = extraer(str(EXCEL), p, {}, m)
    sin, _ = extraer(str(EXCEL), p, {})
    assert json.dumps(con, default=str, sort_keys=True) == json.dumps(sin, default=str, sort_keys=True)
    assert all(c["estado"] == "ok" for c in esquema.verificar(con) if c["obligatoria"])
    viejo = copy.deepcopy(m)
    viejo["hojas"]["CUADRO RESERVA"] = "RESERVA 2025"                     # la hoja cambió de nombre
    lib = Libro(str(EXCEL), viejo)
    with pytest.raises(KeyError):
        lib.hoja("CUADRO RESERVA")
    assert lib.desajustes and "RESERVA 2025" in lib.desajustes[0]


def test_firma_de_formato(esp):
    """Picasso no tiene falsas alarmas de formato; una hoja con otro contenido se rechaza aunque exista."""
    from revisoria import esquema, mapa
    from revisoria.documental import extraer_documental
    from revisoria.excel import Periodo
    from revisoria.extractor import extraer
    m = mapa.cargar(esp.clientes / "picasso")
    p = Periodo(2026, 8)
    modelo, lib = extraer(str(EXCEL), p, {}, m)
    assert esquema.revisar_formato(lib, m["hojas"], extraer_documental(lib, p)) == []
    malo = esquema.revisar_formato(lib, {"VERIFICACION POLIZAS": "EGRESO", "CXP": "PATRIMONIO"})
    assert {x["clave"] for x in malo} == {"VERIFICACION POLIZAS", "CXP"}
    reglas = {h["regla"] for h in __import__("json").load(open(esp.salida("2026-08", "picasso") / "hallazgos.json"))}
    assert not {"M01", "M03"} & reglas


def test_mapa_se_repara_y_aprende(esp, tmp_path):
    """Un mes el cliente renombra hojas: con un nombre parecido se repara solo; con uno distinto, solo si otro
    cliente ya enseñó ese nombre al aprobar su mapa. Las hojas nuevas se avisan."""
    import openpyxl
    from revisoria import mapa
    from revisoria.excel import Libro
    m = mapa.cargar(esp.clientes / "picasso")
    m["sin_usar"] = m.get("sin_usar", [])
    wb = openpyxl.load_workbook(EXCEL)
    wb["CUADRO RESERVA"].title = "RESERVAS PATRIMONIO"           # nombre distinto
    wb["IMPUESTOS"].title = "IMPUESTOS 2026"                      # nombre parecido
    wb.create_sheet("NOTAS GERENCIA")
    x = tmp_path / "mes_siguiente.xlsx"
    wb.save(x)
    lib = Libro(str(x))
    aj = mapa.reparar(lib, m, {})
    assert aj["IMPUESTOS"]["ahora"] == "IMPUESTOS 2026"
    # sin aprendizaje igual se encuentra: es la única hoja libre cuyo contenido tiene la firma de la reserva
    assert aj["CUADRO RESERVA"]["ahora"] == "RESERVAS PATRIMONIO"
    assert "NOTAS GERENCIA" in mapa.hojas_nuevas(lib, m)
    otro = {"hojas": {"CUADRO RESERVA": "Reservas Patrimonio"}}    # otro cliente aprobó ese nombre
    nuevos = mapa.aprender(tmp_path, otro, "otro_cliente")
    assert nuevos and mapa.aprendidos(tmp_path)["reservas patrimonio"]["clave"] == "CUADRO RESERVA"
    aj = mapa.reparar(lib, m, mapa.aprendidos(tmp_path))
    assert aj["CUADRO RESERVA"]["ahora"] == "RESERVAS PATRIMONIO"
    mapa.aprender(tmp_path, {"hojas": {"PATRIMONIO": "Reservas Patrimonio"}}, "tercero")   # contradicción
    assert mapa.aprendidos(tmp_path)["reservas patrimonio"]["clave"] == "AMBIGUO"
    # un sinónimo ambiguo no se usa, pero el contenido de la hoja sigue identificándola
    assert mapa.reparar(lib, m, mapa.aprendidos(tmp_path))["CUADRO RESERVA"]["ahora"] == "RESERVAS PATRIMONIO"
