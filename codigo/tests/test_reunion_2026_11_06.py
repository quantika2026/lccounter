"""Pedidos de la reunión del 6-nov: Auditin sin internet, Word, informe congelado, tiempos, automejora a las 4 a. m."""
import json
from datetime import datetime
from pathlib import Path

import pytest


def test_proveedor_local_no_piensa_y_limpia_el_razonamiento():
    from app import proveedores as pv
    local = {"base_url": "http://127.0.0.1:11434/v1"}
    assert pv.es_local(local) and not pv.es_local({"base_url": "https://api.deepseek.com/v1"})
    assert pv._cuerpo(local, {"model": "m"})["reasoning_effort"] == "none"
    assert "reasoning_effort" not in pv._cuerpo({"base_url": "https://api.openai.com/v1"}, {"model": "m"})
    assert pv._limpiar("<think>pienso…</think>\nRespuesta") == "Respuesta"


def test_auditin_sin_herramientas_recupera_con_python(tmp_path, monkeypatch):
    """Si el modelo local no admite herramientas, Python lee los datos y el modelo solo redacta."""
    from agente.comun import Espacio
    from app import asistente, proveedores as pv
    esp = Espacio(tmp_path)
    (esp.clientes / "8300970309").mkdir(parents=True)
    (esp.clientes / "8300970309" / "ficha.yaml").write_text("copropiedad: Arboleda Park\n", encoding="utf-8")
    (tmp_path / "acuerdos").mkdir()
    (tmp_path / "acuerdos" / "r.md").write_text("Auditin funciona sin internet.", encoding="utf-8")
    pv.guardar_proveedor(tmp_path, "ollama", {"nombre": "Ollama", "tipo": "openai", "base_url": "http://127.0.0.1:11434/v1",
                                              "modelos": ["m"]})
    pv.asignar(tmp_path, "asistente", "ollama", "m")

    def sin_herramientas(*a, **k):
        raise pv.SinHerramientas("no tools")
        yield                                                  # noqa: es un generador
    recibido = {}
    monkeypatch.setattr(pv, "conversar", sin_herramientas)
    monkeypatch.setattr(pv, "texto", lambda raiz, uso, sis, usuario, **k: recibido.setdefault("u", usuario) and "ok")
    evs = list(asistente.conversar(esp, [{"role": "user", "content": "¿Qué se acordó en la reunión sobre Arboleda Park?"}],
                                   "todo", "2026-08", None))
    assert [e["nombre"] for e in evs if e["tipo"] == "herramienta"][0] == "ver_acuerdos"
    assert "Auditin funciona sin internet" in recibido["u"] and any(e["tipo"] == "texto" for e in evs)


def test_word_de_validacion_y_opinion(tmp_path):
    from docx import Document
    from revisoria import word
    from revisoria.validador import Hallazgo
    hs = [Hallazgo("V01", "ALTA", "Liquidez", "Flujo negativo", "Detalle", ["CUADRO EFECTIVO!C5"], "Revisar")]
    r = word.validacion_docx(hs, {"periodo_etiqueta": "Agosto 2026"}, {"copropiedad": "Picasso"}, {"excel": "x.xlsx"},
                             tmp_path / "v.docx")
    texto = "\n".join(p.text for p in Document(r).paragraphs)
    assert "V01" in texto and "CUADRO EFECTIVO!C5" in texto
    r = word.markdown_a_docx("# T\n\n## 1. Hallazgo\n**Nivel:** ALTO\n\n- punto\n\n| a | b |\n|---|---|\n| 1 | 2 |",
                             tmp_path / "o.docx", "Opinión", "Picasso")
    d = Document(r)
    assert any("1. Hallazgo" in p.text for p in d.paragraphs) and d.tables[0].cell(1, 1).text == "2"


def test_automejora_corre_una_vez_al_dia_desde_las_4(tmp_path):
    from agente.comun import Espacio
    from app import rsi
    esp = Espacio(tmp_path)
    assert rsi.config(esp)["hora"] == "04:00"
    assert not rsi.toca_programado(esp, datetime(2026, 11, 7, 3, 59))
    assert rsi.toca_programado(esp, datetime(2026, 11, 7, 4, 0))
    rsi.marcar_programado(esp, datetime(2026, 11, 7, 4, 0))
    assert not rsi.toca_programado(esp, datetime(2026, 11, 7, 9, 0))
    assert rsi.toca_programado(esp, datetime(2026, 11, 8, 4, 1))
    with pytest.raises(ValueError):
        rsi.guardar_config(esp, "proponer", 0.03, 3, "25:00")


def test_tiempo_en_linea_solo_cuenta_latidos_seguidos(tmp_path, monkeypatch):
    from app import seguridad
    U = seguridad.Usuarios(tmp_path)
    t = [1_000_000.0]
    monkeypatch.setattr(seguridad.time, "time", lambda: t[0])
    U.latido("ana")
    t[0] += 60; U.latido("ana")
    t[0] += 60; U.latido("ana")
    t[0] += 3600; U.latido("ana")                             # una hora sin latidos (pestaña cerrada) no cuenta
    assert U.tiempo_en_linea()["ana"]["mes"] == 120


def test_informe_aprobado_queda_congelado(tmp_path):
    import os
    from agente import orquestador
    sal = tmp_path
    (sal / "Informe_Final.pptx").write_bytes(b"x")
    est = {"archivos": {"final": "Informe_Final.pptx"}}
    orquestador._congelar(est, sal)
    assert "Informe_Final.pptx" in est["congelado"]["archivos"]
    assert not os.access(sal / "Informe_Final.pptx", os.W_OK)


def test_concepto_masivo_word_por_copropiedad_con_infografia_y_anexo(tmp_path):
    import zipfile
    from docx import Document
    from agente.comun import Espacio
    from app import conceptos
    esp = Espacio(tmp_path)
    for cid, nombre in (("9001", "Arboleda Park"), ("9002", "Picasso PH")):
        (esp.clientes / cid).mkdir(parents=True)
        (esp.clientes / cid / "ficha.yaml").write_text(f"copropiedad: {nombre}\n", encoding="utf-8")
    frag = [{"documento": "Ley 675 de 2001", "referencia": "Artículo 35", "version": "compilada",
             "texto": "El fondo de imprevistos se constituye con un porcentaje de recargo no inferior al uno por ciento."}]
    vistos = {}

    def redactar(sistema, usuario, esquema):
        vistos["usuario"] = usuario
        return {"asunto": "Fondo de imprevistos", "texto": "## Contexto\nTexto del concepto (F1).",
                "infografia": {"titulo": "Fondo de imprevistos", "norma": ["Mínimo 1 % (F1)"], "opinion": ["Mantenerlo al día"],
                               "recomendacion": "Revisar el presupuesto."}}
    d = conceptos.crear(esp, "Fondo de imprevistos", "Esta Revisoría Fiscal recomienda mantener el fondo al día.", frag, [],
                        "Revisora", redactar=redactar)
    assert "[F0] OPINIÓN DEL REVISOR" in vistos["usuario"] and "[F1] Norma: Ley 675 de 2001" in vistos["usuario"]
    assert d["estado"] == "borrador" and d["clientes"] == ["9001", "9002"]
    d = conceptos.emitir(esp, d["id"], "Revisora")
    assert d["estado"] == "emitido" and len(d["emitido"]["archivos"]) == 2
    z = zipfile.ZipFile(conceptos.archivo(esp, d["id"], d["emitido"]["zip"]))
    assert len(z.namelist()) == 2
    doc = Document(conceptos.archivo(esp, d["id"], d["emitido"]["archivos"][0]))
    texto = "\n".join(p.text for p in doc.paragraphs)
    assert "Anexo: texto de la norma citada" in texto and "porcentaje de recargo" in texto
    assert doc.inline_shapes                                     # la infografía va en el documento
    with pytest.raises(ValueError):
        conceptos.guardar(esp, d["id"], {"texto": "otro"}, "Alguien")      # emitido = congelado


def test_concepto_marca_citas_a_fuentes_inexistentes():
    from app import conceptos
    d = {"fragmentos": [{"documento": "Ley 675"}], "texto": "Según (F2, Ley 675) y (F0)…", "infografia": {"norma": ["x (F1)"]}}
    assert conceptos.citas_invalidas(d) == ["F2"]


def test_sesion_sobrevive_al_reinicio_de_la_app(tmp_path):
    import json as _json
    from app import seguridad
    seguridad.persistir_sesiones(tmp_path)
    try:
        token = seguridad.abrir_sesion("ana")
        guardado = _json.loads((tmp_path / "sesiones.json").read_text(encoding="utf-8"))
        assert token not in guardado and len(guardado) == 1          # en disco solo la huella del token
        seguridad._sesiones.clear()                                  # «reinicio»: la memoria se pierde
        seguridad.persistir_sesiones(tmp_path)
        assert seguridad.sesion(token) == "ana"
        seguridad.cerrar_sesion(token)
        assert seguridad.sesion(token) is None
    finally:
        seguridad._archivo_sesiones = None
        seguridad._sesiones.clear()
