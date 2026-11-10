"""Agente orquestador: corre el lote del periodo y administra la cola de validación del profesional.

Por cada cliente (en paralelo):  normalizar → validar → cuadrar contra balance → generar → borrador de opinión
→ re-auditar → PENDIENTE_VALIDACION.  El único paso secuencial es la validación del revisor fiscal.
Nada sale (PDF) sin APROBADO.
"""
from __future__ import annotations

import shutil
import traceback
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

from .comun import APROBADO, BLOQUEADO, DEVUELTO, PENDIENTE, Espacio, ahora, guardar_json, leer_json


def _conteo(hs) -> dict:
    c = Counter(h.severidad for h in hs)
    return {s: c.get(s, 0) for s in ("ALTA", "MEDIA", "BAJA", "INFO")}


def procesar_cliente(raiz: str, periodo: str, cid: str, usar_llm: bool = False) -> dict:
    """Corre todos los agentes para un cliente. Devuelve el estado final."""
    from revisoria.excel import Periodo
    from revisoria.extractor import extraer
    from revisoria.generador_pptx import Informe
    from revisoria.narrativa import Redactor
    from revisoria.reporte import generar_html
    from revisoria.documental import extraer_documental, validar_documental
    from revisoria.informe_completo import InformeCompleto
    from revisoria.validador import Hallazgo, ordenar, validar_excel
    from revisoria import esquema, mapa as mapa_cliente, word
    from revisoria.excel import Libro

    from . import integridad, normalizador, reauditor, redaccion

    import time as _time
    inicio = _time.monotonic()
    esp = Espacio(Path(raiz))
    try:
        ficha = esp.ficha(cid)
    except FileNotFoundError as e:
        est = {"cliente": cid, "periodo": periodo, "estado": BLOQUEADO, "bitacora": []}
        return esp.evento(est, BLOQUEADO, str(e))
    ent, sal = esp.entrada(periodo, cid), esp.salida(periodo, cid)
    sal.mkdir(parents=True, exist_ok=True)
    est = esp.estado(periodo, cid)
    if est.get("estado") == APROBADO:                        # informe congelado: nada lo vuelve a generar ni a cambiar
        if not integridad.misma_revision(esp, periodo, cid, est.get("manifiesto")):
            esp.evento(est, None, "Alerta: las fuentes cambiaron después de la aprobación; el informe aprobado se "
                                  "conserva congelado y no se regenera")
            esp.guardar_estado(periodo, cid, est)
        return est
    est.pop("error", None)                                   # el error de un intento anterior no debe quedar
    est.update({"cliente": cid, "copropiedad": ficha.get("copropiedad", cid), "revisor": ficha.get("revisor", ""),
                "software": ficha.get("software", "excel_estandar"), "archivos": {}})
    try:
        # 1. Normalizador
        try:
            excel = normalizador.ubicar_excel(ent, ficha)
        except normalizador.ExcelAmbiguo as e:
            esp.evento(est, BLOQUEADO, str(e))
            esp.guardar_estado(periodo, cid, est)
            return est
        if not excel:
            esp.evento(est, BLOQUEADO, f"No llegó el Excel estándar en entradas/{periodo}/{cid}/ (no detiene el lote)")
            esp.guardar_estado(periodo, cid, est)
            return est
        dir_cli = esp.clientes / cid
        libro = Libro(str(excel))
        aprendidos = mapa_cliente.aprendidos(esp.conocimiento)
        mapa = mapa_cliente.cargar(dir_cli)
        h_mapa = []
        if mapa is None:                                     # cliente nuevo: se propone el mapa y queda guardado
            mapa = mapa_cliente.detectar(libro, aprendidos)
            mapa_cliente.guardar(dir_cli, mapa)
            esp.evento(est, None, "Mapa del cliente: propuesto por detección automática (pendiente de revisión)")
        else:
            ajustes = mapa_cliente.reparar(libro, mapa, aprendidos)
            for k, a in ajustes.items():                     # el mapa ya no calza: se ajusta provisionalmente
                antes = ", ".join(mapa_cliente._titulos(a["antes"]))
                if a["ahora"]:
                    h_mapa.append(Hallazgo("M04", "MEDIA", "Mapa del cliente", f"«{k}»: se usó «{a['ahora']}» en lugar de «{antes}»",
                                           f"La hoja «{antes}» ya no está en el Excel; {a['motivo']} («{a['ahora']}»). "
                                           "Se usó de forma provisional: confírmelo en Mapa del Excel.", []))
                else:
                    h_mapa.append(Hallazgo("M01", "MEDIA", "Mapa del cliente", f"«{k}»: falta la hoja «{antes}»",
                                           f"La hoja «{antes}» ya no está en el Excel y {a['motivo']}. "
                                           "La sección sale vacía: revise el Excel o el Mapa del Excel.", []))
            if ajustes:
                mapa = mapa_cliente.aplicar(mapa, ajustes)
                guardar_json(sal / "mapa_ajustes.json", ajustes)
                esp.evento(est, None, f"Mapa del cliente: {len(ajustes)} hoja(s) ya no coinciden; "
                           f"{sum(1 for a in ajustes.values() if a['ahora'])} ajustada(s) provisionalmente")
            nuevas = mapa_cliente.hojas_nuevas(libro, mapa)
            if nuevas:
                h_mapa.append(Hallazgo("M05", "INFO", "Mapa del cliente", f"El Excel trae {len(nuevas)} hoja(s) nueva(s)",
                                       "Hojas que no existían al aprobar el mapa: " + ", ".join(f"«{t}»" for t in nuevas)
                                       + ". Si contienen información del informe, asígnelas en Mapa del Excel.", []))
        h_norm = normalizador.revisar_estructura(excel, mapa)
        balance, h_bal = normalizador.leer_balance(esp, ent, ficha, cid)
        esp.evento(est, None, f"Normalizador: Excel «{excel.name}»"
                   + (f", balance «{balance['archivo']}» (perfil {balance['perfil']})" if balance else ", sin balance"))

        # 2. Validador
        meta = {k: ficha.get(k) for k in ("copropiedad", "nit", "unidades", "numero_informe") if ficha.get(k)}
        completo = str(ficha.get("informe", "completo")).lower() == "completo"

        def leer(mapa_, libro_):
            """Lee y valida el Excel con un mapa (el aprobado o, para el borrador, el que trae las propuestas)."""
            modelo_, libro_ = extraer(str(excel), Periodo.desde_texto(periodo), meta, mapa_, libro_)
            p_ = Periodo.desde_texto(modelo_["meta"]["periodo"])
            h_ = validar_excel(modelo_, libro_, p_)
            doc_ = extraer_documental(libro_, p_) if completo else {}
            h_ += validar_documental(doc_, p_) if completo else []
            cuadre_, h_cuadre_ = normalizador.cuadrar(esp, modelo_, balance)
            return modelo_, libro_, p_, h_, doc_, cuadre_, h_cuadre_, esquema.verificar(modelo_, doc_)

        def h_cobertura(cobertura_, libro_, mapa_, doc_):
            # una sección obligatoria ausente no detiene el informe (sale «pendiente»), pero exige decisión del revisor
            return ([Hallazgo("M02", ("ALTA" if c["estado"] in ("falta", "vacia") else "MEDIA") if c["obligatoria"] else "INFO",
                              "Esquema estándar",
                              f"{c['descripcion']}: {c['estado']}", f"{c['descripcion']} ({c['hoja']}): {c['detalle']}", [])
                     for c in cobertura_ if c["estado"] != "ok"]
                    + [Hallazgo("M03", "MEDIA", "Mapa del cliente", f["detalle"][:90],
                                f["detalle"] + " Esa sección se omite o puede salir mal: corrija el mapa o la hoja.", [])
                       for f in esquema.revisar_formato(libro_, mapa_.get("hojas", {}), doc_, mapa_.get("lectores"))])

        modelo, libro, p, h_excel, doc, cuadre, h_cuadre, cobertura = leer(mapa, libro)
        problemas_identidad = integridad.verificar_identidad(libro, ficha, modelo, periodo)
        guardar_json(sal / "integridad_identidad.json", {"ok": not problemas_identidad,
                                                          "problemas": problemas_identidad})
        if problemas_identidad:
            raise ValueError("Identidad del expediente no confirmada: " + " ".join(problemas_identidad))
        if not any(h.regla in ("M01", "M04") for h in h_mapa):  # los desajustes ya reportados por la reparación
            h_mapa += [Hallazgo("M01", "MEDIA", "Mapa del cliente", d[:90], d + ". Revise el mapa del cliente.", [])
                       for d in dict.fromkeys(libro.desajustes)]
        # 2b. Adaptador: si alguna sección no se pudo leer, propone (sin aprobar) recetas y correcciones del mapa para
        # el Excel de este cliente. Solo el modo automático (sin IA) corre en el lote; con IA se pide desde la app.
        # El BORRADOR ya se arma con lo propuesto (aplicado solo en memoria); el informe final no se puede aprobar
        # mientras queden propuestas sin aprobar (est["propuestas_sin_aprobar"]).
        sin_aprobar = []
        if any(c["estado"] != "ok" for c in cobertura):
            try:
                from . import adaptador
                ad = adaptador.adaptar(esp, cid, periodo)
                n_prop, n_aj = ad["resumen"]["propuestas"], ad["resumen"].get("ajustes_mapa", 0)
                sin_aprobar = adaptador.secciones_con_propuesta(ad)
                if sin_aprobar:
                    mapa_borrador = adaptador.mapa_con_propuestas(mapa, ad)
                    modelo, libro, p, h_excel, doc, cuadre, h_cuadre, cobertura = leer(mapa_borrador, Libro(str(excel)))
                    h_mapa += h_cobertura(cobertura, libro, mapa_borrador, doc)
                    nombres = {x["seccion"]: x["descripcion"] for x in ad["secciones"]}
                    que = " y ".join(x for x in (f"{n_prop} receta(s)" if n_prop else "",
                                                 f"{n_aj} corrección(es) del mapa" if n_aj else "") if x)
                    h_mapa.append(Hallazgo(
                        "A01", "ALTA", "Adaptación al Excel",
                        f"El borrador usa {que} del agente adaptador sin aprobar",
                        "Estas secciones se leyeron con propuestas del agente adaptador que el revisor aún no aprueba: "
                        + ", ".join(nombres.get(x, x) for x in sin_aprobar)
                        + f". Con ellas quedan {ad['resumen']['completas_con_propuestas']} de {ad['resumen']['secciones']} "
                        "secciones completas. Sus diapositivas llevan la etiqueta «Leído con propuesta del agente». Revise "
                        "lo leído y apruébelo en Mapa del Excel → Adaptación al Excel, y vuelva a validar: el informe no se "
                        "puede aprobar antes.", ["adaptacion.json"]))
                else:
                    h_mapa += h_cobertura(cobertura, libro, mapa, doc)
                esp.evento(est, None, f"Adaptador: {n_prop} receta(s) y {n_aj} corrección(es) del mapa propuestas"
                           + (f"; el borrador ya las usa en {len(sin_aprobar)} sección(es), pendientes de aprobación"
                              if sin_aprobar else ""))
            except Exception as e:                           # el adaptador nunca detiene el informe
                esp.evento(est, None, f"Adaptador: no se pudo proponer ({type(e).__name__}: {e})")
                if not any(h.regla == "M02" for h in h_mapa):
                    h_mapa += h_cobertura(cobertura, libro, mapa, doc)
        else:
            h_mapa += h_cobertura(cobertura, libro, mapa, doc)
        est["propuestas_sin_aprobar"] = sin_aprobar
        nit = str(ficha.get("nit") or "").strip()
        if not nit or ficha.get("nit_provisional") or "PROVISIONAL" in nit.upper():
            h_mapa.append(Hallazgo("F01", "ALTA", "Ficha del cliente",
                                   "NIT provisional en la ficha" if nit else "Falta el NIT en la ficha",
                                   f"La ficha tiene el NIT «{nit or 'vacío'}» marcado como provisional. Antes de emitir, "
                                   "reemplácelo por el NIT real (RUT de la copropiedad) con «Cambiar NIT» y vuelva a validar.",
                                   ["ficha.yaml"]))
        hallazgos = ordenar(h_norm + h_bal + h_excel + h_cuadre + h_mapa)
        esp.evento(est, None, f"Validador: {len(hallazgos)} hallazgo(s) antes de generar (el total final está en hallazgos.json) · cuadre contra balance: "
                   + (f"{sum(1 for c in cuadre if c['ok'])}/{sum(1 for c in cuadre if c['ok'] is not None)} OK" if cuadre else "no aplica"))

        # 3. Generador
        base = f"{cid}_{periodo}"
        ruta_pptx = sal / (f"Informe_Revisoria_{base}.pptx" if completo else f"Informe_Financiero_{base}.pptx")
        if completo:
            inf = InformeCompleto(modelo, ficha, doc, Redactor(usar_llm=usar_llm), propuestas=sin_aprobar).construir()
            hallazgos = ordenar(hallazgos + [Hallazgo("G01", "INFO", "Generación", a[:90], a, []) for a in inf.avisos]
                                + [Hallazgo("G02", "ALTA", "Generación", "Diapositiva conservada como pendiente por un error",
                                            a + ". La diapositiva salió como «información pendiente» para mantener la "
                                            "estructura del informe; corrija el dato o reporte el error antes de emitir.", [])
                                   for a in inf.errores])
            guardar_json(sal / "estructura_generada.json", inf.plan)
        else:
            inf = Informe(modelo, ficha, Redactor(usar_llm=usar_llm), propuestas=sin_aprobar).construir()
        inf.guardar(ruta_pptx)
        esp.evento(est, None, f"Generador: {inf.n} diapositivas ({'informe completo' if completo else 'sección financiera'}) "
                   "con tablas y gráficos nativos")

        # 4. Redactor (borrador de opinión)
        texto, items = redaccion.borrador(esp, modelo, ficha, hallazgos, usar_llm)
        (sal / f"Borrador_Opinion_{base}.md").write_text(texto, encoding="utf-8")     # el que se edita en la app
        word.markdown_a_docx(texto, sal / f"Borrador_Opinion_{base}.docx", "Borrador de hallazgos y opinión",
                             f"{ficha.get('copropiedad', cid)} · {modelo['meta']['periodo_etiqueta']}",
                             "BORRADOR — pendiente de revisión y firma del revisor fiscal")
        guardar_json(sal / "borrador_items.json", items)
        esp.evento(est, None, f"Redactor: {len(items)} hallazgo(s) propuestos desde la biblioteca")

        # 5. Re-auditor (independiente)
        abiertos = leer_json(esp.salida(esp.periodo_anterior(periodo), cid) / "hallazgos_abiertos.json", [])
        h_reaud = reauditor.reauditar(ruta_pptx, modelo, libro, p, ficha, h_excel + h_cuadre + h_mapa, abiertos, extra=doc)
        esp.evento(est, None, f"Re-auditor: {len(h_reaud)} observación(es) sobre el informe generado")

        # Salidas
        todos = ordenar(hallazgos + h_reaud)
        ruta_html = sal / f"Validacion_{base}.html"
        ruta_html.write_text(generar_html(todos, modelo["meta"], ficha, {"excel": excel.name}), encoding="utf-8")
        ruta_docx = word.validacion_docx(todos, modelo["meta"], ficha, {"excel": excel.name}, sal / f"Validacion_{base}.docx")
        guardar_json(sal / "modelo.json", modelo)
        guardar_json(sal / "hallazgos.json", hallazgos)
        guardar_json(sal / "reauditoria.json", h_reaud)
        guardar_json(sal / "cuadre_balance.json", cuadre)
        guardar_json(sal / "cobertura_esquema.json", cobertura)
        est["archivos"] = {"excel": str(excel), "pptx": ruta_pptx.name, "validacion": ruta_docx.name,
                           "validacion_html": ruta_html.name, "borrador": f"Borrador_Opinion_{base}.docx",
                           "borrador_md": f"Borrador_Opinion_{base}.md"}
        est["resumen"] = {"diapositivas": inf.n, "hallazgos": _conteo(hallazgos), "reauditoria": _conteo(h_reaud),
                          "propuestas_sin_aprobar": sin_aprobar, "diapositivas_con_propuesta": len(inf.con_propuesta),
                          "cuadre": cuadre, "borrador": [i["titulo"] for i in items],
                          "periodo_etiqueta": modelo["meta"]["periodo_etiqueta"],
                          "cifras": {"liquidez": (modelo.get("liquidez") or {}).get("total"),
                                     "cartera": (modelo.get("cartera") or {}).get("total"),
                                     "cxp": (modelo.get("cxp") or {}).get("total"),
                                     "resultado": (modelo.get("patrimonio") or {}).get("resultado")}}
        est["manifiesto"] = integridad.manifiesto(esp, periodo, cid)
        est["resumen"]["segundos"] = round(_time.monotonic() - inicio, 1)       # tiempo de procesamiento del cliente
        esp.evento(est, PENDIENTE, f"Listo para validación del revisor fiscal · procesado en {est['resumen']['segundos']:.0f} s")
    except Exception as e:                                     # un cliente con problemas no detiene el lote
        est["error"] = traceback.format_exc(limit=3)
        esp.evento(est, BLOQUEADO, f"Error: {e}")
    esp.guardar_estado(periodo, cid, est)
    return est


def correr(esp: Espacio, periodo: str, clientes: list[str] | None = None, paralelo: int = 4,
           usar_llm: bool = False, log=print) -> list[dict]:
    ids = clientes or esp.ids_clientes()
    log(f"Lote {periodo}: {len(ids)} cliente(s) · {paralelo} en paralelo")
    resultados = {}
    if paralelo <= 1 or len(ids) == 1:
        for cid in ids:
            resultados[cid] = procesar_cliente(str(esp.raiz), periodo, cid, usar_llm)
            _log_cliente(log, resultados[cid])
    else:
        with ProcessPoolExecutor(max_workers=paralelo) as ex:
            fut = {ex.submit(procesar_cliente, str(esp.raiz), periodo, cid, usar_llm): cid for cid in ids}
            for f in as_completed(fut):
                cid = fut[f]
                try:
                    resultados[cid] = f.result()
                except Exception as e:                       # el proceso del cliente murió: se bloquea solo ese
                    est = esp.estado(periodo, cid)
                    resultados[cid] = esp.evento(est, BLOQUEADO, f"Error del proceso: {e}")
                    esp.guardar_estado(periodo, cid, est)
                _log_cliente(log, resultados[fut[f]])
    return [resultados[c] for c in ids]


def _log_cliente(log, est):
    r = est.get("resumen") or {}
    h, ra = r.get("hallazgos", {}), r.get("reauditoria", {})
    log(f"  · {est.get('copropiedad', est['cliente'])[:42]:<42} {est['estado']:<22} "
        + (f"Excel A{h.get('ALTA', 0)}/M{h.get('MEDIA', 0)} · re-auditor A{ra.get('ALTA', 0)}/M{ra.get('MEDIA', 0)}"
           if r else est["bitacora"][-1]["detalle"][:60]))


def cola(esp: Espacio, periodo: str) -> list[dict]:
    """Informes pendientes de validación en el orden del calendario de visitas."""
    return [e for e in (esp.estado(periodo, c) for c in esp.ids_clientes()) if e.get("estado") == PENDIENTE]


def _pdf(pptx: Path) -> Path | None:
    from revisoria.libreoffice import convertir
    return convertir(pptx, "pdf", pptx.parent)


def aprobar(esp: Espacio, periodo: str, cid: str, revisor: str, justificacion: str = "",
            pptx_final: str | None = None, borrador_final: str | None = None) -> dict:
    from . import integridad, mejora, reauditor
    est = esp.estado(periodo, cid)
    if est.get("estado") != PENDIENTE:
        raise SystemExit(f"{cid}: solo se aprueba un informe en {PENDIENTE} (estado actual: {est.get('estado')}).")
    if not integridad.misma_revision(esp, periodo, cid, est.get("manifiesto")):
        raise SystemExit(f"{cid}: los archivos, la ficha o el mapa cambiaron después de validar. Vuelva a procesar antes de aprobar.")
    pendientes = est.get("propuestas_sin_aprobar") or (est.get("resumen") or {}).get("propuestas_sin_aprobar") or []
    if pendientes:                                 # no se justifica: el informe final solo usa lecturas aprobadas
        raise SystemExit(f"{cid}: el borrador usa propuestas del agente adaptador que aún no se aprueban "
                         f"({', '.join(pendientes)}). Revíselas y apruébelas en Mapa del Excel → Adaptación al Excel, "
                         "vuelva a validar y después apruebe el informe.")
    sal = esp.salida(periodo, cid)
    r = est["resumen"]
    hallazgos_guardados = leer_json(sal / "hallazgos.json", [])
    formulas_sin_resultado = [h for h in hallazgos_guardados if h.get("regla") == "V21"]
    if formulas_sin_resultado:
        raise SystemExit(f"{cid}: el Excel contiene fórmulas sin resultado calculado. "
                         "Recalcule y guarde el archivo antes de aprobar; este control no se puede justificar.")
    # R06 solo repite los hallazgos ALTA del Excel: no se cuenta dos veces
    reaud = leer_json(sal / "reauditoria.json", None)
    altas_reaud = (sum(1 for h in reaud if h.get("severidad") == "ALTA" and h.get("regla") != "R06")
                   if reaud is not None else r["reauditoria"].get("ALTA", 0))
    altas = r["hallazgos"].get("ALTA", 0) + altas_reaud
    if altas and not justificacion.strip():
        raise SystemExit(f"{cid}: hay {altas} hallazgo(s) ALTA. Corrija el Excel y vuelva a correr, o apruebe con "
                         "--justificacion \"...\" explicando la decisión del revisor fiscal.")

    # Captura de correcciones (automejora)
    generado = sal / est["archivos"]["pptx"]
    final = generado
    cambios = []
    if pptx_final:
        final = Path(pptx_final)
        cambios += mejora.diferencias_pptx(generado, final)
    if borrador_final:
        items = leer_json(sal / "borrador_items.json", [])
        cambios_borrador = mejora.diferencias_borrador(items, Path(borrador_final).read_text(encoding="utf-8"))
        propuesta_claude = sal / mejora.PROPUESTA_CLAUDE
        if propuesta_claude.exists():                        # lo que redactó Claude no es corrección del revisor
            mejora.marcar_origen_claude(cambios_borrador, items, propuesta_claude.read_text(encoding="utf-8"))
        cambios += cambios_borrador
    # El archivo realmente emitido se revisa de nuevo, después de todas las ediciones.
    from revisoria import mapa as mapa_cliente
    from revisoria.documental import extraer_documental
    from revisoria.excel import Libro, Periodo
    modelo = leer_json(sal / "modelo.json", {})
    p = Periodo.desde_texto(periodo)
    libro = Libro(est["archivos"]["excel"], mapa_cliente.cargar(esp.clientes / cid))
    extra = extraer_documental(libro, p) if str(esp.ficha(cid).get("informe", "completo")).lower() == "completo" else {}
    revision_final = reauditor.reauditar(final, modelo, libro, p, esp.ficha(cid), [], [], extra=extra)
    guardar_json(sal / "reauditoria_final.json", revision_final)
    bloqueantes = [h for h in revision_final if h.severidad == "ALTA" and h.regla in ("R01", "R02", "R03")]
    if bloqueantes:
        raise SystemExit(f"{cid}: el informe final tiene {len(bloqueantes)} problema(s) de integridad. Corrija las diapositivas y vuelva a revisar.")

    # Solo después de superar la auditoría se publica la edición como archivo final.
    if pptx_final:
        destino = sal / f"Informe_Final_{cid}_{periodo}.pptx"
        if final.resolve() != destino.resolve():
            shutil.copy2(final, destino)
        final = destino

    n = mejora.registrar(esp, periodo, cid, cambios) if cambios else 0

    pdf = _pdf(final)
    est["archivos"]["final"] = final.name
    if pdf:
        est["archivos"]["pdf"] = pdf.name
    # opinión final del revisor en Word (la editada si la hay; si no, la generada)
    from revisoria import word
    opinion = Path(borrador_final) if borrador_final else sal / f"Borrador_Opinion_{cid}_{periodo}.md"
    if opinion.exists():
        ruta_op = word.markdown_a_docx(opinion.read_text(encoding="utf-8"), sal / f"Opinion_Revisor_{cid}_{periodo}.docx",
                                       "Hallazgos y opinión del revisor fiscal",
                                       f"{esp.ficha(cid).get('copropiedad', cid)} · aprobada por {revisor}")
        est["archivos"]["opinion"] = ruta_op.name
    est["archivos"]["reauditoria_final"] = "reauditoria_final.json"
    est["aprobacion"] = {"revisor": revisor, "justificacion": justificacion, "correcciones": n,
                         "revision": est["manifiesto"]["revision"], "sha256_final": integridad.huella_archivo(final),
                         "version_biblioteca": mejora.version_vigente(esp)}
    esp.evento(est, APROBADO, f"Aprobado por {revisor}" + (f" · {n} corrección(es) capturadas" if n else "")
               + (f" · justificación: {justificacion}" if justificacion else "")
               + ("" if pdf else " · PDF: exportar desde PowerPoint (LibreOffice no disponible)"), actor=revisor)
    _congelar(est, sal)
    esp.guardar_estado(periodo, cid, est)

    # Memoria de hallazgos abiertos para el mes siguiente y golden set
    hs = leer_json(sal / "hallazgos.json", [])
    guardar_json(sal / "hallazgos_abiertos.json", [h for h in hs if h["severidad"] in ("ALTA", "MEDIA")])
    mejora.agregar_golden(esp, periodo, cid, Path(est["archivos"]["excel"]), sal / "modelo.json", esp.ficha(cid))
    return est


def _congelar(est: dict, sal: Path):
    """Informe aprobado = formato de calidad: sus entregables quedan de solo lectura y con su huella registrada.
    Ni la automejora ni nadie los modifica (cargar otro Excel para ese periodo se rechaza)."""
    import stat
    from . import integridad
    huellas = {}
    for clave in ("final", "pdf", "opinion", "validacion", "pptx"):
        nombre = (est.get("archivos") or {}).get(clave)
        ruta = sal / nombre if nombre else None
        if ruta and ruta.exists():
            huellas[nombre] = integridad.huella_archivo(ruta)
            ruta.chmod(stat.S_IRUSR | stat.S_IRGRP)
    est["congelado"] = {"fecha": ahora(), "archivos": huellas}


def devolver(esp: Espacio, periodo: str, cid: str, revisor: str, motivo: str) -> dict:
    from . import mejora
    est = esp.estado(periodo, cid)
    if est.get("estado") not in (PENDIENTE,):
        raise SystemExit(f"{cid}: solo se devuelve un informe pendiente de validación.")
    mejora.registrar(esp, periodo, cid, [{"tipo": "devolucion", "motivo": motivo}])
    esp.evento(est, DEVUELTO, f"Devuelto por {revisor}: {motivo}", actor=revisor)
    esp.guardar_estado(periodo, cid, est)
    return est
