"""Auditin: asistente de consulta sobre los clientes y las bases de conocimiento del espacio de trabajo.

Solo LEE: hallazgos, cifras del modelo con su celda fuente, el Excel del cliente, la bitácora y la memoria
de mejora continua. No aprueba, no edita y no calcula cifras: las cita de las herramientas con su fuente.
"""
from __future__ import annotations

import json
from pathlib import Path

from agente.comun import Espacio, leer_json

import contextvars

LIMITE = 30_000                       # caracteres por resultado de herramienta
LIMITE_LOCAL = 8_000                  # con un modelo local el contexto es más corto
_limite = contextvars.ContextVar("limite", default=LIMITE)

BASES = {
    "todo": "Todas las bases",
    "clientes": "Clientes (ficha, documentos y mapa del Excel)",
    "informes": "Lo que salió en el informe de cada cliente",
    "hallazgos": "Hallazgos y re-auditoría",
    "cifras": "Cifras del informe (modelo con celda fuente)",
    "excel": "Excel de papeles de trabajo",
    "mejora": "Mejora continua (correcciones, propuestas, métricas, versiones)",
    "biblioteca": "Biblioteca de opiniones del revisor",
    "auditoria": "Bitácora, aprobaciones y auditoría mensual",
    "normativa": "Base normativa (leyes, NIIF y manuales cargados)",
    "acuerdos": "Acuerdos de las reuniones con el despacho",
}

SISTEMA = """Eres Auditin, el agente de consulta de LC Counter Contadores y Auditores SAS para la Revisoría \
Fiscal de propiedades horizontales en Colombia. Respondes preguntas del equipo sobre los clientes: la ficha y los \
documentos de cada uno, su Excel de papeles de trabajo, lo que salió en su informe (diapositivas, secciones \
pendientes, propuestas del agente adaptador), los hallazgos, las cifras, la bitácora de aprobaciones, la memoria de \
mejora continua y lo acordado en las reuniones con el despacho. Tu tono es cercano y profesional; te presentas \
como Auditin solo si te lo preguntan.

Reglas:
- Consulta las herramientas antes de responder sobre datos; no respondas de memoria.
- Cada cifra que menciones debe venir de un resultado de herramienta, con su fuente (hoja y celda, o archivo). \
Si haces una operación con cifras (sumar, restar, comparar), dilo y muestra los valores de origen.
- Si la información no está en el espacio de trabajo, dilo claramente. No inventes opiniones del revisor fiscal, \
normas ni cifras.
- Normas, artículos y manuales: cítalos SOLO a partir de `buscar_normativa`, con documento y referencia exactos \
(p. ej. «Ley 675 de 2001, artículo 35»). Si la búsqueda no trae nada pertinente, di que la base normativa cargada \
no lo contiene; no respondas de memoria sobre el contenido de una norma.
- No apruebas ni devuelves informes ni cambias nada: si te lo piden, explica que esa decisión es del revisor \
fiscal y que se hace desde la aplicación.
- Responde en español, breve y concreto, con tablas en Markdown cuando ayuden. Usa los nombres de las reglas \
(V01–V20, D01–D10, R01–R07, N01–N04) tal como aparecen en los datos."""


def _esquema(props: dict, req: list[str]) -> dict:
    return {"type": "object", "properties": props, "required": req, "additionalProperties": False}


P = {"type": "string", "description": "Periodo AAAA-MM"}
C = {"type": "string", "description": "Identificador del cliente: NIT sin puntos ni guion (p. ej. 9001234567)"}

HERRAMIENTAS = {
    "listar_clientes": ({"todo", "hallazgos", "cifras", "excel", "auditoria"},
                        "Lista los clientes del periodo con su estado, conteo de hallazgos y cifras principales.",
                        _esquema({"periodo": P}, ["periodo"])),
    "ver_hallazgos": ({"todo", "hallazgos"},
                      "Hallazgos de validación del Excel y de la re-auditoría de un cliente, con regla, severidad, "
                      "detalle y celdas. Filtra opcionalmente por severidad (ALTA, MEDIA, BAJA, INFO).",
                      _esquema({"periodo": P, "cliente": C,
                                "severidad": {"type": "string", "enum": ["", "ALTA", "MEDIA", "BAJA", "INFO"]}},
                               ["periodo", "cliente", "severidad"])),
    "ver_cifras": ({"todo", "cifras"},
                   "Cifras extraídas del Excel para el informe (modelo), con la celda fuente de cada una. "
                   "Use seccion='' para ver las secciones disponibles.",
                   _esquema({"periodo": P, "cliente": C, "seccion": {"type": "string"}},
                            ["periodo", "cliente", "seccion"])),
    "buscar_en_excel": ({"todo", "excel"},
                        "Busca un texto (sin distinguir mayúsculas) en todas las hojas del Excel del cliente y "
                        "devuelve hoja, celda y valor de hasta 60 coincidencias, con los valores de la misma fila.",
                        _esquema({"periodo": P, "cliente": C, "texto": {"type": "string"}},
                                 ["periodo", "cliente", "texto"])),
    "ver_conocimiento": ({"todo", "mejora", "biblioteca", "auditoria"},
                         "Memoria del sistema: correcciones del revisor, propuestas de mejora, métricas, versiones "
                         "publicadas, biblioteca de opiniones, auditorías mensuales o borradores para la revisora.",
                         _esquema({"tipo": {"type": "string", "enum": ["correcciones", "propuestas", "metricas",
                                                                        "versiones", "biblioteca", "auditorias",
                                                                        "borradores"]}}, ["tipo"])),
    "buscar_normativa": ({"todo", "normativa", "hallazgos"},
                         "Busca en la base normativa cargada por el despacho (leyes, decretos, NIIF, reglamentos y "
                         "manuales de las copropiedades). Devuelve fragmentos con documento, artículo o numeral y "
                         "texto. Use cliente='' para buscar solo en lo general.",
                         _esquema({"consulta": {"type": "string"}, "cliente": {"type": "string"}}, ["consulta", "cliente"])),
    "ver_cliente": ({"todo", "clientes"},
                    "Ficha de un cliente (copropiedad, NIT, unidades, revisor, notas), sus documentos cargados, las "
                    "hojas de su Excel que usa el informe (mapa) y las recetas de lectura aprobadas.",
                    _esquema({"cliente": C}, ["cliente"])),
    "ver_informe": ({"todo", "informes", "cifras"},
                    "Lo que salió en el informe de un cliente en un periodo: estado, número de diapositivas, título "
                    "de cada diapositiva (y si está pendiente o leída con una propuesta sin aprobar), secciones "
                    "completas y pendientes con su motivo, y propuestas del agente adaptador.",
                    _esquema({"periodo": P, "cliente": C}, ["periodo", "cliente"])),
    "ver_acuerdos": ({"todo", "acuerdos"},
                     "Acuerdos y pendientes definidos en las reuniones con el despacho (texto completo).",
                     _esquema({}, [])),
    "ver_bitacora": ({"todo", "auditoria"},
                     "Bitácora de un informe: estados, fechas, quién aprobó o devolvió, justificaciones y archivos.",
                     _esquema({"periodo": P, "cliente": C}, ["periodo", "cliente"])),
}


def disponible(raiz=None) -> dict:
    from .proveedores import disponible as disp
    if raiz is None:
        return {"ok": False, "modelo": "", "ayuda": "Sin espacio de trabajo."}
    return disp(raiz, "asistente")


# ----------------------------------------------------------------------------- herramientas (solo lectura)

def _cortar(o) -> str:
    t = json.dumps(o, ensure_ascii=False, default=str)
    n = _limite.get()
    return t if len(t) <= n else t[:n] + f"… [recortado: {len(t)} caracteres en total]"


def _valido(esp: Espacio, periodo: str, cid: str) -> str | None:
    import re
    if not re.match(r"^\d{4}-\d{2}$", periodo or "") or not re.match(r"^[a-z0-9_]{1,40}$", cid or ""):
        return "Periodo o cliente con formato inválido."
    if cid not in esp.ids_clientes():
        return f"No existe el cliente «{cid}». Clientes: {', '.join(esp.ids_clientes())}."
    return None


def ejecutar(esp: Espacio, nombre: str, a: dict) -> str:
    if nombre == "listar_clientes":
        out = []
        for cid in esp.ids_clientes():
            e = esp.estado(a["periodo"], cid)
            r = e.get("resumen") or {}
            out.append({"cliente": cid, "copropiedad": e.get("copropiedad", cid), "estado": e.get("estado"),
                        "hallazgos_excel": r.get("hallazgos"), "reauditoria": r.get("reauditoria"),
                        "cifras": r.get("cifras"), "diapositivas": r.get("diapositivas")})
        return _cortar(out)
    if nombre == "ver_conocimiento":
        c, t = esp.conocimiento, a["tipo"]
        if t == "correcciones":
            ruta = c / "correcciones.jsonl"
            datos = [json.loads(x) for x in ruta.read_text(encoding="utf-8").splitlines() if x.strip()] if ruta.exists() else []
        elif t == "biblioteca":
            from agente.redaccion import biblioteca
            datos = {"vigente": biblioteca(esp), "nota": "Incluye los textos base y lo aprobado por el revisor."}
        elif t == "metricas":
            datos = leer_json(c / "metricas.json", [])[-6:]
        elif t == "auditorias":
            datos = [leer_json(p) for p in sorted((c / "auditorias").glob("*.json"))] if (c / "auditorias").exists() else []
        elif t == "borradores":
            datos = {p.name: p.read_text(encoding="utf-8") for p in sorted((c / "propuestas").glob("*.md"))} \
                if (c / "propuestas").exists() else {}
        else:
            datos = leer_json(c / f"{t}.json", [])
        return _cortar(datos or f"Sin datos de «{t}» en el espacio de trabajo.")

    if nombre == "ver_acuerdos":
        carpeta = esp.raiz / "acuerdos"
        docs = {x.name: x.read_text(encoding="utf-8") for x in sorted(carpeta.glob("*.md"))} if carpeta.exists() else {}
        return _cortar(docs or "Aún no hay acuerdos de reuniones guardados en el espacio de trabajo.")
    if nombre == "ver_cliente":
        cid = a.get("cliente", "")
        if cid not in esp.ids_clientes():
            return f"No existe el cliente «{cid}». Clientes: {', '.join(esp.ids_clientes())}."
        from revisoria import mapa as mapa_cliente
        d = esp.clientes / cid
        m = mapa_cliente.cargar(d) or {}
        docs = sorted(x.name for x in (d / "documentos").glob("*") if x.is_file()) if (d / "documentos").exists() else []
        return _cortar({"ficha": esp.ficha(cid), "documentos": docs, "tiene_logo": any(d.glob("logo.*")),
                        "normatividad": leer_json(d / "normatividad.json", {}),
                        "mapa": {"aprobado_por": m.get("aprobado_por"), "hojas": m.get("hojas"), "sin_usar": m.get("sin_usar")},
                        "recetas_aprobadas": sorted((m.get("lectores") or {}))})
    if nombre == "buscar_normativa":
        from .normativa import instancia
        cli = a.get("cliente") or None
        if cli and cli not in esp.ids_clientes():
            cli = None
        r = instancia(esp.raiz).buscar(a["consulta"], cli, k=6)
        if not r:
            docs = instancia(esp.raiz).documentos()
            return ("Sin resultados. Documentos cargados: " + "; ".join(f"{d['titulo']} ({d['ambito']})" for d in docs)
                    if docs else "La base normativa está vacía: el despacho aún no ha cargado documentos.")
        return _cortar([{k: x[k] for k in ("cita", "documento", "referencia", "seccion", "ambito", "version", "texto")}
                        for x in r])
    err = _valido(esp, a.get("periodo", ""), a.get("cliente", ""))
    if err:
        return err
    sal = esp.salida(a["periodo"], a["cliente"])
    if nombre == "ver_hallazgos":
        hs = [dict(h, origen="excel") for h in leer_json(sal / "hallazgos.json", [])] + \
             [dict(h, origen="reauditoria") for h in leer_json(sal / "reauditoria.json", [])]
        if a.get("severidad"):
            hs = [h for h in hs if h.get("severidad") == a["severidad"]]
        return _cortar(hs or "Sin hallazgos (o el cliente aún no se ha validado en ese periodo).")
    if nombre == "ver_cifras":
        m = leer_json(sal / "modelo.json")
        if not m:
            return "El cliente no tiene modelo en ese periodo: primero hay que validarlo."
        if not a.get("seccion") or a["seccion"] not in m:
            return _cortar({"secciones": list(m), "meta": m.get("meta")})
        return _cortar(m[a["seccion"]])
    if nombre == "ver_informe":
        e = esp.estado(a["periodo"], a["cliente"])
        arch = e.get("archivos") or {}
        diapositivas = []
        if arch.get("pptx") and (sal / arch["pptx"]).exists():
            from pptx import Presentation
            for i, d in enumerate(Presentation(sal / arch["pptx"]).slides, start=1):
                nombres = {sh.name for sh in d.shapes}
                titulo = next((sh.text_frame.text for sh in d.shapes if sh.name == "Titulo" and sh.has_text_frame), "")
                diapositivas.append({"n": i, "titulo": titulo or "(portada o separador)",
                                     "pendiente": "Pendiente" in nombres, "con_propuesta": "EtiquetaPropuesta" in nombres})
        cob = leer_json(sal / "cobertura_esquema.json", [])
        ad = leer_json(sal / "adaptacion.json", {}) or {}
        r = e.get("resumen") or {}
        # lo más importante primero y compacto (los modelos locales tienen poco contexto)
        return _cortar({"estado": e.get("estado"), "diapositivas_total": len(diapositivas),
                        "secciones_completas": [c["descripcion"] for c in cob if c.get("estado") == "ok"],
                        "secciones_pendientes": [{"seccion": c["descripcion"], "estado": c.get("estado"), "motivo": c.get("detalle")}
                                                 for c in cob if c.get("estado") != "ok"],
                        "propuestas_sin_aprobar": e.get("propuestas_sin_aprobar"),
                        "hallazgos": r.get("hallazgos"), "reauditoria": r.get("reauditoria"), "cifras": r.get("cifras"),
                        "segundos_de_procesamiento": r.get("segundos"),
                        "diapositivas_pendientes": [f"{d['n']}. {d['titulo']}" for d in diapositivas if d["pendiente"]],
                        "diapositivas_con_propuesta": [f"{d['n']}. {d['titulo']}" for d in diapositivas if d["con_propuesta"]],
                        "adaptacion": {"resumen": ad.get("resumen"), "ajustes_mapa": ad.get("ajustes_mapa")},
                        "archivos": arch,
                        "todas_las_diapositivas": [f"{d['n']}. {d['titulo']}" for d in diapositivas]})
    if nombre == "ver_bitacora":
        e = esp.estado(a["periodo"], a["cliente"])
        return _cortar({k: e.get(k) for k in ("estado", "bitacora", "aprobacion", "archivos", "error")})
    if nombre == "buscar_en_excel":
        excel = (esp.estado(a["periodo"], a["cliente"]).get("archivos") or {}).get("excel")
        if not excel:
            cands = sorted(esp.entrada(a["periodo"], a["cliente"]).glob("*.xls*"))
            excel = str(cands[0]) if cands else None
        if not excel or not Path(excel).exists():
            return "No hay Excel cargado para ese cliente y periodo."
        texto = a["texto"].strip().lower()
        if len(texto) < 3:
            return "Use al menos 3 caracteres para buscar."
        import openpyxl
        wb = openpyxl.load_workbook(excel, data_only=True, read_only=True)
        hits = []
        for ws in wb.worksheets:
            for fila in ws.iter_rows():
                for celda in fila:
                    v = celda.value
                    if v is not None and texto in str(v).lower():
                        vecinos = [str(x.value) for x in fila if x.value is not None][:12]
                        hits.append({"hoja": ws.title, "celda": celda.coordinate, "valor": str(v)[:400],
                                     "fila": vecinos})
                        if len(hits) >= 60:
                            break
                if len(hits) >= 60:
                    break
        wb.close()
        return _cortar(hits or f"«{a['texto']}» no aparece en el Excel.")
    return f"Herramienta desconocida: {nombre}"


def _validar_args(nombre: str, args) -> str | None:
    """Con eager_input_streaming el servidor no valida la entrada: se valida aquí contra el esquema."""
    esquema = HERRAMIENTAS[nombre][2]
    if not isinstance(args, dict):
        return "entrada no es un objeto"
    for campo in esquema["required"]:
        if not isinstance(args.get(campo), str):
            return f"falta el campo «{campo}»"
    enum = {k: v["enum"] for k, v in esquema["properties"].items() if "enum" in v}
    for k, opciones in enum.items():
        if args.get(k) not in opciones:
            return f"valor inválido para «{k}»"
    return None


# ----------------------------------------------------------------------------- conversación

def conversar(esp: Espacio, mensajes: list[dict], base: str, periodo: str | None, cliente: str | None):
    """Genera eventos para la interfaz: texto, herramienta, fin o error. Usa el proveedor de IA configurado."""
    from . import proveedores
    if not disponible(esp.raiz)["ok"]:
        yield {"tipo": "error", "mensaje": disponible(esp.raiz)["ayuda"]}
        return
    herramientas = {n: (d, s) for n, (bases, d, s) in HERRAMIENTAS.items() if base == "todo" or base in bases}
    nombres = []
    for cid in esp.ids_clientes():
        try:
            nombres.append(f"{esp.ficha(cid).get('copropiedad') or cid} = {cid}")
        except Exception:
            nombres.append(cid)
    contexto = (f"Base seleccionada por el usuario: {BASES[base]}. Clientes del espacio (nombre = identificador): "
                f"{'; '.join(nombres) or 'ninguno'}."
                + (f" Periodo en pantalla: {periodo}." if periodo else "")
                + (f" Cliente en pantalla: {cliente}." if cliente else ""))
    local = disponible(esp.raiz).get("local")
    if local:                                             # modelo en este Mac: sin internet, contexto más corto
        _limite.set(LIMITE_LOCAL)
        mensajes = mensajes[-6:]
    historial = list(mensajes)
    historial[-1] = {"role": "user", "content": f"{historial[-1]['content']}\n\n[{contexto}]"}
    try:
        yield from proveedores.conversar(esp.raiz, "asistente", SISTEMA, historial, herramientas,
                                         lambda n, a: ejecutar(esp, n, a), _validar_args)
    except proveedores.SinHerramientas:
        yield from _modo_recuperacion(esp, mensajes[-1]["content"], periodo, cliente, contexto)


def _consultas_para(esp: Espacio, pregunta: str, periodo: str | None, cliente: str | None) -> list[tuple[str, dict]]:
    """Modo de recuperación (modelo sin herramientas): Python elige y corre las lecturas que pide la pregunta."""
    from revisoria.excel import norm
    q = norm(pregunta)
    clientes = [cid for cid in esp.ids_clientes()
                if cid in q or (norm(esp.ficha(cid).get("copropiedad", "")) or "\0") in q] or ([cliente] if cliente else [])
    per = periodo or ""
    out = []
    if any(w in q for w in ("reunion", "acuerd", "definio", "pendiente por definir")):
        out.append(("ver_acuerdos", {}))
    if any(w in q for w in ("norma", "ley", "articulo", "decreto", "niif")):
        out.append(("buscar_normativa", {"consulta": pregunta[:200], "cliente": clientes[0] if clientes else ""}))
    for cid in clientes[:2]:
        if any(w in q for w in ("hallazgo", "alta", "celda", "riesgo")):
            out.append(("ver_hallazgos", {"periodo": per, "cliente": cid, "severidad": "ALTA" if "alta" in q else ""}))
        if any(w in q for w in ("informe", "diapositiva", "pendiente", "salio", "propuesta")):
            out.append(("ver_informe", {"periodo": per, "cliente": cid}))
        if any(w in q for w in ("ficha", "documento", "logo", "hoja", "mapa", "quien es", "informacion", "datos")):
            out.append(("ver_cliente", {"cliente": cid}))
        if any(w in q for w in ("cartera", "liquidez", "cifra", "valor", "cuanto", "total", "presupuesto")):
            out.append(("ver_cifras", {"periodo": per, "cliente": cid, "seccion": ""}))
    if not out:
        out = [("listar_clientes", {"periodo": per})] + [("ver_informe", {"periodo": per, "cliente": c}) for c in clientes[:1]]
    return out


def _modo_recuperacion(esp: Espacio, pregunta: str, periodo, cliente, contexto: str):
    from . import proveedores
    datos = []
    for nombre, args in _consultas_para(esp, pregunta, periodo, cliente):
        yield {"tipo": "herramienta", "nombre": nombre, "args": args}
        datos.append(f"### {nombre} {json.dumps(args, ensure_ascii=False)}\n{ejecutar(esp, nombre, args)}")
    usuario = (f"Pregunta: {pregunta}\n\n[{contexto}]\n\nDatos leídos por Python del espacio de trabajo "
               "(responde SOLO con base en ellos y cita hoja y celda o archivo):\n\n" + "\n\n".join(datos))
    try:
        yield {"tipo": "texto", "t": proveedores.texto(esp.raiz, "asistente", SISTEMA, usuario)}
        yield {"tipo": "fin"}
    except proveedores.ErrorIA as e:
        yield {"tipo": "error", "mensaje": str(e)}
