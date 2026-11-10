"""Comunicaciones de la Revisoría Fiscal (al Consejo, a la administración o respuesta a una comunicación recibida).

Principio (mismo del redactor institucional del despacho): la documentación aportada es la ÚNICA fuente.
- Las fuentes se numeran (F1, F2…): ficha, hallazgos elegidos, cifras del informe, borrador de opinión y, si aplica,
  la comunicación recibida. Claude solo puede afirmar lo que está en ellas y cita cada fundamento.
- Nunca inventa normas, artículos, fechas, radicados ni plazos: deja marcadores [entre corchetes].
- Control de cifras: toda cifra del oficio que no aparezca en las fuentes se señala como «sin respaldo».
- El resultado es siempre un BORRADOR. Lo revisa y firma el revisor fiscal; la aplicación nunca lo envía.
"""
from __future__ import annotations

import json
import re
import time
from datetime import date
from pathlib import Path

from agente.comun import Espacio, guardar_json, leer_json


TIPOS = {
    "consejo": "Comunicación al Consejo de Administración sobre los hallazgos del periodo",
    "administracion": "Requerimiento a la administración para subsanar hallazgos y entregar soportes",
    "respuesta": "Respuesta a una comunicación recibida (Consejo, administración o copropietario)",
}
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
         "noviembre", "diciembre"]
RX_NUM = re.compile(r"\d[\d\.,]*\d|\d")

SISTEMA = """Eres especialista en redacción institucional de LC Counter Contadores y Auditores SAS, firma que ejerce \
la Revisoría Fiscal de propiedades horizontales en Colombia. Redactas borradores de comunicaciones oficiales que \
firmará el revisor fiscal.

Principio central: las FUENTES entregadas (F1, F2, …) son la única fuente de verdad. Un oficio compromete a quien \
lo firma; por eso:
- No inventes normas, artículos, numerales, fechas, radicados, plazos, nombres ni cifras. Solo puedes citar una \
norma, artículo o numeral si su texto está en una fuente «Norma: …», y debes citarla con documento y artículo tal \
como aparecen, por ejemplo (F8, Ley 675 de 2001, artículo 35). Usa una norma solo si su texto respalda de verdad lo \
que afirmas; si no hay norma pertinente en las fuentes, no cites ninguna. Si un dato hace falta, deja un marcador entre corchetes, por ejemplo [plazo], [radicado], \
[nombre del presidente del Consejo], y anótalo en «observaciones».
- Cada cifra debe copiarse tal como aparece en las fuentes (mismo formato). No hagas cálculos nuevos.
- Cita únicamente los identificadores de fuente que aparecen en FUENTES (F1, F2, …). Nunca uses un número de fuente \
que no exista.
- Si las fuentes no permiten responder un punto, escribe: «La documentación disponible no contiene información \
suficiente para responder este punto.»
- Si hay contradicción entre fuentes, menciónala sin resolverla.
- Las opiniones y recomendaciones solo pueden ser las que traen las fuentes (hallazgos, borrador de opinión, notas \
del revisor). No agregues opiniones nuevas.

Estilo: formal, técnico, objetivo y respetuoso; tercera persona institucional («esta Revisoría Fiscal»). Sin \
adjetivos innecesarios ni especulación. Escribe en español de Colombia.

Estructura del oficio (en Markdown, omite secciones que no apliquen): ciudad y fecha; destinatario, cargo y \
copropiedad; «Asunto:»; «Referencia:» si responde algo; saludo; 1. Antecedentes; 2. Hallazgos o consideraciones; \
3. Fundamentos (con citas); 4. Solicitudes o respuesta punto por punto; 5. Conclusiones; despedida; firma (nombre, \
«Revisor(a) Fiscal», firma de la sociedad). Cita las fuentes en el texto así: (F3, hoja PATRIMONIO celda C16)."""

ESQUEMA = {
    "type": "object",
    "properties": {
        "resumen": {"type": "string", "description": "Objetivo, fuentes usadas y fundamentos principales, 3 a 5 líneas."},
        "asunto": {"type": "string"},
        "oficio": {"type": "string", "description": "Oficio completo en Markdown, listo para revisión y firma."},
        "trazabilidad": {"type": "array", "items": {
            "type": "object",
            "properties": {"elemento": {"type": "string"}, "fuente": {"type": "string"}, "referencia": {"type": "string"}},
            "required": ["elemento", "fuente", "referencia"], "additionalProperties": False}},
        "observaciones": {"type": "array", "items": {"type": "string"},
                          "description": "Información faltante, marcadores por completar, riesgos y recomendaciones."},
    },
    "required": ["resumen", "asunto", "oficio", "trazabilidad", "observaciones"],
    "additionalProperties": False,
}


def _fecha_larga(d: date) -> str:
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def carpeta(esp: Espacio, periodo: str, cid: str) -> Path:
    return esp.salida(periodo, cid) / "comunicaciones"


def fuentes(esp: Espacio, periodo: str, cid: str, hallazgos_ids: list[str], incluir_borrador: bool,
            recibida: str, incluir_normativa: bool = False) -> list[dict]:
    """Arma las fuentes numeradas. Cada una: {id, titulo, texto}."""
    sal = esp.salida(periodo, cid)
    ficha = esp.ficha(cid)
    est = esp.estado(periodo, cid)
    out = []

    def agregar(titulo, texto):
        out.append({"id": f"F{len(out) + 1}", "titulo": titulo, "texto": texto.strip()})

    datos = {k: ficha.get(k) for k in ("copropiedad", "nit", "unidades", "numero_informe", "revisor", "firma", "ciudad")
             if ficha.get(k)}
    notas = {k: v for k, v in (ficha.get("notas") or {}).items() if v}
    agregar("Ficha del cliente y notas del revisor fiscal",
            json.dumps({"datos": datos, "notas_del_revisor": notas}, ensure_ascii=False, indent=1))
    cif = (est.get("resumen") or {}).get("cifras") or {}
    if cif:
        from revisoria.narrativa import pesos
        etiquetas = {"liquidez": "Total equivalente al efectivo", "cartera": "Cartera total",
                     "cxp": "Cuentas por pagar", "resultado": "Resultado del ejercicio"}
        agregar(f"Cifras principales del informe {ficha.get('numero_informe', '')} ({periodo})",
                "\n".join(f"- {etiquetas.get(k, k)}: {pesos(v)}" for k, v in cif.items() if v is not None))
    todos = {f"{h['regla']}#{i}": h for i, h in enumerate(leer_json(sal / "hallazgos.json", []) + leer_json(sal / "reauditoria.json", []))}
    elegidos = [todos[h] for h in hallazgos_ids if h in todos]
    for h in elegidos:
        if h:
            agregar(f"Hallazgo {h['regla']} ({h['severidad']}): {h['titulo']}",
                    f"Categoría: {h.get('categoria', '')}\nDetalle: {h.get('detalle', '')}\n"
                    f"Ubicación: {', '.join(h.get('ubicacion') or []) or 'no indicada'}\n"
                    f"Recomendación: {h.get('recomendacion') or 'no indicada'}")
    if incluir_borrador:
        editado = sal / f"Borrador_Editado_{cid}_{periodo}.md"
        gen = sal / f"Borrador_Opinion_{cid}_{periodo}.md"
        ruta = editado if editado.exists() else gen
        if ruta.exists():
            agregar("Borrador de opinión" + (" (editado por el revisor)" if ruta == editado else " (generado, sin revisar)"),
                    ruta.read_text(encoding="utf-8"))
    if recibida.strip():
        agregar("Comunicación recibida que se responde", recibida.strip()[:30_000])
    if incluir_normativa:
        # Fundamentos SOLO de la base normativa cargada: hasta 2 por hallazgo (y por la comunicación recibida), máximo 8
        from .normativa import instancia
        base = instancia(esp.raiz)
        consultas = [f"{h['titulo']} {h.get('categoria', '')} {h.get('recomendacion') or ''}" for h in elegidos]
        if recibida.strip():
            consultas.append(recibida[:600])
        vistos = set()
        for q in consultas:
            for r in base.buscar(q, cid, k=2):
                clave = (r["doc_id"], r["fragmento"])
                if clave in vistos or len(vistos) >= 8 or r["puntaje"] < 4:
                    continue
                vistos.add(clave)
                agregar(f"Norma: {r['documento']}, {r['referencia']}" + (f" (versión {r['version']})" if r["version"] else ""),
                        r["texto"][:4000])
    return out


def opciones_hallazgos(esp: Espacio, periodo: str, cid: str) -> list[dict]:
    sal = esp.salida(periodo, cid)
    hs = leer_json(sal / "hallazgos.json", []) + leer_json(sal / "reauditoria.json", [])
    return [{"id": f"{h['regla']}#{i}", "regla": h["regla"], "severidad": h["severidad"], "titulo": h["titulo"]}
            for i, h in enumerate(hs) if h["regla"] not in ("R06", "G01")]


def cifras_sin_respaldo(oficio: str, fuentes_: list[dict], extra: str = "") -> list[str]:
    """Cifras de 3 o más dígitos del oficio que no aparecen en ninguna fuente (ni en la fecha del oficio)."""
    base = " ".join(f["texto"] for f in fuentes_) + " " + extra
    respaldadas = {re.sub(r"[\.,]", "", n) for n in RX_NUM.findall(base)}
    vistas, sin = set(), []
    for n in RX_NUM.findall(oficio):
        if re.fullmatch(r"\d{1,2}(\.\d{1,2}){1,3}\.?", n):      # numeración de secciones (2.5.1), no una cifra
            continue
        limpio = re.sub(r"[\.,]", "", n)
        if len(limpio) >= 3 and limpio not in respaldadas and limpio not in vistas:
            vistas.add(limpio)
            sin.append(n)
    return sin


def citas_invalidas(oficio: str, fuentes_: list[dict]) -> list[str]:
    """Citas a fuentes que no se entregaron (p. ej. «F13» cuando solo hay F1 a F7): señal de una referencia inventada."""
    validas = {f["id"] for f in fuentes_}
    malas = set()
    for m in re.finditer(r"\bF\d{1,3}\b", oficio):
        antes, despues = oficio[max(0, m.start() - 12):m.start()], oficio[m.end():m.end() + 1]
        # una referencia de celda del Excel (F4:F10, PRESUPUESTO!F13, «celda F13», SUM(F4…) no es una cita de fuente
        if re.search(r"[!:$]\s*$|SUM\(\s*$|celdas?\s+$|rango\s+$", antes, re.I) or despues == ":":
            continue
        if m.group(0) not in validas:
            malas.add(m.group(0))
    return sorted(malas, key=lambda x: int(x[1:]))


def generar(esp: Espacio, periodo: str, cid: str, tipo: str, hallazgos_ids: list[str], incluir_borrador: bool,
            recibida: str, instrucciones: str, autor: str, incluir_normativa: bool = True) -> dict:
    from .proveedores import disponible

    if tipo not in TIPOS:
        raise ValueError("Tipo de comunicación inválido.")
    if tipo == "respuesta" and not recibida.strip():
        raise ValueError("Pegue el texto de la comunicación recibida que se va a responder.")
    if not disponible(esp.raiz, "comunicaciones")["ok"]:
        raise RuntimeError(disponible(esp.raiz, "comunicaciones")["ayuda"])
    F = fuentes(esp, periodo, cid, hallazgos_ids, incluir_borrador, recibida, incluir_normativa)
    ficha = esp.ficha(cid)
    hoy = _fecha_larga(date.today())
    pedido = (f"Tipo de comunicación: {TIPOS[tipo]}.\n"
              f"Copropiedad: {ficha.get('copropiedad', cid)}. Periodo del informe: {periodo}.\n"
              f"Ciudad: {ficha.get('ciudad') or '[ciudad]'}. Fecha del oficio: {hoy}.\n"
              f"Firma: {ficha.get('revisor') or '[nombre del revisor fiscal]'}, Revisor(a) Fiscal, "
              f"{ficha.get('firma') or 'LC Counter Contadores y Auditores SAS'}.\n"
              + (f"Instrucciones del revisor (orientan el enfoque; no son fuente de hechos): {instrucciones.strip()}\n"
                 if instrucciones.strip() else "")
              + "\nFUENTES:\n\n" + "\n\n".join(f"[{f['id']}] {f['titulo']}\n{f['texto']}" for f in F))
    from .proveedores import ErrorIA, json_estructurado
    try:
        datos, modelo_resp = json_estructurado(esp.raiz, "comunicaciones", SISTEMA, pedido, ESQUEMA, effort="high")
    except ErrorIA as e:
        raise RuntimeError(str(e)) from None
    cid_doc = time.strftime("%Y%m%d-%H%M%S") + "-" + tipo
    doc = {"id": cid_doc, "tipo": tipo, "tipo_nombre": TIPOS[tipo], "periodo": periodo, "cliente": cid,
           "creado": time.strftime("%Y-%m-%dT%H:%M:%S"), "creado_por": autor, "estado": "borrador",
           "fuentes": [{"id": f["id"], "titulo": f["titulo"]} for f in F], "fuentes_texto": F,
           "instrucciones": instrucciones.strip(), "modelo": modelo_resp, **datos,
           "cifras_sin_respaldo": cifras_sin_respaldo(datos["oficio"], F, hoy),
           "citas_invalidas": citas_invalidas(datos["oficio"], F), "historial": []}
    guardar_json(carpeta(esp, periodo, cid) / f"{cid_doc}.json", doc)
    return doc


def listar(esp: Espacio, periodo: str, cid: str) -> list[dict]:
    d = carpeta(esp, periodo, cid)
    docs = [leer_json(p, {}) for p in sorted(d.glob("*.json"), reverse=True)] if d.exists() else []
    return [{k: x.get(k) for k in ("id", "tipo_nombre", "asunto", "creado", "creado_por", "estado", "revisado_por")}
            for x in docs if x]


def leer(esp: Espacio, periodo: str, cid: str, doc_id: str) -> dict:
    if not re.match(r"^\d{8}-\d{6}-[a-z]+$", doc_id):
        raise ValueError("Identificador de comunicación inválido.")
    d = leer_json(carpeta(esp, periodo, cid) / f"{doc_id}.json")
    if not d:
        raise FileNotFoundError("No existe esa comunicación.")
    return d


def guardar(esp: Espacio, periodo: str, cid: str, doc_id: str, oficio: str, asunto: str, autor: str,
            revisado: bool) -> dict:
    d = leer(esp, periodo, cid, doc_id)
    d["historial"].append({"fecha": time.strftime("%Y-%m-%dT%H:%M:%S"), "por": autor,
                           "accion": "revisado" if revisado else "editado"})
    d.update({"oficio": oficio, "asunto": asunto,
              "cifras_sin_respaldo": cifras_sin_respaldo(oficio, d.get("fuentes_texto", []), d.get("creado", "")),
              "citas_invalidas": citas_invalidas(oficio, d.get("fuentes_texto", []))})
    if revisado:
        d.update({"estado": "revisado", "revisado_por": autor, "revisado": time.strftime("%Y-%m-%dT%H:%M:%S")})
    elif d.get("estado") == "revisado":                 # una edición posterior exige nueva revisión
        d.update({"estado": "borrador", "revisado_por": None})
    guardar_json(carpeta(esp, periodo, cid) / f"{doc_id}.json", d)
    return d


def eliminar(esp: Espacio, periodo: str, cid: str, doc_id: str):
    leer(esp, periodo, cid, doc_id)
    (carpeta(esp, periodo, cid) / f"{doc_id}.json").unlink()
    (carpeta(esp, periodo, cid) / f"{doc_id}.docx").unlink(missing_ok=True)


# ------------------------------------------------------------------ Word

def _linea(par, texto: str):
    """Agrega texto con **negrilla** a un párrafo de python-docx."""
    for i, trozo in enumerate(re.split(r"\*\*(.+?)\*\*", texto)):
        if trozo:
            par.add_run(trozo).bold = bool(i % 2)


def a_word(esp: Espacio, periodo: str, cid: str, doc_id: str, logo: Path) -> Path:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm, Pt, RGBColor
    d = leer(esp, periodo, cid, doc_id)
    w = Document()
    sec = w.sections[0]
    sec.left_margin = sec.right_margin = Cm(2.5)
    estilo = w.styles["Normal"]
    estilo.font.name = "Arial"
    estilo.font.size = Pt(11)
    if logo.exists():
        sec.header.paragraphs[0].add_run().add_picture(str(logo), width=Cm(2.6))
    if d.get("estado") != "revisado":
        p = sec.header.add_paragraph()
        r = p.add_run("BORRADOR — pendiente de revisión y firma del revisor fiscal")
        r.font.size, r.font.color.rgb, r.bold = Pt(9), RGBColor(0xB4, 0x23, 0x18), True
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for linea in d["oficio"].splitlines():
        s = linea.strip()
        if not s:
            continue
        titulo = s.startswith("#") or re.match(r"^\d+\.\s+\S[^.]{0,60}$", s)    # «## …» o «1. Antecedentes»
        if titulo:
            h = w.add_paragraph()
            _linea(h, s.lstrip("#").strip().replace("**", ""))
            for r in h.runs:
                r.bold = True
            h.paragraph_format.space_before = Pt(8)
            continue
        m = re.match(r"^([-*]|\d+\.)\s+(.*)", s)
        if m:
            p = w.add_paragraph(style="List Bullet" if m.group(1) in "-*" else "List Number")
            _linea(p, m.group(2))
            continue
        p = w.add_paragraph()
        _linea(p, s)
        p.paragraph_format.space_after = Pt(6)
    if d.get("trazabilidad"):
        w.add_page_break()
        w.add_paragraph().add_run("Anexo interno: trazabilidad de fuentes (no enviar)").bold = True
        t = w.add_table(rows=1, cols=3)
        t.style = "Table Grid"
        for c, txt in zip(t.rows[0].cells, ("Elemento", "Fuente", "Referencia")):
            c.text = txt
        nombres = {f["id"]: f["titulo"] for f in d.get("fuentes", [])}
        for x in d["trazabilidad"]:
            fila = t.add_row().cells
            fila[0].text, fila[1].text, fila[2].text = x["elemento"], f"{x['fuente']} {nombres.get(x['fuente'], '')}", x["referencia"]
    ruta = carpeta(esp, periodo, cid) / f"{doc_id}.docx"
    w.save(str(ruta))
    return ruta
