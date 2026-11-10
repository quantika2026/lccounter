"""Pantalla «Mapa del Excel»: ver, proponer (con IA), probar y aprobar el mapa de un cliente.

El mapa solo dice QUÉ HOJA del Excel del cliente alimenta cada sección del esquema estándar; el informe no cambia.
La IA propone, nunca guarda: el mapa se guarda cuando un revisor lo aprueba. «Probar» corre la extracción real con
el mapa candidato y muestra qué secciones quedan completas, así nadie aprueba un mapa que el lector no entiende.
"""
from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path

from revisoria import esquema, lectores, mapa as mapa_cliente
from revisoria.documental import extraer_documental
from revisoria.excel import Libro, Periodo, norm
from revisoria.extractor import extraer

from . import proveedores

LISTAS = {"AUDITORIA DOCUMENTAL"}                       # claves que admiten varias hojas
ESPECIALES = {"SEGUIMIENTO DEUDORES", "AUDITORIA DOCUMENTAL", "VERIFICACION POLIZAS"}
CLAVES = sorted(set(esquema.HOJAS_POR_NOMBRE) | ESPECIALES)
AYUDA = {"MAYORES DEUDORES": "Evolución de mayores deudores (respaldo si no hay hoja de seguimiento)",
         "SEGUIMIENTO DEUDORES": "Seguimiento de cobro de mayores deudores (abogado, gestión)"}


def _secciones_de(clave: str) -> list[dict]:
    return [{"clave": s.clave, "descripcion": s.descripcion, "obligatoria": s.obligatoria}
            for s in esquema.SECCIONES if s.hoja == clave]


def _vistazo(ws, filas: int = 10, cols: int = 10) -> list[str]:
    """Primeras filas con texto de la hoja, para que la persona (y la IA) reconozcan qué contiene."""
    out = []
    for fila in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 40), max_col=cols, values_only=True):
        t = " | ".join(str(v).strip()[:40] for v in fila if v not in (None, ""))
        if t:
            out.append(t[:220])
        if len(out) >= filas:
            break
    return out


def _bloque(titulo: str, ws) -> str | None:
    t = norm(" ".join(str(ws.cell(r, 2).value or "") for r in range(1, 4)) + " " + titulo)
    return "Contable" if "contab" in t else "Administrativo" if "admin" in t else None


def normalizar(lib: Libro, hojas: dict) -> dict:
    """Valida un mapa enviado por la pantalla y lo deja en el formato del archivo. Lanza ValueError si no sirve."""
    titulos = lib.valores.sheetnames
    salida = {}
    for clave in CLAVES:
        v = hojas.get(clave)
        if clave in LISTAS:
            v = [x for x in (v or []) if x]
            if any(t not in titulos for t in v):
                raise ValueError(f"«{clave}»: alguna hoja elegida no existe en el Excel.")
            salida[clave] = [[t, _bloque(t, lib.valores[t])] for t in v] or None
        else:
            if v and v not in titulos:
                raise ValueError(f"«{clave}»: la hoja «{v}» no existe en el Excel.")
            salida[clave] = v or None
    desconocidas = set(hojas) - set(CLAVES)
    if desconocidas:
        raise ValueError("Claves desconocidas: " + ", ".join(sorted(desconocidas)))
    return salida


def probar(excel: Path, periodo: str, hojas: dict, lectores: dict | None = None) -> dict:
    """Extrae con el mapa candidato (y las recetas de lectura aprobadas) y devuelve el estado de cada sección."""
    p = Periodo.desde_texto(periodo)
    try:
        modelo, lib = extraer(str(excel), p, {}, {"hojas": hojas, "lectores": lectores or {}})
        doc = extraer_documental(lib, p)
    except Exception as e:                           # el mapa no alcanza para leer lo obligatorio
        return {"ok": False, "error": f"{type(e).__name__}: {e}", "cobertura": []}
    return {"ok": True, "error": "", "cobertura": esquema.verificar(modelo, doc), "desajustes": lib.desajustes,
            "formato": esquema.revisar_formato(lib, hojas, doc, lectores)}


def vista(esp, periodo: str, cid: str, excel: Path) -> dict:
    lib = Libro(str(excel))
    aprendidos = mapa_cliente.aprendidos(esp.conocimiento)
    guardado = mapa_cliente.cargar(esp.clientes / cid)
    mapa = guardado or mapa_cliente.detectar(lib, aprendidos)
    ajustes = mapa_cliente.reparar(lib, mapa, aprendidos) if guardado else {}
    if ajustes:                                      # la pantalla arranca con el ajuste propuesto, para confirmar
        mapa = mapa_cliente.aplicar(mapa, ajustes)
    nuevas = mapa_cliente.hojas_nuevas(lib, mapa) if guardado else []
    hojas_actuales = {k: ([x[0] for x in v] if isinstance(v, list) else v) for k, v in mapa["hojas"].items()}
    usadas = {t for v in hojas_actuales.values() if v for t in (v if isinstance(v, list) else [v])}
    return {
        "excel": excel.name, "guardado": bool(guardado), "origen": mapa.get("origen"),
        "aprobado_por": mapa.get("aprobado_por"), "fecha": mapa.get("fecha"),
        "ajustes": ajustes, "hojas_nuevas": nuevas,
        "claves": [{"clave": k, "lista": k in LISTAS, "ayuda": AYUDA.get(k, ""), "secciones": _secciones_de(k),
                    "hoja": hojas_actuales.get(k)} for k in CLAVES],
        "hojas_excel": [{"titulo": ws.title, "usada": ws.title in usadas, "vistazo": _vistazo(ws)}
                        for ws in lib.valores.worksheets],
        "prueba": probar(excel, periodo, mapa["hojas"], mapa.get("lectores")),
        "lectores": mapa.get("lectores") or {},
        "campos_receta": {sec: {"descripcion": d["descripcion"], "matriz": bool(d.get("matriz")), "por_mes": d["por_mes"],
                                "campos": [{"campo": c, "obligatorio": ob, "que_es": q} for c, (_, ob, q) in d["campos"].items()],
                                "parametros": [{"nombre": k, "obligatorio": ob, "que_es": q}
                                               for k, (ob, q) in (d.get("parametros") or {}).items()]}
                          for sec, d in lectores.CAMPOS.items()},
    }


# ----------------------------------------------------------------------------- propuesta

ESQUEMA_PROPUESTA = {
    "type": "object", "additionalProperties": False, "required": ["propuestas"],
    "properties": {"propuestas": {"type": "array", "items": {
        "type": "object", "additionalProperties": False, "required": ["clave", "hoja", "confianza", "motivo"],
        "properties": {"clave": {"type": "string"}, "hoja": {"type": "string"},
                       "confianza": {"type": "string", "enum": ["alta", "media", "baja"]},
                       "motivo": {"type": "string"}}}}},
}

SISTEMA = """Eres el asistente de mapeo de LCCOUNTER (revisoría fiscal de propiedad horizontal en Colombia).
Cada copropiedad entrega un Excel de papeles de trabajo con nombres de hoja distintos. Tu tarea: proponer qué hoja
del Excel del cliente corresponde a cada clave del esquema estándar que hoy está SIN HOJA o con la sección vacía.
Reglas:
- Usa SOLO los títulos de hoja que aparecen en la lista; cópialos exactamente, con espacios.
- Juzga por el contenido (vistazo de las primeras filas), no solo por el nombre.
- Si ninguna hoja corresponde, no propongas nada para esa clave. Nunca inventes una hoja.
- «aprobado_en_otros_clientes» son títulos que un revisor ya aprobó en otras copropiedades: úsalos como pista,
  pero confirma con el vistazo.
- «motivo»: una frase concreta que cite lo que viste en la hoja (encabezados, títulos).
- confianza «alta» solo si el contenido coincide claramente con la sección."""


def _heuristica(pendientes: list[dict], hojas: list[dict]) -> list[dict]:
    """Sin IA: compara palabras del nombre de la clave/sección con el título y el vistazo de cada hoja libre."""
    def palabras(t):
        return {re.sub(r"(es|s)$", "", w) for w in norm(t).split() if len(w) > 3 and not w.isdigit()}
    libres = [h for h in hojas if not h["usada"]]
    out = []
    for c in pendientes:
        objetivo = palabras(c["clave"] + " " + " ".join(s["descripcion"] for s in c["secciones"]))
        mejor = max(((len(objetivo & palabras(h["titulo"])) * 2 + len(objetivo & palabras(" ".join(h["vistazo"][:4]))), h)
                     for h in libres), default=(0, None), key=lambda x: x[0])
        if mejor[1] and mejor[0] >= 2:
            out.append({"clave": c["clave"], "hoja": mejor[1]["titulo"], "confianza": "baja",
                        "motivo": "Coincidencia de palabras en el nombre o encabezados (sin IA)."})
    return out


def _con_formato(excel: Path, propuestas: list[dict]) -> list[dict]:
    """Marca cada propuesta con la revisión de firma de su hoja (encabezados esperados)."""
    if not propuestas:
        return propuestas
    lib = Libro(str(excel))
    for p in propuestas:
        prob = esquema.revisar_formato(lib, {p["clave"]: p["hoja"]})
        p["formato_ok"], p["formato"] = not prob, prob[0]["detalle"] if prob else ""
    return propuestas


def proponer(raiz: Path, v: dict, excel: Path, aprendidos: dict | None = None) -> dict:
    estados = {c["seccion"]: c["estado"] for c in v["prueba"].get("cobertura", [])}
    pendientes = [c for c in v["claves"]
                  if not c["hoja"] or any(estados.get(s["clave"]) in ("falta", "vacia") for s in c["secciones"])]
    if not pendientes:
        return {"propuestas": [], "modelo": None, "nota": "Todas las secciones tienen hoja y datos."}
    titulos = {h["titulo"] for h in v["hojas_excel"]}
    disp = proveedores.disponible(raiz, "asistente")
    nota, modelo = "", None
    if disp["ok"]:
        usuario = json.dumps({
            "claves_pendientes": [{"clave": c["clave"], "hoja_actual": c["hoja"], "ayuda": c["ayuda"],
                                   "secciones": [s["descripcion"] for s in c["secciones"]]} for c in pendientes],
            "hojas_del_excel": [{"titulo": h["titulo"], "ya_usada": h["usada"], "vistazo": h["vistazo"]}
                                for h in v["hojas_excel"]],
            "aprobado_en_otros_clientes": [{"titulo": a["titulo"], "clave": a["clave"]}
                                           for a in (aprendidos or {}).values() if a["clave"] != "AMBIGUO"][:80]},
            ensure_ascii=False)
        try:
            datos, modelo = proveedores.json_estructurado(raiz, "asistente", SISTEMA, usuario, ESQUEMA_PROPUESTA, "medium")
            claves_ok = {c["clave"] for c in pendientes}
            actual = {c["clave"]: c["hoja"] for c in pendientes}
            propuestas = [p for p in datos["propuestas"] if p["clave"] in claves_ok and p["hoja"] in titulos
                          and p["hoja"] != actual[p["clave"]] and p["hoja"] not in (actual[p["clave"]] or [])]
            if not propuestas:
                nota = "La IA no encontró hojas mejores; las secciones vacías parecen no tener datos este mes."
            return {"propuestas": _con_formato(excel, propuestas), "modelo": modelo, "nota": nota}
        except proveedores.ErrorIA as e:
            nota = f"La IA no respondió ({e}); se usa la comparación por palabras."
    else:
        nota = disp["ayuda"] + " Se usa la comparación por palabras."
    return {"propuestas": _con_formato(excel, _heuristica(pendientes, v["hojas_excel"])), "modelo": None, "nota": nota}


def guardar(esp, cid: str, hojas: dict, revisor: str, excel: Path) -> dict:
    dir_cli = esp.clientes / cid
    anterior = mapa_cliente.ruta(dir_cli)
    if anterior.exists():
        hist = dir_cli / "historial"
        hist.mkdir(exist_ok=True)
        shutil.copy2(anterior, hist / f"mapa_{datetime.now():%Y%m%d_%H%M%S}.json")
    lib = Libro(str(excel))
    previo = mapa_cliente.cargar(dir_cli) or {}
    mapa = {"version": 1, "origen": "revisado", "aprobado_por": revisor,
            "fecha": datetime.now().isoformat(timespec="seconds"), "excel": excel.name,
            "hojas": normalizar(lib, hojas)}
    usadas = {t for v in mapa["hojas"].values() if v for t in ([x[0] for x in v] if isinstance(v, list) else [v])}
    mapa["sin_usar"] = [t for t in lib.valores.sheetnames if t not in usadas]
    if previo.get("lectores"):                      # las recetas aprobadas se conservan al aprobar el mapa
        mapa["lectores"] = previo["lectores"]
    mapa_cliente.guardar(dir_cli, mapa)
    mapa["aprendido"] = mapa_cliente.aprender(esp.conocimiento, mapa, cid)
    return mapa
