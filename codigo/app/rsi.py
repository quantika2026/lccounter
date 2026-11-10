"""Automejora recursiva (RSI) acotada y medible.

Nivel 1 — mejora los artefactos de texto (plantillas de opinión de la biblioteca):
  1. Casos de evaluación: cada vez que el revisor fiscal editó una opinión, su texto final (con las cifras convertidas
     en marcadores) es el objetivo. Se separan casos de entrenamiento y un caso de prueba reservado (el más reciente).
  2. La IA genera candidatos siguiendo la ESTRATEGIA vigente.
  3. Python mide cada candidato (similitud con los textos del revisor) y verifica restricciones: solo marcadores
     permitidos, sin cifras escritas a mano, formateable y prueba de regresión contra el golden set.
  4. Si el mejor supera a la plantilla actual en el caso reservado por el margen configurado: modo «proponer» deja
     una propuesta para el revisor; modo «automático» la publica como versión reversible.

Nivel 2 — recursivo, mejora la propia forma de mejorar:
  5. Con el historial de candidatos ganadores y perdedores, la IA reescribe la ESTRATEGIA (las instrucciones que usa
     para generar candidatos). La nueva estrategia compite contra la vigente en los ciclos siguientes y la reemplaza
     solo si gana la mayoría de las comparaciones. Cada estrategia queda versionada.

Límites fijos (no configurables): la mejora aplica al SISTEMA (plantillas y su estrategia; los mapeos y reglas los
mejora `agente/mejora.py` con aprobación del revisor) y NUNCA a informes ya validados: no lee ni escribe salidas
aprobadas, que quedan congeladas. Nunca modifica cálculos, extracción de cifras, reglas de validación ni código; no
inventa hechos; todo cambio queda versionado y se puede revertir.

Horario: la app corre el ciclo una vez al día a la hora configurada (por defecto 4:00 a. m.), si está encendida.
"""
from __future__ import annotations

import json
import re
import time
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path

from agente.comun import Espacio, ahora, guardar_json, leer_json

MODOS = {"desactivado": "Desactivado", "proponer": "Proponer (el revisor aprueba)",
         "automatico": "Automático (publica lo que mejora en la prueba reservada)"}
CONFIG_BASE = {"modo": "proponer", "margen": 0.03, "candidatos": 3, "max_plantillas": 5, "hora": "04:00"}
RX_HORA = re.compile(r"^([01]\d|2[0-3]):[0-5]\d$")
RX_MARCADOR = re.compile(r"\{([a-z_]+)\}")

ESTRATEGIA_INICIAL = """Reescribe la plantilla para que se parezca lo más posible a la forma en que el revisor fiscal \
redacta esta opinión en los ejemplos: su estructura, su vocabulario, el tono (opinión expresa de la Revisoría Fiscal, \
solicitudes concretas al Consejo o a la administración) y la extensión. Conserva el sentido técnico. Generaliza lo que \
se repite en los ejemplos y omite lo que solo aparece en uno. Usa únicamente los marcadores permitidos para las cifras."""

SISTEMA = """Eres el componente de automejora de un sistema de Revisoría Fiscal para propiedades horizontales en \
Colombia. Generas plantillas de texto (opiniones estándar del revisor fiscal) que luego se llenan con cifras del Excel.
Reglas inviolables:
- Las cifras van SOLO como marcadores entre llaves de la lista permitida, por ejemplo {cartera}. Nunca escribas \
números, fechas, porcentajes ni nombres propios.
- No inventes hechos, normas ni plazos. No agregues opiniones que no estén en los ejemplos del revisor.
- Español formal de Colombia."""

ESQ_CANDIDATOS = {"type": "object", "properties": {"candidatos": {"type": "array", "items": {
    "type": "object", "properties": {"texto": {"type": "string"}, "razon": {"type": "string"}},
    "required": ["texto", "razon"], "additionalProperties": False}}}, "required": ["candidatos"], "additionalProperties": False}
ESQ_ESTRATEGIA = {"type": "object", "properties": {"estrategia": {"type": "string"}, "aprendizajes": {"type": "string"}},
                  "required": ["estrategia", "aprendizajes"], "additionalProperties": False}


# ------------------------------------------------------------------ estado

def _dir(esp: Espacio) -> Path:
    return esp.conocimiento / "rsi"


def config(esp: Espacio) -> dict:
    return {**CONFIG_BASE, **leer_json(_dir(esp) / "config.json", {})}


def guardar_config(esp: Espacio, modo: str, margen: float, candidatos: int, hora: str | None = None):
    if modo not in MODOS:
        raise ValueError("Modo inválido.")
    if not 0.005 <= margen <= 0.5 or not 1 <= candidatos <= 6:
        raise ValueError("Margen (0,5 % a 50 %) o número de candidatos (1 a 6) fuera de rango.")
    hora = hora or config(esp)["hora"]
    if not RX_HORA.match(hora):
        raise ValueError("Hora inválida (use HH:MM, por ejemplo 04:00).")
    guardar_json(_dir(esp) / "config.json", {"modo": modo, "margen": margen, "candidatos": candidatos, "hora": hora,
                                              "max_plantillas": CONFIG_BASE["max_plantillas"]})


def toca_programado(esp: Espacio, ahora_dt) -> bool:
    """¿Corresponde el ciclo diario? Una vez por día, desde la hora configurada (si la app estaba apagada a esa hora,
    corre al encender ese mismo día). Desactivado no corre."""
    cfg = config(esp)
    if cfg["modo"] == "desactivado":
        return False
    hoy = ahora_dt.strftime("%Y-%m-%d")
    return ahora_dt.strftime("%H:%M") >= cfg["hora"] and estado(esp).get("ultimo_programado") != hoy


def marcar_programado(esp: Espacio, ahora_dt):
    e = estado(esp)
    e["ultimo_programado"] = ahora_dt.strftime("%Y-%m-%d")
    _guardar_estado(esp, e)


def estado(esp: Espacio) -> dict:
    e = leer_json(_dir(esp) / "estado.json", {})
    if not e:
        e = {"vigente": "e-001", "retadora": None, "comparaciones": [],
             "estrategias": {"e-001": {"texto": ESTRATEGIA_INICIAL, "creada": ahora(), "origen": "inicial",
                                       "ciclos": 0, "ganancia_total": 0.0, "victorias": 0}}}
    return e


def _guardar_estado(esp: Espacio, e: dict):
    guardar_json(_dir(esp) / "estado.json", e)


def propuestas(esp: Espacio) -> list[dict]:
    return leer_json(_dir(esp) / "propuestas.json", [])


def ciclos(esp: Espacio, n: int = 30) -> list[dict]:
    r = _dir(esp) / "ciclos.jsonl"
    if not r.exists():
        return []
    return [json.loads(x) for x in r.read_text(encoding="utf-8").splitlines() if x.strip()][-n:][::-1]


# ------------------------------------------------------------------ casos y medición

def casos(esp: Espacio) -> dict[str, list[dict]]:
    """Textos finales del revisor (en marcadores) por plantilla, en orden cronológico."""
    from agente.mejora import correcciones
    out = defaultdict(list)
    for c in correcciones(esp):
        # lo que el revisor aceptó tal como lo redactó Claude no es su redacción: no se aprende de ello
        if (c.get("tipo") == "opinion" and c.get("accion") == "editada" and c.get("plantilla")
                and c.get("origen") != "claude"):
            out[c["clave"]].append({"fecha": c["fecha"], "cliente": c["cliente"], "periodo": c["periodo"],
                                    "objetivo": c["plantilla"].replace("{{", "{").replace("}}", "}")})
    return dict(out)


def _palabras(t: str) -> list[str]:
    return re.findall(r"\{[a-z_]+\}|\w+", t.lower())


def puntaje(texto: str, objetivos: list[str]) -> float:
    if not objetivos:
        return 0.0
    a = _palabras(texto)
    return round(sum(SequenceMatcher(None, a, _palabras(o), autojunk=False).ratio() for o in objetivos) / len(objetivos), 4)


def permitidos(clave: str, objetivos: list[str]) -> set[str]:
    from agente.redaccion import BIBLIOTECA_BASE
    base = BIBLIOTECA_BASE.get(clave, {}).get("texto", "")
    return set(RX_MARCADOR.findall(base)) | {m for o in objetivos for m in RX_MARCADOR.findall(o)}


def problemas(texto: str, clave: str, objetivos: list[str]) -> list[str]:
    out = []
    usados = set(RX_MARCADOR.findall(texto))
    otros = usados - permitidos(clave, objetivos)
    if otros:
        out.append(f"marcadores no permitidos: {', '.join(sorted(otros))}")
    sin_marcadores = RX_MARCADOR.sub("", texto)
    numeros_obj = {n for o in objetivos for n in re.findall(r"\d+", RX_MARCADOR.sub("", o))}
    sueltos = [n for n in re.findall(r"\d+", sin_marcadores) if n not in numeros_obj]
    if sueltos:
        out.append(f"cifras escritas a mano: {', '.join(sueltos[:5])}")
    try:
        texto.format(**{m: "X" for m in usados})
    except (KeyError, IndexError, ValueError) as e:
        out.append(f"no se puede formatear: {e}")
    if not 40 <= len(texto) <= 2000:
        out.append("longitud fuera de rango")
    return out


# ------------------------------------------------------------------ generación con IA

def _generar(esp: Espacio, estrategia: str, clave: str, titulo: str, actual: str, entrenamiento: list[str],
             n: int) -> list[dict]:
    from .proveedores import json_estructurado
    pedido = (f"ESTRATEGIA DE MEJORA:\n{estrategia}\n\nPlantilla «{clave}» ({titulo}). Texto actual:\n{actual}\n\n"
              f"Marcadores permitidos: {', '.join('{' + m + '}' for m in sorted(permitidos(clave, entrenamiento)))}\n\n"
              "Textos finales del revisor fiscal para esta opinión (las cifras ya están como marcadores):\n"
              + "\n\n".join(f"Ejemplo {i}:\n{o}" for i, o in enumerate(entrenamiento, 1))
              + f"\n\nPropón {n} candidatos distintos de plantilla, cada uno con una razón breve.")
    datos, _ = json_estructurado(esp.raiz, "rsi", SISTEMA, pedido, ESQ_CANDIDATOS, effort="medium")
    return [c for c in datos["candidatos"] if isinstance(c.get("texto"), str)][:n]


def _nueva_estrategia(esp: Espacio, e: dict, historial: list[dict]) -> dict:
    from .proveedores import json_estructurado
    vig = e["estrategias"][e["vigente"]]["texto"]
    resumen = "\n".join(
        f"- {h['clave']}: ganador «{(h.get('ganador') or {}).get('razon', '—')}» (+{h.get('ganancia', 0):.3f}); "
        f"perdedores: {'; '.join(x['razon'] for x in h.get('perdedores', [])[:3]) or '—'}" for h in historial[:12])
    pedido = (f"Estrategia vigente que usas para generar candidatos:\n{vig}\n\nResultados medidos de los últimos ciclos "
              f"(ganancia = mejora de similitud con los textos reales del revisor en un caso reservado):\n{resumen}\n\n"
              "Escribe una estrategia mejorada (máximo 1.200 caracteres) que aproveche lo que funcionó y evite lo que no. "
              "Debe seguir respetando las reglas inviolables. Explica en «aprendizajes» qué cambiaste y por qué.")
    datos, _ = json_estructurado(esp.raiz, "rsi", SISTEMA, pedido, ESQ_ESTRATEGIA, effort="high")
    return {"texto": datos["estrategia"].strip()[:1500], "aprendizajes": datos["aprendizajes"].strip()[:800]}


# ------------------------------------------------------------------ ciclo

def ciclo(esp: Espacio, autor: str) -> dict:
    """Un ciclo completo de automejora. Devuelve el resumen del ciclo."""
    from agente.mejora import publicar_plantilla, regresion
    from agente.redaccion import biblioteca
    cfg = config(esp)
    if cfg["modo"] == "desactivado":
        raise ValueError("La automejora está desactivada en Configuración.")
    C = casos(esp)
    if not C:
        raise ValueError("Aún no hay casos para aprender: la automejora necesita opiniones editadas por el revisor "
                         "al aprobar informes (pestaña «Borrador de opinión»).")
    e = estado(esp)
    bib = biblioteca(esp)
    estrategias = [e["vigente"]] + ([e["retadora"]] if e.get("retadora") else [])
    registro = {"fecha": ahora(), "autor": autor, "modo": cfg["modo"], "estrategias": estrategias, "plantillas": []}
    victorias = defaultdict(int)
    for clave, lista in list(C.items())[: cfg["max_plantillas"]]:
        if clave not in bib:
            continue
        objetivos = [c["objetivo"] for c in lista]
        independiente = len(objetivos) >= 2
        entren, prueba = (objetivos[:-1], objetivos[-1:]) if independiente else (objetivos, objetivos)
        actual = bib[clave]["texto"]
        base_prueba = puntaje(actual, prueba)
        evaluados = []
        for eid in estrategias:
            for c in _generar(esp, e["estrategias"][eid]["texto"], clave, bib[clave].get("titulo", clave), actual, entren,
                              cfg["candidatos"]):
                prob = problemas(c["texto"], clave, objetivos)
                evaluados.append({"estrategia": eid, "texto": c["texto"].strip(), "razon": c.get("razon", "")[:300],
                                  "problemas": prob, "entrenamiento": puntaje(c["texto"], entren),
                                  "prueba": puntaje(c["texto"], prueba)})
        validos = sorted((x for x in evaluados if not x["problemas"]), key=lambda x: (-x["entrenamiento"], -x["prueba"]))
        # el ganador se elige por entrenamiento y se JUZGA en el caso reservado (evita sobreajuste)
        mejor = validos[0] if validos else None
        ganancia = round(mejor["prueba"] - base_prueba, 4) if mejor else 0.0
        por_estrategia = {eid: max((x["prueba"] for x in validos if x["estrategia"] == eid), default=0) for eid in estrategias}
        if len(estrategias) == 2 and por_estrategia[estrategias[0]] != por_estrategia[estrategias[1]]:
            victorias[max(por_estrategia, key=por_estrategia.get)] += 1      # un empate no cuenta para ninguna
        for eid in estrategias:
            s = e["estrategias"][eid]
            s["ciclos"] += 1
            s["ganancia_total"] = round(s["ganancia_total"] + max(0.0, por_estrategia[eid] - base_prueba), 4)
        resultado = {"clave": clave, "casos": len(objetivos), "prueba_independiente": independiente,
                     "puntaje_actual": base_prueba, "ganador": mejor, "ganancia": ganancia,
                     "perdedores": [x for x in evaluados if x is not mejor][:6], "accion": "sin cambios"}
        if mejor and ganancia >= cfg["margen"]:
            candidata = dict(bib)
            candidata[clave] = {**bib[clave], "texto": mejor["texto"]}
            reg = regresion(esp, candidata)
            if not reg["ok"] and reg["n"]:
                resultado["accion"] = "descartado: falló la prueba de regresión"
            elif cfg["modo"] == "automatico" and independiente and reg["n"]:
                nv = publicar_plantilla(esp, clave, candidata[clave], "RSI (automático)", f"rsi-{clave}",
                                        f"{reg['n']} caso(s) OK", {"rsi": {"ganancia": ganancia, "estrategia": mejor["estrategia"]}})
                resultado["accion"] = f"publicado automáticamente en {nv}"
            else:
                motivo = ("modo proponer" if cfg["modo"] == "proponer" else
                          "sin caso de prueba independiente" if not independiente else "golden set vacío")
                pid = f"RSI-{clave}-{time.strftime('%Y%m%d%H%M%S')}"
                ps = [p for p in propuestas(esp) if not (p["clave"] == clave and p["estado"] == "pendiente")]
                ps.append({"id": pid, "clave": clave, "texto": mejor["texto"], "actual": actual, "razon": mejor["razon"],
                           "ganancia": ganancia, "puntaje_actual": base_prueba, "puntaje_nuevo": mejor["prueba"],
                           "independiente": independiente, "estrategia": mejor["estrategia"], "estado": "pendiente",
                           "fecha": ahora(), "motivo": motivo})
                guardar_json(_dir(esp) / "propuestas.json", ps)
                resultado["accion"] = f"propuesta {pid} ({motivo})"
        registro["plantillas"].append(resultado)

    # ---- nivel 2: la estrategia compite y se reescribe a sí misma
    meta = {}
    if e.get("retadora") and len(estrategias) == 2:
        e["comparaciones"].append({"fecha": ahora(), "vigente": victorias[e["vigente"]], "retadora": victorias[e["retadora"]]})
        recientes = e["comparaciones"][-3:]
        if len(recientes) >= 2:
            gana = sum(1 for c in recientes if c["retadora"] > c["vigente"])
            pierde = sum(1 for c in recientes if c["retadora"] < c["vigente"])
            if gana >= 2:
                meta = {"accion": f"la estrategia {e['retadora']} reemplaza a {e['vigente']} (ganó {gana} de {len(recientes)})"}
                e["estrategias"][e["retadora"]]["victorias"] += 1
                e["vigente"], e["retadora"], e["comparaciones"] = e["retadora"], None, []
            elif pierde >= 2:
                meta = {"accion": f"la estrategia {e['retadora']} se descarta (perdió {pierde} de {len(recientes)})"}
                e["estrategias"][e["retadora"]]["descartada"] = ahora()
                e["retadora"], e["comparaciones"] = None, []
    if not e.get("retadora") and registro["plantillas"]:
        nueva = _nueva_estrategia(esp, e, registro["plantillas"] + [p for c in ciclos(esp, 5) for p in c.get("plantillas", [])])
        nid = f"e-{len(e['estrategias']) + 1:03d}"
        e["estrategias"][nid] = {**nueva, "creada": ahora(), "origen": f"reescrita a partir de {e['vigente']}",
                                 "ciclos": 0, "ganancia_total": 0.0, "victorias": 0}
        e["retadora"] = nid
        meta.setdefault("accion", "")
        meta["accion"] = (meta["accion"] + "; " if meta["accion"] else "") + f"nueva estrategia retadora {nid}"
        meta["aprendizajes"] = nueva["aprendizajes"]
    registro["meta"] = meta
    _guardar_estado(esp, e)
    _dir(esp).mkdir(parents=True, exist_ok=True)
    with (_dir(esp) / "ciclos.jsonl").open("a", encoding="utf-8") as f:
        f.write(json.dumps(registro, ensure_ascii=False) + "\n")
    return registro


def aprobar(esp: Espacio, pid: str, autor: str) -> str:
    from agente.mejora import publicar_plantilla, regresion
    from agente.redaccion import biblioteca
    ps = propuestas(esp)
    p = next((x for x in ps if x["id"] == pid and x["estado"] == "pendiente"), None)
    if not p:
        raise ValueError("No existe esa propuesta pendiente.")
    bib = biblioteca(esp)
    candidata = dict(bib)
    candidata[p["clave"]] = {**bib.get(p["clave"], {}), "texto": p["texto"]}
    reg = regresion(esp, candidata)
    if not reg["ok"] and reg["n"]:
        raise ValueError("La prueba de regresión falló con esta plantilla; no se publica.")
    nv = publicar_plantilla(esp, p["clave"], candidata[p["clave"]], autor, pid,
                            f"{reg['n']} caso(s) OK" if reg["n"] else "golden set vacío", {"rsi": {"ganancia": p["ganancia"]}})
    p.update({"estado": f"aplicada en {nv}", "aprobado_por": autor})
    guardar_json(_dir(esp) / "propuestas.json", ps)
    return nv


def rechazar(esp: Espacio, pid: str, autor: str):
    ps = propuestas(esp)
    p = next((x for x in ps if x["id"] == pid and x["estado"] == "pendiente"), None)
    if not p:
        raise ValueError("No existe esa propuesta pendiente.")
    p.update({"estado": "rechazada", "rechazado_por": autor, "fecha_rechazo": ahora()})
    guardar_json(_dir(esp) / "propuestas.json", ps)


def restaurar_estrategia(esp: Espacio, eid: str):
    e = estado(esp)
    if eid not in e["estrategias"]:
        raise ValueError("No existe esa estrategia.")
    e["vigente"], e["retadora"], e["comparaciones"] = eid, None, []
    _guardar_estado(esp, e)
