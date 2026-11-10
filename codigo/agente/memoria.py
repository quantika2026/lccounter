"""Memoria persistente del flujo con Claude: lo que una sesión aprende y la siguiente necesita saber.

Complementa, sin reemplazar, la memoria del sistema (conocimiento/: correcciones, biblioteca versionada, golden set,
sinónimos de hojas, recetas aprobadas, RSI). Aquí van las notas operativas que no son cifras ni decisiones:
  - por cliente: particularidades de su Excel, pendientes pedidos al cliente (T06), acuerdos con el revisor;
  - del flujo: lecciones (un error y cómo se resolvió) que sirven para todos los clientes.

Se guarda en <espacio>/memoria/<cliente>.jsonl y <espacio>/memoria/flujo.jsonl (una nota por línea, nunca se
borra: un pendiente se cierra con «resolver»). `contexto` junta las notas abiertas con el estado del sistema para
que cada sesión arranque sabiendo dónde quedó la anterior.
"""
from __future__ import annotations

import json
from pathlib import Path

from .comun import Espacio, ahora, leer_json

TIPOS = ("pendiente", "particularidad", "acuerdo", "leccion")
FLUJO = "flujo"


def _ruta(esp: Espacio, cliente: str) -> Path:
    return esp.raiz / "memoria" / f"{cliente}.jsonl"


def notas(esp: Espacio, cliente: str) -> list[dict]:
    r = _ruta(esp, cliente)
    if not r.exists():
        return []
    out = [json.loads(x) for x in r.read_text(encoding="utf-8").splitlines() if x.strip()]
    cerradas = {n["resuelve"] for n in out if n.get("resuelve")}
    return [n | {"abierta": n["id"] not in cerradas} for n in out if not n.get("resuelve")]


def _agregar(esp: Espacio, cliente: str, nota: dict) -> dict:
    r = _ruta(esp, cliente)
    r.parent.mkdir(parents=True, exist_ok=True)
    with r.open("a", encoding="utf-8") as f:
        f.write(json.dumps(nota, ensure_ascii=False) + "\n")
    return nota


def anotar(esp: Espacio, cliente: str, tipo: str, texto: str, periodo: str = "", autor: str = "claude") -> dict:
    if tipo not in TIPOS:
        raise ValueError(f"Tipo inválido: use {', '.join(TIPOS)}.")
    if cliente != FLUJO and not (esp.clientes / cliente / "ficha.yaml").exists():
        raise ValueError(f"No existe el cliente {cliente}.")
    texto = " ".join(texto.split())
    if not texto:
        raise ValueError("La nota está vacía.")
    previas = notas(esp, cliente)
    if any(n["abierta"] and n["tipo"] == tipo and n["texto"] == texto for n in previas):
        return next(n for n in previas if n["texto"] == texto)          # no se duplica
    nid = f"{'F' if cliente == FLUJO else 'N'}{len(previas) + 1:03d}"
    return _agregar(esp, cliente, {"id": nid, "fecha": ahora(), "tipo": tipo, "periodo": periodo,
                                   "texto": texto, "autor": autor})


def resolver(esp: Espacio, cliente: str, nid: str, como: str, autor: str = "claude") -> dict:
    n = next((x for x in notas(esp, cliente) if x["id"] == nid), None)
    if not n or not n["abierta"]:
        raise ValueError(f"No hay una nota abierta {nid} para {cliente}.")
    return _agregar(esp, cliente, {"resuelve": nid, "fecha": ahora(), "texto": " ".join(como.split()), "autor": autor})


def contexto(esp: Espacio, cliente: str, periodo: str) -> dict:
    """Todo lo que una sesión nueva necesita para retomar un cliente en un periodo."""
    anterior = esp.periodo_anterior(periodo)
    est, est_ant = esp.estado(periodo, cliente), esp.estado(anterior, cliente)
    mapa = leer_json(esp.clientes / cliente / "mapa.json", None)
    abiertos = leer_json(esp.salida(anterior, cliente) / "hallazgos_abiertos.json", [])
    return {
        "cliente": cliente, "periodo": periodo,
        "estado_actual": est.get("estado"), "ultimo_evento": (est.get("bitacora") or [{}])[-1].get("detalle"),
        "periodo_anterior": {"periodo": anterior, "estado": est_ant.get("estado"),
                             "aprobado_por": (est_ant.get("aprobacion") or {}).get("revisor")},
        "mapa": None if mapa is None else {"origen": mapa.get("origen"), "aprobado_por": mapa.get("aprobado_por"),
                                           "recetas": sorted(mapa.get("lectores") or {})},
        "hallazgos_abiertos_mes_anterior": [f"{h.get('regla')} {h.get('titulo')}" for h in abiertos],
        "notas_abiertas": [n for n in notas(esp, cliente) if n["abierta"]],
        "lecciones_del_flujo": [n for n in notas(esp, FLUJO) if n["abierta"]][-15:],
    }
