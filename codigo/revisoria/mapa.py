"""Mapa del cliente: dónde está, en SU Excel, cada hoja del esquema estándar (revisoria/esquema.py).

Es lo único que cambia entre clientes. Se guarda en `clientes/<id>/mapa.json`:

    {"version": 1, "origen": "deteccion", "aprobado_por": null, "fecha": "...",
     "hojas": {"CUADRO EFECTIVO": "CUADRO EFECTIVO", "MATRIZ DICTAMEN": null,
               "AUDITORIA DOCUMENTAL": [["REV DOC CONTABLE", "Contable"], ["REV DOC ADMON", "Administrativo"]], ...}}

Un título `null` significa «este cliente no tiene esa hoja» (no es un error). Si el mapa apunta a una hoja que ya
no existe, el Libro lo anota en `desajustes` para avisar que el mapa quedó desactualizado.
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from .documental import detectar_hoja_polizas, detectar_hojas_auditoria
from .esquema import FIRMAS, HOJAS_POR_NOMBRE, revisar_formato
from .excel import Libro, buscar_hoja, norm
from .extractor import detectar_hoja_seguimiento

ARCHIVO = "mapa.json"


def _por_aprendizaje(lib: Libro, clave: str, aprendidos: dict | None) -> str | None:
    """Título de hoja que en otros clientes ya se aprobó para `clave` y que además cumple la firma de formato."""
    for t in lib.valores.sheetnames:
        a = (aprendidos or {}).get(norm(t))
        if a and a["clave"] == clave and not revisar_formato(lib, {clave: t}):
            return t
    return None


def _cumple_firma(lib: Libro, clave: str, titulo: str) -> bool:
    return clave in FIRMAS and not revisar_formato(lib, {clave: titulo})


def _abre_con_firma(lib: Libro, clave: str, titulo: str) -> bool:
    """El primer texto de la hoja (su título, p. ej. «CUENTAS POR PAGAR A AGOSTO 2026») empieza con la firma.
    Distingue la hoja propia de otra que solo menciona la sección en un renglón (una matriz, un resumen)."""
    from .esquema import _texto_cabecera
    cab = _texto_cabecera(lib.valores[titulo])
    return any(cab.startswith(v) for v in FIRMAS[clave][0].split("|"))


def _por_contenido(lib: Libro, hojas: dict, titulos: list[str]) -> dict:
    """Para cada clave sin hoja (o con una hoja que no tiene su firma), busca entre las hojas libres la ÚNICA cuyo
    contenido sí tiene la firma de la sección (p. ej. «PROVEEDORES» con «cuentas por pagar» en sus encabezados).
    Si hay varias candidatas no adivina: la clave queda sin hoja para que el revisor la asigne."""
    usadas = {t for v in hojas.values() if v for t in ([x[0] for x in v] if isinstance(v, list) else [v])}
    cambios = {}
    for clave in HOJAS_POR_NOMBRE:
        actual = hojas.get(clave)
        if clave not in FIRMAS or (actual and _cumple_firma(lib, clave, actual)):
            continue
        libres = [t for t in titulos if t not in usadas or t == actual]
        cand = [t for t in libres if _cumple_firma(lib, clave, t)]
        if len(cand) > 1:                                # desempate: la hoja que ABRE con la firma (su título)
            cand = [t for t in cand if _abre_con_firma(lib, clave, t)] or cand
        if len(cand) == 1:
            cambios[clave] = cand[0]
            usadas.add(cand[0])
        elif actual:                                     # el nombre calzaba pero el contenido no
            cambios[clave] = None
            usadas.discard(actual)
    return cambios


def detectar(lib: Libro, aprendidos: dict | None = None) -> dict:
    """Propone el mapa: primero lo aprendido de mapas aprobados (sinónimos), luego las reglas tolerantes por nombre
    y, para lo que no se encontró por nombre, el contenido de las hojas (firma de la sección)."""
    titulos = [ws.title for ws in lib.valores.worksheets]
    hojas = {n: _por_aprendizaje(lib, n, aprendidos) or buscar_hoja(titulos, n) for n in HOJAS_POR_NOMBRE}
    hojas["SEGUIMIENTO DEUDORES"] = (_por_aprendizaje(lib, "SEGUIMIENTO DEUDORES", aprendidos)
                                     or detectar_hoja_seguimiento(lib))
    hojas["AUDITORIA DOCUMENTAL"] = detectar_hojas_auditoria(lib) or None
    hojas["VERIFICACION POLIZAS"] = (_por_aprendizaje(lib, "VERIFICACION POLIZAS", aprendidos)
                                     or detectar_hoja_polizas(lib))
    hojas.update(_por_contenido(lib, hojas, titulos))     # al final: las hojas ya asignadas no son candidatas
    usadas = {t for v in hojas.values() if v for t in ([x[0] for x in v] if isinstance(v, list) else [v])}
    return {"version": 1, "origen": "deteccion", "aprobado_por": None,
            "fecha": datetime.now().isoformat(timespec="seconds"), "excel": Path(lib.ruta).name,
            "hojas": dict(sorted(hojas.items())),
            "sin_usar": [t for t in titulos if t not in usadas]}   # hojas del cliente que el informe no lee


def ruta(dir_cliente: Path) -> Path:
    return Path(dir_cliente) / ARCHIVO


def cargar(dir_cliente: Path) -> dict | None:
    r = ruta(dir_cliente)
    return json.loads(r.read_text(encoding="utf-8")) if r.exists() else None


def guardar(dir_cliente: Path, mapa: dict) -> Path:
    r = ruta(dir_cliente)
    r.write_text(json.dumps(mapa, ensure_ascii=False, indent=2), encoding="utf-8")
    return r


def _titulos(v) -> list[str]:
    return [x[0] for x in v] if isinstance(v, list) else [v] if v else []


# ---------------------------------------------------------------- mapa vs. Excel del mes

def reparar(lib: Libro, mapa: dict, aprendidos: dict | None = None) -> dict:
    """Claves cuyo título ya no está en el Excel (renombraron o quitaron la hoja). Para cada una busca un reemplazo
    que cumpla la firma de formato. Devuelve {clave: {"antes", "ahora" (o None), "motivo"}}. No toca el mapa."""
    titulos = set(lib.valores.sheetnames)
    rotas = [k for k, v in mapa.get("hojas", {}).items() if any(t not in titulos for t in _titulos(v))]
    if not rotas:
        return {}
    nuevo = detectar(lib, aprendidos)["hojas"]
    usadas = {t for k, v in mapa["hojas"].items() if k not in rotas for t in _titulos(v)}
    salida = {}
    for k in rotas:
        cand = nuevo.get(k)
        ok = cand and not (set(_titulos(cand)) & usadas) and not revisar_formato(lib, {k: cand})
        salida[k] = {"antes": mapa["hojas"][k], "ahora": cand if ok else None,
                     "motivo": ("la hoja cambió de nombre; se encontró una equivalente con el mismo formato" if ok
                                else "no se encontró una hoja equivalente con el formato esperado")}
    return salida


def aplicar(mapa: dict, ajustes: dict) -> dict:
    """Copia del mapa con los ajustes encontrados (provisional: lo confirma el revisor en la pantalla del mapa)."""
    nuevo = json.loads(json.dumps(mapa))
    for k, a in ajustes.items():
        if a["ahora"]:
            nuevo["hojas"][k] = a["ahora"]
    return nuevo


def hojas_nuevas(lib: Libro, mapa: dict) -> list[str]:
    """Hojas que el Excel trae este mes y que no existían cuando se aprobó el mapa (ni usadas ni conocidas)."""
    conocidas = set(mapa.get("sin_usar", [])) | {t for v in mapa.get("hojas", {}).values() for t in _titulos(v)}
    return [t for t in lib.valores.sheetnames if t not in conocidas] if "sin_usar" in mapa else []


# ---------------------------------------------------------------- aprendizaje entre clientes

SINONIMOS = "sinonimos_hojas.json"


def aprendidos(conocimiento: Path) -> dict:
    r = Path(conocimiento) / SINONIMOS
    return json.loads(r.read_text(encoding="utf-8")) if r.exists() else {}


def aprender(conocimiento: Path, mapa: dict, cid: str) -> list[str]:
    """Cada mapa aprobado enseña sus títulos de hoja: {título normalizado: {clave, clientes}}. Devuelve los nuevos.
    Si un mismo título se aprobó para claves distintas en clientes distintos, queda marcado como ambiguo y no se usa."""
    datos = aprendidos(conocimiento)
    nuevos = []
    for clave, v in mapa.get("hojas", {}).items():
        for t in _titulos(v):
            n = norm(t)
            if n == norm(clave):
                continue                                 # el nombre estándar no enseña nada
            a = datos.get(n)
            if a is None:
                datos[n] = {"clave": clave, "titulo": t, "clientes": [cid]}
                nuevos.append(f"«{t}» → {clave}")
            elif a["clave"] == clave:
                if cid not in a["clientes"]:
                    a["clientes"].append(cid)
            elif a["clave"] != "AMBIGUO":
                a.update({"clave": "AMBIGUO", "opciones": sorted({a["clave"], clave})})
    r = Path(conocimiento) / SINONIMOS
    r.parent.mkdir(parents=True, exist_ok=True)
    r.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")
    return nuevos
