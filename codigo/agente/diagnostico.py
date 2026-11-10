"""Diagnóstico de recepción (E1, T03–T06) de un cliente y periodo. Solo lee: no escribe mapa, ficha ni salidas.

Lo usa el subagente recepcion-ph a través del comando `diagnosticar`, para no depender de funciones internas.
"""
from __future__ import annotations

from revisoria import esquema, mapa as mapa_cliente
from revisoria.documental import extraer_documental
from revisoria.excel import Libro, Periodo
from revisoria.extractor import extraer

from . import integridad, normalizador
from .comun import Espacio


def diagnosticar(esp: Espacio, cid: str, periodo: str) -> dict:
    p = Periodo.desde_texto(periodo)
    ficha = esp.ficha(cid)
    ent = esp.entrada(periodo, cid)
    out = {"cliente": cid, "periodo": periodo,
           "archivos": sorted(x.name for x in ent.iterdir()) if ent.is_dir() else [], "bloqueos": []}
    try:
        excel = normalizador.ubicar_excel(ent, ficha)
    except normalizador.ExcelAmbiguo as e:
        out["bloqueos"].append(str(e))
        return out
    if not excel:
        out["bloqueos"].append(f"No llegó el Excel en entradas/{periodo}/{cid}/")
        return out
    m = mapa_cliente.cargar(esp.clientes / cid) or {}
    lib = Libro(str(excel), m)
    out.update({"excel": excel.name, "excel_del_mapa": m.get("excel"), "recetas": sorted(m.get("lectores") or {}),
                "hojas": dict(m.get("hojas") or {}), "hojas_sin_usar": m.get("sin_usar", [])})
    try:
        meta = {k: ficha.get(k) for k in ("copropiedad", "nit", "unidades") if ficha.get(k)}
        mod, _ = extraer(str(excel), p, meta, m, lib)
    except Exception as e:                                   # p. ej. el Excel es de otro mes
        out["bloqueos"].append(f"La extracción falló: {e}")
        return out
    doc = extraer_documental(lib, p)
    identidad = integridad.verificar_identidad(lib, ficha, mod, periodo)
    out.update({"identidad": identidad or "OK",
                "cobertura": esquema.verificar(mod, doc),
                "formato": esquema.revisar_formato(lib, m.get("hojas", {}), doc, m.get("lectores")),
                "desajustes": lib.desajustes})
    out["bloqueos"] += identidad
    return out
