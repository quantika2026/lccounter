"""Agente normalizador: lleva lo que entrega cada cliente al Excel estándar y al balance de prueba.

Cada cliente usa un software contable distinto; los informes que exporta son diferentes, pero el Excel de
papeles de trabajo es estándar. El normalizador:

1. Ubica el Excel estándar del periodo y verifica que tenga las hojas que usa el generador.
2. Lee el balance de prueba exportado del software contable con el *perfil* de ese software
   (nombres de columnas, separador, formato numérico). Los perfiles viven en
   conocimiento/perfiles_software.yaml y los aprueba el profesional.
3. Si el formato es desconocido, PROPONE un mapeo (heurístico; opcionalmente con IA) y lo deja en
   conocimiento/propuestas/ para aprobación. Nunca usa un mapeo no aprobado.
4. Cuadra las cifras del Excel contra el balance (regla V19).
"""
from __future__ import annotations

import csv
import io
import re
import unicodedata
from pathlib import Path

from revisoria.validador import Hallazgo

from .comun import Espacio, ahora, guardar_json, leer_yaml

HOJAS_REQUERIDAS = ["CUADRO EFECTIVO", "GRAFICA DEUDORES", "Rangos de Cartera", "MAYORES DEUDORES",
                    "ANTIC. E INGRESOS", "CXP", "CUADRO RESERVA", "PATRIMONIO", "PRESUPUESTO",
                    "CONCILIACIONES", "IMPUESTOS"]

PERFILES_BASE = {
    "generico": {
        "descripcion": "Balance de prueba en CSV o Excel con columnas cuenta / nombre / saldo final.",
        "columnas": {"cuenta": ["cuenta", "codigo", "código", "codigo cuenta"],
                     "nombre": ["nombre", "nombre cuenta", "descripcion", "descripción"],
                     "saldo": ["saldo final", "saldo", "nuevo saldo", "saldo actual"]},
        "separador": "auto", "decimal": "auto", "estado": "aprobado",
    },
}

# Qué cifra del modelo se cuadra contra qué cuentas del PUC. Editable en conocimiento/perfiles_software.yaml
CUADRES_BASE = [
    {"concepto": "Anticipos y avances (1330)", "prefijos": ["1330"], "modelo": "anticipos.total_hoja"},
    {"concepto": "Deudores varios (1380)", "prefijos": ["1380"], "modelo": "deudores_varios.total_hoja"},
    {"concepto": "Cuentas por pagar (23)", "prefijos": ["23"], "modelo": "cxp.total"},
    {"concepto": "Ingresos del ejercicio (4)", "prefijos": ["4"], "modelo": "patrimonio.ingresos"},
    {"concepto": "Gastos del ejercicio (5)", "prefijos": ["5"], "modelo": "patrimonio.gastos"},
]


def _n(t: str) -> str:
    t = unicodedata.normalize("NFKD", str(t)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", t).strip()


def perfiles(esp: Espacio) -> dict:
    datos = leer_yaml(esp.conocimiento / "perfiles_software.yaml", {}) or {}
    p = dict(PERFILES_BASE)
    p.update(datos.get("perfiles") or {})
    return {"perfiles": p, "cuadres": datos.get("cuadres") or CUADRES_BASE,
            "tolerancia": float(datos.get("tolerancia", 1000))}


# --------------------------------------------------------------------------- Excel estándar

class ExcelAmbiguo(ValueError):
    """Hay más de un Excel candidato en la entrada: el sistema no elige por su cuenta."""


def candidatos_excel(carpeta: Path, ficha: dict) -> list[Path]:
    patron = (ficha.get("entradas") or {}).get("excel", "*.xlsx")
    return sorted(p for p in carpeta.glob(patron) if not p.name.startswith(("~$", ".")))


def ubicar_excel(carpeta: Path, ficha: dict) -> Path | None:
    """El Excel del expediente. Con dos o más candidatos (p. ej. una copia «… 2.xlsx» de Finder) lanza ExcelAmbiguo:
    tomar uno en orden alfabético podría validar el archivo equivocado."""
    candidatos = candidatos_excel(carpeta, ficha)
    if len(candidatos) > 1:
        raise ExcelAmbiguo("Hay " + str(len(candidatos)) + " archivos Excel en la entrada (" +
                           ", ".join(f"«{p.name}»" for p in candidatos) + "). Deje solo el que corresponde y vuelva a validar.")
    return candidatos[0] if candidatos else None


def revisar_estructura(ruta: Path, mapa: dict | None = None) -> list[Hallazgo]:
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True)
    from revisoria.excel import buscar_hoja
    titulos = list(wb.sheetnames)
    wb.close()
    mapeadas = (mapa or {}).get("hojas", {})
    faltan = [h for h in HOJAS_REQUERIDAS
              if (mapeadas[h] not in titulos if h in mapeadas else buscar_hoja(titulos, h) is None)]
    if not faltan:
        return []
    return [Hallazgo("N01", "MEDIA", "Normalización", f"Faltan {len(faltan)} hoja(s) del Excel estándar",
                     "No se encontraron: " + ", ".join(faltan) + ". Las secciones que dependen de ellas no se "
                     "generarán; la información faltante no bloquea el resto del informe.",
                     [ruta.name], "Solicitar al cliente o completar las hojas antes de emitir.")]


# --------------------------------------------------------------------------- Balance de prueba

def _leer_tabla(ruta: Path) -> list[list[str]]:
    if ruta.suffix.lower() in (".xlsx", ".xlsm"):
        import openpyxl
        ws = openpyxl.load_workbook(ruta, data_only=True, read_only=True).active
        return [["" if v is None else str(v) for v in fila] for fila in ws.iter_rows(values_only=True)]
    texto = ruta.read_bytes().decode("utf-8-sig", errors="replace")
    try:
        dialecto = csv.Sniffer().sniff(texto[:4000], delimiters=";,\t|")
    except csv.Error:
        dialecto = csv.excel
    return [fila for fila in csv.reader(io.StringIO(texto), dialecto)]


def _numero(t: str, decimal: str = "auto") -> float | None:
    s = str(t).strip().replace("$", "").replace(" ", "")
    if not s:
        return None
    neg = s.startswith("(") and s.endswith(")") or s.startswith("-")
    s = s.strip("()-")
    if decimal == "auto":
        if "," in s and "." in s:
            decimal = "," if s.rfind(",") > s.rfind(".") else "."
        elif s.count(",") == 1 and len(s.split(",")[1]) <= 2:
            decimal = ","
        elif s.count(".") == 1 and len(s.split(".")[1]) <= 2:
            decimal = "."
        else:
            decimal = "."
            s = s.replace(",", "").replace(".", "")
    if decimal == ",":
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def _ubicar_columnas(filas, perfil):
    for i, fila in enumerate(filas[:30]):
        enc = [_n(c) for c in fila]
        idx = {}
        for campo, alias in perfil["columnas"].items():
            for a in alias:
                if _n(a) in enc:
                    idx[campo] = enc.index(_n(a))
                    break
        if "cuenta" in idx and "saldo" in idx:
            return i, idx
    return None, None


def proponer_mapeo(encabezados: list[str]) -> dict:
    """Propone qué columna es cuenta / nombre / saldo para un formato desconocido.

    Heurística por palabras clave. Es solo una PROPUESTA: queda en conocimiento/propuestas/ y el profesional
    la aprueba copiándola a perfiles_software.yaml.
    """
    claves = {"cuenta": ["cuenta", "codigo", "cod"], "nombre": ["nombre", "descrip", "detalle"],
              "saldo": ["saldo final", "nuevo saldo", "saldo", "final"]}
    out = {}
    for campo, ks in claves.items():
        for k in ks:
            hit = next((h for h in encabezados if k in _n(h) and h not in out.values()), None)
            if hit:
                out[campo] = [hit]
                break
    return out


def leer_balance(esp: Espacio, carpeta: Path, ficha: dict, cid: str) -> tuple[dict | None, list[Hallazgo]]:
    cfg = (ficha.get("entradas") or {}).get("balance")
    if not cfg:
        return None, [Hallazgo("N02", "INFO", "Normalización", "Sin balance de prueba para cuadrar",
                               "La ficha del cliente no define el balance de prueba del software contable; no se "
                               "ejecuta el cuadre contra balance (V19).", [], "Agregar el balance en la ficha.")]
    archivos = sorted(carpeta.glob(cfg.get("archivo", "balance*")))
    if not archivos:
        return None, [Hallazgo("N02", "MEDIA", "Normalización", "No llegó el balance de prueba del periodo",
                               f"Se esperaba un archivo «{cfg.get('archivo')}» en la carpeta de entrada.",
                               [str(carpeta.name)], "Solicitar el balance de prueba al cliente.")]
    ruta = archivos[0]
    P = perfiles(esp)
    nombre_perfil = cfg.get("perfil") or ficha.get("software") or "generico"
    perfil = P["perfiles"].get(nombre_perfil)
    filas = _leer_tabla(ruta)
    if perfil is None or perfil.get("estado") != "aprobado":
        fila_enc = next((f for f in filas[:30] if sum(1 for c in f if c.strip()) >= 3), [])
        prop = {"software": nombre_perfil, "archivo": ruta.name, "encabezados": fila_enc,
                "mapeo_propuesto": proponer_mapeo(fila_enc), "fecha": ahora(), "estado": "propuesto",
                "instruccion": "Revise el mapeo; si es correcto, cópielo a perfiles_software.yaml con estado: aprobado."}
        guardar_json(esp.conocimiento / "propuestas" / f"mapeo_{nombre_perfil}_{cid}.json", prop)
        return None, [Hallazgo("N03", "MEDIA", "Normalización", f"Formato de balance sin perfil aprobado («{nombre_perfil}»)",
                               "Se propuso un mapeo de columnas que debe aprobar el profesional antes de usarse. "
                               "Mientras tanto no se cuadra contra balance.",
                               [ruta.name], "Aprobar la propuesta en conocimiento/propuestas/.")]
    i, idx = _ubicar_columnas(filas, perfil)
    if i is None:
        return None, [Hallazgo("N04", "MEDIA", "Normalización", "No se reconocen las columnas del balance",
                               f"El archivo {ruta.name} no tiene las columnas del perfil «{nombre_perfil}».",
                               [ruta.name], "Verificar el exporte o actualizar el perfil del software.")]
    saldos = {}
    for fila in filas[i + 1:]:
        if len(fila) <= max(idx.values()):
            continue
        cta = re.sub(r"\D", "", fila[idx["cuenta"]])
        v = _numero(fila[idx["saldo"]], perfil.get("decimal", "auto"))
        if cta and v is not None:
            saldos[cta] = saldos.get(cta, 0) + v
    return {"archivo": ruta.name, "perfil": nombre_perfil, "saldos": saldos}, []


def _saldo_prefijo(saldos: dict, prefijo: str) -> float | None:
    if prefijo in saldos:
        return saldos[prefijo]
    hijos = [c for c in saldos if c.startswith(prefijo)]
    if not hijos:
        return None
    hojas = [c for c in hijos if not any(o != c and o.startswith(c) for o in hijos)]
    return sum(saldos[c] for c in hojas)


def _valor_modelo(m: dict, ruta: str):
    o = m
    for k in ruta.split("."):
        if not isinstance(o, dict) or o.get(k) is None:
            return None
        o = o[k]
    return o


def cuadrar(esp: Espacio, modelo: dict, balance: dict | None) -> tuple[list[dict], list[Hallazgo]]:
    if not balance:
        return [], []
    P = perfiles(esp)
    filas, hs = [], []
    for c in P["cuadres"]:
        vm = _valor_modelo(modelo, c["modelo"])
        vb = sum(v for v in (_saldo_prefijo(balance["saldos"], p) for p in c["prefijos"]) if v is not None) \
            if any(_saldo_prefijo(balance["saldos"], p) is not None for p in c["prefijos"]) else None
        if vm is None or vb is None:
            filas.append({"concepto": c["concepto"], "excel": vm, "balance": vb, "diferencia": None, "ok": None})
            continue
        dif = abs(vm) - abs(vb)
        ok = abs(dif) <= P["tolerancia"]
        filas.append({"concepto": c["concepto"], "excel": abs(vm), "balance": abs(vb), "diferencia": dif, "ok": ok})
        if not ok:
            from revisoria.narrativa import pesos
            hs.append(Hallazgo("V19", "ALTA", "Cuadre contra balance", f"{c['concepto']}: el Excel no cuadra con el balance",
                               f"Excel {pesos(abs(vm))} vs balance de prueba {pesos(abs(vb))}; diferencia {pesos(dif)}.",
                               [c["modelo"], f"{balance['archivo']} · cuentas {', '.join(c['prefijos'])}"],
                               "Conciliar la hoja del Excel con el auxiliar contable antes de emitir."))
    return filas, hs
