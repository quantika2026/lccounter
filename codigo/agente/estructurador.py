"""Subagente ESTRUCTURADOR: lee un Excel con otra estructura y propone la receta de lectura de una sección.

Entrenamiento (sin reentrenar ningún modelo): aprendizaje por ejemplos aprobados.
  - Ejemplos base: las recetas del informe modelo (config/lectores_modelo.json), cada una con el «vistazo» de su hoja.
    tests/test_lectores.py prueba que con ellas el lector genérico reproduce exactamente al nativo.
  - Ejemplos aprendidos: cada receta que un revisor aprueba se guarda en conocimiento/lectores_aprendidos.json con su
    vistazo; el próximo cliente con una hoja parecida se resuelve con menos intentos.

Ciclo del subagente: mira la hoja → propone la receta (IA, con los ejemplos) → el sistema la ejecuta y la prueba
(revisoria/lectores.probar) → si hay errores se los devuelve al agente para que corrija (hasta 3 intentos) →
el revisor ve lo leído y aprueba. El agente nunca aprueba ni inventa datos: solo dice dónde están.
"""
from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

from revisoria import lectores
from revisoria.excel import Libro, Periodo, norm

RAIZ_CODIGO = Path(__file__).resolve().parent.parent
MODELO = RAIZ_CODIGO / "config" / "lectores_modelo.json"
APRENDIDOS = "lectores_aprendidos.json"
MAX_INTENTOS = 3


# ------------------------------------------------------------------ lo que el agente ve de una hoja

def vistazo(ws, p: Periodo | None = None, filas_inicio: int = 22, max_col: int = 12) -> list[str]:
    """Filas con coordenadas («12 | B: CUENTAS POR PAGAR… | D: 8817802»): las primeras filas y, en hojas largas,
    las filas alrededor de los títulos del mes del periodo y de las filas TOTAL."""
    def fila(r):
        celdas = []
        for c in range(1, min(ws.max_column, max_col) + 1):
            v = ws.cell(r, c).value
            if v in (None, ""):
                continue
            if isinstance(v, (datetime, date)):
                v = v.strftime("%Y-%m-%d")
            elif isinstance(v, float) and v.is_integer():
                v = int(v)
            celdas.append(f"{ws.cell(r, c).column_letter}: {str(v).strip()[:45]}")
        return f"{r} | " + " | ".join(celdas) if celdas else None

    elegidas = list(range(1, min(ws.max_row, filas_inicio) + 1))
    if p is not None and ws.max_row > filas_inicio:
        mes = p.nombre_mes
        for r in range(filas_inicio + 1, ws.max_row + 1):
            textos = [norm(ws.cell(r, c).value) for c in range(1, min(ws.max_column, max_col) + 1)
                      if isinstance(ws.cell(r, c).value, str)]
            if any(mes in t and len(t) < 90 for t in textos):
                elegidas += range(max(1, r - 1), min(ws.max_row, r + 4) + 1)
            elif any(t.startswith("total") for t in textos):
                elegidas += [r]
    salida, previa = [], 0
    for r in sorted(set(elegidas)):
        t = fila(r)
        if not t:
            continue
        if previa and r > previa + 1:
            salida.append("…")
        salida.append(t)
        previa = r
    return salida[:90]


# ------------------------------------------------------------------ ejemplos (entrenamiento)

def ejemplos(conocimiento: Path | None, seccion: str, n: int = 3) -> list[dict]:
    """Ejemplos para el agente: primero los aprobados de la misma sección (más recientes), luego el informe modelo."""
    base = json.loads(MODELO.read_text(encoding="utf-8"))["recetas"]
    aprendidos = json.loads((Path(conocimiento) / APRENDIDOS).read_text(encoding="utf-8")) if conocimiento and (Path(conocimiento) / APRENDIDOS).exists() else []
    propios = [a for a in reversed(aprendidos) if a["receta"]["seccion"] == seccion]
    modelo = [{"origen": "informe modelo", "vistazo": r.get("vistazo", []), "receta": {k: v for k, v in r.items() if k != "vistazo"}}
              for r in base if r["seccion"] == seccion]
    otros = [{"origen": "informe modelo", "vistazo": r.get("vistazo", [])[:12],
              "receta": {k: v for k, v in r.items() if k != "vistazo"}} for r in base if r["seccion"] != seccion][:2]
    return (propios + modelo)[:n] + otros


def aprender(conocimiento: Path, cliente: str, receta: dict, vista: list[str], revisor: str) -> None:
    r = Path(conocimiento) / APRENDIDOS
    datos = json.loads(r.read_text(encoding="utf-8")) if r.exists() else []
    datos = [d for d in datos if not (d["cliente"] == cliente and d["receta"]["seccion"] == receta["seccion"])]
    datos.append({"cliente": cliente, "aprobado_por": revisor, "fecha": datetime.now().isoformat(timespec="seconds"),
                  "vistazo": vista[:40], "receta": receta})
    r.parent.mkdir(parents=True, exist_ok=True)
    r.write_text(json.dumps(datos[-200:], ensure_ascii=False, indent=1), encoding="utf-8")


# ------------------------------------------------------------------ propuesta con IA

ESQUEMA = {
    "type": "object", "additionalProperties": False,
    "required": ["corresponde", "hoja", "titulo", "fila_encabezado", "columnas", "niveles", "parametros", "motivo"],
    "properties": {
        "corresponde": {"type": "boolean"},
        "hoja": {"type": "string"},
        "titulo": {"type": "string"},
        "fila_encabezado": {"type": "integer"},
        "columnas": {"type": "array", "items": {
            "type": "object", "additionalProperties": False, "required": ["campo", "columnas"],
            "properties": {"campo": {"type": "string"}, "columnas": {"type": "array", "items": {"type": "string"}}}}},
        "niveles": {"type": "array", "items": {"type": "string"}},
        "parametros": {"type": "array", "items": {
            "type": "object", "additionalProperties": False, "required": ["nombre", "valor"],
            "properties": {"nombre": {"type": "string"}, "valor": {"type": "string"}}}},
        "motivo": {"type": "string"},
    },
}

SISTEMA = """Eres el subagente ESTRUCTURADOR de LCCOUNTER (revisoría fiscal de propiedad horizontal en Colombia).
Cada copropiedad organiza su Excel de papeles de trabajo a su manera, pero el informe en PowerPoint es siempre el
mismo. Tu tarea: escribir la RECETA que le dice al lector genérico dónde está cada campo estándar de UNA sección.
No extraes datos ni calculas cifras: solo indicas filas y columnas.

Receta:
- hoja: título exacto de la hoja.
- titulo: si la hoja tiene UN BLOQUE POR MES (cada uno con un título como «CUENTAS POR PAGAR A AGOSTO 2026»), la
  plantilla de ese título en minúsculas con {mes} y {anio} en lugar del mes y el año (p. ej.
  «cuentas por pagar a {mes} {anio}»). Si el título del bloque no trae año, omite {anio}. Si la hoja tiene UNA sola
  tabla, deja titulo vacío ("").
- fila_encabezado: si titulo está vacío, el número de la fila de encabezados de la tabla; si hay titulo, 0.
  (Con titulo, el encabezado se asume en la fila siguiente al título.)
- columnas: para cada campo estándar, la letra (o letras, si el dato está repartido en dos columnas que se unen) de
  la columna donde está. Usa SOLO los campos de la lista. Los obligatorios deben ir siempre.
- Columnas relativas: en CARTERA con titulo, si cada mes es una lista con su propio título y las columnas se
  repiten al lado, da apto/valor relativos a la columna del título («+1», «+2»).
- parametros: solo para LIQUIDEZ, CARTERA y PRESUPUESTO (ver «parametros» de la sección): textos con que empiezan
  las filas clave (p. ej. encabezado «análisis de liquidez», total «total equivalente al efectivo»), en minúsculas.
  En liquidez y cartera sin titulo la columna del mes se busca sola por su encabezado (fecha o «ago-26»): no la des.
- niveles: solo para matrices: las letras de las columnas de «nivel de cumplimiento» por periodo; si no hay, [].
- corresponde: false si la hoja NO contiene esta sección (p. ej. una lista de PAGOS de pólizas no es la
  VERIFICACIÓN de pólizas). En ese caso deja el resto vacío y explica en motivo.
- motivo: una frase que cite los encabezados que viste.
Fíjate en los ejemplos: muestran cómo se ve una hoja (vistazo con coordenadas) y la receta correcta para ella.
Si el encabezado de una columna no coincide con lo que realmente contienen los datos (columnas corridas),
manda lo que contienen los DATOS."""


def _a_receta(seccion: str, r: dict) -> dict:
    rec = {"seccion": seccion, "hoja": r["hoja"],
           "columnas": {c["campo"]: (c["columnas"][0] if len(c["columnas"]) == 1 else c["columnas"])
                        for c in r["columnas"] if c["columnas"]}}
    if r.get("titulo"):
        rec["titulo"] = r["titulo"]
    else:
        rec["fila_encabezado"] = r.get("fila_encabezado")
    if r.get("niveles"):
        rec["niveles"] = r["niveles"]
    validos = (lectores.CAMPOS[seccion].get("parametros") or {})
    for prm in r.get("parametros") or []:
        if prm.get("nombre") in validos and str(prm.get("valor", "")).strip():
            v = str(prm["valor"]).strip()
            rec[prm["nombre"]] = int(v) if prm["nombre"] == "signo_total" and v.lstrip("-").isdigit() else v
    return rec


def proponer(raiz: Path, lib: Libro, seccion: str, hoja: str, p: Periodo, conocimiento: Path | None = None,
             llamar=None, log=None) -> dict:
    """Propone, prueba y corrige la receta de `seccion` en `hoja`. `llamar(sistema, usuario, esquema)` permite
    inyectar el modelo (pruebas); por defecto usa el proveedor asignado al uso «estructurador»."""
    if seccion not in lectores.CAMPOS:
        raise ValueError(f"La sección «{seccion}» no admite receta. Disponibles: {', '.join(lectores.CAMPOS)}.")
    if hoja not in lib.valores.sheetnames:
        raise ValueError(f"La hoja «{hoja}» no está en el Excel.")
    if llamar is None:
        from app import proveedores
        disp = proveedores.disponible(raiz, "estructurador")
        if not disp["ok"]:
            raise proveedores.ErrorIA(disp["ayuda"] + " También puede escribir la receta a mano y probarla.")

        def llamar(sistema, usuario, esquema):
            return proveedores.json_estructurado(raiz, "estructurador", sistema, usuario, esquema, "high")
    ws = lib.valores[hoja]
    vista = vistazo(ws, p)
    def_sec = lectores.CAMPOS[seccion]
    contexto = {
        "seccion": seccion, "descripcion": def_sec["descripcion"],
        "campos": [{"campo": c, "tipo": {"t": "texto", "n": "número", "f": "fecha", "p": "mes"}[t], "obligatorio": ob,
                    "que_es": d} for c, (t, ob, d) in def_sec["campos"].items()],
        "parametros": [{"nombre": k, "obligatorio": ob, "que_es": q} for k, (ob, q) in (def_sec.get("parametros") or {}).items()],
        "es_matriz": bool(def_sec.get("matriz")), "periodo": f"{p.nombre_mes} {p.anio}",
        "ejemplos": ejemplos(conocimiento, seccion),
        "hoja": hoja, "vistazo_de_la_hoja": vista,
    }
    intentos, mensaje, modelo = [], json.dumps(contexto, ensure_ascii=False), None
    for n in range(1, MAX_INTENTOS + 1):
        datos, modelo = llamar(SISTEMA, mensaje, ESQUEMA)
        if not datos.get("corresponde", True):
            intentos.append({"intento": n, "corresponde": False, "motivo": datos.get("motivo", "")})
            return {"corresponde": False, "motivo": datos.get("motivo", ""), "intentos": intentos, "modelo": modelo,
                    "vistazo": vista}
        datos["hoja"] = hoja                                   # la hoja la eligió la persona, no el modelo
        receta = _a_receta(seccion, datos)
        prueba = lectores.probar(lib, receta, p)
        intentos.append({"intento": n, "receta": receta, "errores": prueba["errores"], "avisos": prueba["avisos"],
                         "motivo": datos.get("motivo", "")})
        if log:
            log(f"intento {n}: {'OK' if prueba['ok'] else '; '.join(prueba['errores'])}")
        if prueba["ok"]:
            return {"corresponde": True, "receta": receta, "prueba": prueba, "motivo": datos.get("motivo", ""),
                    "intentos": intentos, "modelo": modelo, "vistazo": vista}
        mensaje = json.dumps({**contexto, "tu_receta_anterior": receta, "errores_al_probarla": prueba["errores"],
                              "avisos": prueba["avisos"], "lo_que_leyo": prueba["muestra"],
                              "instruccion": "Corrige la receta para que no tenga esos errores."}, ensure_ascii=False)
    return {"corresponde": True, "receta": receta, "prueba": prueba, "motivo": datos.get("motivo", ""),
            "intentos": intentos, "modelo": modelo, "vistazo": vista}


def aprobar(esp, cid: str, receta: dict, p: Periodo, excel: Path, revisor: str) -> dict:
    """Guarda la receta en el mapa del cliente (y la hoja en el mapa) y la suma a los ejemplos aprendidos."""
    from revisoria import esquema, mapa as mapa_cliente
    lib = Libro(str(excel))
    receta = lectores.validar_receta(dict(receta))
    prueba = lectores.probar(lib, receta, p)
    if not prueba["ok"]:
        raise ValueError("La receta tiene errores: " + "; ".join(prueba["errores"]))
    dir_cli = esp.clientes / cid
    mapa = mapa_cliente.cargar(dir_cli) or mapa_cliente.detectar(lib)
    mapa.setdefault("lectores", {})[receta["seccion"]] = receta
    clave = next(s.hoja for s in esquema.SECCIONES if s.clave == receta["seccion"])
    mapa.setdefault("hojas", {})[clave] = receta["hoja"]
    mapa_cliente.guardar(dir_cli, mapa)
    aprender(esp.conocimiento, cid, receta, vistazo(lib.valores[receta["hoja"]], p), revisor)
    return {"receta": receta, "prueba": prueba}
