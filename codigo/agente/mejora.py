"""Agente de mejora continua (automejora controlada).

Ciclo: Captura → Clasificación → Propuesta → Prueba → Aprobación y versión.

- Captura: al aprobar, se compara el informe generado con el que el revisor dejó como final (PPTX) y el
  borrador de opinión con el editado. Cada diferencia se guarda en conocimiento/correcciones.jsonl.
- Clasificación: cifra · texto · estructura · opinión · devolución.
- Propuesta: agrupa correcciones y propone ajustes (plantilla de opinión, revisión de regla o extractor).
- Prueba: toda propuesta se prueba contra el golden set (informes ya aprobados) antes de publicarse.
- Aprobación y versión: el revisor aprueba; se publica una versión nueva, reversible.

Límites: la lógica de cálculo nunca se modifica automáticamente (las correcciones de cifras solo generan
una alerta para revisar el extractor) y las reglas de validación nunca se relajan.
"""
from __future__ import annotations

import json
import re
import shutil
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import yaml

from revisoria.excel import norm

from .comun import APROBADO, DEVUELTO, PENDIENTE, Espacio, ahora, guardar_json, leer_json

RX_DIG = re.compile(r"\d[\d\.,]*")


def _h(t: str) -> str:
    import hashlib
    return hashlib.md5(t.encode()).hexdigest()[:6]


# ------------------------------------------------------------------ captura

def _unidades(ruta) -> list[dict]:
    from pptx import Presentation
    prs = Presentation(str(ruta))
    out = []
    for n, s in enumerate(prs.slides, 1):
        titulo = ""
        u = {}
        for k, sh in enumerate(s.shapes):
            if sh.has_text_frame:
                for j, par in enumerate(sh.text_frame.paragraphs):
                    t = "".join(r.text for r in par.runs).strip()
                    if t:
                        u[f"s{k}.p{j}"] = t
                        titulo = titulo or t
            elif getattr(sh, "has_table", False) and sh.has_table:
                for r, fila in enumerate(sh.table.rows):
                    for c, celda in enumerate(fila.cells):
                        if celda.text.strip():
                            u[f"s{k}.t{r}.{c}"] = celda.text.strip()
            elif getattr(sh, "has_chart", False) and sh.has_chart:
                for ns, serie in enumerate(sh.chart.series):
                    try:
                        for nv, valor in enumerate(serie.values):
                            if valor is not None:
                                u[f"s{k}.g{ns}.{nv}"] = str(valor)
                    except AttributeError:
                        pass
        out.append({"n": n, "titulo": titulo, "u": u})
    return out


def tipo_diapositiva(titulo: str) -> str:
    t = norm(titulo)
    t = re.sub(r"\(\d+/\d+\)", "", t)
    t = RX_DIG.sub("#", t)
    return re.sub(r"\s+", " ", t).strip()[:60]


def diferencias_pptx(generado, final) -> list[dict]:
    a, b = _unidades(generado), _unidades(final)
    cambios = []
    if len(a) != len(b):
        cambios.append({"tipo": "estructura", "diapositiva": None, "tipo_diapositiva": "(presentación)",
                        "antes": f"{len(a)} diapositivas", "despues": f"{len(b)} diapositivas"})
    for da, db in zip(a, b):
        td = tipo_diapositiva(da["titulo"])
        for k in sorted(set(da["u"]) | set(db["u"])):
            va, vb = da["u"].get(k), db["u"].get(k)
            if va == vb:
                continue
            if va is None or vb is None:
                tipo = "estructura"
            elif Counter(RX_DIG.findall(va)) != Counter(RX_DIG.findall(vb)):
                tipo = "cifra"
            else:
                tipo = "texto"
            cambios.append({"tipo": tipo, "diapositiva": da["n"], "tipo_diapositiva": td, "elemento": k,
                            "antes": va, "despues": vb})
    return cambios


def _secciones_md(texto: str) -> dict[str, str]:
    secs, actual = {}, None
    for linea in texto.splitlines():
        m = re.match(r"##\s+\d+\.\s+(.*)", linea)
        if m:
            actual = m.group(1).strip()
            secs[actual] = []
        elif actual and not linea.startswith("**Nivel") and not linea.startswith("## "):
            secs[actual].append(linea)
        elif linea.startswith("## "):
            actual = None
    return {k: "\n".join(v).strip() for k, v in secs.items()}


def plantillar(texto: str, cifras: dict) -> tuple[str, list[str]]:
    """Convierte un texto final del revisor en plantilla: reemplaza las cifras del mes por {marcadores}."""
    t = texto.replace("{", "{{").replace("}", "}}")
    for clave, valor in sorted(cifras.items(), key=lambda kv: -len(str(kv[1]))):
        t = re.sub(r"(?<![\d.,])" + re.escape(str(valor)) + r"(?![\d]|[.,]\d)", "{" + clave + "}", t)
    sueltas = [x for x in RX_DIG.findall(re.sub(r"\{[a-z_]+\}", "", t)) if len(re.sub(r"\D", "", x)) >= 4]
    return t, sueltas


def diferencias_borrador(items: list[dict], texto_final: str) -> list[dict]:
    final = _secciones_md(texto_final)
    out = []
    for it in items:
        nuevo = final.get(it["titulo"])
        if nuevo is None:
            out.append({"tipo": "opinion", "clave": it["clave"], "accion": "descartada", "antes": it["texto"]})
        elif nuevo.strip() != it["texto"].strip():
            plantilla, sueltas = plantillar(nuevo.strip(), it["cifras"])
            out.append({"tipo": "opinion", "clave": it["clave"], "accion": "editada", "antes": it["texto"],
                        "despues": nuevo.strip(), "plantilla": plantilla, "cifras_literales": sueltas})
    return out


PROPUESTA_CLAUDE = "Borrador_propuesta_claude.md"     # lo escribe el subagente redactor-ph


def marcar_origen_claude(cambios: list[dict], items: list[dict], texto_claude: str) -> None:
    """Marca con origen «claude» las opiniones cuyo texto final es exactamente el que propuso Claude.
    Así la mejora continua aprende solo de la redacción del revisor y no de la de Claude."""
    propuesta = _secciones_md(texto_claude)
    titulos = {it["clave"]: it["titulo"] for it in items}
    for c in cambios:
        if c.get("tipo") == "opinion" and c.get("accion") == "editada":
            sugerido = propuesta.get(titulos.get(c["clave"], ""))
            if sugerido is not None and sugerido.strip() == c["despues"]:
                c["origen"] = "claude"


def registrar(esp: Espacio, periodo: str, cid: str, cambios: list[dict]) -> int:
    ruta = esp.conocimiento / "correcciones.jsonl"
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with ruta.open("a", encoding="utf-8") as f:
        for c in cambios:
            f.write(json.dumps({"fecha": ahora(), "periodo": periodo, "cliente": cid, **c}, ensure_ascii=False) + "\n")
    return len(cambios)


def correcciones(esp: Espacio) -> list[dict]:
    ruta = esp.conocimiento / "correcciones.jsonl"
    if not ruta.exists():
        return []
    return [json.loads(l) for l in ruta.read_text(encoding="utf-8").splitlines() if l.strip()]


# ------------------------------------------------------------------ golden set y regresión

def agregar_golden(esp: Espacio, periodo: str, cid: str, ruta_excel: Path, modelo_json: Path, ficha: dict) -> Path:
    d = esp.conocimiento / "golden" / f"{periodo}_{cid}"
    d.mkdir(parents=True, exist_ok=True)
    shutil.copy2(ruta_excel, d / "excel.xlsx")
    shutil.copy2(modelo_json, d / "modelo_aprobado.json")
    (d / "ficha.yaml").write_text(yaml.safe_dump(ficha, allow_unicode=True, sort_keys=False), encoding="utf-8")
    mapa = esp.clientes / cid / "mapa.json"
    if mapa.exists():
        shutil.copy2(mapa, d / "mapa.json")
    return d


def _planos(o, pref="", out=None):
    out = {} if out is None else out
    if isinstance(o, dict):
        for k, v in o.items():
            if not k.startswith("_") and k not in ("soportes",):
                _planos(v, f"{pref}.{k}" if pref else k, out)
    elif isinstance(o, list):
        for i, v in enumerate(o):
            _planos(v, f"{pref}[{i}]", out)
    elif isinstance(o, (int, float)) and not isinstance(o, bool):
        out[pref] = round(o, 2)
    return out


def regresion(esp: Espacio, biblioteca_candidata: dict | None = None) -> dict:
    """Re-extrae cada informe aprobado y compara cifra por cifra. Prueba también la biblioteca candidata."""
    from revisoria.excel import Periodo
    from revisoria.extractor import extraer

    from . import redaccion
    casos = []
    for d in sorted((esp.conocimiento / "golden").glob("*_*")):
        aprobado = leer_json(d / "modelo_aprobado.json", {})
        if not aprobado:
            continue
        try:
            nuevo, _ = extraer(str(d / "excel.xlsx"), Periodo.desde_texto(aprobado["meta"]["periodo"]),
                               {k: aprobado["meta"].get(k) for k in ("copropiedad", "nit", "unidades", "numero_informe")},
                               leer_json(d / "mapa.json", {}))
            nuevo = json.loads(json.dumps(nuevo, default=str))
            a, b = _planos(aprobado), _planos(nuevo)
            difs = [k for k in a if k in b and abs(a[k] - b[k]) > 0.5] + [k for k in a if k not in b]
            errores_bib = []
            if biblioteca_candidata:
                for clave, datos in redaccion.condiciones(nuevo).items():
                    e = biblioteca_candidata.get(clave)
                    if e:
                        try:
                            e["texto"].format(**datos)
                        except (KeyError, IndexError, ValueError) as ex:
                            errores_bib.append(f"{clave}: marcador inválido {ex}")
            casos.append({"caso": d.name, "cifras": len(a), "diferencias": difs[:20], "n_diferencias": len(difs),
                          "errores_biblioteca": errores_bib, "ok": not difs and not errores_bib})
        except Exception as ex:                                  # un caso roto no detiene la prueba
            casos.append({"caso": d.name, "ok": False, "error": str(ex)})
    res = {"fecha": ahora(), "casos": casos, "ok": all(c["ok"] for c in casos), "n": len(casos)}
    guardar_json(esp.conocimiento / "ultima_regresion.json", res)
    return res


# ------------------------------------------------------------------ propuestas

def proponer(esp: Espacio) -> list[dict]:
    corr = correcciones(esp)
    previas = {p["id"]: p for p in leer_json(esp.conocimiento / "propuestas.json", [])}
    nuevas = []

    # 1. Opiniones editadas por el revisor → nueva plantilla de la biblioteca (la última versión manda)
    por_clave = defaultdict(list)
    for c in corr:
        if c["tipo"] == "opinion" and c.get("accion") == "editada" and c.get("origen") != "claude":
            por_clave[c["clave"]].append(c)
    for clave, cs in por_clave.items():
        ult = cs[-1]
        nuevas.append({"id": f"BIB-{clave}-{len(cs)}", "tipo": "biblioteca", "clave": clave,
                       "titulo": f"Actualizar la plantilla de opinión «{clave}» con la redacción del revisor",
                       "evidencia": f"{len(cs)} edición(es); última de {ult['cliente']} {ult['periodo']}",
                       "plantilla": ult["plantilla"], "advertencias": ult.get("cifras_literales") or [],
                       "riesgo": "bajo (solo texto; las cifras siguen saliendo del Excel)"})

    # 2. Cambios de texto repetidos en el mismo tipo de diapositiva en 2+ clientes → ajustar plantilla
    grupos = defaultdict(set)
    for c in corr:
        if c["tipo"] == "texto":
            grupos[c["tipo_diapositiva"]].add(c["cliente"])
    for td, cli in grupos.items():
        if len(cli) >= 2:
            nuevas.append({"id": f"TXT-{_h(td)}", "tipo": "plantilla_texto",
                           "titulo": f"Ajustar el texto estándar de «{td}»",
                           "evidencia": f"Corregido en {len(cli)} clientes: {', '.join(sorted(cli))}",
                           "riesgo": "bajo"})

    # 3. Cifras corregidas → alerta para revisar extractor o regla (NUNCA se cambia sola)
    cifras = [c for c in corr if c["tipo"] == "cifra"]
    por_td = Counter(c["tipo_diapositiva"] for c in cifras)
    for td, n in por_td.items():
        ejemplos = [c for c in cifras if c["tipo_diapositiva"] == td][:3]
        nuevas.append({"id": f"CIF-{_h(td)}", "tipo": "revisar_calculo",
                       "titulo": f"El revisor corrigió cifras en «{td}» ({n} vez/veces)",
                       "evidencia": "; ".join(f"{e['cliente']} {e['periodo']}: «{e['antes']}» → «{e['despues']}»"
                                              for e in ejemplos),
                       "riesgo": "alto: requiere revisión humana del extractor o del Excel; no se automatiza"})

    # 4. Estructura y devoluciones
    est = [c for c in corr if c["tipo"] == "estructura"]
    if est:
        nuevas.append({"id": f"EST-{len(est)}", "tipo": "plantilla_presentacion",
                       "titulo": f"{len(est)} cambio(s) de estructura en las presentaciones",
                       "evidencia": "; ".join(sorted({c['tipo_diapositiva'] for c in est})[:5]), "riesgo": "medio"})
    devs = [c for c in corr if c["tipo"] == "devolucion"]
    for c in devs:
        nuevas.append({"id": f"DEV-{c['cliente']}-{c['periodo']}", "tipo": "devolucion",
                       "titulo": f"Motivo de devolución: {c['motivo']}",
                       "evidencia": f"{c['cliente']} {c['periodo']}", "riesgo": "a evaluar"})
    for grupo in _agrupar_motivos(devs):
        if len(grupo) >= 2:
            nuevas.append(_borrador_regla(esp, grupo))

    # 5. Auditoría mensual: lo que el segundo revisor encontró y el sistema no
    for c in corr:
        if c["tipo"] == "auditoria" and c["conclusion"] != "sin_observaciones":
            nuevas.append({"id": f"AUD-{c['cliente']}-{c['periodo']}", "tipo": "auditoria",
                           "titulo": f"Auditoría mensual: {c['observacion'] or c['conclusion']}",
                           "evidencia": f"{c['auditor']} revisó {c['cliente']} {c['periodo']} (aprobó {c['aprobado_por']}); "
                                        f"conclusión «{c['conclusion']}»",
                           "riesgo": "alto: el informe salió con un problema" if c["conclusion"] == "reproceso" else "medio"})

    out = []
    for p in nuevas:
        prev = previas.get(p["id"])
        p["estado"] = prev["estado"] if prev else "pendiente"
        p["fecha"] = prev.get("fecha", ahora()) if prev else ahora()
        out.append(p)
    guardar_json(esp.conocimiento / "propuestas.json", out)
    return out


PALABRAS_VACIAS = {"para", "como", "esta", "este", "esto", "corregir", "favor", "debe", "sobre", "desde", "entre",
                   "tiene", "falta", "informe", "excel", "hoja", "valor"}


def _palabras(motivo: str) -> set[str]:
    return {w for w in re.findall(r"[a-z0-9]+", norm(motivo)) if len(w) >= 3 and w not in PALABRAS_VACIAS}


def _agrupar_motivos(devs: list[dict]) -> list[list[dict]]:
    """Agrupa devoluciones con motivos parecidos (comparten al menos la mitad de sus palabras clave)."""
    grupos: list[list[dict]] = []
    for d in devs:
        pd = _palabras(d["motivo"])
        for g in grupos:
            pg = _palabras(g[0]["motivo"])
            if pd and pg and len(pd & pg) / min(len(pd), len(pg)) >= 0.5:
                g.append(d)
                break
        else:
            grupos.append([d])
    return grupos


def _borrador_regla(esp: Espacio, grupo: list[dict]) -> dict:
    """Borrador de regla nueva para un desarrollador: evidencia, reglas que ya lo detectaron y esqueleto de prueba."""
    rid = f"REG-{_h(' | '.join(sorted(d['motivo'] for d in grupo)))}"
    detectadas = Counter()
    for d in grupo:
        for h in leer_json(esp.salida(d["periodo"], d["cliente"]) / "hallazgos.json", []):
            if h["severidad"] in ("ALTA", "MEDIA"):
                detectadas[f"{h['regla']} {h['titulo']}"] += 1
    ya = [k for k, n in detectadas.items() if n == len(grupo)]
    casos = "\n".join(f"- {d['cliente']} {d['periodo']}: «{d['motivo']}»" for d in grupo)
    texto = f"""# Borrador de regla de validación · {rid}

> Generado por la mejora continua a partir de {len(grupo)} devoluciones con el mismo motivo.
> Es un BORRADOR para un desarrollador: no se aplica solo y no relaja ninguna regla existente.

## Evidencia (devoluciones del revisor fiscal)
{casos}

## Reglas actuales que ya marcaron estos informes (ALTA o MEDIA en todos)
{chr(10).join('- ' + k for k in ya) or '- Ninguna: el problema no lo detecta hoy ninguna regla.'}

## Qué decidir
{'- Una regla existente ya lo detecta: evaluar si su severidad debe subir o si el revisor necesita verlo antes.' if ya else
 '- Escribir una regla nueva en `revisoria/validador.py` (siguiente código V libre) que detecte el motivo.'}
- Agregar como caso de prueba el Excel de cada devolución (están en `entradas/<periodo>/<cliente>/`).
- Correr `regresion`: la regla no debe marcar los informes del golden set que se aprobaron sin este problema.

## Esqueleto
```python
# revisoria/validador.py
def _regla_nueva(m, lib, p):
    out = []
    # condición que reproduce el motivo de devolución
    return out

# tests/test_mvp.py
def test_regla_nueva_detecta_devoluciones():
    ...  # cargar el Excel devuelto y verificar que aparece el hallazgo
```
"""
    ruta = esp.conocimiento / "propuestas" / f"regla_{rid}.md"
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(texto, encoding="utf-8")
    return {"id": rid, "tipo": "regla_borrador",
            "titulo": f"Regla de validación nueva: «{grupo[0]['motivo']}» ({len(grupo)} devoluciones)",
            "evidencia": ", ".join(f"{d['cliente']} {d['periodo']}" for d in grupo)
                         + (f" · ya lo marca: {', '.join(k.split()[0] for k in ya)}" if ya else " · ninguna regla lo detecta hoy"),
            "archivo": str(ruta.relative_to(esp.raiz)), "riesgo": "medio: requiere desarrollo y prueba de regresión"}


def aprobar_propuesta(esp: Espacio, pid: str, aprobado_por: str) -> dict:
    props = leer_json(esp.conocimiento / "propuestas.json", [])
    p = next((x for x in props if x["id"] == pid), None)
    if not p:
        raise SystemExit(f"No existe la propuesta {pid}")
    if p["tipo"] == "revisar_calculo":
        raise SystemExit("Las correcciones de cifras no se aplican automáticamente: revise el extractor o el Excel "
                         "con el revisor y corra `regresion` después de cualquier cambio en el código.")
    if p["tipo"] != "biblioteca":
        p["estado"] = "aprobada (aplicación manual)"
        p["aprobado_por"] = aprobado_por
        guardar_json(esp.conocimiento / "propuestas.json", props)
        return {"ok": True, "detalle": "Registrada; este tipo de ajuste se aplica manualmente en la plantilla."}

    from . import redaccion
    candidata = redaccion.biblioteca(esp)
    entrada = dict(candidata.get(p["clave"], {}))
    entrada["texto"] = p["plantilla"]
    candidata[p["clave"]] = entrada
    prueba = regresion(esp, candidata)
    if not prueba["n"]:
        return {"ok": False, "detalle": "El golden set está vacío (aún no hay informes aprobados): no hay contra qué "
                                        "probar la propuesta; no se publica."}
    if not prueba["ok"]:
        return {"ok": False, "detalle": "La prueba contra el golden set falló; no se publica.", "regresion": prueba}
    if p.get("advertencias"):
        return {"ok": False, "detalle": "La plantilla trae cifras escritas a mano que no salen del Excel ("
                + ", ".join(p["advertencias"][:5]) + "): quedarían fijas en todos los meses. Reemplácelas por "
                "marcadores o retírelas y vuelva a correr `mejorar`; no se publica."}

    nv = publicar_plantilla(esp, p["clave"], entrada, aprobado_por, pid, f"{prueba['n']} caso(s) OK")
    p["estado"] = f"aplicada en {nv}"
    p["aprobado_por"] = aprobado_por
    guardar_json(esp.conocimiento / "propuestas.json", props)
    return {"ok": True, "detalle": f"Biblioteca actualizada (versión {nv}). Revertir con: revertir {nv}"}


def publicar_plantilla(esp: Espacio, clave: str, entrada: dict, aprobado_por: str, origen: str, regresion: str,
                       extra: dict | None = None) -> str:
    """Escribe una plantilla en la biblioteca como versión nueva y reversible (guarda la anterior en historial/)."""
    ruta_bib = esp.conocimiento / "biblioteca_opiniones.yaml"
    actual = (yaml.safe_load(ruta_bib.read_text(encoding="utf-8")) if ruta_bib.exists() else {}) or {}
    versiones = leer_json(esp.conocimiento / "versiones.json", [])
    nv = f"bib-{sum(1 for v in versiones if not v['version'].startswith('revertir-')) + 1:03d}"
    if ruta_bib.exists():
        (esp.conocimiento / "historial").mkdir(parents=True, exist_ok=True)
        shutil.copy2(ruta_bib, esp.conocimiento / "historial" / f"biblioteca_antes_de_{nv}.yaml")
    else:
        (esp.conocimiento / "historial").mkdir(parents=True, exist_ok=True)
        (esp.conocimiento / "historial" / f"biblioteca_antes_de_{nv}.yaml").write_text("{}\n", encoding="utf-8")
    actual[clave] = entrada
    ruta_bib.write_text(yaml.safe_dump(actual, allow_unicode=True, sort_keys=False, width=110), encoding="utf-8")
    versiones.append({"version": nv, "fecha": ahora(), "propuesta": origen, "aprobado_por": aprobado_por,
                      "regresion": regresion, **(extra or {})})
    guardar_json(esp.conocimiento / "versiones.json", versiones)
    return nv


def revertir(esp: Espacio, version: str) -> str:
    versiones = leer_json(esp.conocimiento / "versiones.json", [])
    if version.startswith("revertir-") or not any(v["version"] == version and not version.startswith("revertir-") for v in versiones):
        publicadas = [v["version"] for v in versiones if not v["version"].startswith("revertir-")]
        raise SystemExit(f"No existe la versión «{version}». Publicadas: {', '.join(publicadas) or 'ninguna'}.")
    resp = esp.conocimiento / "historial" / f"biblioteca_antes_de_{version}.yaml"
    ruta_bib = esp.conocimiento / "biblioteca_opiniones.yaml"
    if not resp.exists():
        raise SystemExit(f"No existe el respaldo necesario para revertir «{version}»; no se modificó la biblioteca.")
    shutil.copy2(resp, ruta_bib)
    versiones.append({"version": f"revertir-{version}", "fecha": ahora()})
    guardar_json(esp.conocimiento / "versiones.json", versiones)
    return f"Biblioteca restaurada al estado anterior a {version}."


# ------------------------------------------------------------------ versión vigente

def version_vigente(esp: Espacio) -> str:
    """Versión de la biblioteca en uso: «base» si no hay ninguna publicada o se revirtieron todas."""
    pila = []
    for v in leer_json(esp.conocimiento / "versiones.json", []):
        if v["version"].startswith("revertir-"):
            objetivo = v["version"][len("revertir-"):]
            if objetivo in pila:                          # revertir restaura el estado anterior a esa versión
                del pila[pila.index(objetivo):]
        else:
            pila.append(v["version"])
    return pila[-1] if pila else "base"


# ------------------------------------------------------------------ métricas

def _por_version(aprobados: list[dict], corr: list[dict]) -> list[dict]:
    """¿Sirvió cada versión de la biblioteca? Correcciones por informe aprobado con esa versión vigente."""
    por_informe = Counter((c["cliente"], c["periodo"]) for c in corr if c["tipo"] not in ("devolucion", "auditoria"))
    opiniones = Counter((c["cliente"], c["periodo"]) for c in corr if c["tipo"] == "opinion" and c.get("origen") != "claude")
    grupos = defaultdict(list)
    for e in aprobados:
        grupos[(e.get("aprobacion") or {}).get("version_biblioteca", "base")].append((e["cliente"], e["periodo"]))
    return [{"version": v, "informes": len(ks),
             "correcciones_por_informe": round(sum(por_informe[k] for k in ks) / len(ks), 2),
             "opiniones_editadas_por_informe": round(sum(opiniones[k] for k in ks) / len(ks), 2)}
            for v, ks in grupos.items()]


def _auditoria(esp: Espacio) -> dict | None:
    concl = [c for c in correcciones(esp) if c["tipo"] == "auditoria"]
    if not concl:
        return None
    n = Counter(c["conclusion"] for c in concl)
    return {"revisados": len(concl), "sin_observaciones": n.get("sin_observaciones", 0),
            "observaciones_menores": n.get("menores", 0), "requiere_reproceso": n.get("reproceso", 0)}


def _mapas(esp: Espacio) -> dict:
    """Aprendizaje de mapas del Excel: cuántos clientes tienen mapa aprobado, sinónimos aprendidos y ajustes."""
    from revisoria import mapa as mapa_cliente
    mapas = [leer_json(p, {}) for p in esp.clientes.glob("*/mapa.json")]
    sin = mapa_cliente.aprendidos(esp.conocimiento)
    return {"clientes": len(mapas), "aprobados": sum(1 for m in mapas if m.get("aprobado_por")),
            "sinonimos": sum(1 for a in sin.values() if a["clave"] != "AMBIGUO"),
            "ambiguos": sum(1 for a in sin.values() if a["clave"] == "AMBIGUO"),
            "ajustes": len(list((esp.raiz / "salidas").glob("*/*/mapa_ajustes.json")))}


def metricas(esp: Espacio) -> dict:
    estados = [leer_json(p, {}) for p in (esp.raiz / "salidas").glob("*/*/estado.json")]
    corr = correcciones(esp)
    aprobados = [e for e in estados if e.get("estado") == APROBADO]
    diap_total = sum(e.get("resumen", {}).get("diapositivas", 0) for e in aprobados)
    diap_corr = len({(c["cliente"], c["periodo"], c.get("diapositiva")) for c in corr
                     if c["tipo"] in ("cifra", "texto", "estructura") and c.get("diapositiva")})
    tiempos = []
    for e in aprobados:
        b = e.get("bitacora", [])
        t0 = next((x["fecha"] for x in b if x["estado"] == PENDIENTE), None)
        t1 = next((x["fecha"] for x in reversed(b) if x["estado"] == APROBADO), None)
        if t0 and t1:
            tiempos.append((datetime.fromisoformat(t1) - datetime.fromisoformat(t0)).total_seconds() / 3600)
    reaud = [e.get("resumen", {}).get("reauditoria", {}) for e in estados if e.get("resumen")]
    m = {
        "fecha": ahora(),
        "informes": len(estados), "aprobados": len(aprobados),
        "devueltos": sum(1 for e in estados if e.get("estado") == DEVUELTO),
        "pct_diapositivas_sin_correccion": round(1 - diap_corr / diap_total, 3) if diap_total else None,
        "correcciones_por_informe": round(len([c for c in corr if c["tipo"] not in ("devolucion", "auditoria")]) / len(aprobados), 2)
        if aprobados else None,
        "horas_validacion_promedio": round(sum(tiempos) / len(tiempos), 2) if tiempos else None,
        "hallazgos_reauditor_por_informe": round(sum(r.get("ALTA", 0) + r.get("MEDIA", 0) for r in reaud) / len(reaud), 2)
        if reaud else None,
        "correcciones_por_tipo": dict(Counter(c["tipo"] for c in corr)),
        "por_version": _por_version(aprobados, corr),
        "auditoria_mensual": _auditoria(esp),
        "mapas": _mapas(esp),
    }
    hist = leer_json(esp.conocimiento / "metricas.json", [])
    hist.append(m)
    guardar_json(esp.conocimiento / "metricas.json", hist[-60:])
    return m
