"""Auditoría periódica del sistema.

- Por informe: la hace el re-auditor automáticamente (ver reauditor.py).
- Mensual: muestra aleatoria de 3–5 informes APROBADOS para que un profesional distinto al que validó los
  revise a fondo. Este módulo elige la muestra, vuelve a re-auditar el informe FINAL y arma la lista de chequeo.
- Trimestral: revisión humana de métricas y adaptadores (el panel trae los datos).
- Antes de cada versión: pruebas de regresión contra el golden set (ver mejora.regresion).
"""
from __future__ import annotations

import html
import random
import zlib
from pathlib import Path

from .comun import APROBADO, Espacio, ahora, guardar_json, leer_json

CHEQUEO = [
    "Las cifras de los títulos coinciden con el Excel del periodo",
    "Las tablas de cartera, cuentas por pagar y patrimonio cuadran con sus totales",
    "Las opiniones y recomendaciones son coherentes con las cifras",
    "Los hallazgos ALTA aprobados con justificación están bien sustentados",
    "El cuadre contra el balance de prueba quedó conciliado",
    "Los hallazgos abiertos del mes anterior tienen seguimiento",
    "El PDF entregado corresponde a la versión aprobada",
]


CONCLUSIONES = {"sin_observaciones": "Sin observaciones", "menores": "Observaciones menores",
                "reproceso": "Requiere reproceso"}


def registrar_conclusion(esp: Espacio, periodo: str, cid: str, auditor: str, conclusion: str,
                         observacion: str = "") -> dict:
    """Guarda la conclusión del segundo revisor y la lleva a la mejora continua."""
    from . import mejora
    if conclusion not in CONCLUSIONES:
        raise SystemExit(f"Conclusión inválida «{conclusion}». Use: {', '.join(CONCLUSIONES)}.")
    ruta = esp.conocimiento / "auditorias" / f"muestra_{periodo}.json"
    muestra = leer_json(ruta)
    if not muestra:
        raise SystemExit(f"No hay muestra del periodo {periodo}. Corra primero: muestra --periodo {periodo}")
    fila = next((f for f in muestra["muestra"] if f["cliente"] == cid), None)
    if not fila:
        raise SystemExit(f"{cid} no está en la muestra de {periodo}: "
                         + ", ".join(f["cliente"] for f in muestra["muestra"]))
    if auditor.strip().lower() == str(fila["aprobado_por"]).strip().lower():
        raise SystemExit(f"{auditor} aprobó este informe: la auditoría mensual la hace un profesional distinto.")
    if conclusion != "sin_observaciones" and not observacion.strip():
        raise SystemExit("Describa la observación con --observacion \"...\" (alimenta la mejora continua).")
    fila["revision"] = {"auditor": auditor, "conclusion": conclusion, "observacion": observacion, "fecha": ahora()}
    guardar_json(ruta, muestra)
    mejora.registrar(esp, periodo, cid, [{"tipo": "auditoria", "auditor": auditor, "aprobado_por": fila["aprobado_por"],
                                          "conclusion": conclusion, "observacion": observacion}])
    return fila


def muestra_mensual(esp: Espacio, periodo: str, k: int = 4, semilla: int | None = None) -> Path:
    from revisoria.excel import Periodo
    from revisoria.extractor import extraer
    from revisoria import mapa as mapa_cliente

    from . import reauditor
    aprobados = [e for e in (esp.estado(periodo, c) for c in esp.ids_clientes()) if e.get("estado") == APROBADO]
    rnd = random.Random(semilla if semilla is not None else zlib.crc32(periodo.encode()))   # reproducible
    elegidos = rnd.sample(aprobados, min(k, len(aprobados))) if aprobados else []
    filas = []
    for e in elegidos:
        cid = e["cliente"]
        sal = esp.salida(periodo, cid)
        final = sal / e["archivos"].get("final", e["archivos"]["pptx"])
        modelo, lib = extraer(e["archivos"]["excel"], Periodo.desde_texto(periodo),
                              {k: v for k, v in esp.ficha(cid).items() if k in ("copropiedad", "nit", "unidades", "numero_informe")},
                              mapa_cliente.cargar(esp.clientes / cid))
        p = Periodo.desde_texto(periodo)
        hs = reauditor.reauditar(final, modelo, lib, p, esp.ficha(cid), [], None)
        hs = [h for h in hs if h.regla in ("R01", "R02", "R03", "R04")]
        filas.append({"cliente": cid, "copropiedad": e.get("copropiedad"), "aprobado_por": e["aprobacion"]["revisor"],
                      "justificacion": e["aprobacion"].get("justificacion", ""), "hallazgos": [h.dict() for h in hs]})
    guardar_json(esp.conocimiento / "auditorias" / f"muestra_{periodo}.json",
                 {"fecha": ahora(), "periodo": periodo, "universo": len(aprobados), "muestra": filas})
    e_ = html.escape
    tarjetas = []
    for f in filas:
        hs = "".join(f"<li><b>{e_(h['regla'])}</b> {e_(h['titulo'])}: {e_(h['detalle'])}</li>" for h in f["hallazgos"]) \
            or "<li>Sin observaciones automáticas sobre el informe final.</li>"
        chk = "".join(f"<label><input type='checkbox'> {e_(c)}</label>" for c in CHEQUEO)
        tarjetas.append(f"""<section><h2>{e_(f['copropiedad'] or f['cliente'])}</h2>
<p class="sub">Aprobó: {e_(f['aprobado_por'])}{' · Justificación: ' + e_(f['justificacion']) if f['justificacion'] else ''}</p>
<h3>Re-auditoría del informe final</h3><ul>{hs}</ul><h3>Lista de chequeo del segundo revisor</h3><div class="chk">{chk}</div>
<h3>Registrar la conclusión</h3><p class="sub">Al terminar, registre el resultado para que alimente la mejora continua
(conclusión: sin_observaciones, menores o reproceso):</p>
<code>python3 -m agente --raiz "{e_(str(esp.raiz))}" auditar {e_(f['cliente'])} --periodo {periodo} --auditor "Nombre" --conclusion menores --observacion "..."</code></section>""")
    ruta = esp.raiz / f"Auditoria_Muestra_{periodo}.html"
    ruta.write_text(f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Auditoría mensual {periodo}</title><style>
body{{margin:0;font:15px/1.5 -apple-system,Segoe UI,Aptos,Roboto,sans-serif;color:#334155;background:#F6F7FA}}
main{{max-width:900px;margin:0 auto;padding:28px 16px 60px}} h1{{color:#002060}} h2{{color:#002060;margin:0}} h3{{font-size:15px;margin:16px 0 6px}}
section{{background:#fff;border:1px solid #D5DBE5;border-radius:10px;padding:18px;margin:16px 0}} .sub{{color:#64748B;margin:4px 0}}
.chk label{{display:block;margin:4px 0}}
code{{display:block;background:#F6F7FA;border:1px solid #D5DBE5;border-radius:4px;padding:6px 8px;font-size:12px;overflow-x:auto}}</style></head><body><main>
<h1>Auditoría mensual por muestreo · {periodo}</h1>
<p class="sub">{len(filas)} de {len(aprobados)} informes aprobados, elegidos al azar. Debe revisarlos un profesional distinto al que los validó.</p>
{''.join(tarjetas) or '<p>No hay informes aprobados en el periodo.</p>'}
</main></body></html>""", encoding="utf-8")
    return ruta
