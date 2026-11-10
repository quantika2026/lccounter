"""Panel del lote: estado de cada cliente, cola de validación, hallazgos y mejora continua."""
from __future__ import annotations

import html
from collections import Counter
from datetime import datetime

from revisoria.narrativa import pesos

from .comun import APROBADO, BLOQUEADO, DEVUELTO, PENDIENTE, Espacio, leer_json

ETIQUETAS = {PENDIENTE: ("Por validar", "#C77700"), APROBADO: ("Aprobado", "#1E7B3A"),
             DEVUELTO: ("Devuelto", "#C01E1E"), BLOQUEADO: ("Bloqueado", "#64748B"), "RECIBIDO": ("Recibido", "#64748B")}


def _chip(estado):
    t, c = ETIQUETAS.get(estado, (estado, "#64748B"))
    return f"<span class='chip' style='background:{c}'>{t}</span>"


def _sev(d):
    if not d:
        return "—"
    return (f"<span class='a'>{d.get('ALTA', 0)}</span> / <span class='m'>{d.get('MEDIA', 0)}</span>")


def generar(esp: Espacio, periodo: str) -> str:
    e = html.escape
    estados = [esp.estado(periodo, c) | {"cliente": c} for c in esp.ids_clientes()]
    cnt = Counter(x.get("estado") for x in estados)
    kpis = [("Clientes", len(estados), "#002060"), ("Por validar", cnt.get(PENDIENTE, 0), "#C77700"),
            ("Aprobados", cnt.get(APROBADO, 0), "#1E7B3A"), ("Devueltos", cnt.get(DEVUELTO, 0), "#C01E1E"),
            ("Bloqueados", cnt.get(BLOQUEADO, 0), "#64748B")]
    kpi_html = "".join(f"<div class='kpi'><span style='color:{c}'>{v}</span><small>{t}</small></div>" for t, v, c in kpis)

    cola = [x for x in estados if x.get("estado") == PENDIENTE]
    cola_html = "".join(
        f"<li><b>{i}.</b> {e(x.get('copropiedad', x['cliente']))} <small>· {e(x.get('revisor', ''))}</small>"
        f"<code>python -m agente aprobar {x['cliente']} --periodo {periodo} --revisor \"…\"</code></li>"
        for i, x in enumerate(cola, 1)) or "<li>No hay informes esperando validación.</li>"

    filas = []
    for x in estados:
        r = x.get("resumen") or {}
        a = x.get("archivos") or {}
        base = f"salidas/{periodo}/{x['cliente']}/"
        links = " · ".join(f"<a href='{base}{a[k]}'>{t}</a>" for k, t in
                           (("validacion", "Validación"), ("pptx", "PPTX"), ("borrador", "Borrador"),
                            ("final", "Final"), ("pdf", "PDF")) if a.get(k))
        cuadre = r.get("cuadre") or []
        evaluados = [c for c in cuadre if c["ok"] is not None]
        cuadre_txt = (f"{sum(1 for c in evaluados if c['ok'])}/{len(evaluados)}" if evaluados else "—")
        cif = r.get("cifras") or {}
        ult = (x.get("bitacora") or [{}])[-1]
        filas.append(f"""<tr><td><b>{e(x.get('copropiedad', x['cliente']))}</b><br><small>{e(x.get('software', ''))} · {e(x.get('revisor', ''))}</small></td>
<td>{_chip(x.get('estado'))}</td><td class='n'>{pesos(cif['cartera']) if cif.get('cartera') is not None else '—'}</td>
<td class='n'>{pesos(cif['resultado']) if cif.get('resultado') is not None else '—'}</td>
<td class='c'>{_sev(r.get('hallazgos'))}</td><td class='c'>{cuadre_txt}</td><td class='c'>{_sev(r.get('reauditoria'))}</td>
<td>{links or '—'}<br><small>{e(ult.get('detalle', ''))[:90]}</small></td></tr>""")

    # Mejora continua
    met = (leer_json(esp.conocimiento / "metricas.json", []) or [{}])[-1]
    props = [p for p in leer_json(esp.conocimiento / "propuestas.json", []) if p.get("estado") == "pendiente"]
    vers = leer_json(esp.conocimiento / "versiones.json", [])
    reg = leer_json(esp.conocimiento / "ultima_regresion.json", {})
    golden = len(list((esp.conocimiento / "golden").glob("*_*"))) if (esp.conocimiento / "golden").exists() else 0

    def fmt(v, pctj=False):
        if v is None:
            return "—"
        return f"{v * 100:.0f}%" if pctj else str(v)
    met_html = "".join(f"<div class='kpi'><span>{v}</span><small>{t}</small></div>" for t, v in [
        ("Diapositivas sin corrección", fmt(met.get("pct_diapositivas_sin_correccion"), True)),
        ("Correcciones por informe", fmt(met.get("correcciones_por_informe"))),
        ("Horas de validación (prom.)", fmt(met.get("horas_validacion_promedio"))),
        ("Hallazgos re-auditor / informe", fmt(met.get("hallazgos_reauditor_por_informe"))),
        ("Golden set", golden)])
    props_html = "".join(f"<li><b>{e(p['id'])}</b> {e(p['titulo'])}<br><small>{e(p.get('evidencia', ''))} · riesgo {e(p.get('riesgo', ''))}</small></li>"
                         for p in props) or "<li>Sin propuestas pendientes.</li>"
    vers_html = "".join(f"<li>{e(v['version'])} · {e(v['fecha'][:16])} · {e(v.get('aprobado_por', ''))} · {e(v.get('regresion', ''))}</li>"
                        for v in vers[-5:]) or "<li>Sin versiones publicadas.</li>"
    porv_html = "".join(f"<li><b>{e(v['version'])}</b> · {v['informes']} informe(s) · {v['correcciones_por_informe']} "
                        f"correcciones y {v['opiniones_editadas_por_informe']} opiniones editadas por informe</li>"
                        for v in met.get("por_version") or []) or "<li>Aún no hay informes aprobados.</li>"
    aud = met.get("auditoria_mensual")
    aud_html = (f"{aud['revisados']} revisado(s): {aud['sin_observaciones']} sin observaciones, "
                f"{aud['observaciones_menores']} con observaciones menores, {aud['requiere_reproceso']} requieren reproceso"
                if aud else "Sin conclusiones registradas.")
    reg_html = (f"{'OK' if reg.get('ok') else 'FALLÓ'} · {reg.get('n', 0)} caso(s) · {reg.get('fecha', '')[:16]}" if reg else "No se ha corrido")

    return f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Panel Revisoría Fiscal {periodo}</title><style>
:root{{--navy:#002060;--txt:#334155;--gris:#64748B;--borde:#D5DBE5;--fondo:#F6F7FA}}
*{{box-sizing:border-box}} body{{margin:0;font:14.5px/1.5 -apple-system,Segoe UI,Aptos,Roboto,sans-serif;color:var(--txt);background:var(--fondo)}}
main{{max-width:1180px;margin:0 auto;padding:26px 16px 60px}} h1{{color:var(--navy);margin:0;font-size:25px}} h2{{color:var(--navy);font-size:18px;margin:30px 0 10px}}
.sub{{color:var(--gris);margin:4px 0 0}} .kpis{{display:flex;gap:12px;flex-wrap:wrap;margin:18px 0}}
.kpi{{background:#fff;border:1px solid var(--borde);border-radius:10px;padding:10px 16px;min-width:120px;flex:1}}
.kpi span{{display:block;font-size:26px;font-weight:700;color:var(--navy)}} .kpi small{{color:var(--gris)}}
.tabla{{overflow-x:auto;background:#fff;border:1px solid var(--borde);border-radius:10px}}
table{{border-collapse:collapse;width:100%;min-width:900px}} th{{background:var(--navy);color:#fff;text-align:left;font-weight:600;padding:9px 10px;font-size:13px}}
td{{border-top:1px solid var(--borde);padding:9px 10px;vertical-align:top}} td.n{{text-align:right;white-space:nowrap}} td.c{{text-align:center;white-space:nowrap}}
.chip{{color:#fff;border-radius:999px;padding:2px 10px;font-size:12px;font-weight:600;white-space:nowrap}}
.a{{color:#C01E1E;font-weight:700}} .m{{color:#C77700;font-weight:700}} small{{color:var(--gris)}} a{{color:#2F5FA8}}
.cols{{display:grid;grid-template-columns:1fr 1fr;gap:16px}} .caja{{background:#fff;border:1px solid var(--borde);border-radius:10px;padding:14px 18px}}
.caja h3{{margin:0 0 8px;color:var(--navy);font-size:15px}} ul{{margin:0;padding-left:18px}} li{{margin:6px 0}}
code{{display:block;background:var(--fondo);border:1px solid var(--borde);border-radius:4px;padding:2px 6px;margin-top:4px;font-size:12px;overflow-x:auto}}
.principio{{background:#E8EEF8;border-left:4px solid var(--navy);padding:10px 14px;border-radius:6px;margin-top:16px}}
@media (max-width:760px){{.cols{{grid-template-columns:1fr}}}}
</style></head><body><main>
<h1>Revisoría Fiscal · Panel del lote {periodo}</h1>
<p class="sub">Actualizado {datetime.now():%d/%m/%Y %H:%M} · Nada sale sin la aprobación del revisor fiscal.</p>
<div class="kpis">{kpi_html}</div>
<h2>Cola de validación del profesional</h2><div class="caja"><ul>{cola_html}</ul></div>
<h2>Clientes</h2>
<div class="tabla"><table><tr><th>Copropiedad</th><th>Estado</th><th>Cartera</th><th>Resultado</th><th>Excel A/M</th><th>Cuadre balance</th><th>Re-auditor A/M</th><th>Archivos y último evento</th></tr>
{''.join(filas)}</table></div>
<h2>Mejora continua</h2>
<div class="kpis">{met_html}</div>
<div class="cols"><div class="caja"><h3>Propuestas de mejora pendientes</h3><ul>{props_html}</ul></div>
<div class="caja"><h3>Versiones y prueba de regresión</h3><p>Última regresión: {e(reg_html)}</p><ul>{vers_html}</ul></div>
<div class="caja"><h3>¿Sirvió cada versión de la biblioteca?</h3><ul>{porv_html}</ul></div>
<div class="caja"><h3>Auditoría mensual (segundo revisor)</h3><p>{e(aud_html)}</p></div></div>
<p class="principio"><b>Python calcula, la IA organiza y redacta, el profesional decide.</b> Las cifras salen del Excel; los textos con IA pasan por un candado de cifras; cada corrección del revisor alimenta la mejora continua.</p>
</main></body></html>"""


def escribir(esp: Espacio, periodo: str):
    ruta = esp.raiz / f"Panel_{periodo}.html"
    ruta.write_text(generar(esp, periodo), encoding="utf-8")
    return ruta
