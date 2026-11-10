"""Reporte de validación en HTML autocontenido (se abre en cualquier navegador)."""
from __future__ import annotations

import html
from collections import Counter
from datetime import datetime

COLORES = {"ALTA": "#C01E1E", "MEDIA": "#C77700", "BAJA": "#2F5FA8", "INFO": "#64748B"}


def generar_html(hallazgos, meta: dict, cfg: dict, archivos: dict) -> str:
    c = Counter(h.severidad for h in hallazgos)
    e = html.escape

    def tarjeta(h):
        ubic = "".join(f"<code>{e(u)}</code>" for u in h.ubicacion)
        rec = f"<p class='rec'><b>Recomendación:</b> {e(h.recomendacion)}</p>" if h.recomendacion else ""
        return f"""
        <article class="h" data-sev="{h.severidad}" data-origen="{h.origen}">
          <header><span class="sev" style="background:{COLORES[h.severidad]}">{h.severidad}</span>
          <span class="regla">{h.regla} · {e(h.categoria)}</span></header>
          <h3>{e(h.titulo)}</h3>
          <p>{e(h.detalle)}</p>{rec}
          <div class="ubic">{ubic}</div>
        </article>"""

    excel = [h for h in hallazgos if h.origen == "excel"]
    pptx = [h for h in hallazgos if h.origen == "pptx"]
    seccion_pptx = ""
    if pptx:
        seccion_pptx = (f"<h2>Informe PPTX revisado <small>{e(archivos.get('pptx_revisado', ''))}</small></h2>"
                        + "".join(tarjeta(h) for h in pptx))
    resumen = "".join(
        f"<div class='kpi'><span style='color:{COLORES[s]}'>{c.get(s, 0)}</span><small>{s}</small></div>"
        for s in ("ALTA", "MEDIA", "BAJA", "INFO"))
    return f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Validación {e(cfg.get('copropiedad', ''))} {e(meta['periodo_etiqueta'])}</title>
<style>
:root{{--navy:#002060;--txt:#334155;--gris:#64748B;--borde:#D5DBE5;--fondo:#F6F7FA}}
*{{box-sizing:border-box}} body{{margin:0;font:15px/1.5 -apple-system,Segoe UI,Aptos,Roboto,sans-serif;color:var(--txt);background:var(--fondo)}}
main{{max-width:960px;margin:0 auto;padding:28px 16px 60px}}
h1{{color:var(--navy);margin:0 0 4px;font-size:26px}} h2{{color:var(--navy);margin:34px 0 12px;font-size:19px}}
h2 small{{color:var(--gris);font-weight:400;font-size:13px}}
.sub{{color:var(--gris);margin:0 0 20px}}
.kpis{{display:flex;gap:12px;flex-wrap:wrap;margin:18px 0}}
.kpi{{background:#fff;border:1px solid var(--borde);border-radius:10px;padding:10px 18px;min-width:110px}}
.kpi span{{font-size:28px;font-weight:700;display:block}} .kpi small{{color:var(--gris);letter-spacing:.05em}}
.filtros{{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px}}
.filtros button{{border:1px solid var(--borde);background:#fff;border-radius:999px;padding:5px 14px;cursor:pointer;font:inherit;font-size:13px}}
.filtros button.on{{background:var(--navy);color:#fff;border-color:var(--navy)}}
.h{{background:#fff;border:1px solid var(--borde);border-radius:10px;padding:14px 18px;margin:10px 0}}
.h header{{display:flex;gap:10px;align-items:center}} .sev{{color:#fff;font-size:11px;font-weight:700;padding:2px 8px;border-radius:4px}}
.regla{{color:var(--gris);font-size:12.5px}} .h h3{{margin:8px 0 4px;font-size:16px;color:#1F2937}} .h p{{margin:4px 0}}
.rec{{color:#1F2937}} .ubic code{{display:inline-block;background:var(--fondo);border:1px solid var(--borde);border-radius:4px;padding:1px 6px;margin:6px 6px 0 0;font-size:12px;word-break:break-all}}
footer{{color:var(--gris);font-size:12.5px;margin-top:30px}}
</style></head><body><main>
<h1>Validación del informe de Revisoría Fiscal</h1>
<p class="sub">{e(cfg.get('copropiedad', ''))} · Periodo {e(meta['periodo_etiqueta'])} · Excel: {e(archivos.get('excel', ''))}</p>
<div class="kpis">{resumen}</div>
<div class="filtros"><button class="on" data-f="TODAS">Todas</button><button data-f="ALTA">Alta</button>
<button data-f="MEDIA">Media</button><button data-f="BAJA">Baja</button><button data-f="INFO">Info</button></div>
<h2>Excel de papeles de trabajo</h2>
{''.join(tarjeta(h) for h in excel) or '<p>Sin hallazgos.</p>'}
{seccion_pptx}
<footer>Generado el {datetime.now():%d/%m/%Y %H:%M}. El validador señala; la conclusión es del revisor fiscal.</footer>
</main><script>
document.querySelectorAll('.filtros button').forEach(b=>b.onclick=()=>{{
 document.querySelectorAll('.filtros button').forEach(x=>x.classList.toggle('on',x===b));
 const f=b.dataset.f; document.querySelectorAll('.h').forEach(h=>h.style.display=(f==='TODAS'||h.dataset.sev===f)?'':'none');
}});
</script></body></html>"""
