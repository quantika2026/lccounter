#!/usr/bin/env python3
"""Genera la sección financiera del informe de Revisoría Fiscal y su reporte de validación.

Uso:
  python generar_informe.py --excel "EXCEL PICASSO AGOSTO 2026 .xlsx" --config config/picasso.yaml
  python generar_informe.py --excel ... --config ... --revisar-pptx "Inf 006-2026 PICASSO AGOSTO 2026.pptx"
  python generar_informe.py --excel ... --config ... --periodo 2026-08 --llm

Salidas (carpeta --salida, por defecto ./salida):
  Seccion_Financiera_<copropiedad>_<periodo>.pptx   diapositivas con tablas y gráficos nativos
  Validacion_<copropiedad>_<periodo>.html           hallazgos del Excel (y del PPTX revisado)
  modelo_<periodo>.json                             todas las cifras con su celda fuente
  hallazgos_<periodo>.json                          hallazgos en formato máquina
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date
from pathlib import Path

import yaml

from revisoria.auditor_pptx import auditar
from revisoria.excel import Periodo, norm
from revisoria.extractor import extraer
from revisoria.generador_pptx import Informe
from revisoria.narrativa import Redactor
from revisoria.reporte import generar_html
from revisoria.validador import ordenar, validar_excel


def _slug(t: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", norm(t)).strip("_")[:40] or "copropiedad"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--excel", required=True, help="Excel de papeles de trabajo del periodo")
    ap.add_argument("--config", required=True, help="YAML de la copropiedad (meta y notas del revisor)")
    ap.add_argument("--periodo", help="AAAA-MM. Si se omite, se toma el mes más reciente del Excel")
    ap.add_argument("--revisar-pptx", help="Informe PPTX ya elaborado para auditar contra el Excel")
    ap.add_argument("--salida", default="salida")
    ap.add_argument("--llm", action="store_true", help="Pulir redacción con Claude (requiere ANTHROPIC_API_KEY)")
    ap.add_argument("--completo", action="store_true",
                    help="Informe completo con la estructura del modelo (regulatoria, financiera y control interno)")
    a = ap.parse_args(argv)

    cfg = yaml.safe_load(Path(a.config).read_text(encoding="utf-8")) or {}
    periodo = Periodo.desde_texto(a.periodo) if a.periodo else None
    meta = {k: cfg.get(k) for k in ("copropiedad", "nit", "unidades", "numero_informe") if cfg.get(k)}

    print("1/4  Leyendo Excel…")
    modelo, libro = extraer(a.excel, periodo, meta)
    p = Periodo.desde_texto(modelo["meta"]["periodo"])
    print(f"     Periodo: {p.etiqueta}")

    print("2/4  Validando…")
    hallazgos = validar_excel(modelo, libro, p)
    doc = None
    if a.completo:
        from revisoria.documental import extraer_documental, validar_documental
        doc = extraer_documental(libro, p)
        hallazgos = ordenar(hallazgos + validar_documental(doc, p))
    if a.revisar_pptx:
        hallazgos = ordenar(hallazgos + auditar(a.revisar_pptx, modelo, libro, p))

    print("3/4  Generando diapositivas…")
    redactor = Redactor(usar_llm=a.llm)
    if a.llm and not redactor.usar_llm:
        print("     Aviso: --llm pedido pero no hay ANTHROPIC_API_KEY o paquete 'anthropic'; se usan textos estándar.")
    salida = Path(a.salida)
    salida.mkdir(parents=True, exist_ok=True)
    slug = _slug(cfg.get("copropiedad", "copropiedad"))
    if a.completo:
        from revisoria.informe_completo import InformeCompleto
        ruta_pptx = salida / f"Informe_Revisoria_Fiscal_{slug}_{p.clave}.pptx"
        inf = InformeCompleto(modelo, cfg, doc, redactor).construir()
        for aviso in inf.avisos:
            print("     Aviso:", aviso)
    else:
        ruta_pptx = salida / f"Seccion_Financiera_{slug}_{p.clave}.pptx"
        inf = Informe(modelo, cfg, redactor).construir()
    inf.guardar(ruta_pptx)

    print("4/4  Escribiendo reporte…")
    archivos = {"excel": Path(a.excel).name, "pptx_revisado": Path(a.revisar_pptx).name if a.revisar_pptx else ""}
    ruta_html = salida / f"Validacion_{slug}_{p.clave}.html"
    ruta_html.write_text(generar_html(hallazgos, modelo["meta"], cfg, archivos), encoding="utf-8")

    def ser(o):
        if isinstance(o, date):
            return o.isoformat()
        if isinstance(o, (bytes, bytearray)):
            return f"<imagen {len(o)} bytes>"
        raise TypeError(type(o))
    (salida / f"modelo_{p.clave}.json").write_text(json.dumps(modelo, default=ser, ensure_ascii=False, indent=1),
                                                    encoding="utf-8")
    (salida / f"hallazgos_{p.clave}.json").write_text(
        json.dumps([h.dict() for h in hallazgos], ensure_ascii=False, indent=1), encoding="utf-8")

    from collections import Counter
    c = Counter(h.severidad for h in hallazgos)
    print(f"\nListo: {inf.n} diapositivas · hallazgos ALTA {c.get('ALTA', 0)} · MEDIA {c.get('MEDIA', 0)} · "
          f"BAJA {c.get('BAJA', 0)} · INFO {c.get('INFO', 0)}")
    print(f"  {ruta_pptx}\n  {ruta_html}")
    if redactor.registro:
        print("  Redacción con Claude:", Counter(r["estado"].split(":")[0] for r in redactor.registro))
    return 0


if __name__ == "__main__":
    sys.exit(main())
