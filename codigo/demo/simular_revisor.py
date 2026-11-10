#!/usr/bin/env python3
"""Simula las ediciones que haría la revisora antes de aprobar (solo para la demostración).

  python demo/simular_revisor.py <carpeta_salida_cliente>

Deja en la carpeta:  Informe_editado.pptx  y  Borrador_editado.md
- En el PPTX cambia la redacción del subtítulo de una diapositiva (corrección de TEXTO).
- En el borrador reescribe la recomendación de concentración de cartera con su propio estilo (OPINIÓN).
"""
import re
import sys
from pathlib import Path

from pptx import Presentation

d = Path(sys.argv[1])
pptx = next(x for x in sorted(d.glob("Informe_*.pptx")) if not x.name.startswith("Informe_Final"))
prs = Presentation(pptx)
hecho = False
for s in prs.slides:
    for sh in s.shapes:
        if sh.has_text_frame and "Principales acreedores" in sh.text_frame.text and not hecho:
            r = sh.text_frame.paragraphs[0].runs[0]
            r.text = r.text.replace("Principales acreedores", "Acreedores con saldo")
            hecho = True
prs.save(d / "Informe_editado.pptx")

md = next(d.glob("Borrador_Opinion_*.md")).read_text(encoding="utf-8")
md = re.sub(r"(## \d+\. Concentración de la cartera en pocos inmuebles\n\*\*Nivel de riesgo propuesto:\*\* \w+\n\n)(.*?)(\n\n)",
            lambda m: m.group(1) + re.sub(r"Se recomienda .*", "En opinión de la Revisoría Fiscal, la administración debe "
                                          "presentar al Consejo, en la próxima sesión, el estado de cada proceso jurídico de "
                                          "estos inmuebles y el plan de cobro con fechas y responsables.", m.group(2)) + m.group(3),
            md, flags=re.S)
(d / "Borrador_editado.md").write_text(md, encoding="utf-8")
print("Ediciones simuladas:", d / "Informe_editado.pptx", d / "Borrador_editado.md")
