#!/usr/bin/env python3
"""Demo: la coordinadora corrige el total digitado de CXP del cliente devuelto (vuelve a ser fórmula)."""
import sys
import warnings
warnings.filterwarnings("ignore")
import openpyxl
ruta = sys.argv[1]
wb = openpyxl.load_workbook(ruta, rich_text=True)
ws = next(w for w in wb.worksheets if w.title.strip().lower() == "cxp")
ws["D95"].value = "=+SUM(D85:D94)"
wb.save(ruta)
sys.path.insert(0, __file__.rsplit("/", 2)[0])
from demo.preparar_demo import recalcular  # noqa: E402
from pathlib import Path  # noqa: E402
recalcular(Path(ruta))
print("Total de CXP corregido en", ruta)
