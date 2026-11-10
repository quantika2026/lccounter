#!/usr/bin/env python3
"""Arma un espacio de trabajo de demostración con 4 clientes para el periodo 2026-08.

  python demo/preparar_demo.py --excel "<EXCEL PICASSO AGOSTO 2026 .xlsx>" [--destino demo/trabajo]

Clientes:
  picasso   Caso real (Excel de Picasso). Balance de prueba SIMULADO a partir del Excel (cuadra).
  mirador   FICTICIO: Excel de Picasso con cifras escaladas ×0,62. Software «contable_x» con perfil aprobado.
  arboleda  FICTICIO: cifras ×1,35, un valor alterado en CXP (error sembrado) y balance de un software
            sin perfil aprobado («contable_y»): el normalizador propone el mapeo y no cuadra hasta aprobarlo.
  sanrafael FICTICIO: no entregó el Excel del mes → BLOQUEADO, sin detener el lote.

Los clientes ficticios existen solo para mostrar el flujo por lote; sus cifras no son reales.
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import warnings
from pathlib import Path

import openpyxl
import yaml

warnings.filterwarnings("ignore")
RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

PERIODO = "2026-08"
import re
RX_CONST = re.compile(r"(?<![A-Za-z$:!\d.])\d{4,}(?:\.\d+)?(?![\d:!])")
HOJAS = {"cuadro efectivo", "grafica flujo de efectivo", "grafica deudores propietari", "rangos de cartera",
         "mayores deudores", "may deudores", "antic. e ingresos x c", "cxp", "cuadro reserva", "patrimonio",
         "presupuesto", "conciliaciones", "impuestos", "caja menor"}

PERFILES = {
    "tolerancia": 1000,
    "perfiles": {
        "contable_x": {"descripcion": "Software contable X (demo): CSV con ';' y coma decimal",
                       "columnas": {"cuenta": ["Código"], "nombre": ["Descripción"], "saldo": ["Saldo Final"]},
                       "decimal": ",", "estado": "aprobado"},
    },
}


def recalcular(ruta: Path):
    script = Path("/mnt/skills/public/xlsx/scripts/recalc.py")
    if script.exists():
        subprocess.run([sys.executable, str(script), str(ruta), "120"], capture_output=True)
        return
    from revisoria import libreoffice
    if not libreoffice.buscar():
        sys.exit("No se encontró LibreOffice (soffice): se necesita para recalcular los Excel de la demo.")
    out = ruta.parent / "_recalc"                  # recálculo vía conversión xlsx → xlsx
    recalculado = libreoffice.convertir(ruta, "xlsx", out, timeout=120)
    if not recalculado:
        sys.exit(f"LibreOffice no pudo recalcular {ruta.name}.")
    shutil.move(recalculado, ruta)
    shutil.rmtree(out, ignore_errors=True)


def escalar(origen: Path, destino: Path, factor: float, alterar: tuple | None = None):
    wb = openpyxl.load_workbook(origen, rich_text=True)
    for ws in wb.worksheets:
        if ws.title.strip().lower() not in HOJAS:
            continue
        for fila in ws.iter_rows():
            for c in fila:
                v = c.value
                if c.data_type == "n" and isinstance(v, (int, float)) and not isinstance(v, bool) \
                        and abs(v) >= 1000 and not (1900 <= v <= 2100 and float(v).is_integer()):
                    c.value = round(v * factor)
                elif isinstance(v, str) and v.startswith("="):       # constantes digitadas dentro de fórmulas
                    c.value = RX_CONST.sub(lambda mm: str(round(float(mm.group(0)) * factor)), v)
    if alterar:
        alterar(wb)
    wb.save(destino)
    recalcular(destino)


def sembrar_error(wb):
    """Error sembrado para la demo: el total de CXP queda digitado y $2.500.000 por encima de la suma."""
    ws = next(w for w in wb.worksheets if w.title.strip().lower() == "cxp")
    total = 0
    for r in range(85, 95):
        v = ws.cell(r, 4).value
        total += eval(v.lstrip("=+")) if isinstance(v, str) else (v or 0)
    ws["D95"].value = round(total) + 2_500_000


def _por(v, k):
    return v * k if isinstance(v, (int, float)) else None


def balance_desde_modelo(modelo: dict, estilo: str) -> str:
    """Balance de prueba SIMULADO con las cuentas que se cuadran (para la demostración)."""
    filas = [("1330", "Anticipos y avances", modelo["anticipos"]["total_hoja"]),
             ("1380", "Deudores varios", modelo["deudores_varios"]["total_hoja"]),
             ("2335", "Costos y gastos por pagar", _por(modelo["cxp"]["total"], 0.8)),
             ("2365", "Retención en la fuente", _por(modelo["cxp"]["total"], 0.2)),
             ("4", "Ingresos", _por(modelo["patrimonio"]["ingresos"], -1)),
             ("5", "Gastos", _por(modelo["patrimonio"]["gastos"], -1))]
    # una cifra que el Excel no trae (None) no se inventa en el balance simulado: la cuenta se omite
    filas = [(c, n, v) for c, n, v in filas if isinstance(v, (int, float))]
    if estilo == "generico":
        return "cuenta,nombre,saldo final\n" + "\n".join(f"{c},{n},{v:.2f}" for c, n, v in filas)
    if estilo == "contable_x":
        def co(v):
            return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return "Código;Descripción;Saldo Inicial;Saldo Final\n" + "\n".join(f"{c};{n};0;{co(v)}" for c, n, v in filas)
    # contable_y: formato que el sistema no conoce
    return "CTA PUC\tNOMBRE DE LA CUENTA\tMOV DEBITO\tMOV CREDITO\tSALDO A LA FECHA\n" + "\n".join(
        f"{c}\t{n}\t0\t0\t{v:.0f}" for c, n, v in filas)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    ap.add_argument("--destino", default=str(RAIZ / "demo" / "trabajo"))
    ap.add_argument("--config", default=str(RAIZ / "config" / "picasso.yaml"))
    a = ap.parse_args()
    from revisoria.excel import Periodo
    from revisoria.extractor import extraer

    dst = Path(a.destino)
    if dst.exists():
        shutil.rmtree(dst)
    base = yaml.safe_load(Path(a.config).read_text(encoding="utf-8"))
    (dst / "conocimiento").mkdir(parents=True)
    (dst / "conocimiento" / "perfiles_software.yaml").write_text(
        "# Perfiles de exporte por software contable. Solo se usan los que tienen estado: aprobado.\n"
        + yaml.safe_dump(PERFILES, allow_unicode=True, sort_keys=False), encoding="utf-8")

    clientes = {
        "picasso": dict(base, revisor="Revisora fiscal A", orden=1, software="generico", factor=None),
        "mirador": {"copropiedad": "Edificio Mirador del Parque PH (DEMO)", "nit": "900.000.001-1", "unidades": 408,
                    "firma": base["firma"], "numero_informe": "014-2026", "revisor": "Revisora fiscal A", "orden": 2,
                    "software": "contable_x", "factor": 0.62,
                    "notas": {"concentracion_cartera": "La administración informa que la abogada radicó tres demandas en el mes."}},
        "arboleda": {"copropiedad": "Conjunto Arboleda Real PH (DEMO)", "nit": "900.000.002-2", "unidades": 408,
                     "firma": base["firma"], "numero_informe": "021-2026", "revisor": "Revisor fiscal B", "orden": 3,
                     "software": "contable_y", "factor": 1.35, "notas": {}},
        "sanrafael": {"copropiedad": "Torres de San Rafael PH (DEMO)", "nit": "900.000.003-3", "unidades": 120,
                      "firma": base["firma"], "numero_informe": "008-2026", "revisor": "Revisor fiscal B", "orden": 4,
                      "software": "generico", "factor": "sin_excel", "notas": {}},
    }
    for cid, f in clientes.items():
        factor = f.pop("factor")
        f["entradas"] = {"excel": "*.xlsx", "balance": {"archivo": "balance*", "perfil": f["software"]}}
        (dst / "clientes" / cid).mkdir(parents=True)
        (dst / "clientes" / cid / "ficha.yaml").write_text(
            yaml.safe_dump(f, allow_unicode=True, sort_keys=False, width=110), encoding="utf-8")
        ent = dst / "entradas" / PERIODO / cid
        ent.mkdir(parents=True)
        if factor == "sin_excel":
            continue
        excel = ent / f"Excel_{cid}_{PERIODO}.xlsx"
        if factor is None:
            shutil.copy2(a.excel, excel)
        else:
            alterar = sembrar_error if cid == "arboleda" else None
            escalar(Path(a.excel), excel, factor, alterar)
        modelo, _ = extraer(str(excel), Periodo.desde_texto(PERIODO), {})
        modelo = json.loads(json.dumps(modelo, default=str))
        (ent / f"balance_prueba_{PERIODO}.{'tsv' if f['software'] == 'contable_y' else 'csv'}").write_text(
            balance_desde_modelo(modelo, f["software"]), encoding="utf-8")
        print(f"  {cid:<10} {excel.name}")
    print(f"Espacio de demostración listo en {dst}")


if __name__ == "__main__":
    main()
