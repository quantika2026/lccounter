"""Utilidades de lectura del Excel de papeles de trabajo.

La búsqueda se hace por ETIQUETAS (texto de la fila, título del bloque, mes del
encabezado) y no por coordenadas fijas, para que el sistema tolere que las filas
se muevan de un mes a otro.
"""
from __future__ import annotations

import re
import unicodedata
import warnings
from dataclasses import dataclass
from datetime import date, datetime

import openpyxl
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", module="openpyxl")

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
MESES_CORTOS = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
ERRORES_EXCEL = ("#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!")


def norm(texto) -> str:
    """Minúsculas, sin tildes, espacios colapsados."""
    if texto is None:
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().lower()


def es_numero(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


@dataclass(frozen=True)
class Periodo:
    anio: int
    mes: int

    @classmethod
    def desde_texto(cls, texto: str) -> "Periodo":
        a, m = texto.split("-")
        return cls(int(a), int(m))

    @property
    def clave(self) -> str:
        return f"{self.anio:04d}-{self.mes:02d}"

    @property
    def nombre_mes(self) -> str:
        return MESES[self.mes - 1]

    @property
    def etiqueta(self) -> str:          # "Agosto 2026"
        return f"{self.nombre_mes.capitalize()} {self.anio}"

    @property
    def corto(self) -> str:             # "ago-26"
        return f"{MESES_CORTOS[self.mes - 1]}-{str(self.anio)[2:]}"

    def anterior(self) -> "Periodo":
        return Periodo(self.anio - 1, 12) if self.mes == 1 else Periodo(self.anio, self.mes - 1)

    def cierre_anterior(self) -> "Periodo":
        return Periodo(self.anio - 1, 12)

    @property
    def fin(self) -> date:
        sig = Periodo(self.anio + 1, 1) if self.mes == 12 else Periodo(self.anio, self.mes + 1)
        return date.fromordinal(date(sig.anio, sig.mes, 1).toordinal() - 1)


def periodo_de(v) -> Periodo | None:
    if isinstance(v, (datetime, date)):
        return Periodo(v.year, v.month)
    return None


def buscar_hoja(titulos: list[str], nombre: str) -> str | None:
    """Título de la hoja que corresponde a `nombre`: exacta, por prefijo o por palabras significativas
    («GRAFICA DE DEUDORES PROPIETARIO» por «GRAFICA DEUDORES», «ACTAS CONSEJO» por «ACTAS DE CONSEJO 2026»,
    «EGRESOS» por «EGRESO»). Ante varias, la más parecida (menos palabras extra)."""
    objetivo = norm(nombre)
    for t in titulos:
        if norm(t) == objetivo:
            return t
    for t in titulos:
        if norm(t).startswith(objetivo):
            return t

    def palabras(x):
        # palabras de 1–2 letras («e», «y», «x») no identifican la hoja: «EGRESO» no debe calzar con
        # «FIRMAS E INF DE GESTIÓN» por la «e»
        return [re.sub(r"(es|s)$", "", w) for w in norm(x).split()
                if len(w) > 2 and w not in ("del", "las", "los") and not re.fullmatch(r"\d{4}|\(\d+\)", w)]
    buscadas = palabras(nombre)
    candidatas = [(len(palabras(t)), t) for t in titulos
                  if buscadas and all(any(w.startswith(b) or b.startswith(w) for w in palabras(t)) for b in buscadas)]
    return min(candidatas)[1] if candidatas else None


class Libro:
    """Envuelve el libro con valores calculados y con fórmulas."""

    def __init__(self, ruta: str, mapa: dict | None = None):
        self.ruta = ruta
        self.mapa = mapa or {}          # mapa del cliente (revisoria/mapa.py): nombre lógico → título real
        self.desajustes: list[str] = []  # entradas del mapa que ya no calzan con este Excel
        self.valores = openpyxl.load_workbook(ruta, data_only=True)
        self.formulas = openpyxl.load_workbook(ruta, data_only=False)

    # ------------------------------------------------------------------ hojas
    def hoja(self, nombre: str):
        """Hoja por su nombre lógico. Si el cliente tiene mapa, manda el mapa; si no, se busca tolerando espacios,
        mayúsculas, tildes y variantes entre copropiedades."""
        hojas = self.mapa.get("hojas", {})
        if nombre in hojas:
            return self.hoja_mapeada(nombre, hojas[nombre])
        titulo = buscar_hoja([ws.title for ws in self.valores.worksheets], nombre)
        if titulo is None:
            raise KeyError(f"No se encontró la hoja '{nombre}' en {self.ruta}")
        return self.valores[titulo]

    def hoja_mapeada(self, nombre: str, titulo):
        if titulo is None:
            raise KeyError(f"Según el mapa del cliente, no tiene la hoja «{nombre}»")
        if titulo not in self.valores.sheetnames:
            self.desajustes.append(f"El mapa apunta a la hoja «{titulo}» para «{nombre}», pero no está en el Excel")
            raise KeyError(f"Mapa desactualizado: no está la hoja «{titulo}» («{nombre}»)")
        return self.valores[titulo]

    def resolver(self, nombre: str, detectar):
        """Hojas que no se ubican por nombre (p. ej. varias hojas de auditoría): usa el mapa si lo tiene,
        si no `detectar()`. El mapa guarda títulos (o listas de títulos)."""
        hojas = self.mapa.get("hojas", {})
        return hojas[nombre] if nombre in hojas else detectar()

    @property
    def ricos(self):
        """Libro cargado con texto enriquecido (negrillas y colores por fragmento). Se carga al pedirlo."""
        if not hasattr(self, "_ricos"):
            self._ricos = openpyxl.load_workbook(self.ruta, rich_text=True)
        return self._ricos

    def segmentos(self, ws, coord: str) -> list[dict] | None:
        """Fragmentos con formato de una celda: [{'t': texto, 'b': negrilla, 'rojo': bool}].
        None si la celda no tiene énfasis (todo el texto con el mismo formato normal)."""
        from openpyxl.cell.rich_text import CellRichText
        c = self.ricos[ws.title][coord]
        # celda combinada: el formato está en la celda de origen
        for rng in self.ricos[ws.title].merged_cells.ranges:
            if c.coordinate in rng:
                c = self.ricos[ws.title].cell(rng.min_row, rng.min_col)
                break

        def rojo(font):
            try:
                rgb = font.color.rgb if font is not None and font.color is not None else None
                return isinstance(rgb, str) and rgb[-6:].upper() in ("FF0000", "C00000", "E00000", "FF0101")
            except Exception:
                return False

        segs = []
        if isinstance(c.value, CellRichText):
            for t in c.value:
                if hasattr(t, "text"):
                    f = t.font
                    segs.append({"t": str(t.text), "b": bool(f is not None and f.b), "rojo": rojo(f)})
                else:
                    segs.append({"t": str(t), "b": bool(c.font and c.font.b), "rojo": rojo(c.font)})
        elif isinstance(c.value, str):
            segs.append({"t": c.value, "b": bool(c.font and c.font.b), "rojo": rojo(c.font)})
        else:
            return None
        if not any(x["b"] or x["rojo"] for x in segs):
            return None
        # une fragmentos contiguos con el mismo formato
        unidos = []
        for x in segs:
            if unidos and unidos[-1]["b"] == x["b"] and unidos[-1]["rojo"] == x["rojo"]:
                unidos[-1]["t"] += x["t"]
            else:
                unidos.append(dict(x))
        return unidos

    def hoja_formulas(self, ws):
        return self.formulas[ws.title]

    def formula(self, ws, coord: str):
        v = self.hoja_formulas(ws)[coord].value
        return v if isinstance(v, str) and v.startswith("=") else None

    def formulas_sin_resultado(self) -> list[tuple[str, str, str]]:
        """Fórmulas cuyo valor calculado no fue guardado por Excel/LibreOffice."""
        faltan = []
        for wf in self.formulas.worksheets:
            if wf.title not in self.valores.sheetnames:
                continue
            wv = self.valores[wf.title]
            for fila in wf.iter_rows():
                for c in fila:
                    if isinstance(c.value, str) and c.value.startswith("=") and wv[c.coordinate].value is None:
                        faltan.append((wf.title, c.coordinate, c.value))
        return faltan


# --------------------------------------------------------------- búsquedas

def buscar_fila(ws, etiqueta: str, columnas=("A", "B"), desde: int = 1, hasta: int | None = None,
                exacta: bool = False) -> int | None:
    objetivo = norm(etiqueta)
    hasta = hasta or ws.max_row
    for r in range(desde, hasta + 1):
        for col in columnas:
            t = norm(ws[f"{col}{r}"].value)
            if not t:
                continue
            if (t == objetivo) if exacta else t.startswith(objetivo):
                return r
    return None


def buscar_celdas(ws, patron: str):
    """Todas las celdas cuyo texto normalizado cumple la regex."""
    rx = re.compile(patron)
    for fila in ws.iter_rows():
        for c in fila:
            if isinstance(c.value, str) and rx.search(norm(c.value)):
                yield c


def columna_mes(ws, fila_encabezado: int, periodo: Periodo, fila_dato: int | None = None,
                desde_col: int = 1, hasta_col: int | None = None) -> int | None:
    """Columna cuyo encabezado es el mes buscado. Si hay varias, prefiere la que tiene dato."""
    candidatas = []
    hasta_col = hasta_col or ws.max_column
    for c in range(desde_col, hasta_col + 1):
        if periodo_de(ws.cell(fila_encabezado, c).value) == periodo:
            candidatas.append(c)
    if not candidatas:
        return None
    if fila_dato:
        for c in candidatas:                      # primero: dato numérico distinto de cero
            v = ws.cell(fila_dato, c).value
            if es_numero(v) and v != 0:
                return c
        for c in candidatas:
            if es_numero(ws.cell(fila_dato, c).value):
                return c
    return candidatas[0]


def meses_en_fila(ws, fila: int, desde_col: int = 1, hasta_col: int | None = None):
    """[(periodo, columna)] de los encabezados de mes en una fila (sin duplicados)."""
    vistos, salida = set(), []
    for c in range(desde_col, (hasta_col or ws.max_column) + 1):
        p = periodo_de(ws.cell(fila, c).value)
        if p and p not in vistos:
            vistos.add(p)
            salida.append((p, c))
    return salida


def valor(ws, fila: int | None, col: int | None):
    if fila is None or col is None:
        return None
    return ws.cell(fila, col).value


def num(v, defecto=0.0) -> float:
    return float(v) if es_numero(v) else defecto


def coord(fila: int, col: int) -> str:
    return f"{get_column_letter(col)}{fila}"


def leer_tabla(ws, fila_encabezado: int, col_inicio: int, fin_si=("total",), max_filas: int = 200,
               columnas: list[str] | None = None):
    """Lee filas bajo un encabezado hasta encontrar una fila 'TOTAL' o vacía."""
    encabezados = []
    c = col_inicio
    while c <= ws.max_column:
        v = ws.cell(fila_encabezado, c).value
        if v is None and columnas is None:
            break
        encabezados.append(str(v).strip() if v is not None else f"col{c}")
        c += 1
        if columnas and len(encabezados) >= len(columnas):
            break
    if columnas:
        encabezados = columnas
    filas = []
    fila_total = None
    for r in range(fila_encabezado + 1, fila_encabezado + 1 + max_filas):
        primera = ws.cell(r, col_inicio).value
        if any(norm(primera).startswith(f) for f in fin_si):
            fila_total = r
            break
        celdas = [ws.cell(r, col_inicio + i).value for i in range(len(encabezados))]
        if all(v is None for v in celdas):
            break
        filas.append({"_fila": r, **dict(zip(encabezados, celdas))})
    return encabezados, filas, fila_total
