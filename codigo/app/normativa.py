"""Base normativa: leyes, normas contables y manuales que el despacho carga, divididos por artículo o numeral.

- Solo contiene lo que se carga (PDF con texto, Word, TXT o Markdown). Nunca se completa de memoria: un artículo
  mal citado en un oficio firmado es un riesgo para el revisor fiscal.
- Ámbito: «general» (Ley 675 de 2001, NIIF para pymes, decretos…) o un cliente (reglamento de PH, manual de
  cartera, manual de contratación…).
- Fragmentos por artículo («Artículo 32»), numeral («4.3») o, si no hay estructura, por página.
- Búsqueda BM25 sin dependencias externas, insensible a tildes. Cada resultado trae su cita: documento + referencia.

Estructura: <raiz>/normativa/<ambito>/<doc_id>/{original.<ext>, meta.json, fragmentos.json}
"""
from __future__ import annotations

import math
import re
import shutil
import time
import unicodedata
from collections import Counter
from pathlib import Path

from agente.comun import guardar_json, leer_json

TIPOS = {"ley": "Ley", "decreto": "Decreto", "resolucion": "Resolución", "norma_contable": "Norma contable (NIIF)",
         "reglamento": "Reglamento de propiedad horizontal", "manual": "Manual de la copropiedad",
         "contrato": "Contrato", "otro": "Otro documento"}
JERARQUIA = ["ley", "decreto", "resolucion", "norma_contable", "reglamento", "contrato", "manual", "otro"]
EXTENSIONES = {".pdf", ".docx", ".txt", ".md", ".html", ".htm"}
VACIAS = set("""a al ante bajo con contra de del desde durante e el en entre es esta este hacia hasta la las le les lo los
mas o para pero por que se segun sin sobre su sus u un una uno unos y ya como cual cuando donde ser son fue han ha
articulo articulos paragrafo""".split())

RX_ARTICULO = re.compile(r"(?im)^\s*(?:art[íi]culo|art\.)\s+(\d+[a-z]?(?:\s*(?:bis|ter))?)\s*[\.\-–:º°o]*")
RX_NUMERAL = re.compile(r"(?m)^\s*(\d{1,2}(?:\.\d{1,3}){1,3})\.?\s+(?=[A-ZÁÉÍÓÚÑ])")
RX_SECCION = re.compile(r"(?im)^\s*(secci[óo]n|cap[íi]tulo|t[íi]tulo)\s+([0-9IVXLC]+)\b")


def _norm(t: str) -> str:
    return unicodedata.normalize("NFKD", t.lower()).encode("ascii", "ignore").decode()


def tokens(t: str) -> list[str]:
    return [w for w in re.findall(r"[a-z0-9]+", _norm(t)) if len(w) > 2 and w not in VACIAS]


# ------------------------------------------------------------------ extracción de texto

def _texto_html(crudo: bytes) -> str:
    """Texto de una página web guardada: un salto de línea por bloque, sin scripts ni estilos ni menús."""
    from html.parser import HTMLParser
    m = re.search(rb'charset=["\']?([A-Za-z0-9_-]+)', crudo[:4000])
    for cod in ([m.group(1).decode()] if m else []) + ["utf-8", "latin-1"]:
        try:
            fuente = crudo.decode(cod)
            break
        except (UnicodeDecodeError, LookupError):
            continue
    BLOQUES = {"p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "table", "section", "article", "blockquote"}

    class P(HTMLParser):
        def __init__(self):
            super().__init__(convert_charrefs=True)
            self.partes, self.saltar = [], 0

        def handle_starttag(self, tag, attrs):
            if tag in ("script", "style", "nav", "head", "noscript"):
                self.saltar += 1
            elif tag in BLOQUES:
                self.partes.append("\n")

        def handle_endtag(self, tag):
            if tag in ("script", "style", "nav", "head", "noscript"):
                self.saltar = max(0, self.saltar - 1)
            elif tag in BLOQUES:
                self.partes.append("\n")

        def handle_data(self, data):
            if not self.saltar:
                self.partes.append(data)
    pa = P()
    pa.feed(fuente)
    texto = "".join(pa.partes).replace("\xa0", " ")
    return re.sub(r"\n[ \t]*\n[\s]*", "\n\n", re.sub(r"[ \t]+", " ", texto))


def extraer_paginas(ruta: Path) -> list[str]:
    ext = ruta.suffix.lower()
    if ext in (".html", ".htm"):
        return [_texto_html(ruta.read_bytes())]
    if ext == ".pdf":
        import pymupdf
        with pymupdf.open(ruta) as doc:
            return [p.get_text() for p in doc]
    if ext == ".docx":
        from docx import Document
        d = Document(str(ruta))
        partes = [p.text for p in d.paragraphs]
        for t in d.tables:
            for fila in t.rows:
                partes.append(" | ".join(c.text for c in fila.cells))
        return ["\n".join(partes)]
    return [ruta.read_text(encoding="utf-8", errors="replace")]


def fragmentar(paginas: list[str]) -> list[dict]:
    """Divide por artículos; si no hay, por numerales; si tampoco, por página (o bloques de ~2.500 caracteres)."""
    texto, inicio_pag = "", []
    for i, p in enumerate(paginas, 1):
        inicio_pag.append((len(texto), i))
        texto += p.replace("\x0c", "") + "\n"

    def pagina(pos):
        return max((n for ini, n in inicio_pag if ini <= pos), default=1)

    def seccion_previa(pos):
        m = None
        for m in RX_SECCION.finditer(texto, 0, pos):
            pass
        return f"{m.group(1).capitalize()} {m.group(2)}" if m else ""

    for rx, nombre in ((RX_ARTICULO, "Artículo"), (RX_NUMERAL, "Numeral")):
        marcas = list(rx.finditer(texto))
        if len(marcas) >= 3:
            frags = []
            if marcas[0].start() > 200:
                frags.append({"ref": "Encabezado", "texto": texto[:marcas[0].start()]})
                frags[-1]["pagina"] = 1
            for i, m in enumerate(marcas):
                fin = marcas[i + 1].start() if i + 1 < len(marcas) else len(texto)
                numero = re.sub(r"(?<=\d)[oº°]$", "", m.group(1).strip())     # «2o» → «2»
                ref = f"{nombre} {numero}"
                sec = seccion_previa(m.start()) if nombre == "Artículo" else ""
                frags.append({"ref": ref, "seccion": sec, "pagina": pagina(m.start()), "texto": texto[m.start():fin]})
            break
    else:
        frags = []
        for i, p in enumerate(paginas, 1):
            for j in range(0, max(len(p), 1), 2500):
                trozo = p[j:j + 2500]
                if trozo.strip():
                    frags.append({"ref": f"Página {i}" + (f" (parte {j // 2500 + 1})" if len(p) > 2500 else ""),
                                  "pagina": i, "texto": trozo})
    out = []
    for f in frags:
        t = re.sub(r"[ \t]+", " ", re.sub(r"\n{3,}", "\n\n", f["texto"])).strip()
        # el título del capítulo siguiente queda al final del artículo anterior: se quita
        t = re.sub(r"(?is)(\n\s*(secci[óo]n|cap[íi]tulo|t[íi]tulo)\s+[0-9IVXLC]+\b[^\n]*)+\s*$", "", t).strip()
        if len(t) >= 20:
            out.append({**f, "texto": t[:12_000], "id": len(out)})
    return out


def revisar_numeracion(frags: list[dict]) -> list[str]:
    """Avisa si faltan, se repiten o se desordenan artículos (p. ej. porque no se cargó una de las partes)."""
    nums = [int(m.group(1)) for f in frags if (m := re.match(r"Artículo (\d+)$", f["ref"]))]
    if len(nums) < 3:
        return []
    out = []
    faltan = sorted(set(range(min(nums), max(nums) + 1)) - set(nums))
    if faltan:
        rangos, ini = [], faltan[0]
        for a, b in zip(faltan, faltan[1:] + [None]):
            if b != a + 1:
                rangos.append(f"{ini}" if ini == a else f"{ini} a {a}")
                if b is not None:
                    ini = b
        out.append(f"Faltan los artículos {', '.join(rangos[:8])}{'…' if len(rangos) > 8 else ''}: verifique que cargó "
                   "todas las partes y que el PDF no omitió páginas.")
    repetidos = sorted({n for n in nums if nums.count(n) > 1})
    if repetidos:
        out.append(f"Artículos repetidos: {', '.join(map(str, repetidos[:10]))}. Puede que una parte se haya cargado dos veces "
                   "o que el texto cite artículos al inicio de una línea.")
    if any(b < a for a, b in zip(nums, nums[1:])) and not repetidos:
        out.append("Los artículos no siguen el orden: revise el orden en que se cargaron las partes.")
    return out


# ------------------------------------------------------------------ biblioteca

_instancias: dict[str, "Normativa"] = {}


def instancia(raiz: Path) -> "Normativa":
    """Una por espacio de trabajo, para reutilizar el índice entre búsquedas."""
    clave = str(Path(raiz).resolve())
    if clave not in _instancias:
        _instancias[clave] = Normativa(Path(raiz))
    return _instancias[clave]


class Normativa:
    def __init__(self, raiz: Path):
        self.raiz = Path(raiz) / "normativa"
        self._indice = None

    def _dir(self, ambito: str) -> Path:
        return self.raiz / ("general" if ambito == "general" else f"clientes/{ambito}")

    def documentos(self, ambito: str | None = None) -> list[dict]:
        metas = [leer_json(p, {}) for p in sorted(self.raiz.glob("**/meta.json"))] if self.raiz.exists() else []
        metas = [m for m in metas if m and (ambito is None or m["ambito"] == ambito)]
        return sorted(metas, key=lambda m: (m["ambito"] != "general", JERARQUIA.index(m.get("tipo", "otro")), m["titulo"]))

    def cargar(self, archivos: list[tuple[Path, str]], titulo: str, tipo: str, ambito: str, version: str,
               origen: str, por: str) -> dict:
        """archivos = [(ruta temporal, nombre original), …] en orden: se unen como UN documento (p. ej. una ley
        publicada en varias páginas)."""
        if not archivos:
            raise ValueError("Seleccione al menos un archivo.")
        for _, nombre in archivos:
            if Path(nombre).suffix.lower() not in EXTENSIONES:
                raise ValueError(f"{nombre}: formato no admitido ({', '.join(sorted(EXTENSIONES))}).")
        if tipo not in TIPOS:
            raise ValueError("Tipo de documento inválido.")
        if not titulo.strip():
            raise ValueError("Indique el título del documento (p. ej. «Ley 675 de 2001»).")
        paginas = []
        for i, (ruta, nombre) in enumerate(archivos):
            partes = extraer_paginas(ruta)
            if i and RX_ARTICULO.search("".join(partes)):
                # de la segunda parte en adelante, el encabezado repetido de la página (título, menús, fecha de
                # actualización) se descarta para que no quede pegado al último artículo de la parte anterior
                unido = "".join(partes)
                partes = [unido[RX_ARTICULO.search(unido).start():]]
            if len("".join(partes).strip()) < 100:
                raise ValueError(f"{nombre}: casi no tiene texto extraíble (¿PDF escaneado?). Cárguelo con texto "
                                 "seleccionable o páselo antes por OCR.")
            paginas += partes
        texto = "".join(paginas)
        if len(texto.strip()) < 200:
            raise ValueError("El documento casi no tiene texto extraíble (¿PDF escaneado?). Cárguelo con texto "
                             "seleccionable o pase antes el PDF por OCR.")
        frags = fragmentar(paginas)
        advertencias = revisar_numeracion(frags)
        doc_id = time.strftime("%Y%m%d%H%M%S") + "-" + re.sub(r"[^a-z0-9]+", "-", _norm(titulo))[:40].strip("-")
        d = self._dir(ambito) / doc_id
        d.mkdir(parents=True)
        for i, (ruta, nombre) in enumerate(archivos, 1):
            shutil.copy2(ruta, d / f"original_{i}{Path(nombre).suffix.lower()}")
        estructura = Counter(f["ref"].split()[0] for f in frags).most_common(1)[0][0]
        meta = {"id": doc_id, "titulo": titulo.strip(), "tipo": tipo, "ambito": ambito, "version": version.strip(),
                "origen": origen.strip(), "archivos": [n for _, n in archivos], "archivo": archivos[0][1],
                "cargado": time.strftime("%Y-%m-%dT%H:%M:%S"), "advertencias": advertencias,
                "cargado_por": por, "paginas": len(paginas), "fragmentos": len(frags), "estructura": estructura,
                "caracteres": len(texto)}
        guardar_json(d / "fragmentos.json", frags)
        guardar_json(d / "meta.json", meta)
        self._indice = None
        return meta

    def eliminar(self, ambito: str, doc_id: str):
        d = self._dir(ambito) / doc_id
        if not re.match(r"^\d{14}-[a-z0-9-]*$", doc_id) or not (d / "meta.json").exists():
            raise FileNotFoundError("No existe ese documento.")
        shutil.rmtree(d)
        self._indice = None

    def fragmentos(self, ambito: str, doc_id: str) -> tuple[dict, list[dict]]:
        d = self._dir(ambito) / doc_id
        if not re.match(r"^\d{14}-[a-z0-9-]*$", doc_id) or not (d / "meta.json").exists():
            raise FileNotFoundError("No existe ese documento.")
        return leer_json(d / "meta.json"), leer_json(d / "fragmentos.json", [])

    # -------------------------------------------------------------- búsqueda
    def _construir(self):
        docs = []
        for m in self.documentos():
            for f in leer_json(self._dir(m["ambito"]) / m["id"] / "fragmentos.json", []):
                tk = tokens(f"{m['titulo']} {f['ref']} {f.get('seccion', '')} {f['texto']}")
                docs.append((m, f, Counter(tk), len(tk)))
        df = Counter()
        for _, _, tf, _ in docs:
            df.update(tf.keys())
        prom = sum(n for *_, n in docs) / len(docs) if docs else 1
        self._indice = (docs, df, prom)

    def buscar(self, consulta: str, cliente: str | None = None, k: int = 8, tipos: list[str] | None = None) -> list[dict]:
        """BM25 sobre lo general + lo del cliente indicado. Un número de artículo en la consulta lo prioriza."""
        if self._indice is None:
            self._construir()
        docs, df, prom = self._indice
        q = tokens(consulta)
        art = re.search(r"art[íi]culo\s+(\d+[a-z]?)", consulta, re.I)
        if (not q and not art) or not docs:
            return []
        n = len(docs)
        res = []
        for m, f, tf, largo in docs:
            if m["ambito"] not in ("general", cliente) or (tipos and m["tipo"] not in tipos):
                continue
            s = 0.0
            for w in q:
                if w in tf:
                    idf = math.log(1 + (n - df[w] + 0.5) / (df[w] + 0.5))
                    s += idf * tf[w] * 2.2 / (tf[w] + 1.2 * (0.25 + 0.75 * largo / prom))
            if art and f["ref"].lower() == f"artículo {art.group(1).lower()}":
                s += 50
            if s > 0:
                res.append((s, m, f))
        res.sort(key=lambda x: -x[0])
        return [{"documento": m["titulo"], "tipo": TIPOS[m["tipo"]], "ambito": m["ambito"], "doc_id": m["id"],
                 "version": m.get("version", ""), "referencia": f["ref"], "seccion": f.get("seccion", ""),
                 "pagina": f.get("pagina"), "fragmento": f["id"], "texto": f["texto"], "puntaje": round(s, 2),
                 "cita": f"({m['titulo']}, {f['ref'].lower() if f['ref'].startswith('Artículo') else f['ref']})"}
                for s, m, f in res[:k]]
