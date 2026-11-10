"""Informe COMPLETO de Revisoría Fiscal, con la misma estructura del informe modelo (006-2026 Picasso).

Estructura y textos fijos: config/estructura_informe.yaml. Contenido: Excel de papeles de trabajo
(sección financiera con revisoria.extractor; secciones regulatoria y de control interno con revisoria.documental).

Todo es nativo y editable: tablas, textos y semáforos. Las fotos de soporte que el equipo pega en el Excel se
llevan como imágenes junto al campo o comprobante al que están ancladas.
"""
from __future__ import annotations

import io
import math
import re
from pathlib import Path

import yaml
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from .excel import MESES_CORTOS, Periodo, norm
from .generador_pptx import (ALTO, ANCHO, AZUL, BLANCO, FUENTE, FUENTE_T, GRIS, GRIS_CLARO, LINEA, NAVY, NEG,
                             PLANTILLA, ROJO, SECCIONES_DIAPOSITIVA, TEXTO, Y0, Y_MAX, Informe, _rect, _sin_sombra,
                             _triangulo, _txt,
                             fmt_pesos)

ESTRUCTURA = Path(__file__).resolve().parent.parent / "config" / "estructura_informe.yaml"
TAMANOS = (9, 8.5, 8, 7.5, 7, 6.5)


def _clave(t) -> str:
    return re.sub(r"[^a-z0-9]+", " ", norm(t)).strip()


def _rgb(hexa: str) -> RGBColor:
    return RGBColor.from_string(hexa.upper())


def _fecha_corta(iso) -> str:
    if not iso:
        return ""
    if hasattr(iso, "strftime"):
        return iso.strftime("%d/%m/%Y")
    iso = str(iso)
    try:
        a, m, d = iso[:10].split("-")
        return f"{d}/{m}/{a}"
    except ValueError:
        return iso


def _motivo(seccion: str, detalle: str | None) -> str:
    """Motivo legible de una sección que no se pudo leer (sin rutas de archivo ni nombres de excepción)."""
    nombres = {"liquidez": "el análisis de liquidez", "cartera": "la cartera por inmueble", "deudores": "los deudores",
               "presupuesto": "la ejecución presupuestal", "patrimonio": "el patrimonio y los excedentes"}
    d = re.sub(r" en /\S+.*$", "", detalle or "")
    if "formato distinto" in d:
        d = "la hoja tiene un formato distinto; se puede leer con una receta de lectura aprobada por el revisor"
    return f"No se pudo leer {nombres.get(seccion, seccion)}: {d or 'no está en el Excel'}."


# secciones del modelo que usa cada tipo de ítem de la estructura (las financieras van en SECCIONES_DIAPOSITIVA)
SECCIONES_TIPO = {"egresos": ("egresos",), "conciliaciones": ("conciliaciones",), "retencion": ("retencion",),
                  "actas_consejo": ("actas_consejo",), "asamblea": ("asamblea",), "contratos": ("contratos",),
                  "polizas_manual": ("polizas", "matriz_manuales"), "matriz_riesgo": ("matriz_riesgo",)}


class InformeCompleto(Informe):
    def __init__(self, m: dict, cfg: dict, doc: dict, redactor=None, estructura: dict | None = None, propuestas=None):
        super().__init__(m, cfg, redactor, propuestas)
        self.doc = doc or {}
        self.est = estructura or yaml.safe_load(ESTRUCTURA.read_text(encoding="utf-8"))
        self.sem = self.est.get("semaforo", {})
        self.usados: set[str] = set()          # campos del Excel que ya salieron en el informe
        self.avisos: list[str] = []            # información pendiente en el Excel (para el reporte)
        self.errores: list[str] = []           # diapositivas que fallaron al armarse (se conservan como pendientes)
        self.plan: list[dict] = []             # cada ítem de la estructura y las diapositivas que ocupó

    # ---------------------------------------------------------------- utilidades
    def campos(self, claves) -> list[dict]:
        claves = claves if isinstance(claves, list) else [claves]
        out = []
        disponibles = ((self.doc.get("auditoria") or {}).get("campos")) or []
        for k in claves:
            kk = _clave(k)
            c = next((c for c in disponibles if _clave(c["campo"]) == kk), None) or \
                next((c for c in disponibles if _clave(c["campo"]).startswith(kk) or kk.startswith(_clave(c["campo"]))), None)
            if c:
                out.append(c)
                self.usados.add(c["campo"])
            else:
                self.avisos.append(f"No se encontró el campo «{k}» en la hoja AUD CONT Y ADMON.")
        return out

    def badge(self, s, nivel: str, x=8.05, y=0.60, w=0.82, h=0.28):
        nivel = (nivel or "").upper().strip()
        if nivel not in self.sem:
            return
        color = _rgb(self.sem[nivel]["color"])
        for sh in s.shapes:                      # el subtítulo no debe pasar por debajo del semáforo
            if sh.name == "Subtitulo":
                sh.width = Inches(x - 0.42 - 0.1)
        b = _rect(s, x, y, w, h, color, MSO_SHAPE.ROUNDED_RECTANGLE)
        tf = b.text_frame
        tf.margin_left = tf.margin_right = Inches(0.02)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = nivel
        r.font.size, r.font.bold, r.font.name = Pt(10), True, FUENTE_T
        r.font.color.rgb = TEXTO if nivel == "MEDIO" else BLANCO

    def _vinetas(self, s, x, y, w, h, items, size, titulo=None, color_titulo=NAVY):
        """items: [{'texto', 'tipo', 'seg'}]. Hallazgos en rojo; énfasis del Excel respetado."""
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.03)
        primero = True
        if titulo:
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = titulo
            r.font.size, r.font.bold, r.font.name = Pt(size + 1.5), True, FUENTE_T
            r.font.color.rgb = color_titulo
            primero = False
        for it in items:
            p = tf.paragraphs[0] if primero else tf.add_paragraph()
            primero = False
            p.space_before = Pt(3)
            rojo_todo = it.get("tipo") == "hallazgo"
            segs = it.get("seg") or [{"t": it["texto"], "b": rojo_todo, "rojo": rojo_todo}]
            r = p.add_run()
            r.text = "•  "
            r.font.size, r.font.name = Pt(size), FUENTE
            r.font.color.rgb = ROJO
            for k, sg in enumerate(segs):
                t = sg["t"]
                if k == 0:
                    t = re.sub(r"^\s*\((ev|op)\)\s*:?\s*", "", t, flags=re.I)
                t = t.replace("\n", " ")
                if not t:
                    continue
                r = p.add_run()
                r.text = t
                r.font.size, r.font.name = Pt(size), FUENTE
                r.font.bold = bool(sg["b"])
                r.font.color.rgb = NEG if sg["rojo"] else TEXTO
        return tb

    def _alto_items(self, items, w, size):
        return sum(self.alto_texto(it["texto"], w - 0.2, size) - 0.1 for it in items) + 0.3

    def _imagenes(self, s, imgs, x, y, w, h, max_cols=4):
        if not imgs:
            return
        n = len(imgs)
        cols = min(max_cols, n)
        filas = math.ceil(n / cols)
        cw, ch = (w - 0.1 * (cols - 1)) / cols, (h - 0.1 * (filas - 1)) / filas
        for k, im in enumerate(imgs):
            r, c = divmod(k, cols)
            ancho, alto = im.get("ancho") or 400, im.get("alto") or 300
            esc = min(cw / (ancho / 96), ch / (alto / 96))
            ww, hh = ancho / 96 * esc, alto / 96 * esc
            xx = x + c * (cw + 0.1) + (cw - ww) / 2
            yy = y + r * (ch + 0.1) + (ch - hh) / 2
            try:
                pic = s.shapes.add_picture(io.BytesIO(im["bytes"]), Inches(xx), Inches(yy), Inches(ww), Inches(hh))
                pic.line.color.rgb = LINEA
            except Exception:
                self.avisos.append("Una imagen del Excel no se pudo insertar (formato no soportado).")

    def _pagina_imagenes(self, titulo, subtitulo, imgs, fuentes=None, leyendas=None):
        por_pag = 4
        for k in range(0, len(imgs), por_pag):
            s = self._nueva(titulo + (f" ({k // por_pag + 1}/{math.ceil(len(imgs) / por_pag)})" if len(imgs) > por_pag else ""),
                            subtitulo, fuentes)
            grupo = imgs[k:k + por_pag]
            cols = 2 if len(grupo) > 1 else 1
            filas = math.ceil(len(grupo) / cols)
            alto_celda = (Y_MAX - Y0 - 0.1 * (filas - 1)) / filas
            for j, im in enumerate(grupo):
                r, c = divmod(j, cols)
                x = 0.42 + c * 4.65
                y = Y0 + r * (alto_celda + 0.1)
                ley = (leyendas or [None] * len(imgs))[k + j]
                if ley:
                    _txt(s, x, y, 4.5, 0.32, ley, size=7.5, color=NAVY, bold=True)
                self._imagenes(s, [im], x, y + (0.34 if ley else 0), 4.5, alto_celda - (0.36 if ley else 0), 1)

    # ---------------------------------------------------------------- portada y textos fijos
    def d_portada(self):
        s = self.prs.slides.add_slide(self.blank)
        self.n += 1
        s.shapes.add_picture(str(PLANTILLA / "fondo_seccion.jpg"), 0, 0, Inches(ANCHO), Inches(ALTO))
        s.shapes.add_picture(str(PLANTILLA / "logo.png"), Inches(8.93), Inches(0.125), Inches(0.883), Inches(0.625))
        c, meta = self.cfg, self.m["meta"]
        lineas = [" ".join(x for x in ("Informe", str(c.get("numero_informe") or "").strip(), "de Revisoría Fiscal.") if x), f"{c.get('copropiedad', '')}.",
                  f"Nit. {c.get('nit', '')}"]
        unidades = meta.get("unidades") or c.get("unidades")
        if unidades:                                      # ficha nueva sin unidades: no se imprime «None Aptos»
            lineas.append(f"{unidades} Aptos")
        _txt(s, 3.9, 1.35, 5.6, 1.8, lineas, size=19, bold=True, color=NAVY, font=FUENTE_T)
        _rect(s, 5.75, 3.37, 0.45, 0.03, ROJO)
        _txt(s, 6.3, 3.25, 3.0, 0.3, f"Período: {self.p.etiqueta}", size=10.5, bold=True, color=NAVY)
        return s

    def d_metodologia(self):
        s = self._nueva("Metodología de Evaluación y Semáforo de Cumplimiento", "Auditoría integral")
        ejes = ["Auditoría de cumplimiento (regulatoria contable, tributaria y administrativa)",
                "Auditoría financiera", "Auditoría de control interno y actos administrativos"]
        for k, e in enumerate(ejes):
            y = Y0 + 0.35 + k * 1.15
            b = _rect(s, 0.6, y, 3.4, 0.95, NAVY if k != 1 else AZUL, MSO_SHAPE.ROUNDED_RECTANGLE)
            tf = b.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = e
            r.font.size, r.font.bold, r.font.name, r.font.color.rgb = Pt(10), True, FUENTE_T, BLANCO
        for k, (nivel, d) in enumerate(self.sem.items()):
            y = Y0 + 0.2 + k * 0.98
            flecha = _rect(s, 4.5, y, 1.35, 0.7, _rgb(d["color"]), MSO_SHAPE.PENTAGON)
            tf = flecha.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = nivel
            r.font.size, r.font.bold, r.font.name = Pt(14), True, FUENTE_T
            r.font.color.rgb = TEXTO if nivel == "MEDIO" else BLANCO
            _rect(s, 5.95, y, 3.65, 0.7, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE)
            _txt(s, 6.08, y + 0.05, 3.45, 0.6, d["texto"], size=9, color=TEXTO, anchor=MSO_ANCHOR.MIDDLE)

    def d_objetivo_alcance(self):
        tf_ = self.est.get("textos_fijos", {})
        s = self._nueva("Objetivo y Alcance de la Auditoría", "Revisión integral del cumplimiento normativo contable")
        _txt(s, 0.6, Y0 + 0.35, 4.1, 0.3, "Objetivo principal", size=12, bold=True, color=ROJO, font=FUENTE_T)
        _txt(s, 0.6, Y0 + 0.75, 4.1, 2.4, tf_.get("objetivo", "").strip(), size=11, color=TEXTO)
        _rect(s, 4.95, Y0 + 0.35, 0.03, 3.2, ROJO)
        _txt(s, 5.25, Y0 + 0.35, 4.2, 0.3, "Alcance de la revisión", size=12, bold=True, color=ROJO, font=FUENTE_T)
        self._vinetas(s, 5.25, Y0 + 0.75, 4.3, 3.0, [{"texto": a} for a in tf_.get("alcance", [])], 11)

    def d_texto_fijo(self, clave):
        t = self.est.get("textos_fijos", {}).get(clave) or {}
        s = self._nueva(t.get("titulo", clave), None)
        _txt(s, 0.6, Y0, 9.0, 0.6, t.get("intro", "").strip(), size=11, color=TEXTO)
        items = t.get("items", [])
        mitad = math.ceil(len(items) / 2)
        for col, grupo in enumerate((items[:mitad], items[mitad:])):
            tb = s.shapes.add_textbox(Inches(0.6 + col * 4.6), Inches(Y0 + 0.7), Inches(4.4), Inches(3.4))
            tf = tb.text_frame
            tf.word_wrap = True
            for k, it in enumerate(grupo):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                p.space_before = Pt(3)
                r = p.add_run()
                r.text = f"{k + 1 + col * mitad}. "
                r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(10.5), True, ROJO, FUENTE
                r = p.add_run()
                r.text = it
                r.font.size, r.font.color.rgb, r.font.name = Pt(10.5), TEXTO, FUENTE

    def d_cierre(self):
        t = self.est.get("textos_fijos", {}).get("cierre") or {}
        s = self.prs.slides.add_slide(self.blank)
        self.n += 1
        _rect(s, 0, 0, ANCHO, ALTO, RGBColor(0xEE, 0xEF, 0xF2))
        _triangulo(s, [(0, 0), (2.2, 0), (0, 1.3)], RGBColor(0xF6, 0xD5, 0xCF))
        s.shapes.add_picture(str(PLANTILLA / "logo.png"), Inches(8.93), Inches(0.125), Inches(0.883), Inches(0.625))
        _txt(s, 0.5, 1.45, 9.0, 0.9, "GRACIAS", size=40, bold=True, color=NAVY, font=FUENTE_T, align=PP_ALIGN.CENTER)
        _txt(s, 0.5, 2.55, 9.0, 0.4, self.cfg.get("firma", "LC Counter Contadores y Auditores SAS"), size=15, bold=True,
             color=ROJO, align=PP_ALIGN.CENTER)
        _txt(s, 0.5, 2.95, 9.0, 0.3, t.get("lema", ""), size=10.5, color=TEXTO, align=PP_ALIGN.CENTER)
        _txt(s, 0.5, 3.55, 9.0, 0.3, f"{t.get('correo', '')}      ·      {t.get('web', '')}", size=9.5, color=GRIS,
             align=PP_ALIGN.CENTER)

    # ---------------------------------------------------------------- auditoría regulatoria
    def d_egresos(self):
        E = self.doc.get("egresos") or {}
        if not E.get("items"):
            self.avisos.append(E.get("faltante") or "Sin comprobantes de egreso del periodo.")
            self.d_sin_datos(
                f"Comprobantes de Egreso {self.p.etiqueta}", "Revisión de soportes y observaciones",
                "No se encontraron comprobantes de egreso del periodo en los papeles de trabajo."
            )
            return
        cols = [("CE", 0.45, "c"), ("Fecha", 0.68, "c"), ("Tercero", 1.6, "l"), ("Descripción", 2.45, "l"),
                ("Valor", 0.95, "r"), ("Observaciones", 3.03, "l")]
        filas = []
        for i in E["items"]:
            obs = i["observacion"] or ""
            estilo = {"seg": i["seg_observacion"]} if i["seg_observacion"] else \
                ({} if i["correcto"] else {"color": NEG, "bold": False})
            filas.append([i["ce"], i["fecha"], i["tercero"], i["descripcion"],
                          fmt_pesos(i["valor"]) if isinstance(i["valor"], (int, float)) else str(i["valor"] or ""),
                          (obs, estilo) if estilo else obs])
        total = ["", "", (E.get("etiqueta_total") or f"Total egresos {self.p.etiqueta}").upper(), "", fmt_pesos(E["total_hoja"]), ""]
        self.tabla_paginada(f"Comprobantes de Egreso {self.p.etiqueta}",
                            f"{len(E['items'])} comprobantes revisados · {E['con_observacion']} con observación",
                            cols, filas, [E.get("fuente_total")], size=6.5, total=total)
        con_img = [i for i in E["items"] if i["imagenes"]]
        if con_img:
            imgs, ley = [], []
            for i in con_img:
                for im in i["imagenes"]:
                    imgs.append(im)
                    ley.append(f"CE {i['ce']} · {i['tercero']} — {i['observacion'][:80]}")
            self._pagina_imagenes("Soportes de comprobantes con observación", f"Comprobantes de egreso {self.p.etiqueta}",
                                  imgs, [i["fuente"] for i in con_img], ley)

    def d_campo(self, item):
        cs = self.campos(item.get("campo"))
        if not cs:
            self.d_sin_datos(
                item.get("titulo") or str(item.get("campo")), item.get("subtitulo") or "Auditoría regulatoria",
                f"No se encontró el ítem «{item.get('campo')}» en la hoja AUD CONT Y ADMON."
            )
            return
        titulo = item.get("titulo") or cs[0]["campo"]
        nivel = self._nivel_peor([c["nivel"] for c in cs])
        normas = [n for c in cs for n in c["normatividad"]]
        textos = [t for c in cs for t in c["textos"]]
        evid = [t for t in textos if t["tipo"] in ("evidencia", "hallazgo")]
        opin = [t for c in cs for t in c["textos"] if t["tipo"] == "opinion"]
        imgs = [im for c in cs for im in c["imagenes"]]
        if not imgs:                      # soportes del comprobante citado (p. ej. la prefactura del CE 4295)
            docs = [d for c in cs for d in re.split(r"[-,;\s]+", c.get("documentos") or "") if d.isdigit()]
            if 0 < len(docs) <= 2:
                imgs = [im for i in (self.doc.get("egresos") or {}).get("items", []) if i["ce"] in docs
                        for im in i["imagenes"]][:4]
        dato = self._dato_documentos(cs[0])
        fuentes = [c["fuente"] for c in cs] + [t["fuente"] for t in textos]
        self._diapositiva_campo(titulo, item.get("subtitulo"), nivel, normas, evid, opin, imgs, dato, fuentes)

    def _nivel_peor(self, niveles):
        orden = ["BAJO", "MEDIO", "ODM", "ALTO"]
        validos = [n for n in niveles if n in orden]
        return min(validos, key=orden.index) if validos else ""

    def _dato_documentos(self, c):
        E = self.doc.get("egresos") or {}
        n_total = len(E.get("items") or [])
        docs = [d for d in re.split(r"[-,;\s]+", c.get("documentos") or "") if d.isdigit()]
        if docs and n_total:
            return f"{len(docs)} de {n_total} comprobantes de egreso ({len(docs) / n_total * 100:.1f}%)".replace(".", ",") + \
                f": CE {', '.join(docs)}"
        return None

    def _diapositiva_campo(self, titulo, subtitulo, nivel, normas, evid, opin, imgs, dato, fuentes):
        # reparte evidencias en páginas si no caben ni con la letra mínima
        ancho_izq = 5.45 if opin else 9.15
        y_ev = Y0 + 0.62 + (0.3 if dato else 0)
        alto_disp = Y_MAX - y_ev - 0.05
        for size in TAMANOS:
            if self._alto_items(evid, ancho_izq, size) <= alto_disp and \
                    self._alto_items(opin, 3.5, size) <= Y_MAX - (Y0 + 0.62) - 0.1:
                break
        paginas, actual, acum = [], [], 0.3
        for it in evid:
            h = self.alto_texto(it["texto"], ancho_izq - 0.2, size) - 0.1
            if actual and acum + h > alto_disp:
                paginas.append(actual)
                actual, acum = [], 0.3
            actual.append(it)
            acum += h
        paginas.append(actual)
        for k, pag in enumerate(paginas):
            t = titulo + (f" ({k + 1}/{len(paginas)})" if len(paginas) > 1 else "")
            s = self._nueva(t, subtitulo, fuentes)
            self.badge(s, nivel)
            # normatividad
            _txt(s, 0.42, Y0, 1.75, 0.25, "Normatividad aplicable", size=8.5, bold=True, color=ROJO)
            _txt(s, 2.2, Y0, 7.4, 0.5, " · ".join(normas) or "—", size=8.5, color=TEXTO)
            _rect(s, 0.42, Y0 + 0.52, 9.18, 0.012, LINEA)
            if dato and k == 0:
                _txt(s, 0.42, Y0 + 0.6, 9.18, 0.25, dato, size=8, bold=True, color=NEG)
            y = y_ev if k == 0 else Y0 + 0.62
            if pag:
                self._vinetas(s, 0.42, y, ancho_izq, Y_MAX - y, pag, size, titulo="Evidencias principales")
            if opin and k == len(paginas) - 1:
                h = min(Y_MAX - (Y0 + 0.62), self._alto_items(opin, 3.5, size) + 0.15)
                _rect(s, 6.05, Y0 + 0.62, 3.55, h, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE)
                _rect(s, 6.05, Y0 + 0.74, 0.05, h - 0.24, ROJO)
                self._vinetas(s, 6.18, Y0 + 0.68, 3.38, h - 0.1, opin, size, titulo="Opinión y recomendación",
                              color_titulo=ROJO)
        if imgs:
            self._pagina_imagenes(f"Evidencias: {titulo}", subtitulo, imgs, fuentes)

    def d_conciliaciones(self, item):
        cuentas = [c for c in (self.m.get("conciliaciones") or []) if not c.get("sin_datos")]
        if not cuentas:
            self.avisos.append("Sin conciliaciones bancarias del periodo.")
            self.d_sin_datos(
                item.get("titulo", "Conciliaciones Bancarias"), f"Corte a {self.p.etiqueta}",
                "No se encontraron conciliaciones bancarias con saldos de libros, extractos y partidas conciliatorias."
            )
            return
        ok = all(abs((c.get("diferencia_hoja") or 0) + sum((pp.get("valor") or 0) for pp in c.get("partidas", [])
                                                        if "pendientes" in norm(pp["concepto"]))) < 1 for c in cuentas)
        s = self._nueva(item.get("titulo", "Conciliaciones Bancarias"), f"Corte a {self.p.etiqueta}",
                        [c.get("fuente") for c in cuentas])
        self.badge(s, "ALTO" if ok else "ODM")
        cols, w = 3, 3.02
        alto_fila0 = 0.0
        for k, c in enumerate(cuentas):
            r, col = divmod(k, cols)
            x = 0.42 + col * (w + 0.06)
            filas = [[a.replace("anexo ", "Anexo ").capitalize(), v or "—"] for a, v in c.get("anexos", {}).items()]
            filas += [["Saldo en libros", fmt_pesos(c.get("libros"))], ["Saldo extracto", fmt_pesos(c.get("extracto"))]]
            dif = c.get("diferencia_hoja") or 0
            filas.append(["Diferencia", (fmt_pesos(dif), {"color": NEG if abs(dif) > 1 else TEXTO, "bold": True})])
            for pp in c.get("partidas", []):
                if pp.get("valor"):
                    filas.append([pp["concepto"], fmt_pesos(pp["valor"])])
            y = Y0 + (0 if r == 0 else alto_fila0 + 0.12 * r)
            _, h = self.tabla(s, x, y, [(c["cuenta"][:70], 1.92, "l"), (self.p.corto, 1.1, "r")], filas, size=6,
                              alto_fila=0.17, alto_enc=0.36)
            if r == 0:
                alto_fila0 = max(alto_fila0, h)

    def d_retencion(self, item):
        cs = self.campos(item.get("campo")) if item.get("campo") else []
        R = self.m.get("retencion") or []
        desde = self.p.cierre_anterior().clave
        filas = []
        for x in R:
            if desde <= x.get("periodo", "") < self.p.clave:
                pa, me = x["periodo"].split("-")
                etiqueta = f"{MESES_CORTOS[int(me) - 1].capitalize()}-{pa[2:]}"
                obs = x.get("observacion") or ""
                filas.append([etiqueta, _fecha_corta(x.get("presentado")), _fecha_corta(x.get("pagado")),
                              fmt_pesos(x.get("valor")), (obs, {} if norm(obs).startswith("correcto") else {"color": NEG})])
        s = self._nueva(item.get("titulo", "Retención en la Fuente"), item.get("subtitulo"),
                        [x.get("fuente") for x in R] + [c["fuente"] for c in cs])
        self.badge(s, self._nivel_peor([c["nivel"] for c in cs]) if cs else "")
        normas = [n for c in cs for n in c["normatividad"]]
        _txt(s, 0.42, Y0, 1.75, 0.25, "Normatividad aplicable", size=8.5, bold=True, color=ROJO)
        _txt(s, 2.2, Y0, 7.4, 0.45, " · ".join(normas) or "—", size=8.5, color=TEXTO)
        _txt(s, 0.42, Y0 + 0.5, 5.6, 0.25, "Presentación y pago del período auditado", size=9, bold=True, color=NAVY)
        self.tabla(s, 0.42, Y0 + 0.8, [("Periodo", 0.75, "c"), ("Presentado", 0.95, "c"), ("Pagado", 0.95, "c"),
                                       ("Valor", 1.05, "r"), ("Observación", 1.9, "l")], filas, size=8, alto_fila=0.24)
        textos = [t for c in cs for t in c["textos"]]
        if textos:
            self._vinetas(s, 6.25, Y0 + 0.5, 3.35, Y_MAX - Y0 - 0.5, textos, 8, titulo="Evidencias")

    def d_actas_consejo(self, item):
        A = self.doc.get("actas_consejo") or {}
        cs = self.campos(item.get("campo")) if item.get("campo") else []
        if not A.get("actas"):
            self.avisos.append("Sin actas de consejo del año en la hoja ACTAS DE CONSEJO.")
            self.d_sin_datos(
                item.get("titulo", "Libros Actas de Consejo de Administración"),
                f"Revisión de actas del año {self.p.anio}",
                "La hoja ACTAS DE CONSEJO no contiene actas para revisar temas, fechas y firmas."
            )
            return
        norma = " · ".join(n for c in cs for n in c["normatividad"]) or "la Ley 675 de 2001"
        sub = f"Conforme al {norma}, se auditó la documentación de actas de consejo de administración del año {self.p.anio}."
        filas = []
        for a in A["actas"]:
            pres = a["firma_presidente"] or "Sin firma"
            filas.append([f"{a['numero']}\n{a['fecha']}".strip(), a["temas"],
                          (pres, {} if "sin firma" not in norm(pres) else {"color": NEG, "bold": True}),
                          a["firma_administracion"] or "—"])
        cols = [("N. acta y fecha", 0.95, "c"), ("Temas", 5.5, "l"), ("Firma presidente", 1.35, "c"),
                ("Firma administración", 1.35, "c")]
        n0 = self.n
        self.tabla_paginada(item.get("titulo", "Libros Actas de Consejo de Administración"), sub, cols, filas,
                            [a["fuente"] for a in A["actas"]], y=Y0 + 0.05, size=6.5)
        nivel = self._nivel_peor([c["nivel"] for c in cs]) if cs else ""
        for s in list(self.prs.slides)[n0:self.n]:
            self.badge(s, nivel)

    def d_asamblea(self, item):
        bloques = (self.doc.get("asamblea") or {}).get("bloques") or []
        clave = norm(item.get("bloque", "ordinaria"))
        b = next((b for b in reversed(bloques) if clave in norm(b["titulo"])), None)
        if not b:
            self.avisos.append("No se encontró el bloque de decisiones de asamblea.")
            self.d_sin_datos(
                "Decisiones de la Asamblea", "Votación y seguimiento",
                "No se encontró un bloque de decisiones de asamblea para el periodo auditado."
            )
            return
        filas = [[f["tema"], (f["detalle"], {"color": NEG, "bold": True} if norm(f["detalle"]).startswith("no aprueba")
                              else {"bold": True})] for f in b["filas"]]
        self.tabla_paginada(b["titulo"][0].upper() + b["titulo"][1:], "Votación", [("Pregunta", 5.6, "l"),
                                                                                     ("Resultado", 3.55, "l")],
                            filas, [b["fuente"]], size=8)

    # ---------------------------------------------------------------- control interno
    def d_contratos(self, item):
        C = self.doc.get("contratos") or {}
        if not C.get("contratos"):
            self.avisos.append("Sin contratos en la hoja CONTRATOS 2026.")
            self.d_sin_datos(
                item.get("titulo", "Contratos Vigentes"), f"Contratos de {self.cfg.get('copropiedad', '')}",
                "La hoja de contratos no contiene registros para revisar objeto, vigencia, valor, pólizas y observaciones."
            )
            return
        filas = []
        for c in C["contratos"]:
            for k, pt in enumerate(c["partes"]):
                def plata(v):
                    return fmt_pesos(v) if isinstance(v, (int, float)) else str(v or "")
                obs = pt.get("observaciones") or ""
                seg = (pt.get("seg") or {}).get("observaciones")
                filas.append([c["n"] if k == 0 else "", pt["proveedor"], pt["vigencia"],
                              (plata(pt["valor"]), {"seg": pt["seg"].get("valor")} if pt["seg"].get("valor") else {}),
                              plata(pt["total"]), (obs, {"seg": seg}) if seg else obs])
        cols = [("N.", 0.3, "c"), ("Proveedor / objeto contractual", 2.0, "l"), ("Vigencia", 1.15, "l"),
                ("Valor", 1.15, "l"), ("Total contrato", 0.95, "r"), ("Pólizas y garantías – observaciones", 3.6, "l")]
        self.tabla_paginada(item.get("titulo", "Contratos Vigentes"), f"{C['n']} contratos · {self.cfg.get('copropiedad', '')}",
                            cols, filas, [c["fuente"] for c in C["contratos"]], size=6)

    def d_polizas_manual(self, item):
        P = self.doc.get("polizas") or {}
        if P.get("filas"):
            filas = [[f["n"], f["contrato"], (f["verificacion"], {} if f["cumple"] else {"color": NEG, "bold": True})]
                     for f in P["filas"]]
            self.tabla_paginada(item.get("titulo", "Verificación de Pólizas"),
                                f"{len(P['filas']) - len(P['no_cumplen'])} de {len(P['filas'])} contratos cumplen las pólizas "
                                "mínimas del manual de contratación",
                                [("N.", 0.35, "c"), ("Contrato", 4.9, "l"), ("Verificación de pólizas de acuerdo al manual", 3.9, "l")],
                                filas, [f["fuente"] for f in P["filas"]], size=7)
        else:
            self.d_sin_datos(
                item.get("titulo", "Verificación de Pólizas frente al Manual de Contratación"),
                "Coberturas y vigencias exigidas",
                "No se encontraron contratos con información suficiente para verificar sus pólizas."
            )
        M = P.get("matriz") or {}
        if M.get("filas"):
            s = self._nueva((M.get("titulo") or "Matriz de lineamientos por tipo de contrato").strip(" 0123456789").capitalize(),
                            item.get("subtitulo_matriz", "Matriz extraída del manual de adquisición y contratación"), [])
            n = len(M["columnas"])
            anchos = [1.45, 1.75, 1.05, 1.35, 1.55, 1.25, 0.75][:n]
            self.tabla(s, 0.42, Y0, [(t, w, "l") for t, w in zip(M["columnas"], anchos)],
                       [[x for x in f[:n]] for f in M["filas"]], size=7.5, alto_fila=0.26, alto_enc=0.3)
        else:
            self.d_sin_datos(
                "Matriz de lineamientos por tipo de contrato",
                item.get("subtitulo_matriz", "Matriz extraída del manual de adquisición y contratación"),
                "No se encontró la matriz del manual con los requisitos aplicables a cada tipo de contrato."
            )

    def d_matriz_riesgo(self, item):
        M = self.doc.get("matriz_riesgo") or {}
        if not M.get("filas"):
            self.d_sin_datos(
                "Matriz de riesgo, Ambiental, Mantenimiento y Salubridad",
                "Seguimiento semestral del nivel de cumplimiento",
                "No se encontró la matriz de riesgos y acciones de control en los papeles de trabajo."
            )
            return
        filas = []
        for f in M["filas"]:
            segs = []
            for k, d in enumerate(f["descripcion"]):
                if k:
                    segs.append({"t": "\n", "b": False, "rojo": False})
                segs += d["seg"] or [{"t": d["texto"], "b": False, "rojo": False}]
            texto = "\n".join(d["texto"] for d in f["descripcion"])
            fila = [(f["objetivo"], {"bold": True, "color": NAVY}), (texto, {"seg": segs}), f["valoracion"]]
            for per in M["periodos"]:
                nv = f["niveles"].get(per, "")
                col = self.sem.get(nv, {}).get("color")
                fila.append((nv or "—", {"fondo": _rgb(col), "bold": True, "color": TEXTO if nv == "MEDIO" else BLANCO}
                             if col else {}))
            filas.append(fila)
        cols = [("Objetivo de control", 1.25, "l"), ("Descripción del riesgo y acciones recomendadas", 5.45, "l"),
                ("Valoración del riesgo", 0.85, "c")] + [(f"Nivel {p}", 0.8, "c") for p in M["periodos"]]
        self.tabla_paginada(M.get("titulo") or "Matriz de riesgo", "Seguimiento semestral del nivel de cumplimiento",
                            cols, filas, [f["fuente"] for f in M["filas"]], size=6)

    # ---------------------------------------------------------------- orquestación
    def _recortar(self, n: int):
        """Quita las diapositivas que una sección fallida alcanzó a agregar después de la n-ésima."""
        lista = self.prs.slides._sldIdLst
        while len(lista) > n:
            sld = lista[-1]
            self.prs.part.drop_rel(sld.rId)
            lista.remove(sld)
        self.n = n

    def financiera(self):
        self.seccion(f"Auditoría Financiera a {self.p.etiqueta}.")
        nombres = {"d_monetizacion": "monetización del fondo de imprevistos", "d_liquidez": "liquidez",
                   "d_cartera": "cartera", "d_concentracion": "concentración de cartera",
                   "d_mayores_deudores": "mayores deudores", "d_anticipos": "anticipos", "d_cxp": "cuentas por pagar",
                   "d_resultado": "resultado del ejercicio", "d_excedentes": "excedentes acumulados"}
        # secciones que cada diapositiva necesita: si el Excel no las trae (o no se pudieron leer), la diapositiva
        # se conserva como pendiente con el motivo, para que el informe mantenga la estructura estándar
        requiere = {"d_monetizacion": ("liquidez",), "d_liquidez": ("liquidez",), "d_cartera": ("cartera", "deudores"),
                    "d_concentracion": ("cartera",), "d_resultado": ("presupuesto",), "d_excedentes": ("patrimonio",)}
        faltantes = self.m.get("faltantes") or {}
        for nombre in nombres:
            faltan = [k for k in requiere.get(nombre, ()) if not self.m.get(k)]
            if faltan:
                self.d_sin_datos(nombres[nombre][:1].upper() + nombres[nombre][1:], f"Auditoría financiera a {self.p.etiqueta}",
                                 " ".join(_motivo(k, faltantes.get(k)) for k in faltan))
                self.avisos.append(f"Diapositiva de {nombres[nombre]} pendiente: falta {', '.join(faltan)}.")
                continue
            antes = self.n
            try:                                       # una diapositiva que falla no tumba la sección financiera
                getattr(self, nombre)()
                self.marcar(antes, SECCIONES_DIAPOSITIVA.get(nombre, ()))
            except Exception as e:                     # …y tampoco desaparece: queda como pendiente, en su lugar
                self._recortar(antes)
                self.d_sin_datos(nombres[nombre][:1].upper() + nombres[nombre][1:], f"Auditoría financiera a {self.p.etiqueta}",
                                 "Los datos de esta sección se leyeron, pero no se pudo armar la diapositiva con ellos. "
                                 "Revise la sección en el Excel o informe a soporte.")
                self.errores.append(f"No se generó la diapositiva de {nombres[nombre]}: {type(e).__name__} {e}"[:200])

    def construir(self):
        for item in self.est["estructura"]:
            t = item["tipo"]
            antes = self.n
            try:
                if t == "portada":
                    self.d_portada()
                elif t == "metodologia":
                    self.d_metodologia()
                elif t == "seccion":
                    self.seccion(item["titulo"].format(periodo=self.p.etiqueta))
                elif t == "objetivo_alcance":
                    self.d_objetivo_alcance()
                elif t == "egresos":
                    self.d_egresos()
                elif t == "campo":
                    self.d_campo(item)
                elif t == "conciliaciones":
                    self.d_conciliaciones(item)
                elif t == "retencion":
                    self.d_retencion(item)
                elif t == "actas_consejo":
                    self.d_actas_consejo(item)
                elif t == "asamblea":
                    self.d_asamblea(item)
                elif t == "financiera":
                    self.financiera()
                elif t == "contratos":
                    self.d_contratos(item)
                elif t == "polizas_manual":
                    self.d_polizas_manual(item)
                elif t == "matriz_riesgo":
                    self.d_matriz_riesgo(item)
                elif t == "texto_fijo":
                    self.d_texto_fijo(item["clave"])
                elif t == "cierre":
                    self.d_cierre()
            except Exception as e:                       # una sección con problemas no tumba el informe ni desaparece
                self._recortar(antes)
                self.d_sin_datos(item.get("titulo") or t.replace("_", " ").capitalize(), item.get("subtitulo") or "",
                                 "No se pudo armar esta diapositiva con los datos leídos del Excel. Se conserva para mantener "
                                 "la estructura estándar del informe.")
                self.errores.append(f"No se pudo generar «{item.get('titulo') or t}»: {type(e).__name__} {e}"[:200])
            if t in SECCIONES_TIPO:
                self.marcar(antes, SECCIONES_TIPO[t])
            self.plan.append({"tipo": t, "titulo": item.get("titulo") or item.get("campo") or t,
                              "desde": antes + 1, "diapositivas": self.n - antes})
        omitidos = {_clave(x) for x in self.est.get("omitir_campos", [])}
        for c in ((self.doc.get("auditoria") or {}).get("campos") or []):
            if c["campo"] not in self.usados and _clave(c["campo"]) not in omitidos:
                self.avisos.append(f"El campo «{c['campo']}» está en el Excel pero no tiene lugar en la estructura del informe.")
        return self
