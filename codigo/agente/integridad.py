"""Huellas, versiones e identidad del expediente mensual."""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import unicodedata
from datetime import datetime
from pathlib import Path

from .comun import RECIBIDO, ahora, guardar_json, leer_json


def huella_archivo(ruta: Path) -> str:
    h = hashlib.sha256()
    with Path(ruta).open("rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def _huella_objeto(valor) -> str:
    crudo = json.dumps(valor, ensure_ascii=False, sort_keys=True, default=str).encode()
    return hashlib.sha256(crudo).hexdigest()


def manifiesto(esp, periodo: str, cid: str) -> dict:
    """Identidad inmutable de los insumos que alimentan una validación."""
    ent = esp.entrada(periodo, cid)
    archivos = []
    for p in sorted(x for x in ent.iterdir() if x.is_file() and not x.name.startswith(".")) if ent.exists() else []:
        archivos.append({"nombre": p.name, "bytes": p.stat().st_size, "sha256": huella_archivo(p)})
    ficha = esp.ficha(cid)
    mapa = leer_json(esp.clientes / cid / "mapa.json", {})
    base = {"periodo": periodo, "cliente": cid, "archivos": archivos,
            "ficha_sha256": _huella_objeto(ficha), "mapa_sha256": _huella_objeto(mapa)}
    base["revision"] = _huella_objeto(base)[:16]
    return base


def misma_revision(esp, periodo: str, cid: str, esperada: dict | None) -> bool:
    return bool(esperada and manifiesto(esp, periodo, cid)["revision"] == esperada.get("revision"))


def archivar_revision(esp, periodo: str, cid: str) -> Path | None:
    """Conserva entradas, salidas y decisión previas antes de recibir nuevos archivos."""
    ent, sal = esp.entrada(periodo, cid), esp.salida(periodo, cid)
    existentes = [p for carpeta in (ent, sal) if carpeta.exists() for p in carpeta.rglob("*") if p.is_file()]
    if not existentes:
        return None
    est = esp.estado(periodo, cid)
    revision = (est.get("manifiesto") or {}).get("revision") or datetime.now().strftime("%Y%m%d%H%M%S")
    destino = esp.raiz / "historial" / periodo / cid / revision
    destino.mkdir(parents=True, exist_ok=True)
    if ent.exists():
        shutil.copytree(ent, destino / "entradas", dirs_exist_ok=True)
    if sal.exists():
        shutil.copytree(sal, destino / "salidas", dirs_exist_ok=True)
    guardar_json(destino / "archivo_revision.json", {"archivada": ahora(), "estado": est.get("estado"),
                                                       "revision": revision})
    return destino


def reiniciar_por_carga(esp, periodo: str, cid: str, actor: str, nombres: list[str], archivar: bool = True) -> dict:
    if archivar:
        archivar_revision(esp, periodo, cid)
    est = {"cliente": cid, "periodo": periodo, "estado": RECIBIDO, "bitacora": []}
    esp.evento(est, RECIBIDO, "Nueva versión recibida: " + ", ".join(nombres), actor=actor)
    est["manifiesto_carga"] = manifiesto(esp, periodo, cid)
    esp.guardar_estado(periodo, cid, est)
    return est


def _norm(t) -> str:
    s = unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def _nits_declarados(libro) -> list[tuple[str, str]]:
    """Extrae NIT solo cuando el libro lo identifica expresamente como tal."""
    encontrados: list[tuple[str, str]] = []

    def numeros(valor) -> list[str]:
        salida = []
        for m in re.finditer(r"(?<!\d)\d(?:[\d.\- ]{6,18})\d(?!\d)", str(valor or "")):
            digitos = re.sub(r"\D", "", m.group(0))
            if 8 <= len(digitos) <= 10:
                salida.append(digitos)
        return salida

    for ws in libro.valores.worksheets:
        max_fila, max_col = min(ws.max_row, 80), min(ws.max_column, 20)
        for fila in ws.iter_rows(min_row=1, max_row=max_fila, max_col=max_col):
            for c in fila:
                if not isinstance(c.value, str) or not re.search(r"\bnit\b", _norm(c.value)):
                    continue
                candidatos = [(c.coordinate, c.value)]
                candidatos.extend(
                    (ws.cell(c.row, col).coordinate, ws.cell(c.row, col).value)
                    for col in range(c.column + 1, min(c.column + 3, max_col) + 1)
                )
                for coordenada, valor in candidatos:
                    encontrados.extend((nit, f"{ws.title}!{coordenada}") for nit in numeros(valor))
    return list(dict.fromkeys(encontrados))


def verificar_identidad(libro, ficha: dict, modelo: dict, periodo: str) -> list[str]:
    """Devuelve contradicciones que deben impedir generar un informe."""
    textos = []
    for ws in libro.valores.worksheets:
        for fila in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 80), max_col=min(ws.max_column, 20)):
            for c in fila:
                if isinstance(c.value, str) and c.value.strip():
                    textos.append((ws.title, c.coordinate, c.value.strip()))
    universo = _norm("\n".join(t for _, _, t in textos))
    problemas = []
    nombre = _norm(ficha.get("copropiedad"))
    palabras = [w for w in nombre.split() if len(w) >= 4 and w not in {"conjunto", "residencial", "edificio", "propiedad", "horizontal"}]
    coincide_nombre = bool(palabras and all(w in universo for w in palabras[:4]))
    candidatos = [(h, c, t) for h, c, t in textos if re.search(r"\b(conjunto|edificio|agrupacion)\b.*\b(ph|propiedad horizontal)\b", _norm(t))]
    es_demo = "(demo)" in str(ficha.get("copropiedad", "")).lower()
    if palabras and not coincide_nombre and candidatos and not es_demo:
        # se muestra el nombre más corto que declara el archivo (un título, no un párrafo de evidencia)
        cortos = [(h, c, t) for h, c, t in textos if len(t) <= 60
                  and re.match(r"(conjunto|edificio|agrupacion)\b", _norm(t))]
        muestra = min(cortos or candidatos, key=lambda x: len(x[2]))
        problemas.append(f"El archivo no identifica a «{ficha.get('copropiedad')}»; declara «{muestra[2]}» en {muestra[0]}!{muestra[1]}.")
    nit = re.sub(r"\D", "", str(ficha.get("nit") or ""))
    if len(nit) >= 8:
        declarados = _nits_declarados(libro)
        coincide = any(nit == encontrado or (len(nit) == 10 and nit[:-1] == encontrado)
                       or (len(encontrado) == 10 and encontrado[:-1] == nit)
                       for encontrado, _ in declarados)
        if declarados and not coincide:
            valores = ", ".join(f"{valor} ({fuente})" for valor, fuente in declarados[:3])
            problemas.append(
                f"El NIT de la ficha ({ficha.get('nit')}) no coincide con el NIT declarado en el archivo: {valores}."
            )
    esperadas = ficha.get("unidades")
    encontradas = (modelo.get("cartera") or {}).get("unidades_hoja")
    if esperadas and encontradas and int(esperadas) != int(encontradas):
        problemas.append(f"La ficha registra {esperadas} unidades y el Excel declara {int(encontradas)}.")
    if modelo.get("meta", {}).get("periodo") != periodo:
        problemas.append(f"El archivo se extrajo para {modelo.get('meta', {}).get('periodo')} y el expediente corresponde a {periodo}.")
    return problemas


def cambiar_nit(esp, cid: str, nit_nuevo: str, actor: str) -> str:
    """Reemplaza un NIT provisional por el real. Como el NIT es el identificador del cliente, mueve su carpeta y la de
    cada expediente (entradas, salidas, historial) y conserva todo. No se permite si ya hay informes aprobados: esos
    entregables se emitieron con el NIT anterior y no se reescriben."""
    import yaml
    from .comun import APROBADO, mostrar_nit, normalizar_nit
    nuevo = normalizar_nit(nit_nuevo)
    origen = esp.clientes / cid
    if not origen.exists():
        raise ValueError(f"No existe el cliente «{cid}».")
    if nuevo == cid:
        raise ValueError("El NIT nuevo es igual al actual.")
    if (esp.clientes / nuevo).exists():
        raise ValueError(f"Ya existe un cliente con NIT {mostrar_nit(nuevo)}.")
    aprobados = [p.parent.parent.name for p in (esp.raiz / "salidas").glob(f"*/{cid}/estado.json")
                 if leer_json(p, {}).get("estado") == APROBADO]
    if aprobados:
        raise ValueError("El cliente tiene informes aprobados con el NIT anterior (" + ", ".join(sorted(aprobados))
                         + "). Esos entregables no se reescriben; cree un cliente nuevo con el NIT correcto.")
    movidos = []
    for base in ("entradas", "salidas", "historial"):
        for d in sorted((esp.raiz / base).glob(f"*/{cid}")):
            d.rename(d.parent / nuevo)
            movidos.append(f"{base}/{d.parent.name}")
    origen.rename(esp.clientes / nuevo)
    ruta = esp.clientes / nuevo / "ficha.yaml"
    ficha = yaml.safe_load(ruta.read_text(encoding="utf-8")) or {}
    anterior = ficha.get("nit")
    ficha.update({"id": nuevo, "nit": mostrar_nit(nuevo), "nit_provisional": False,
                  "nit_anterior": anterior, "nit_cambiado": ahora()})
    ruta.write_text(f"# Ficha del cliente (NIT cambiado por {actor}: {anterior} → {mostrar_nit(nuevo)}).\n"
                    + yaml.safe_dump(ficha, allow_unicode=True, sort_keys=False, width=110), encoding="utf-8")
    for p in (esp.raiz / "salidas").glob(f"*/{nuevo}/estado.json"):
        est = leer_json(p, {})
        est["cliente"] = nuevo
        esp.evento(est, None, f"NIT cambiado de {anterior} a {mostrar_nit(nuevo)}: vuelva a validar", actor=actor)
        guardar_json(p, est)
    return nuevo
