"""Utilidades compartidas: rutas del espacio de trabajo, JSON y bitácora."""
from __future__ import annotations

import json
import os
import re
import tempfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import yaml

# Estados del ciclo de vida de un informe
RECIBIDO = "RECIBIDO"
BLOQUEADO = "BLOQUEADO"
PENDIENTE = "PENDIENTE_VALIDACION"
APROBADO = "APROBADO"
DEVUELTO = "DEVUELTO"
ESTADOS = [RECIBIDO, BLOQUEADO, PENDIENTE, DEVUELTO, APROBADO]


def ahora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _ser(o):
    if isinstance(o, (date, datetime)):
        return o.isoformat()
    if isinstance(o, (bytes, bytearray)):
        return f"<imagen {len(o)} bytes>"
    if hasattr(o, "dict"):
        return o.dict()
    raise TypeError(type(o))


def guardar_json(ruta: Path, datos) -> Path:
    """Escribe JSON de forma atómica: nunca deja un archivo parcial como estado válido."""
    ruta.parent.mkdir(parents=True, exist_ok=True)
    contenido = json.dumps(datos, default=_ser, ensure_ascii=False, indent=1)
    fd, temporal = tempfile.mkstemp(prefix=f".{ruta.name}.", suffix=".tmp", dir=ruta.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(contenido)
            f.flush()
            os.fsync(f.fileno())
        Path(temporal).replace(ruta)
    except Exception:
        Path(temporal).unlink(missing_ok=True)
        raise
    return ruta


def leer_json(ruta: Path, defecto=None):
    try:
        return json.loads(Path(ruta).read_text(encoding="utf-8"))
    except FileNotFoundError:
        return defecto
    except json.JSONDecodeError as e:
        raise ValueError(f"El archivo de estado está dañado: {ruta} ({e})") from e


def leer_yaml(ruta: Path, defecto=None):
    try:
        return yaml.safe_load(Path(ruta).read_text(encoding="utf-8")) or (defecto if defecto is not None else {})
    except FileNotFoundError:
        return defecto


def slug(t: str) -> str:
    import unicodedata
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "_", t).strip("_")[:40] or "cliente"


def normalizar_nit(nit: str) -> str:
    """Identificador estable del cliente: nueve dígitos del NIT y su dígito de verificación."""
    valor = re.sub(r"\D", "", str(nit or ""))
    if len(valor) != 10:
        raise ValueError("El NIT debe incluir nueve dígitos y el dígito de verificación (ejemplo: 900.123.456-7).")
    return valor


def mostrar_nit(nit: str) -> str:
    valor = normalizar_nit(nit)
    return f"{valor[:3]}.{valor[3:6]}.{valor[6:9]}-{valor[9]}"


@dataclass
class Espacio:
    """Estructura de carpetas del despacho.

    raiz/
      clientes/<nit>/ficha.yaml         ficha permanente del cliente (software, notas, revisor)
      entradas/<periodo>/<nit>/         lo que llega del cliente: Excel estándar y/o exportes contables
      salidas/<periodo>/<nit>/          informe, validación, re-auditoría, borrador y estado
      conocimiento/                     memoria del sistema: correcciones, golden set, biblioteca, versiones
    """
    raiz: Path

    def __post_init__(self):
        self.raiz = Path(self.raiz).resolve()

    @property
    def clientes(self) -> Path:
        return self.raiz / "clientes"

    @property
    def conocimiento(self) -> Path:
        return self.raiz / "conocimiento"

    def entrada(self, periodo: str, cid: str) -> Path:
        return self.raiz / "entradas" / periodo / cid

    def salida(self, periodo: str, cid: str) -> Path:
        return self.raiz / "salidas" / periodo / cid

    def ficha(self, cid: str) -> dict:
        f = leer_yaml(self.clientes / cid / "ficha.yaml")
        if f is None:
            raise FileNotFoundError(f"No existe la ficha del cliente «{cid}» en {self.clientes / cid}")
        f.setdefault("id", cid)
        return f

    def ids_clientes(self) -> list[str]:
        fichas = sorted(self.clientes.glob("*/ficha.yaml"))
        # orden del calendario de visitas si la ficha lo trae
        return [p.parent.name for p in sorted(fichas, key=lambda p: ((leer_yaml(p) or {}).get("orden", 999), p.parent.name))]

    def estado(self, periodo: str, cid: str) -> dict:
        return leer_json(self.salida(periodo, cid) / "estado.json",
                         {"cliente": cid, "periodo": periodo, "estado": RECIBIDO, "bitacora": []})

    def guardar_estado(self, periodo: str, cid: str, est: dict) -> None:
        guardar_json(self.salida(periodo, cid) / "estado.json", est)

    def evento(self, est: dict, estado: str | None, detalle: str, actor: str = "agente") -> dict:
        if estado:
            est["estado"] = estado
        est.setdefault("bitacora", []).append({"fecha": ahora(), "actor": actor, "estado": est["estado"],
                                               "detalle": detalle})
        return est

    def periodo_anterior(self, periodo: str) -> str:
        a, m = map(int, periodo.split("-"))
        a, m = (a - 1, 12) if m == 1 else (a, m - 1)
        return f"{a:04d}-{m:02d}"
