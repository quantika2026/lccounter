"""Usuarios, roles y sesiones de la aplicación LCCOUNTER.

- Usuarios en `<raiz>/usuarios.json` con la contraseña derivada (PBKDF2-SHA256, sal por usuario). Nunca en claro.
- Roles: asistente, revisor, auditor, admin (un usuario puede tener varios).
- Sesión: cookie HttpOnly + SameSite=Strict con un token aleatorio; vence a las 12 horas de inactividad.
- Cada acción relevante queda en `<raiz>/registro_accesos.jsonl` (quién, qué, cuándo).

Administración por consola (recuperación si se pierde el acceso del administrador):
  python -m app.seguridad --raiz <espacio> crear <usuario> --nombre "..." --roles admin,revisor
  python -m app.seguridad --raiz <espacio> clave <usuario>
"""
from __future__ import annotations

import hashlib
import hmac
import json
import re
import secrets
import threading
import time
from pathlib import Path

ROLES = {
    "asistente": "Carga archivos, valida, crea clientes y consulta",
    "revisor": "Revisor fiscal: edita ficha, borrador y diapositivas; aprueba o devuelve; aprueba mejoras",
    "auditor": "Segundo revisor: muestra mensual y conclusión de auditoría",
    "admin": "Administra usuarios",
}
# Qué puede hacer cada rol (es lo que la API exige en cada ruta; se muestra en la pantalla de usuarios)
PERMISOS = [
    ("Consultar clientes, informes y a Auditin", ["asistente", "revisor", "auditor", "admin"]),
    ("Crear clientes, subir su logo y sus documentos", ["asistente", "revisor"]),
    ("Cargar el Excel y validar (correr el flujo)", ["asistente", "revisor"]),
    ("Editar ficha, ficha de normatividad, borrador de opinión y diapositivas", ["revisor"]),
    ("Aprobar o devolver el informe (queda congelado al aprobar)", ["revisor"]),
    ("Aprobar recetas, correcciones del mapa y mejoras del sistema", ["revisor"]),
    ("Emitir conceptos del revisor fiscal a todos los clientes", ["revisor"]),
    ("Auditoría mensual por muestreo (no la del informe que aprobó)", ["auditor", "revisor"]),
    ("Ver el registro de acciones y los tiempos", ["revisor", "admin"]),
    ("Administrar usuarios y la configuración de IA", ["admin"]),
]
RX_USUARIO = re.compile(r"^[a-z0-9._-]{3,32}$")
ITERACIONES = 390_000
DURACION = 12 * 3600
COOKIE = "lccounter_sesion"

_candado = threading.Lock()
_sesiones: dict[str, dict] = {}            # token → {usuario, vence}
_fallos: dict[str, list[float]] = {}       # usuario → instantes de intentos fallidos


def _derivar(clave: str, sal: bytes) -> str:
    return hashlib.pbkdf2_hmac("sha256", clave.encode(), sal, ITERACIONES).hex()


def validar_clave(clave: str) -> str | None:
    if len(clave) < 10:
        return "La contraseña debe tener al menos 10 caracteres."
    if clave.lower() == clave or not re.search(r"\d", clave):
        return "La contraseña debe combinar mayúsculas, minúsculas y números."
    return None


class Usuarios:
    def __init__(self, raiz: Path):
        self.ruta = Path(raiz) / "usuarios.json"
        self.registro = Path(raiz) / "registro_accesos.jsonl"

    # --------------------------------------------------------------- almacenamiento
    def _leer(self) -> dict:
        try:
            return json.loads(self.ruta.read_text(encoding="utf-8"))
        except FileNotFoundError:
            return {}

    def _guardar(self, datos: dict):
        self.ruta.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.ruta.with_suffix(".tmp")
        tmp.write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
        tmp.chmod(0o600)
        tmp.replace(self.ruta)

    def hay_usuarios(self) -> bool:
        return bool(self._leer())

    def publico(self, u: str) -> dict | None:
        d = self._leer().get(u)
        return None if not d else {"usuario": u, "nombre": d["nombre"], "roles": d["roles"], "activo": d.get("activo", True),
                                   "creado": d.get("creado"), "ultimo_acceso": d.get("ultimo_acceso")}

    def listar(self) -> list[dict]:
        return [self.publico(u) for u in sorted(self._leer())]

    # --------------------------------------------------------------- gestión
    def crear(self, usuario: str, nombre: str, roles: list[str], clave: str) -> dict:
        usuario = usuario.strip().lower()
        if not RX_USUARIO.match(usuario):
            raise ValueError("Usuario inválido: 3 a 32 caracteres (minúsculas, números, punto, guion).")
        if not nombre.strip():
            raise ValueError("Indique el nombre completo.")
        roles = sorted(set(roles))
        if not roles or any(r not in ROLES for r in roles):
            raise ValueError(f"Roles válidos: {', '.join(ROLES)}.")
        err = validar_clave(clave)
        if err:
            raise ValueError(err)
        with _candado:
            datos = self._leer()
            if usuario in datos:
                raise ValueError(f"Ya existe el usuario «{usuario}».")
            if any(d["nombre"].strip().lower() == nombre.strip().lower() for d in datos.values()):
                raise ValueError("Ya hay un usuario con ese nombre: el nombre identifica quién aprueba y audita.")
            sal = secrets.token_bytes(16)
            datos[usuario] = {"nombre": nombre.strip(), "roles": roles, "sal": sal.hex(), "clave": _derivar(clave, sal),
                              "activo": True, "creado": time.strftime("%Y-%m-%dT%H:%M:%S"), "cambiar_clave": True}
            self._guardar(datos)
        return self.publico(usuario)

    def actualizar(self, usuario: str, roles: list[str] | None = None, activo: bool | None = None) -> dict:
        with _candado:
            datos = self._leer()
            if usuario not in datos:
                raise ValueError(f"No existe el usuario «{usuario}».")
            if roles is not None:
                if not roles or any(r not in ROLES for r in roles):
                    raise ValueError(f"Roles válidos: {', '.join(ROLES)}.")
                datos[usuario]["roles"] = sorted(set(roles))
            if activo is not None:
                datos[usuario]["activo"] = activo
            admins = [u for u, d in datos.items() if "admin" in d["roles"] and d.get("activo", True)]
            if not admins:
                raise ValueError("Debe quedar al menos un administrador activo.")
            self._guardar(datos)
        if activo is False:
            cerrar_sesiones(usuario)
        return self.publico(usuario)

    def poner_clave(self, usuario: str, clave: str, temporal: bool):
        err = validar_clave(clave)
        if err:
            raise ValueError(err)
        with _candado:
            datos = self._leer()
            if usuario not in datos:
                raise ValueError(f"No existe el usuario «{usuario}».")
            sal = secrets.token_bytes(16)
            datos[usuario].update({"sal": sal.hex(), "clave": _derivar(clave, sal), "cambiar_clave": temporal})
            self._guardar(datos)
        cerrar_sesiones(usuario)

    def verificar(self, usuario: str, clave: str) -> dict | None:
        """Devuelve el usuario si la contraseña es correcta. Frena intentos repetidos (5 por 15 minutos)."""
        usuario = usuario.strip().lower()
        ahora = time.time()
        recientes = [t for t in _fallos.get(usuario, []) if ahora - t < 900]
        if len(recientes) >= 5:
            raise PermissionError("Demasiados intentos fallidos. Espere 15 minutos o pida al administrador que restablezca la contraseña.")
        d = self._leer().get(usuario)
        ok = bool(d) and d.get("activo", True) and hmac.compare_digest(_derivar(clave, bytes.fromhex(d["sal"])), d["clave"])
        if not d:
            _derivar(clave, b"0" * 16)                          # mismo tiempo de respuesta exista o no el usuario
        if not ok:
            _fallos[usuario] = recientes + [ahora]
            return None
        _fallos.pop(usuario, None)
        with _candado:
            datos = self._leer()
            datos[usuario]["ultimo_acceso"] = time.strftime("%Y-%m-%dT%H:%M:%S")
            self._guardar(datos)
        return {**self.publico(usuario), "cambiar_clave": d.get("cambiar_clave", False)}

    def debe_cambiar_clave(self, usuario: str) -> bool:
        return bool(self._leer().get(usuario, {}).get("cambiar_clave"))

    def anotar(self, usuario: str, accion: str, detalle: str = ""):
        with _candado, self.registro.open("a", encoding="utf-8") as f:
            f.write(json.dumps({"fecha": time.strftime("%Y-%m-%dT%H:%M:%S"), "usuario": usuario, "accion": accion,
                                "detalle": detalle[:300]}, ensure_ascii=False) + "\n")

    # ------------------------------------------------------------ tiempo en línea
    def latido(self, usuario: str, visible: bool = True):
        """La app abierta y visible avisa cada minuto. Se suma el tiempo entre latidos seguidos (máx. 2,5 min):
        una pestaña cerrada o en segundo plano no cuenta."""
        ruta = self.ruta.parent / "tiempo_en_linea.json"
        with _candado:
            try:
                datos = json.loads(ruta.read_text(encoding="utf-8"))
            except (FileNotFoundError, ValueError):
                datos = {}
            ahora_ = time.time()
            u = datos.setdefault(usuario, {"ultimo": 0, "dias": {}})
            delta = ahora_ - u.get("ultimo", 0)
            if visible and 0 < delta <= 150:
                dia = time.strftime("%Y-%m-%d", time.localtime(ahora_))
                u["dias"][dia] = round(u["dias"].get(dia, 0) + delta)
            u["ultimo"] = ahora_ if visible else 0
            tmp = ruta.with_suffix(".tmp")
            tmp.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")
            tmp.replace(ruta)

    def tiempo_en_linea(self, dias: int = 30) -> dict:
        """{usuario: {"hoy": s, "semana": s, "mes": s, "dias": {fecha: s}}} de los últimos `dias` días."""
        ruta = self.ruta.parent / "tiempo_en_linea.json"
        try:
            datos = json.loads(ruta.read_text(encoding="utf-8"))
        except (FileNotFoundError, ValueError):
            return {}
        hoy = time.time()
        fecha = lambda d: time.strftime("%Y-%m-%d", time.localtime(hoy - d * 86400))
        out = {}
        for usuario, u in datos.items():
            ds = u.get("dias", {})
            out[usuario] = {"hoy": ds.get(fecha(0), 0), "semana": sum(ds.get(fecha(i), 0) for i in range(7)),
                            "mes": sum(ds.get(fecha(i), 0) for i in range(dias)),
                            "en_linea": hoy - u.get("ultimo", 0) <= 150}
        return out

    def ultimos_registros(self, n: int = 200) -> list[dict]:
        if not self.registro.exists():
            return []
        return [json.loads(x) for x in self.registro.read_text(encoding="utf-8").splitlines()[-n:] if x.strip()][::-1]


# ------------------------------------------------------------------- sesiones
# Se guardan en disco (solo la huella SHA-256 del token, permisos 600) para que reiniciar la app no cierre la sesión
# de quien la está usando. Sin `persistir_sesiones()` (pruebas, consola) viven solo en memoria.
_archivo_sesiones: Path | None = None


def _huella(token: str) -> str:
    import hashlib
    return hashlib.sha256(token.encode()).hexdigest()


def persistir_sesiones(raiz: Path):
    global _archivo_sesiones
    _archivo_sesiones = Path(raiz) / "sesiones.json"
    try:
        datos = json.loads(_archivo_sesiones.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        datos = {}
    ahora_ = time.time()
    with _candado:
        _sesiones.update({h: s for h, s in datos.items() if s.get("vence", 0) > ahora_})


def _guardar_sesiones():
    """Se llama con _candado tomado."""
    if not _archivo_sesiones:
        return
    tmp = _archivo_sesiones.with_suffix(".tmp")
    tmp.write_text(json.dumps(_sesiones), encoding="utf-8")
    tmp.chmod(0o600)
    tmp.replace(_archivo_sesiones)


def abrir_sesion(usuario: str) -> str:
    token = secrets.token_urlsafe(32)
    with _candado:
        _sesiones[_huella(token)] = {"usuario": usuario, "vence": time.time() + DURACION}
        _guardar_sesiones()
    return token


def sesion(token: str | None) -> str | None:
    if not token:
        return None
    h = _huella(token)
    with _candado:
        s = _sesiones.get(h)
        if not s or s["vence"] < time.time():
            if _sesiones.pop(h, None):
                _guardar_sesiones()
            return None
        nuevo = time.time() + DURACION
        if nuevo - s["vence"] > 600:                     # renovar en disco cada 10 min, no en cada solicitud
            s["vence"] = nuevo
            _guardar_sesiones()
        return s["usuario"]


def cerrar_sesion(token: str | None):
    with _candado:
        if _sesiones.pop(_huella(token or ""), None):
            _guardar_sesiones()


def cerrar_sesiones(usuario: str):
    with _candado:
        for t in [t for t, s in _sesiones.items() if s["usuario"] == usuario]:
            del _sesiones[t]
        _guardar_sesiones()


# ------------------------------------------------------------------- consola

def main():
    import argparse
    import getpass
    ap = argparse.ArgumentParser(description="Usuarios de la aplicación LCCOUNTER")
    ap.add_argument("--raiz", required=True)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("crear"); s.add_argument("usuario"); s.add_argument("--nombre", required=True)
    s.add_argument("--roles", required=True, help=f"separados por coma: {', '.join(ROLES)}")
    s = sub.add_parser("clave"); s.add_argument("usuario")
    sub.add_parser("listar")
    a = ap.parse_args()
    U = Usuarios(Path(a.raiz))
    try:
        if a.cmd == "listar":
            for u in U.listar():
                print(f"{u['usuario']:<16} {u['nombre']:<32} {','.join(u['roles']):<28} {'activo' if u['activo'] else 'inactivo'}")
            return
        clave = getpass.getpass("Contraseña: ")
        if clave != getpass.getpass("Repita la contraseña: "):
            raise ValueError("Las contraseñas no coinciden.")
        if a.cmd == "crear":
            U.crear(a.usuario, a.nombre, [r.strip() for r in a.roles.split(",")], clave)
            print(f"Usuario «{a.usuario}» creado. Deberá cambiar la contraseña al entrar.")
        else:
            U.poner_clave(a.usuario, clave, temporal=True)
            print(f"Contraseña de «{a.usuario}» restablecida.")
        U.anotar("consola", f"usuarios.{a.cmd}", a.usuario)
    except ValueError as e:
        raise SystemExit(str(e)) from None


if __name__ == "__main__":
    main()
