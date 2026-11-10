"""LCCOUNTER · aplicación web local del agente de Revisoría Fiscal PH.

  python -m app.servidor [--raiz ../espacio-trabajo] [--puerto 8790]

Sirve la interfaz (app/static) y una API sobre el mismo espacio de trabajo que usa `python -m agente`.
Nada nuevo se calcula aquí: cada acción llama a los agentes existentes. El asistente conversacional
(Claude) solo lee el espacio de trabajo con herramientas de lectura; no aprueba, no edita y no calcula cifras.

Acceso con usuario y contraseña (app/seguridad.py). Quien aprueba, devuelve o audita queda registrado con el
nombre de su sesión, nunca con un nombre escrito a mano.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import threading
from datetime import datetime
from pathlib import Path

RAIZ_CODIGO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ_CODIGO))

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, Response, UploadFile  # noqa: E402
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse  # noqa: E402
from fastapi.staticfiles import StaticFiles  # noqa: E402
from pydantic import BaseModel  # noqa: E402

from agente import auditoria, mejora, orquestador  # noqa: E402
from agente.__main__ import FICHA_NUEVA  # noqa: E402
from agente.comun import Espacio, ahora, guardar_json, leer_json, leer_yaml, mostrar_nit, normalizar_nit  # noqa: E402

from . import asistente, comunicaciones, conceptos, diapositivas, mapeo, normativa, proveedores, seguridad  # noqa: E402

RX_ID = re.compile(r"^[a-z0-9_]{1,40}$")
RX_PERIODO = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
DESCARGABLES = {".pptx", ".pdf", ".html", ".md", ".json", ".xlsx", ".docx"}
PUBLICAS = {"/api/sesion", "/api/login", "/api/configurar"}

app = FastAPI(title="LCCOUNTER", docs_url=None, redoc_url=None, openapi_url=None)
ESP: Espacio = Espacio(Path(os.environ.get("REVISORIA_RAIZ", RAIZ_CODIGO.parent / "espacio-trabajo")))
USU = seguridad.Usuarios(ESP.raiz)
_candado = threading.Lock()          # una corrida de lote a la vez


def _id(v: str) -> str:
    if not RX_ID.match(v or ""):
        raise HTTPException(400, f"Identificador inválido «{v}» (minúsculas, números y _).")
    return v


def _periodo(v: str) -> str:
    if not RX_PERIODO.match(v or ""):
        raise HTTPException(400, f"Periodo inválido «{v}» (AAAA-MM).")
    return v


def _salir(fn, *a, **k):
    """Los agentes señalan reglas de negocio con SystemExit: se devuelven como 409 con su mensaje."""
    try:
        return fn(*a, **k)
    except SystemExit as e:
        raise HTTPException(409, str(e)) from None


# ----------------------------------------------------------------------------- acceso

@app.middleware("http")
async def control_acceso(request: Request, call_next):
    ruta = request.url.path
    if ruta.startswith("/api/"):
        # Acciones solo desde la propia aplicación: otra página abierta en el navegador no puede actuar por el usuario
        if request.method not in ("GET", "HEAD"):
            origen = request.headers.get("origin") or request.headers.get("referer") or ""
            propio = f"{request.url.scheme}://{request.headers.get('host', '')}"
            if origen and not origen.startswith(propio):
                return JSONResponse({"detail": "Origen no permitido."}, status_code=403)
        if ruta not in PUBLICAS:
            usuario = seguridad.sesion(request.cookies.get(seguridad.COOKIE))
            if not usuario or not USU.publico(usuario) or not USU.publico(usuario)["activo"]:
                return JSONResponse({"detail": "Sesión vencida o inexistente. Ingrese de nuevo."}, status_code=401)
            if USU.debe_cambiar_clave(usuario) and ruta not in ("/api/clave", "/api/logout"):
                return JSONResponse({"detail": "Debe cambiar su contraseña antes de continuar."}, status_code=403)
            request.state.usuario = USU.publico(usuario)
    respuesta = await call_next(request)
    if not ruta.startswith("/api/"):                 # la interfaz siempre en su versión actual (sin caché vieja)
        respuesta.headers["Cache-Control"] = "no-cache"
    respuesta.headers["X-Content-Type-Options"] = "nosniff"
    respuesta.headers["X-Frame-Options"] = "DENY"
    respuesta.headers["Referrer-Policy"] = "same-origin"
    return respuesta


def requiere(*roles: str):
    """Dependencia: el usuario de la sesión debe tener alguno de los roles indicados."""
    def dep(request: Request) -> dict:
        u = getattr(request.state, "usuario", None)
        if not u:
            raise HTTPException(401, "Ingrese de nuevo.")
        if roles and not set(roles) & set(u["roles"]):
            nombres = " o ".join(roles)
            raise HTTPException(403, f"Su usuario no tiene permiso para esta acción (requiere rol {nombres}).")
        return u
    return dep


CUALQUIERA = requiere()
OPERA = requiere("asistente", "revisor")
REVISOR = requiere("revisor")
AUDITA = requiere("auditor", "revisor")
ADMIN = requiere("admin")


def _cookie(resp: Response, token: str):
    resp.set_cookie(seguridad.COOKIE, token, httponly=True, samesite="strict", max_age=seguridad.DURACION, path="/")


class Login(BaseModel):
    usuario: str
    clave: str


@app.get("/api/sesion")
def ver_sesion(request: Request):
    usuario = seguridad.sesion(request.cookies.get(seguridad.COOKIE))
    u = USU.publico(usuario) if usuario else None
    if u:
        u["avatar"] = bool(_avatar(u["usuario"]))
    return {"configurado": USU.hay_usuarios(), "usuario": u if u and u["activo"] else None,
            "cambiar_clave": bool(u) and USU.debe_cambiar_clave(usuario), "roles": seguridad.ROLES}


class Configurar(BaseModel):
    usuario: str
    nombre: str
    clave: str


@app.post("/api/configurar")
def configurar(c: Configurar, response: Response):
    """Primer uso: crea el administrador. Solo funciona mientras no exista ningún usuario."""
    if USU.hay_usuarios():
        raise HTTPException(409, "La aplicación ya está configurada.")
    try:
        u = USU.crear(c.usuario, c.nombre, ["admin", "revisor"], c.clave)
        USU.poner_clave(u["usuario"], c.clave, temporal=False)
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "configurar", "administrador inicial")
    _cookie(response, seguridad.abrir_sesion(u["usuario"]))
    return {"ok": True, "usuario": USU.publico(u["usuario"])}


@app.post("/api/login")
def login(l: Login, response: Response):
    try:
        u = USU.verificar(l.usuario, l.clave)
    except PermissionError as e:
        raise HTTPException(429, str(e)) from None
    if not u:
        USU.anotar(l.usuario.strip().lower()[:40], "login_fallido")
        raise HTTPException(401, "Usuario o contraseña incorrectos.")
    USU.anotar(u["usuario"], "login")
    _cookie(response, seguridad.abrir_sesion(u["usuario"]))
    return {"ok": True, "usuario": u}


@app.post("/api/logout")
def logout(request: Request, response: Response):
    seguridad.cerrar_sesion(request.cookies.get(seguridad.COOKIE))
    response.delete_cookie(seguridad.COOKIE, path="/")
    return {"ok": True}


class CambioClave(BaseModel):
    actual: str
    nueva: str


@app.post("/api/clave")
def cambiar_clave(c: CambioClave, request: Request, response: Response, u=Depends(CUALQUIERA)):
    if c.actual == c.nueva:
        raise HTTPException(400, "La nueva contraseña debe ser distinta de la actual.")
    try:
        if not USU.verificar(u["usuario"], c.actual):
            raise HTTPException(400, "La contraseña actual no es correcta.")
        USU.poner_clave(u["usuario"], c.nueva, temporal=False)
    except PermissionError as e:
        raise HTTPException(429, str(e)) from None
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "cambio_clave")
    _cookie(response, seguridad.abrir_sesion(u["usuario"]))       # las demás sesiones quedaron cerradas
    return {"ok": True}


# --- administración de usuarios

class NuevoUsuario(BaseModel):
    usuario: str
    nombre: str
    roles: list[str]
    clave: str


class CambioUsuario(BaseModel):
    roles: list[str] | None = None
    activo: bool | None = None
    clave: str | None = None


@app.get("/api/usuarios")
def usuarios(u=Depends(ADMIN)):
    return {"usuarios": USU.listar(), "roles": seguridad.ROLES, "registro": USU.ultimos_registros(150),
            "tiempo": USU.tiempo_en_linea(), "permisos": seguridad.PERMISOS}


class Latido(BaseModel):
    visible: bool = True


@app.post("/api/latido")
def latido(l: Latido, u=Depends(CUALQUIERA)):
    USU.latido(u["usuario"], l.visible)
    return {"ok": True}


# ----------------------------------------------------------------------------- conceptos del revisor (a todos los clientes)

class BuscarFragmentos(BaseModel):
    consulta: str


@app.post("/api/conceptos/fragmentos")
def fragmentos_concepto(b: BuscarFragmentos, u=Depends(REVISOR)):
    if len(b.consulta.strip()) < 3:
        raise HTTPException(400, "Escriba qué norma o tema buscar.")
    return {"fragmentos": normativa.instancia(ESP.raiz).buscar(b.consulta, None, k=12)}


class NuevoConcepto(BaseModel):
    tema: str
    opinion: str
    fragmentos: list[dict]                 # [{ambito, doc_id, fragmento}]: el texto se toma de la base normativa
    clientes: list[str] = []


def _fragmentos_de_base(elegidos: list[dict]) -> list[dict]:
    base = normativa.instancia(ESP.raiz)
    out = []
    for e in elegidos[:20]:
        try:
            meta, frs = base.fragmentos(str(e.get("ambito", "")), str(e.get("doc_id", "")))
        except FileNotFoundError:
            raise HTTPException(400, "Un fragmento elegido ya no está en la base normativa.") from None
        f = next((x for x in frs if str(x["id"]) == str(e.get("fragmento"))), None)
        if not f:
            raise HTTPException(400, "Un fragmento elegido ya no está en la base normativa.")
        out.append({"documento": meta["titulo"], "version": meta.get("version", ""), "referencia": f["ref"],
                    "texto": f["texto"], "doc_id": meta["id"], "ambito": meta["ambito"], "fragmento": f["id"]})
    return out


@app.get("/api/conceptos")
def ver_conceptos(u=Depends(CUALQUIERA)):
    return {"conceptos": conceptos.listar(ESP), "disponible": proveedores.disponible(ESP.raiz, "comunicaciones")}


@app.post("/api/conceptos")
def crear_concepto(c: NuevoConcepto, u=Depends(REVISOR)):
    try:
        d = conceptos.crear(ESP, c.tema, c.opinion, _fragmentos_de_base(c.fragmentos), [_id(x) for x in c.clientes], u["nombre"])
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    except proveedores.ErrorIA as e:
        raise HTTPException(503, str(e)) from None
    USU.anotar(u["usuario"], "concepto.crear", f"{d['id']}: {d['tema'][:120]}")
    return d


@app.get("/api/conceptos/{cid}")
def ver_concepto(cid: str, u=Depends(CUALQUIERA)):
    try:
        return conceptos.leer(ESP, cid)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e)) from None


class EdicionConcepto(BaseModel):
    asunto: str | None = None
    texto: str | None = None
    infografia: dict | None = None
    clientes: list[str] | None = None


@app.put("/api/conceptos/{cid}")
def editar_concepto(cid: str, c: EdicionConcepto, u=Depends(REVISOR)):
    try:
        d = conceptos.guardar(ESP, cid, c.model_dump(), u["nombre"])
    except FileNotFoundError as e:
        raise HTTPException(404, str(e)) from None
    except ValueError as e:
        raise HTTPException(409, str(e)) from None
    USU.anotar(u["usuario"], "concepto.editar", cid)
    return d


@app.get("/api/conceptos/{cid}/infografia.png")
def previa_infografia(cid: str, u=Depends(CUALQUIERA)):
    try:
        d = conceptos.leer(ESP, cid)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e)) from None
    ruta = Path(tempfile.mkdtemp()) / "infografia.png"
    conceptos.infografia(d, ruta, Path(__file__).parent / "static" / "logo.png")
    return FileResponse(ruta, media_type="image/png", headers={"Cache-Control": "no-store"})


@app.post("/api/conceptos/{cid}/emitir")
def emitir_concepto(cid: str, u=Depends(REVISOR)):
    try:
        d = conceptos.emitir(ESP, cid, u["nombre"])
    except FileNotFoundError as e:
        raise HTTPException(404, str(e)) from None
    USU.anotar(u["usuario"], "concepto.emitir", f"{cid}: {len(d['emitido']['archivos'])} copropiedad(es)")
    return d


@app.get("/api/conceptos/{cid}/archivo/{nombre}")
def archivo_concepto(cid: str, nombre: str, u=Depends(CUALQUIERA)):
    try:
        ruta = conceptos.archivo(ESP, cid, nombre)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e)) from None
    return FileResponse(ruta, filename=ruta.name)


@app.get("/api/registro")
def registro(u=Depends(requiere("admin", "revisor"))):
    return {"registro": USU.ultimos_registros(300)}


@app.get("/api/tiempos")
def tiempos(u=Depends(requiere("admin", "revisor"))):
    """Tiempo en línea por usuario y tiempo de procesamiento por cliente en cada periodo."""
    proc = []
    for carpeta in sorted((ESP.raiz / "salidas").glob("*")):
        if not carpeta.is_dir():
            continue
        for cid in ESP.ids_clientes():
            e = ESP.estado(carpeta.name, cid)
            seg = (e.get("resumen") or {}).get("segundos")
            if seg is not None:
                proc.append({"periodo": carpeta.name, "cliente": cid, "copropiedad": e.get("copropiedad", cid),
                             "segundos": seg, "estado": e.get("estado")})
    nombres = {x["usuario"]: x["nombre"] for x in USU.listar()}
    return {"en_linea": [{"usuario": k, "nombre": nombres.get(k, k), **v} for k, v in USU.tiempo_en_linea().items()],
            "procesamiento": proc}


@app.post("/api/usuarios")
def crear_usuario(n: NuevoUsuario, u=Depends(ADMIN)):
    try:
        nuevo = USU.crear(n.usuario, n.nombre, n.roles, n.clave)
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "usuario.crear", f"{nuevo['usuario']} ({','.join(nuevo['roles'])})")
    return {"ok": True, "usuario": nuevo}


@app.patch("/api/usuarios/{usuario}")
def cambiar_usuario(usuario: str, c: CambioUsuario, u=Depends(ADMIN)):
    if usuario == u["usuario"] and (c.activo is False or (c.roles is not None and "admin" not in c.roles)):
        raise HTTPException(400, "No puede desactivarse ni quitarse el rol de administrador a sí mismo.")
    try:
        if c.clave:
            USU.poner_clave(usuario, c.clave, temporal=True)
        r = USU.actualizar(usuario, c.roles, c.activo)
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "usuario.cambiar", f"{usuario}: roles={c.roles} activo={c.activo} clave={'sí' if c.clave else 'no'}")
    return {"ok": True, "usuario": r}


# ----------------------------------------------------------------------------- lectura

@app.get("/api/resumen")
def resumen(u=Depends(CUALQUIERA)):
    periodos = sorted({p.name for p in (ESP.raiz / "salidas").glob("*") if p.is_dir()}
                      | {p.name for p in (ESP.raiz / "entradas").glob("*") if p.is_dir()}, reverse=True)
    ids = ESP.ids_clientes()
    return {"raiz": str(ESP.raiz), "periodos": periodos, "clientes": ids,
            "nombres": {cid: ESP.ficha(cid).get("copropiedad") or cid for cid in ids},
            "logos": [cid for cid in ids if _logo_cliente(cid)],
            "asistente": asistente.disponible(ESP.raiz), "usuario": u}


# Foto o avatar de cada usuario: se guarda en el espacio de trabajo y solo la cambia su dueño.
AVATAR_TIPOS = {"image/png": ".png", "image/jpeg": ".jpg", "image/webp": ".webp", "image/gif": ".gif"}
RX_USUARIO = re.compile(r"^[a-z0-9][a-z0-9._\-]{2,31}$")


def _avatar(usuario: str) -> Path | None:
    if not RX_USUARIO.match(usuario or ""):
        raise HTTPException(400, "Usuario inválido.")
    return next((p for p in (ESP.raiz / "avatares").glob(usuario + ".*") if p.suffix in AVATAR_TIPOS.values()), None)


@app.get("/api/avatar/{usuario}")
def ver_avatar(usuario: str, u=Depends(CUALQUIERA)):
    ruta = _avatar(usuario)
    if not ruta:
        raise HTTPException(404, "Sin avatar.")
    return FileResponse(ruta, headers={"Cache-Control": "no-cache"})


@app.post("/api/avatar")
async def subir_avatar(imagen: UploadFile = File(...), u=Depends(CUALQUIERA)):
    ext = AVATAR_TIPOS.get(imagen.content_type or "")
    if not ext:
        raise HTTPException(400, "Use una imagen PNG, JPG, WEBP o GIF.")
    datos = await imagen.read(2 * 1024 * 1024 + 1)
    if len(datos) > 2 * 1024 * 1024:
        raise HTTPException(400, "La imagen pesa más de 2 MB.")
    previa = _avatar(u["usuario"])
    if previa:
        previa.unlink()
    carpeta = ESP.raiz / "avatares"
    carpeta.mkdir(parents=True, exist_ok=True)
    (carpeta / (u["usuario"] + ext)).write_bytes(datos)
    USU.anotar(u["usuario"], "usuario.avatar", "cambió su imagen")
    return {"ok": True}


@app.delete("/api/avatar")
def borrar_avatar(u=Depends(CUALQUIERA)):
    previa = _avatar(u["usuario"])
    if previa:
        previa.unlink()
    return {"ok": True}


@app.get("/api/lote/{periodo}")
def lote(periodo: str, u=Depends(CUALQUIERA)):
    _periodo(periodo)
    filas = []
    for cid in ESP.ids_clientes():
        e = ESP.estado(periodo, cid)
        f = ESP.ficha(cid)
        ent = ESP.entrada(periodo, cid)
        filas.append({"cliente": cid, "nit": f.get("nit", ""), "copropiedad": f.get("copropiedad", cid), "revisor": f.get("revisor", ""),
                      "software": f.get("software", ""), "estado": e.get("estado"),
                      "resumen": e.get("resumen"), "archivos": e.get("archivos", {}),
                      "ultimo": (e.get("bitacora") or [{}])[-1],
                      "entradas": sorted(p.name for p in ent.glob("*") if p.is_file()) if ent.exists() else []})
    cola = [e["cliente"] for e in orquestador.cola(ESP, periodo)]
    return {"periodo": periodo, "clientes": filas, "cola": cola}


def _rutas_borrador(periodo: str, cid: str) -> tuple[Path, Path, Path]:
    sal = ESP.salida(periodo, cid)
    return (sal / f"Borrador_Opinion_{cid}_{periodo}.md", sal / f"Borrador_Editado_{cid}_{periodo}.md",
            sal / "borrador_editado.json")


def _huella(ruta: Path) -> str:
    import hashlib
    return hashlib.sha256(ruta.read_bytes()).hexdigest()[:16] if ruta.exists() else ""


def _borrador_editado(periodo: str, cid: str) -> dict | None:
    """El borrador editado y si sigue basado en el borrador generado actual (una nueva validación lo invalida)."""
    gen, edit, meta = _rutas_borrador(periodo, cid)
    if not edit.exists():
        return None
    m = leer_json(meta, {})
    return {"texto": edit.read_text(encoding="utf-8"), "guardado": m.get("fecha"), "por": m.get("por", ""),
            "desactualizado": m.get("base") != _huella(gen)}


@app.get("/api/cliente/{periodo}/{cid}")
def cliente(periodo: str, cid: str, u=Depends(CUALQUIERA)):
    _periodo(periodo), _id(cid)
    sal = ESP.salida(periodo, cid)
    gen = _rutas_borrador(periodo, cid)[0]
    return {"ficha": ESP.ficha(cid), "estado": ESP.estado(periodo, cid),
            "hallazgos": leer_json(sal / "hallazgos.json", []), "reauditoria": leer_json(sal / "reauditoria.json", []),
            "cuadre": leer_json(sal / "cuadre_balance.json", []),
            "borrador": gen.read_text(encoding="utf-8") if gen.exists() else "",
            "borrador_editado": _borrador_editado(periodo, cid)}


@app.get("/api/archivo/{periodo}/{cid}/{nombre}")
def archivo(periodo: str, cid: str, nombre: str, u=Depends(CUALQUIERA)):
    _periodo(periodo), _id(cid)
    sal = ESP.salida(periodo, cid).resolve()
    ruta = (sal / nombre).resolve()
    if ruta.parent != sal or ruta.suffix.lower() not in DESCARGABLES or not ruta.exists():
        raise HTTPException(404, "Archivo no disponible.")
    return FileResponse(ruta, filename=ruta.name)


@app.get("/api/mejora")
def ver_mejora(u=Depends(CUALQUIERA)):
    c = ESP.conocimiento
    return {"propuestas": leer_json(c / "propuestas.json", []), "versiones": leer_json(c / "versiones.json", []),
            "metricas": (leer_json(c / "metricas.json", []) or [None])[-1],
            "regresion": leer_json(c / "ultima_regresion.json"),
            "biblioteca": leer_yaml(c / "biblioteca_opiniones.yaml", {}) or {},
            "borradores": sorted(p.name for p in (c / "propuestas").glob("*.md")) if (c / "propuestas").exists() else []}


# ----------------------------------------------------------------------------- carga y validación

LOGO_TIPOS = {"image/png": ".png", "image/jpeg": ".jpg", "image/webp": ".webp", "image/svg+xml": ".svg"}


def _logo_cliente(cid: str) -> Path | None:
    return next((p for p in (ESP.clientes / cid).glob("logo.*") if p.suffix in LOGO_TIPOS.values()), None)


async def _guardar_logo(cid: str, logo: UploadFile):
    ext = LOGO_TIPOS.get(logo.content_type or "")
    if not ext:
        raise HTTPException(400, "El logo debe ser PNG, JPG, WEBP o SVG.")
    datos = await logo.read(3 * 1024 * 1024 + 1)
    if len(datos) > 3 * 1024 * 1024:
        raise HTTPException(400, "El logo pesa más de 3 MB.")
    previo = _logo_cliente(cid)
    if previo:
        previo.unlink()
    (ESP.clientes / cid / ("logo" + ext)).write_bytes(datos)


@app.get("/api/clientes/{cid}/logo")
def ver_logo(cid: str, u=Depends(CUALQUIERA)):
    ruta = _logo_cliente(_id(cid))
    if not ruta:
        raise HTTPException(404, "El cliente no tiene logo.")
    # un SVG se sirve como imagen y sin ejecutar nada: la página lo muestra en <img>
    return FileResponse(ruta, headers={"Cache-Control": "no-cache", "Content-Security-Policy": "script-src 'none'"})


@app.post("/api/clientes/{cid}/logo")
async def cambiar_logo(cid: str, logo: UploadFile = File(...), u=Depends(OPERA)):
    ESP.ficha(_id(cid))
    await _guardar_logo(cid, logo)
    USU.anotar(u["usuario"], "cliente.logo", cid)
    return {"ok": True}


# Documentos y ficha de normatividad de cada cliente (existen desde que se crea el cliente)
DOC_EXT = {".pdf", ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".txt", ".md", ".png", ".jpg", ".jpeg", ".csv", ".zip"}
NORMATIVIDAD_BASE = {"reglamento": "", "escritura": "", "marco_contable": "", "normas": [], "manuales": [],
                     "obligaciones": "", "observaciones": "", "actualizada_por": "", "fecha": ""}


def _dir_docs(cid: str) -> Path:
    ESP.ficha(_id(cid))
    d = ESP.clientes / cid / "documentos"
    d.mkdir(exist_ok=True)
    return d


def _nombre_doc(nombre: str) -> str:
    limpio = re.sub(r"[^\w.\- ()áéíóúÁÉÍÓÚñÑ]", "_", Path(nombre or "").name).strip(" .")
    if not limpio or Path(limpio).suffix.lower() not in DOC_EXT:
        raise HTTPException(400, "Tipo de archivo no permitido. Use PDF, Word, Excel, PowerPoint, imagen, texto o ZIP.")
    return limpio


@app.get("/api/clientes/{cid}/documentos")
def ver_documentos(cid: str, u=Depends(CUALQUIERA)):
    d = _dir_docs(cid)
    return {"documentos": [{"nombre": x.name, "bytes": x.stat().st_size,
                            "fecha": datetime.fromtimestamp(x.stat().st_mtime).isoformat(timespec="minutes")}
                           for x in sorted(d.iterdir(), key=lambda x: x.name.lower()) if x.is_file()]}


@app.post("/api/clientes/{cid}/documentos")
async def subir_documentos(cid: str, archivos: list[UploadFile] = File(...), u=Depends(OPERA)):
    d = _dir_docs(cid)
    guardados = []
    for up in archivos:
        nombre = _nombre_doc(up.filename)
        datos = await up.read(60 * 1024 * 1024 + 1)
        if len(datos) > 60 * 1024 * 1024:
            raise HTTPException(400, f"«{nombre}» pesa más de 60 MB.")
        (d / nombre).write_bytes(datos)
        guardados.append(nombre)
    USU.anotar(u["usuario"], "cliente.documentos", f"{cid}: {', '.join(guardados)}")
    return {"ok": True, "guardados": guardados}


@app.get("/api/clientes/{cid}/documentos/{nombre}")
def bajar_documento(cid: str, nombre: str, u=Depends(CUALQUIERA)):
    ruta = _dir_docs(cid) / _nombre_doc(nombre)
    if not ruta.exists():
        raise HTTPException(404, "No existe ese documento.")
    return FileResponse(ruta, filename=ruta.name)


@app.delete("/api/clientes/{cid}/documentos/{nombre}")
def borrar_documento(cid: str, nombre: str, u=Depends(REVISOR)):
    ruta = _dir_docs(cid) / _nombre_doc(nombre)
    if ruta.exists():
        hist = ESP.clientes / cid / "historial" / "documentos"           # no se pierde: pasa al historial
        hist.mkdir(parents=True, exist_ok=True)
        ruta.replace(hist / f"{datetime.now():%Y%m%d_%H%M%S}_{ruta.name}")
    USU.anotar(u["usuario"], "cliente.documento.retirar", f"{cid}: {nombre}")
    return {"ok": True}


@app.get("/api/clientes/{cid}/normatividad")
def ver_normatividad(cid: str, u=Depends(CUALQUIERA)):
    ESP.ficha(_id(cid))
    return {**NORMATIVIDAD_BASE, **leer_json(ESP.clientes / cid / "normatividad.json", {})}


class Normatividad(BaseModel):
    reglamento: str = ""
    escritura: str = ""
    marco_contable: str = ""
    normas: list[str] = []
    manuales: list[str] = []
    obligaciones: str = ""
    observaciones: str = ""


@app.put("/api/clientes/{cid}/normatividad")
def guardar_normatividad(cid: str, n: Normatividad, u=Depends(REVISOR)):
    ESP.ficha(_id(cid))
    datos = {**n.model_dump(), "actualizada_por": u["nombre"], "fecha": ahora()}
    datos["normas"] = [x.strip() for x in datos["normas"] if x.strip()][:80]
    datos["manuales"] = [x.strip() for x in datos["manuales"] if x.strip()][:80]
    guardar_json(ESP.clientes / cid / "normatividad.json", datos)
    USU.anotar(u["usuario"], "cliente.normatividad", cid)
    return datos


@app.post("/api/clientes")
async def nuevo_cliente(nit: str = Form(...), copropiedad: str = Form(...), software: str = Form("generico"),
                        logo: UploadFile | None = File(None), u=Depends(OPERA)):
    import yaml

    try:
        cid = normalizar_nit(nit)
        nit_visible = mostrar_nit(cid)
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    d = ESP.clientes / cid
    if (d / "ficha.yaml").exists():
        raise HTTPException(409, f"Ya existe el cliente con NIT {nit_visible}.")
    d.mkdir(parents=True, exist_ok=True)
    ficha = {"id": cid, "copropiedad": copropiedad.strip(), "nit": nit_visible, "unidades": None,
             "firma": "LC Counter Contadores y Auditores SAS", "numero_informe": "", "revisor": "",
             "orden": 99, "software": _id(software),
             "entradas": {"excel": "*.xlsx", "balance": {"archivo": "balance*", "perfil": _id(software)}},
             "notas": {"concentracion_cartera": "", "opinion_anticipos": "", "presupuesto": {}}}
    (d / "ficha.yaml").write_text(yaml.safe_dump(ficha, allow_unicode=True, sort_keys=False), encoding="utf-8")
    (d / "documentos").mkdir(exist_ok=True)               # carpeta de documentos y ficha de normatividad desde su creación
    guardar_json(d / "normatividad.json", {**NORMATIVIDAD_BASE, "actualizada_por": u["nombre"], "fecha": ahora()})
    if logo and logo.filename:
        await _guardar_logo(cid, logo)
    USU.anotar(u["usuario"], "cliente.crear", f"{cid} ({nit_visible})")
    return {"ok": True, "id": cid, "nit": nit_visible, "ficha": str(d / "ficha.yaml")}


@app.post("/api/cargar")
async def cargar(periodo: str = Form(...), cliente: str = Form(...), excel: UploadFile | None = File(None),
                 balance: UploadFile | None = File(None), u=Depends(OPERA)):
    from agente import integridad

    _periodo(periodo), _id(cliente)
    ESP.ficha(cliente)
    if ESP.estado(periodo, cliente).get("estado") == "APROBADO":
        raise HTTPException(409, "El informe de este periodo ya está aprobado y congelado: no se puede cargar otra "
                                 "versión ni modificarlo. Si hay un error, se corrige en el informe del periodo siguiente.")
    ent = ESP.entrada(periodo, cliente)
    ent.mkdir(parents=True, exist_ok=True)
    recibidos = [(up, tipo) for up, tipo in ((excel, "excel"), (balance, "balance")) if up and up.filename]
    if not recibidos:
        raise HTTPException(400, "No se recibió ningún archivo.")
    preparados = []
    for up, tipo in recibidos:
        ext = Path(up.filename).suffix.lower()
        permitidas = {".xlsx", ".xlsm"} if tipo == "excel" else {".csv", ".tsv", ".txt", ".xlsx"}
        if ext not in permitidas:
            raise HTTPException(400, f"{up.filename}: formato no admitido para {tipo} ({', '.join(sorted(permitidas))}).")
        nombre = f"Excel_{cliente}_{periodo}{ext}" if tipo == "excel" else f"balance_prueba_{periodo}{ext}"
        preparados.append((up, tipo, nombre))

    cargas = ESP.raiz / ".cargas"
    cargas.mkdir(parents=True, exist_ok=True)
    guardados = []
    with tempfile.TemporaryDirectory(prefix="carga-", dir=cargas) as td:
        temporal = Path(td)
        for up, _, nombre in preparados:
            with (temporal / nombre).open("wb") as f:
                shutil.copyfileobj(up.file, f)
        excel_temporal = next((temporal / nombre for _, tipo, nombre in preparados if tipo == "excel"), None)
        if excel_temporal:
            try:
                from revisoria.excel import Libro
                from revisoria.extractor import detectar_periodo
                from revisoria import mapa as mapa_cliente
                periodo_excel = detectar_periodo(Libro(str(excel_temporal), mapa_cliente.cargar(ESP.clientes / cliente)))
            except (KeyError, ValueError, TypeError):
                periodo_excel = None
            if periodo_excel and periodo_excel.clave != periodo:
                raise HTTPException(
                    409,
                    f"El Excel corresponde a {periodo_excel.etiqueta}, pero el periodo seleccionado es "
                    f"{periodo}. Seleccione {periodo_excel.etiqueta} y vuelva a cargar el archivo."
                )
        integridad.archivar_revision(ESP, periodo, cliente)
        for _, tipo, nombre in preparados:
            for viejo in ent.glob("*.xls*"):
                if tipo == "excel" and not viejo.name.startswith("balance"):
                    viejo.unlink()
            for viejo in ent.glob("balance*"):
                if tipo == "balance":
                    viejo.unlink()
            (temporal / nombre).replace(ent / nombre)
            guardados.append(nombre)
    integridad.reiniciar_por_carga(ESP, periodo, cliente, u["nombre"], guardados, archivar=False)
    USU.anotar(u["usuario"], "cargar", f"{cliente} {periodo}: {', '.join(guardados)}")
    return {"ok": True, "guardados": guardados, "revision": integridad.manifiesto(ESP, periodo, cliente)["revision"]}


class Validar(BaseModel):
    periodo: str
    clientes: list[str] | None = None


@app.post("/api/validar")
def validar(v: Validar, u=Depends(OPERA)):
    _periodo(v.periodo)
    ids = [_id(c) for c in v.clientes] if v.clientes else None
    if not _candado.acquire(blocking=False):
        raise HTTPException(409, "Ya hay un lote corriendo; espere a que termine.")
    try:
        log = []
        res = orquestador.correr(ESP, v.periodo, ids, paralelo=4, log=log.append)
    finally:
        _candado.release()
    from agente import panel
    panel.escribir(ESP, v.periodo)
    USU.anotar(u["usuario"], "validar", f"{v.periodo}: {', '.join(ids) if ids else 'todo el lote'}")
    return {"log": log, "estados": {r["cliente"]: r["estado"] for r in res}}


# ----------------------------------------------------------------------------- decisión del revisor

class Decision(BaseModel):
    periodo: str
    cliente: str
    justificacion: str = ""
    motivo: str = ""


@app.post("/api/aprobar")
def aprobar(d: Decision, u=Depends(REVISOR)):
    _periodo(d.periodo), _id(d.cliente)
    editado = _borrador_editado(d.periodo, d.cliente)
    if editado and editado["desactualizado"]:
        raise HTTPException(409, "El borrador editado se basa en una validación anterior: el informe se volvió a "
                                 "validar después. Revíselo y guárdelo de nuevo, o descártelo, antes de aprobar.")
    final = str(_rutas_borrador(d.periodo, d.cliente)[1]) if editado else None
    rp = _rutas_informe(d.periodo, d.cliente)
    inf = diapositivas.estado_edicion(rp) if rp else None
    if inf and inf["desactualizado"]:
        raise HTTPException(409, "Las ediciones de las diapositivas se hicieron sobre un informe anterior: el informe "
                                 "se volvió a validar después. Revíselas o descártelas antes de aprobar.")
    e = _salir(orquestador.aprobar, ESP, d.periodo, d.cliente, u["nombre"], d.justificacion.strip(),
               pptx_final=str(rp["editado"]) if inf else None, borrador_final=final)
    USU.anotar(u["usuario"], "aprobar", f"{d.cliente} {d.periodo}")
    return {"ok": True, "detalle": e["bitacora"][-1]["detalle"]}


@app.post("/api/devolver")
def devolver(d: Decision, u=Depends(REVISOR)):
    _periodo(d.periodo), _id(d.cliente)
    if not d.motivo.strip():
        raise HTTPException(400, "Indique el motivo de la devolución.")
    e = _salir(orquestador.devolver, ESP, d.periodo, d.cliente, u["nombre"], d.motivo.strip())
    USU.anotar(u["usuario"], "devolver", f"{d.cliente} {d.periodo}: {d.motivo.strip()}")
    return {"ok": True, "detalle": e["bitacora"][-1]["detalle"]}


class Borrador(BaseModel):
    periodo: str
    cliente: str
    texto: str


@app.put("/api/borrador")
def guardar_borrador(b: Borrador, u=Depends(REVISOR)):
    _periodo(b.periodo), _id(b.cliente)
    if ESP.estado(b.periodo, b.cliente).get("estado") != "PENDIENTE_VALIDACION":
        raise HTTPException(409, "Solo se edita el borrador de un informe pendiente de validación.")
    gen, edit, meta = _rutas_borrador(b.periodo, b.cliente)
    if not gen.exists():
        raise HTTPException(404, "No hay borrador generado para este cliente y periodo.")
    if len(b.texto) > 200_000:
        raise HTTPException(400, "El borrador es demasiado largo.")
    edit.write_text(b.texto.replace("\r\n", "\n"), encoding="utf-8")
    guardar_json(meta, {"fecha": ahora(), "por": u["nombre"], "base": _huella(gen)})
    USU.anotar(u["usuario"], "borrador.guardar", f"{b.cliente} {b.periodo}")
    return {"ok": True, "editado": _borrador_editado(b.periodo, b.cliente)}


@app.delete("/api/borrador/{periodo}/{cid}")
def descartar_borrador(periodo: str, cid: str, u=Depends(REVISOR)):
    _periodo(periodo), _id(cid)
    for r in _rutas_borrador(periodo, cid)[1:]:
        r.unlink(missing_ok=True)
    USU.anotar(u["usuario"], "borrador.descartar", f"{cid} {periodo}")
    return {"ok": True}


# ----------------------------------------------------------------------------- visor de diapositivas

def _rutas_informe(periodo: str, cid: str) -> dict | None:
    gen = (ESP.estado(periodo, cid).get("archivos") or {}).get("pptx")
    return diapositivas.rutas(ESP.salida(periodo, cid), cid, periodo, gen) if gen else None


@app.get("/api/diapositivas/{periodo}/{cid}")
def ver_diapositivas(periodo: str, cid: str, version: str = "actual", u=Depends(CUALQUIERA)):
    """Imágenes y textos editables. version: «actual» (editado si existe), «generado» o «editado»."""
    _periodo(periodo), _id(cid)
    r = _rutas_informe(periodo, cid)
    if not r or not r["generado"].exists():
        raise HTTPException(404, "Este cliente aún no tiene informe en el periodo: valídelo primero.")
    ed = diapositivas.estado_edicion(r)
    usar = "editado" if (version == "editado" or version == "actual") and ed else "generado"
    try:
        h, n = diapositivas.render(r[usar], r["vista"])
    except RuntimeError as ex:
        raise HTTPException(503, str(ex)) from None
    unidades = diapositivas.unidades(r[usar])
    originales = {x["n"]: x["u"] for x in diapositivas.unidades(r["generado"])} if usar == "editado" else None
    return {"version": usar, "huella": h, "n": n, "edicion": ed,
            "editable": ESP.estado(periodo, cid).get("estado") == "PENDIENTE_VALIDACION" and "revisor" in u["roles"],
            "diapositivas": [{"n": x["n"], "titulo": x["titulo"], "imagen": f"/api/diapositiva/{periodo}/{cid}/{h}/{x['n']}",
                              "unidades": x["u"],
                              "editadas": sorted(k for k in set(x["u"]) | set(originales.get(x["n"], {}))
                                                 if x["u"].get(k) != originales.get(x["n"], {}).get(k)) if originales else []}
                             for x in unidades]}


@app.get("/api/diapositiva/{periodo}/{cid}/{h}/{n}")
def imagen_diapositiva(periodo: str, cid: str, h: str, n: int, u=Depends(CUALQUIERA)):
    _periodo(periodo), _id(cid)
    if not re.match(r"^[0-9a-f]{16}$", h):
        raise HTTPException(400, "Huella inválida.")
    ruta = ESP.salida(periodo, cid) / "vista" / h / f"{n:03d}.png"
    if not ruta.exists():
        raise HTTPException(404, "Imagen no disponible.")
    return FileResponse(ruta, media_type="image/png", headers={"Cache-Control": "private, max-age=86400"})


class EdicionDiapositivas(BaseModel):
    periodo: str
    cliente: str
    cambios: dict[str, dict[str, str]]


@app.put("/api/diapositivas")
def editar_diapositivas(e: EdicionDiapositivas, u=Depends(REVISOR)):
    _periodo(e.periodo), _id(e.cliente)
    if ESP.estado(e.periodo, e.cliente).get("estado") != "PENDIENTE_VALIDACION":
        raise HTTPException(409, "Solo se editan las diapositivas de un informe pendiente de validación.")
    r = _rutas_informe(e.periodo, e.cliente)
    if not r or not r["generado"].exists():
        raise HTTPException(404, "No hay informe generado.")
    if not e.cambios:
        raise HTTPException(400, "No hay cambios para guardar.")
    ed = diapositivas.estado_edicion(r)
    if ed and ed["desactualizado"]:
        raise HTTPException(409, "La edición anterior se hizo sobre un informe que ya cambió. Descártela primero.")
    base = r["editado"] if ed else r["generado"]
    try:
        n = diapositivas.aplicar(base, r["editado"], e.cambios)
    except ValueError as ex:
        raise HTTPException(400, str(ex)) from None
    # cuántos elementos difieren del generado y de qué tipo (lo mismo que capturará la mejora continua)
    difs = mejora.diferencias_pptx(r["generado"], r["editado"])
    guardar_json(r["meta"], {"fecha": ahora(), "por": u["nombre"], "base": diapositivas.huella(r["generado"]),
                             "cambios": len(difs)})
    USU.anotar(u["usuario"], "diapositivas.guardar", f"{e.cliente} {e.periodo}: {n} elemento(s)")
    return {"ok": True, "aplicados": n, "diferencias": len(difs),
            "cifras": sum(1 for x in difs if x["tipo"] == "cifra")}


@app.delete("/api/diapositivas/{periodo}/{cid}")
def descartar_diapositivas(periodo: str, cid: str, u=Depends(REVISOR)):
    _periodo(periodo), _id(cid)
    r = _rutas_informe(periodo, cid)
    if r:
        r["editado"].unlink(missing_ok=True)
        r["meta"].unlink(missing_ok=True)
    USU.anotar(u["usuario"], "diapositivas.descartar", f"{cid} {periodo}")
    return {"ok": True}


# ----------------------------------------------------------------------------- ficha del cliente

CAMPOS_TEXTO = ("copropiedad", "nit", "numero_informe", "revisor", "firma", "ciudad")


class Ficha(BaseModel):
    copropiedad: str
    nit: str = ""
    unidades: int | None = None
    numero_informe: str = ""
    revisor: str = ""
    firma: str = ""
    ciudad: str = ""
    orden: int = 99
    software: str = "generico"
    informe: str = "completo"
    notas: dict = {}


class CambioNit(BaseModel):
    nit: str


@app.post("/api/clientes/{cid}/nit")
def cambiar_nit(cid: str, c: CambioNit, u=Depends(REVISOR)):
    from agente import integridad
    _id(cid)
    with _candado:                                   # no mover carpetas mientras corre un lote
        try:
            nuevo = integridad.cambiar_nit(ESP, cid, c.nit, u["nombre"])
        except ValueError as e:
            raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "cliente.cambiar_nit", f"{cid} → {nuevo}")
    return {"ok": True, "id": nuevo, "nit": mostrar_nit(nuevo)}


@app.put("/api/ficha/{cid}")
def guardar_ficha(cid: str, f: Ficha, u=Depends(REVISOR)):
    """Actualiza los campos editables y conserva el resto (entradas, claves propias). Guarda la versión anterior."""
    import yaml
    from datetime import datetime
    _id(cid), _id(f.software)
    if not f.copropiedad.strip():
        raise HTTPException(400, "La copropiedad no puede quedar vacía.")
    if cid.isdigit():
        try:
            nit_ficha = normalizar_nit(f.nit)
        except ValueError as e:
            raise HTTPException(400, str(e)) from None
        if nit_ficha != cid:
            raise HTTPException(409, "El NIT identifica al cliente y no se cambia desde «Guardar ficha»: "
                                     "use el botón «Cambiar NIT…».")
    if f.informe not in ("completo", "financiero"):
        raise HTTPException(400, "El tipo de informe debe ser «completo» o «financiero».")
    if (f.unidades is not None and f.unidades < 0) or not 0 <= f.orden <= 999:
        raise HTTPException(400, "Unidades u orden fuera de rango.")
    notas = f.notas or {}
    pres = notas.get("presupuesto") or {}
    if not isinstance(pres, dict) or not all(isinstance(k, str) and isinstance(v, str) for k, v in pres.items()):
        raise HTTPException(400, "Las observaciones de presupuesto deben ser pares rubro → texto.")
    ruta = ESP.clientes / cid / "ficha.yaml"
    if not ruta.exists():
        raise HTTPException(404, f"No existe el cliente «{cid}».")
    actual = leer_yaml(ruta, {}) or {}
    hist = ruta.parent / "historial"
    hist.mkdir(exist_ok=True)
    shutil.copy2(ruta, hist / f"ficha_{datetime.now():%Y%m%d_%H%M%S}.yaml")
    for k in CAMPOS_TEXTO:
        actual[k] = getattr(f, k).strip()
    actual.update({"unidades": f.unidades, "orden": f.orden, "software": f.software, "informe": f.informe})
    entradas = actual.setdefault("entradas", {"excel": "*.xlsx", "balance": {"archivo": "balance*"}})
    if isinstance(entradas.get("balance"), dict):
        entradas["balance"]["perfil"] = f.software
    actual["notas"] = {"concentracion_cartera": str(notas.get("concentracion_cartera") or "").strip(),
                       "opinion_anticipos": str(notas.get("opinion_anticipos") or "").strip(),
                       "presupuesto": {k.strip(): v.strip() for k, v in pres.items() if k.strip() and v.strip()},
                       **{k: v for k, v in notas.items()
                          if k not in ("concentracion_cartera", "opinion_anticipos", "presupuesto")}}
    ruta.write_text(f"# Ficha del cliente (editada en LCCOUNTER por {u['nombre']}; versiones previas en historial/).\n"
                    + yaml.safe_dump(actual, allow_unicode=True, sort_keys=False, width=110), encoding="utf-8")
    USU.anotar(u["usuario"], "ficha.guardar", cid)
    return {"ok": True, "ficha": ESP.ficha(cid)}


# ----------------------------------------------------------------------------- mapa del Excel del cliente

def _excel_cliente(periodo: str, cid: str) -> Path:
    from agente import normalizador
    _periodo(periodo), _id(cid)
    try:
        ficha = ESP.ficha(cid)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))
    try:
        excel = normalizador.ubicar_excel(ESP.entrada(periodo, cid), ficha)
    except normalizador.ExcelAmbiguo as e:
        raise HTTPException(409, str(e)) from None
    if not excel:
        raise HTTPException(404, f"No hay Excel cargado para «{cid}» en {periodo}. Cárguelo primero.")
    return excel


def mapa_cliente_aprendidos() -> dict:
    from revisoria import mapa as mapa_cliente
    return mapa_cliente.aprendidos(ESP.conocimiento)


class MapaHojas(BaseModel):
    hojas: dict


@app.get("/api/mapa/{periodo}/{cid}")
def ver_mapa(periodo: str, cid: str, u=Depends(CUALQUIERA)):
    return mapeo.vista(ESP, periodo, cid, _excel_cliente(periodo, cid))


@app.post("/api/mapa/{periodo}/{cid}/proponer")
def proponer_mapa(periodo: str, cid: str, u=Depends(OPERA)):
    excel = _excel_cliente(periodo, cid)
    r = mapeo.proponer(ESP.raiz, mapeo.vista(ESP, periodo, cid, excel), excel,
                       mapa_cliente_aprendidos())
    USU.anotar(u["usuario"], "mapa.proponer", f"{cid} {periodo}: {len(r['propuestas'])} propuesta(s)")
    return r


@app.post("/api/mapa/{periodo}/{cid}/probar")
def probar_mapa(periodo: str, cid: str, m: MapaHojas, u=Depends(OPERA)):
    from revisoria.excel import Libro
    excel = _excel_cliente(periodo, cid)
    try:
        hojas = mapeo.normalizar(Libro(str(excel)), m.hojas)
    except ValueError as e:
        raise HTTPException(400, str(e))
    from revisoria import mapa as mapa_cliente
    return mapeo.probar(excel, periodo, hojas, (mapa_cliente.cargar(ESP.clientes / cid) or {}).get("lectores"))


class PedidoReceta(BaseModel):
    seccion: str
    hoja: str


class RecetaLectura(BaseModel):
    receta: dict


@app.post("/api/receta/{periodo}/{cid}/proponer")
def proponer_receta(periodo: str, cid: str, r: PedidoReceta, u=Depends(OPERA)):
    """Subagente estructurador: propone, prueba y corrige la receta de lectura de una sección (no la guarda)."""
    from agente import estructurador
    from revisoria.excel import Libro, Periodo
    excel = _excel_cliente(periodo, cid)
    try:
        res = estructurador.proponer(ESP.raiz, Libro(str(excel)), r.seccion, r.hoja, Periodo.desde_texto(periodo),
                                     ESP.conocimiento)
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    except proveedores.ErrorIA as e:
        raise HTTPException(502, f"El estructurador no respondió: {e}") from None
    guardar_json(ESP.salida(periodo, cid) / f"receta_propuesta_{r.seccion}.json", res)
    USU.anotar(u["usuario"], "receta.proponer", f"{cid} {periodo} {r.seccion} ← {r.hoja}: "
               + ("OK" if res.get("prueba", {}).get("ok") else "no corresponde" if not res["corresponde"] else "con errores"))
    return json.loads(json.dumps(res, default=str))


# ----------------------------------------------------------------------------- agente adaptador

@app.get("/api/adaptacion/{periodo}/{cid}")
def ver_adaptacion(periodo: str, cid: str, u=Depends(CUALQUIERA)):
    from agente import adaptador
    _periodo(periodo), _id(cid)
    return leer_json(ESP.salida(periodo, cid) / adaptador.ARCHIVO, None) or {"secciones": [], "resumen": None}


class PedidoAdaptacion(BaseModel):
    ia: bool = False


@app.post("/api/adaptacion/{periodo}/{cid}")
def correr_adaptacion(periodo: str, cid: str, r: PedidoAdaptacion, u=Depends(OPERA)):
    """Agente adaptador: propone recetas para todas las secciones que no se pudieron leer. No aprueba nada."""
    from agente import adaptador
    _periodo(periodo), _id(cid)
    try:
        res = adaptador.adaptar(ESP, cid, periodo, usar_ia=r.ia)
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "adaptacion.proponer", f"{cid} {periodo}: {res['resumen']['propuestas']} propuesta(s)"
               + (" con IA" if res.get("ia") else ""))
    return json.loads(json.dumps(res, default=str))


class AprobarAdaptacion(BaseModel):
    secciones: list[str]


@app.put("/api/adaptacion/{periodo}/{cid}")
def aprobar_adaptacion(periodo: str, cid: str, r: AprobarAdaptacion, u=Depends(REVISOR)):
    from agente import adaptador
    from revisoria import lectores
    _periodo(periodo), _id(cid)
    try:
        hechas = adaptador.aprobar(ESP, cid, periodo, [_id(s) for s in r.secciones], u["nombre"])
    except (ValueError, lectores.RecetaInvalida) as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "adaptacion.aprobar", f"{cid} {periodo}: {', '.join(hechas)}")
    return {"ok": True, "aprobadas": hechas}


@app.post("/api/receta/{periodo}/{cid}/probar")
def probar_receta(periodo: str, cid: str, r: RecetaLectura, u=Depends(OPERA)):
    from revisoria import lectores
    from revisoria.excel import Libro, Periodo
    excel = _excel_cliente(periodo, cid)
    try:
        lectores.validar_receta(dict(r.receta))
    except lectores.RecetaInvalida as e:
        return {"ok": False, "errores": [str(e)], "avisos": [], "n": 0, "muestra": []}
    return json.loads(json.dumps(lectores.probar(Libro(str(excel)), r.receta, Periodo.desde_texto(periodo)), default=str))


@app.put("/api/receta/{periodo}/{cid}")
def aprobar_receta(periodo: str, cid: str, r: RecetaLectura, u=Depends(REVISOR)):
    from agente import estructurador
    from revisoria import lectores
    from revisoria.excel import Periodo
    excel = _excel_cliente(periodo, cid)
    try:
        res = estructurador.aprobar(ESP, cid, r.receta, Periodo.desde_texto(periodo), excel, u["nombre"])
    except (ValueError, lectores.RecetaInvalida) as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "receta.aprobar", f"{cid} {r.receta.get('seccion')} ← {r.receta.get('hoja')}")
    return {"ok": True, "receta": res["receta"]}


@app.delete("/api/receta/{periodo}/{cid}/{seccion}")
def quitar_receta(periodo: str, cid: str, seccion: str, u=Depends(REVISOR)):
    from revisoria import mapa as mapa_cliente
    _periodo(periodo), _id(cid), _id(seccion)
    mapa = mapa_cliente.cargar(ESP.clientes / cid) or {}
    if seccion not in (mapa.get("lectores") or {}):
        raise HTTPException(404, "Esa sección no tiene receta.")
    del mapa["lectores"][seccion]
    mapa_cliente.guardar(ESP.clientes / cid, mapa)
    USU.anotar(u["usuario"], "receta.quitar", f"{cid} {seccion}")
    return {"ok": True}


@app.put("/api/mapa/{periodo}/{cid}")
def guardar_mapa(periodo: str, cid: str, m: MapaHojas, u=Depends(REVISOR)):
    excel = _excel_cliente(periodo, cid)
    try:
        mapa = mapeo.guardar(ESP, cid, m.hojas, u["nombre"], excel)
    except ValueError as e:
        raise HTTPException(400, str(e))
    USU.anotar(u["usuario"], "mapa.aprobar", cid + (f" · aprendió {len(mapa['aprendido'])}" if mapa.get("aprendido") else ""))
    return {"ok": True, "mapa": mapa}


# ----------------------------------------------------------------------------- mejora continua y auditoría

@app.post("/api/mejora/analizar")
def analizar(u=Depends(REVISOR)):
    props = mejora.proponer(ESP)
    return {"propuestas": props, "metricas": mejora.metricas(ESP)}


@app.post("/api/propuestas/{pid}/aprobar")
def aprobar_propuesta(pid: str, u=Depends(REVISOR)):
    r = _salir(mejora.aprobar_propuesta, ESP, pid, u["nombre"])
    USU.anotar(u["usuario"], "propuesta.aprobar", f"{pid}: {r.get('detalle', '')}")
    return r


@app.post("/api/regresion")
def regresion(u=Depends(AUDITA)):
    return mejora.regresion(ESP)


class Muestra(BaseModel):
    periodo: str
    k: int = 4


@app.post("/api/muestra")
def muestra(m: Muestra, u=Depends(AUDITA)):
    _periodo(m.periodo)
    ruta = auditoria.muestra_mensual(ESP, m.periodo, max(1, min(m.k, 20)))
    USU.anotar(u["usuario"], "muestra", m.periodo)
    return {"ok": True, "archivo": ruta.name, "muestra": leer_json(ESP.conocimiento / "auditorias" / f"muestra_{m.periodo}.json")}


@app.get("/api/muestra/{periodo}")
def ver_muestra(periodo: str, u=Depends(CUALQUIERA)):
    _periodo(periodo)
    return {"muestra": leer_json(ESP.conocimiento / "auditorias" / f"muestra_{periodo}.json")}


class Conclusion(BaseModel):
    periodo: str
    cliente: str
    conclusion: str
    observacion: str = ""


@app.post("/api/auditar")
def auditar(c: Conclusion, u=Depends(requiere("auditor"))):
    """Conclusión del segundo revisor: el agente rechaza que la registre quien aprobó el informe."""
    _periodo(c.periodo), _id(c.cliente)
    f = _salir(auditoria.registrar_conclusion, ESP, c.periodo, c.cliente, u["nombre"], c.conclusion, c.observacion.strip())
    USU.anotar(u["usuario"], "auditar", f"{c.cliente} {c.periodo}: {c.conclusion}")
    return {"ok": True, "fila": f}


# ----------------------------------------------------------------------------- informe en el navegador (PDF)

@app.get("/api/pdf/{periodo}/{cid}")
def ver_pdf(periodo: str, cid: str, version: str = "actual", u=Depends(CUALQUIERA)):
    """El informe (editado si existe, o el generado) como PDF para abrirlo en una pestaña del navegador."""
    _periodo(periodo), _id(cid)
    r = _rutas_informe(periodo, cid)
    if not r or not r["generado"].exists():
        raise HTTPException(404, "Este cliente aún no tiene informe en el periodo.")
    ed = diapositivas.estado_edicion(r)
    origen = r["editado"] if version in ("actual", "editado") and ed else r["generado"]
    from revisoria.libreoffice import convertir
    destino = r["vista"] / f"pdf_{diapositivas.huella(origen)}"
    pdf = next(iter(destino.glob("*.pdf")), None) if destino.exists() else None
    if not pdf:
        with diapositivas._candado:
            pdf = convertir(origen, "pdf", destino, timeout=300)
    if not pdf:
        raise HTTPException(503, "No se pudo convertir el informe a PDF (¿LibreOffice instalado?).")
    return FileResponse(pdf, media_type="application/pdf",
                        headers={"Content-Disposition": f'inline; filename="{origen.stem}.pdf"'})


# ----------------------------------------------------------------------------- configuración de IA y RSI

@app.get("/api/configuracion")
def ver_configuracion(u=Depends(requiere("admin", "revisor"))):
    from . import rsi
    return {"ia": proveedores.publico(ESP.raiz), "rsi": {"config": rsi.config(ESP), "modos": rsi.MODOS},
            "puede_editar": "admin" in u["roles"]}


class Proveedor(BaseModel):
    id: str
    nombre: str
    tipo: str
    base_url: str = ""
    api_key: str | None = None
    modelos: list[str] = []


@app.put("/api/configuracion/proveedores")
def guardar_proveedor(p: Proveedor, u=Depends(ADMIN)):
    try:
        proveedores.guardar_proveedor(ESP.raiz, p.id, p.model_dump())
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "ia.proveedor", f"{p.id} ({p.tipo}){' · clave actualizada' if p.api_key and '•' not in p.api_key else ''}")
    return proveedores.publico(ESP.raiz)


@app.delete("/api/configuracion/proveedores/{pid}")
def borrar_proveedor(pid: str, u=Depends(ADMIN)):
    try:
        proveedores.eliminar_proveedor(ESP.raiz, pid)
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "ia.proveedor.eliminar", pid)
    return proveedores.publico(ESP.raiz)


class Asignacion(BaseModel):
    uso: str
    proveedor: str
    modelo: str


@app.put("/api/configuracion/uso")
def asignar_uso(a: Asignacion, u=Depends(ADMIN)):
    try:
        proveedores.asignar(ESP.raiz, a.uso, a.proveedor, a.modelo)
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "ia.uso", f"{a.uso} → {a.proveedor}/{a.modelo}")
    return proveedores.publico(ESP.raiz)


class Prueba(BaseModel):
    proveedor: str
    modelo: str


@app.post("/api/configuracion/probar")
def probar_proveedor(p: Prueba, u=Depends(ADMIN)):
    try:
        return {"ok": True, "detalle": proveedores.probar(ESP.raiz, p.proveedor, p.modelo)}
    except proveedores.ErrorIA as e:
        return {"ok": False, "detalle": str(e)}


class ConfigRSI(BaseModel):
    modo: str
    margen: float
    candidatos: int
    hora: str | None = None


@app.put("/api/configuracion/rsi")
def configurar_rsi(c: ConfigRSI, u=Depends(ADMIN)):
    from . import rsi
    try:
        rsi.guardar_config(ESP, c.modo, c.margen, c.candidatos, c.hora)
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "rsi.config", f"modo={c.modo} margen={c.margen} candidatos={c.candidatos} hora={c.hora}")
    return rsi.config(ESP)


@app.get("/api/rsi")
def ver_rsi(u=Depends(requiere("admin", "revisor", "auditor"))):
    from . import rsi
    e = rsi.estado(ESP)
    return {"config": rsi.config(ESP), "estado": e, "casos": {k: len(v) for k, v in rsi.casos(ESP).items()},
            "propuestas": rsi.propuestas(ESP), "ciclos": rsi.ciclos(ESP, 15),
            "disponible": proveedores.disponible(ESP.raiz, "rsi")}


_candado_rsi = threading.Lock()


@app.post("/api/rsi/ciclo")
def ciclo_rsi(u=Depends(requiere("admin", "revisor"))):
    from . import rsi
    if not _candado_rsi.acquire(blocking=False):
        raise HTTPException(409, "Ya hay un ciclo de automejora en curso.")
    try:
        r = rsi.ciclo(ESP, u["nombre"])
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    except proveedores.ErrorIA as e:
        raise HTTPException(503, str(e)) from None
    finally:
        _candado_rsi.release()
    USU.anotar(u["usuario"], "rsi.ciclo", "; ".join(f"{p['clave']}: {p['accion']}" for p in r["plantillas"])[:300])
    return r


def _reloj_rsi():
    """Ciclo diario de automejora a la hora configurada (4:00 a. m. por defecto). Mejora el sistema (plantillas);
    nunca toca informes validados. Si no hay proveedor de IA o no hay casos, lo anota y sigue."""
    import time as _t
    from datetime import datetime as _dt
    from . import rsi
    while True:
        _t.sleep(60)
        try:
            ahora_dt = _dt.now()
            if not rsi.toca_programado(ESP, ahora_dt) or not _candado_rsi.acquire(blocking=False):
                continue
            try:
                rsi.marcar_programado(ESP, ahora_dt)
                r = rsi.ciclo(ESP, "Automejora programada")
                USU.anotar("sistema", "rsi.programado", "; ".join(f"{p['clave']}: {p['accion']}" for p in r["plantillas"])[:300]
                           or "sin cambios")
            finally:
                _candado_rsi.release()
        except Exception as e:                                # el reloj nunca tumba la app
            try:
                USU.anotar("sistema", "rsi.programado", f"no corrió: {type(e).__name__}: {e}"[:300])
            except Exception:
                pass


@app.on_event("startup")
def _iniciar_reloj():
    if os.environ.get("LCCOUNTER_SIN_RELOJ") != "1":
        threading.Thread(target=_reloj_rsi, daemon=True, name="reloj-rsi").start()


@app.post("/api/rsi/propuestas/{pid}/{accion}")
def decidir_rsi(pid: str, accion: str, u=Depends(REVISOR)):
    from . import rsi
    try:
        if accion == "aprobar":
            det = f"Plantilla publicada en {rsi.aprobar(ESP, pid, u['nombre'])}."
        elif accion == "rechazar":
            rsi.rechazar(ESP, pid, u["nombre"])
            det = "Propuesta rechazada."
        else:
            raise HTTPException(400, "Acción inválida.")
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], f"rsi.{accion}", pid)
    return {"ok": True, "detalle": det}


@app.post("/api/rsi/estrategias/{eid}/restaurar")
def restaurar_estrategia(eid: str, u=Depends(ADMIN)):
    from . import rsi
    try:
        rsi.restaurar_estrategia(ESP, eid)
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    USU.anotar(u["usuario"], "rsi.estrategia.restaurar", eid)
    return {"ok": True}


@app.post("/api/versiones/{version}/revertir")
def revertir_version(version: str, u=Depends(REVISOR)):
    if not re.match(r"^bib-\d{3}$", version):
        raise HTTPException(400, "Versión inválida.")
    r = _salir(mejora.revertir, ESP, version)
    USU.anotar(u["usuario"], "biblioteca.revertir", version)
    return {"ok": True, "detalle": r}


# ----------------------------------------------------------------------------- base normativa

NORMATIVA_EDITA = requiere("revisor", "admin")


def _ambito(a: str) -> str:
    if a != "general":
        _id(a)
        if a not in ESP.ids_clientes():
            raise HTTPException(404, f"No existe el cliente «{a}».")
    return a


@app.get("/api/normativa")
def ver_normativa(u=Depends(CUALQUIERA)):
    return {"documentos": normativa.instancia(ESP.raiz).documentos(), "tipos": normativa.TIPOS,
            "clientes": {c: ESP.ficha(c).get("copropiedad", c) for c in ESP.ids_clientes()}}


@app.get("/api/normativa/buscar")
def buscar_normativa(q: str, cliente: str | None = None, u=Depends(CUALQUIERA)):
    if cliente:
        _id(cliente)
    if len(q) > 500:
        raise HTTPException(400, "Consulta demasiado larga.")
    return {"resultados": normativa.instancia(ESP.raiz).buscar(q, cliente, k=12)}


@app.post("/api/normativa")
async def cargar_normativa(archivos: list[UploadFile] = File(...), titulo: str = Form(...), tipo: str = Form(...),
                           ambito: str = Form("general"), version: str = Form(""), origen: str = Form(""),
                           u=Depends(NORMATIVA_EDITA)):
    """Uno o varios archivos (en orden) que forman UN documento, p. ej. una ley publicada en tres páginas."""
    _ambito(ambito)
    import tempfile
    if not archivos or len(archivos) > 20:
        raise HTTPException(400, "Cargue entre 1 y 20 archivos.")
    temporales, total = [], 0
    try:
        for up in archivos:
            sufijo = Path(up.filename or "").suffix.lower()
            with tempfile.NamedTemporaryFile(suffix=sufijo, delete=False) as tmp:
                temporales.append((Path(tmp.name), up.filename or f"parte{len(temporales) + 1}{sufijo}"))
                while bloque := await up.read(1 << 20):
                    total += len(bloque)
                    if total > 60 << 20:
                        raise HTTPException(400, "Los archivos superan 60 MB en total.")
                    tmp.write(bloque)
        meta = normativa.instancia(ESP.raiz).cargar(temporales, titulo, tipo, ambito, version, origen, u["nombre"])
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    finally:
        for ruta, _ in temporales:
            ruta.unlink(missing_ok=True)
    USU.anotar(u["usuario"], "normativa.cargar",
               f"{ambito}: {meta['titulo']} ({len(temporales)} archivo(s), {meta['fragmentos']} fragmentos)")
    return meta


@app.get("/api/normativa/{ambito}/{doc_id}")
def ver_documento_normativo(ambito: str, doc_id: str, u=Depends(CUALQUIERA)):
    try:
        meta, frags = normativa.instancia(ESP.raiz).fragmentos(_ambito(ambito), doc_id)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e)) from None
    return {"meta": meta, "fragmentos": frags}


@app.get("/api/normativa/{ambito}/{doc_id}/original")
def original_normativo(ambito: str, doc_id: str, parte: int = 1, u=Depends(CUALQUIERA)):
    try:
        meta, _ = normativa.instancia(ESP.raiz).fragmentos(_ambito(ambito), doc_id)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e)) from None
    nombres = meta.get("archivos") or [meta["archivo"]]
    if not 1 <= parte <= len(nombres):
        raise HTTPException(404, "No existe esa parte.")
    carpeta = normativa.instancia(ESP.raiz)._dir(ambito) / doc_id
    ruta = next(iter(sorted(carpeta.glob(f"original_{parte}.*"))), None) or next(iter(carpeta.glob("original.*")), None)
    if not ruta:
        raise HTTPException(404, "Original no disponible.")
    return FileResponse(ruta, filename=nombres[parte - 1])


@app.delete("/api/normativa/{ambito}/{doc_id}")
def eliminar_normativo(ambito: str, doc_id: str, u=Depends(NORMATIVA_EDITA)):
    try:
        normativa.instancia(ESP.raiz).eliminar(_ambito(ambito), doc_id)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e)) from None
    USU.anotar(u["usuario"], "normativa.eliminar", f"{ambito}: {doc_id}")
    return {"ok": True}


# ----------------------------------------------------------------------------- comunicaciones al Consejo

@app.get("/api/comunicaciones/{periodo}/{cid}")
def ver_comunicaciones(periodo: str, cid: str, u=Depends(CUALQUIERA)):
    _periodo(periodo), _id(cid)
    return {"lista": comunicaciones.listar(ESP, periodo, cid), "tipos": comunicaciones.TIPOS,
            "hallazgos": comunicaciones.opciones_hallazgos(ESP, periodo, cid), "asistente": proveedores.disponible(ESP.raiz, "comunicaciones")}


class NuevaComunicacion(BaseModel):
    periodo: str
    cliente: str
    tipo: str
    hallazgos: list[str] = []
    incluir_borrador: bool = True
    incluir_normativa: bool = True
    recibida: str = ""
    instrucciones: str = ""


@app.post("/api/comunicaciones")
def crear_comunicacion(n: NuevaComunicacion, u=Depends(OPERA)):
    _periodo(n.periodo), _id(n.cliente)
    if len(n.recibida) > 30_000 or len(n.instrucciones) > 3_000:
        raise HTTPException(400, "El texto recibido o las instrucciones son demasiado largos.")
    try:
        doc = comunicaciones.generar(ESP, n.periodo, n.cliente, n.tipo, n.hallazgos[:30], n.incluir_borrador,
                                     n.recibida, n.instrucciones, u["nombre"], n.incluir_normativa)
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    except RuntimeError as e:
        raise HTTPException(503, str(e)) from None
    USU.anotar(u["usuario"], "comunicacion.generar", f"{n.cliente} {n.periodo}: {n.tipo} ({len(n.hallazgos)} hallazgos)")
    return doc


def _doc(periodo, cid, doc_id):
    _periodo(periodo), _id(cid)
    try:
        return comunicaciones.leer(ESP, periodo, cid, doc_id)
    except ValueError as e:
        raise HTTPException(400, str(e)) from None
    except FileNotFoundError as e:
        raise HTTPException(404, str(e)) from None


@app.get("/api/comunicaciones/{periodo}/{cid}/{doc_id}")
def ver_comunicacion(periodo: str, cid: str, doc_id: str, u=Depends(CUALQUIERA)):
    return _doc(periodo, cid, doc_id)


class EdicionComunicacion(BaseModel):
    oficio: str
    asunto: str = ""
    revisado: bool = False


@app.put("/api/comunicaciones/{periodo}/{cid}/{doc_id}")
def editar_comunicacion(periodo: str, cid: str, doc_id: str, e: EdicionComunicacion, u=Depends(OPERA)):
    _doc(periodo, cid, doc_id)
    if e.revisado and "revisor" not in u["roles"]:
        raise HTTPException(403, "Solo un revisor fiscal puede marcar la comunicación como revisada.")
    if not e.oficio.strip() or len(e.oficio) > 100_000:
        raise HTTPException(400, "El texto del oficio está vacío o es demasiado largo.")
    d = comunicaciones.guardar(ESP, periodo, cid, doc_id, e.oficio.replace("\r\n", "\n"), e.asunto.strip(),
                               u["nombre"], e.revisado)
    USU.anotar(u["usuario"], "comunicacion." + ("revisar" if e.revisado else "editar"), f"{cid} {periodo}: {doc_id}")
    return d


@app.delete("/api/comunicaciones/{periodo}/{cid}/{doc_id}")
def eliminar_comunicacion(periodo: str, cid: str, doc_id: str, u=Depends(OPERA)):
    _doc(periodo, cid, doc_id)
    comunicaciones.eliminar(ESP, periodo, cid, doc_id)
    USU.anotar(u["usuario"], "comunicacion.eliminar", f"{cid} {periodo}: {doc_id}")
    return {"ok": True}


@app.get("/api/comunicaciones/{periodo}/{cid}/{doc_id}/word")
def comunicacion_word(periodo: str, cid: str, doc_id: str, u=Depends(CUALQUIERA)):
    _doc(periodo, cid, doc_id)
    ruta = comunicaciones.a_word(ESP, periodo, cid, doc_id, Path(__file__).parent / "static" / "logo.png")
    return FileResponse(ruta, filename=f"Comunicacion_{cid}_{doc_id}.docx",
                        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ----------------------------------------------------------------------------- asistente (Claude)

class Pregunta(BaseModel):
    mensajes: list[dict]          # historial [{role, content}] solo con texto
    base: str = "todo"
    periodo: str | None = None
    cliente: str | None = None


@app.post("/api/preguntar")
def preguntar(p: Pregunta, u=Depends(CUALQUIERA)):
    if p.periodo:
        _periodo(p.periodo)
    if p.cliente:
        _id(p.cliente)
    if p.base not in asistente.BASES:
        raise HTTPException(400, f"Base desconocida «{p.base}».")
    mensajes = [{"role": m.get("role"), "content": str(m.get("content", ""))} for m in p.mensajes[-20:]
                if m.get("role") in ("user", "assistant") and str(m.get("content", "")).strip()]
    if not mensajes or mensajes[-1]["role"] != "user":
        raise HTTPException(400, "Falta la pregunta.")
    USU.anotar(u["usuario"], "asistente", f"[{p.base}] {mensajes[-1]['content'][:200]}")

    def eventos():
        for ev in asistente.conversar(ESP, mensajes, p.base, p.periodo, p.cliente):
            yield f"data: {json.dumps(ev, ensure_ascii=False)}\n\n"
    return StreamingResponse(eventos(), media_type="text/event-stream")


app.mount("/", StaticFiles(directory=Path(__file__).parent / "static", html=True), name="static")


def main():
    global ESP, USU
    ap = argparse.ArgumentParser(description="Aplicación web LCCOUNTER")
    ap.add_argument("--raiz", default=str(ESP.raiz))
    ap.add_argument("--puerto", type=int, default=8790)
    a = ap.parse_args()
    ESP = Espacio(Path(a.raiz))
    USU = seguridad.Usuarios(ESP.raiz)
    seguridad.persistir_sesiones(ESP.raiz)          # reiniciar la app no cierra la sesión de quien la usa
    import uvicorn
    print(f"LCCOUNTER en http://127.0.0.1:{a.puerto} · espacio {ESP.raiz}")
    if not USU.hay_usuarios():
        print("Primer uso: abra la aplicación para crear el usuario administrador.")
    uvicorn.run(app, host="127.0.0.1", port=a.puerto)       # solo local: el espacio tiene datos de clientes


if __name__ == "__main__":
    main()
