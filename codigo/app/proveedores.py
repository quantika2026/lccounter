"""Proveedores de IA configurables: Claude (SDK oficial de Anthropic) y cualquier API compatible con OpenAI
(OpenAI, Google Gemini, DeepSeek, Mistral, Groq, Ollama local u otro servidor propio).

Cada función de la aplicación usa el proveedor y el modelo que el administrador elija en Configuración:
  asistente       consultas con herramientas de solo lectura
  comunicaciones  borradores de oficios (salida JSON estructurada)
  rsi             automejora: genera y evalúa candidatos

Configuración en <raiz>/configuracion_ia.json (permisos 600). Las claves nunca se devuelven a la interfaz: solo
se muestran enmascaradas. Para Claude, si no hay clave en la configuración, se usan las credenciales del entorno
(ANTHROPIC_API_KEY o `ant auth login`).
"""
from __future__ import annotations

import json
import os
import re
import threading
import urllib.error
import urllib.request
from pathlib import Path

USOS = {"asistente": "Asistente de consulta (con herramientas)", "comunicaciones": "Redacción de comunicaciones",
        "rsi": "Automejora recursiva (RSI)", "estructurador": "Subagente estructurador (Excel con otra estructura)"}
PRESETS = {
    "anthropic": {"nombre": "Claude (Anthropic)", "tipo": "anthropic", "base_url": "", "modelos": ["claude-opus-5-5", "claude-sonnet-5-5", "claude-haiku-5-5"]},
    "openai": {"nombre": "OpenAI", "tipo": "openai", "base_url": "https://api.openai.com/v1", "modelos": []},
    "gemini": {"nombre": "Google Gemini", "tipo": "openai", "base_url": "https://generativelanguage.googleapis.com/v1beta/openai", "modelos": []},
    "deepseek": {"nombre": "DeepSeek", "tipo": "openai", "base_url": "https://api.deepseek.com/v1", "modelos": []},
    "mistral": {"nombre": "Mistral", "tipo": "openai", "base_url": "https://api.mistral.ai/v1", "modelos": []},
    "groq": {"nombre": "Groq", "tipo": "openai", "base_url": "https://api.groq.com/openai/v1", "modelos": []},
    "ollama": {"nombre": "Ollama (local)", "tipo": "openai", "base_url": "http://localhost:11434/v1", "modelos": []},
    "personalizado": {"nombre": "Otro (compatible con OpenAI)", "tipo": "openai", "base_url": "", "modelos": []},
}
POR_DEFECTO = {"proveedores": {"anthropic": {**PRESETS["anthropic"], "api_key": ""}},
               "uso": {u: {"proveedor": "anthropic", "modelo": "claude-opus-5-5"} for u in USOS}}
_candado = threading.Lock()


class ErrorIA(RuntimeError):
    """Error de un proveedor, con un mensaje apto para mostrar al usuario."""


LIMITE_LOCAL_ENTRADA = 24_000           # caracteres de fuentes que se le pasan a un modelo local


class SinHerramientas(ErrorIA):
    """El modelo no admite llamadas a herramientas (algunos modelos locales)."""


def es_local(p: dict) -> bool:
    """Proveedor que corre en este equipo (Ollama u otro servidor local): los datos no salen del Mac."""
    import re
    return bool(re.match(r"^http://(localhost|127\.0\.0\.1)(:\d+)?(/|$)", p.get("base_url") or ""))


def _limpiar(texto: str) -> str:
    """Quita el razonamiento interno que algunos modelos locales devuelven entre <think>…</think>."""
    import re
    return re.sub(r"<think>.*?</think>\s*", "", texto or "", flags=re.S)


def _cuerpo(p: dict, cuerpo: dict) -> dict:
    """Ajustes para modelos locales: sin «pensamiento» largo (responde en segundos en vez de minutos)."""
    return {**cuerpo, "reasoning_effort": "none"} if es_local(p) else cuerpo


# ------------------------------------------------------------------ configuración

def _ruta(raiz: Path) -> Path:
    return Path(raiz) / "configuracion_ia.json"


def leer(raiz: Path) -> dict:
    try:
        cfg = json.loads(_ruta(raiz).read_text(encoding="utf-8"))
    except FileNotFoundError:
        cfg = json.loads(json.dumps(POR_DEFECTO))
    for u in USOS:                                   # cada uso nuevo arranca con Claude (claude-opus-5-5)
        cfg.setdefault("uso", {}).setdefault(u, dict(POR_DEFECTO["uso"][u]))
    return cfg


def guardar(raiz: Path, cfg: dict):
    with _candado:
        r = _ruta(raiz)
        tmp = r.with_suffix(".tmp")
        tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
        tmp.chmod(0o600)
        tmp.replace(r)


def _mascara(k: str) -> str:
    return "" if not k else ("•" * 8 + k[-4:] if len(k) > 8 else "••••")


def publico(raiz: Path) -> dict:
    cfg = leer(raiz)
    return {"proveedores": {pid: {**{k: v for k, v in p.items() if k != "api_key"}, "api_key": _mascara(p.get("api_key", "")),
                                  "con_clave": bool(p.get("api_key")) or (p["tipo"] == "anthropic" and _credencial_entorno())}
                            for pid, p in cfg["proveedores"].items()},
            "uso": cfg["uso"], "usos": USOS, "presets": PRESETS}


def _credencial_entorno() -> bool:
    return bool(os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")
                or (Path.home() / ".config" / "anthropic").exists())


def guardar_proveedor(raiz: Path, pid: str, datos: dict):
    import re
    if not re.match(r"^[a-z0-9_-]{2,30}$", pid):
        raise ValueError("Identificador de proveedor inválido.")
    tipo = datos.get("tipo")
    if tipo not in ("anthropic", "openai"):
        raise ValueError("Tipo de proveedor inválido.")
    base = (datos.get("base_url") or "").strip().rstrip("/")
    if tipo == "openai" and not re.match(r"^https?://", base):
        raise ValueError("Indique la URL base de la API (https://…).")
    if base.startswith("http://") and not re.match(r"^http://(localhost|127\.0\.0\.1)(:\d+)?(/|$)", base):
        raise ValueError("Solo se admite http:// para servidores locales; use https://.")
    cfg = leer(raiz)
    actual = cfg["proveedores"].get(pid, {})
    clave = datos.get("api_key")
    cfg["proveedores"][pid] = {
        "nombre": (datos.get("nombre") or pid).strip()[:60], "tipo": tipo, "base_url": base,
        "modelos": [m.strip() for m in (datos.get("modelos") or []) if m.strip()][:20],
        "api_key": actual.get("api_key", "") if clave in (None, "") or "•" in str(clave) else str(clave).strip(),
    }
    guardar(raiz, cfg)


def eliminar_proveedor(raiz: Path, pid: str):
    cfg = leer(raiz)
    if any(u["proveedor"] == pid for u in cfg["uso"].values()):
        raise ValueError("Ese proveedor está en uso; asigne otro a esas funciones antes de eliminarlo.")
    cfg["proveedores"].pop(pid, None)
    guardar(raiz, cfg)


def asignar(raiz: Path, uso: str, proveedor: str, modelo: str):
    cfg = leer(raiz)
    if uso not in USOS or proveedor not in cfg["proveedores"] or not modelo.strip():
        raise ValueError("Función, proveedor o modelo inválido.")
    cfg["uso"][uso] = {"proveedor": proveedor, "modelo": modelo.strip()[:80]}
    guardar(raiz, cfg)


def disponible(raiz: Path, uso: str) -> dict:
    cfg = leer(raiz)
    u = cfg["uso"][uso]
    p = cfg["proveedores"].get(u["proveedor"])
    if not p:
        return {"ok": False, "modelo": u["modelo"], "proveedor": u["proveedor"],
                "ayuda": "El proveedor asignado no existe. Revíselo en Configuración."}
    ok = bool(p.get("api_key")) or (p["tipo"] == "anthropic" and _credencial_entorno()) or es_local(p)
    return {"ok": ok, "modelo": u["modelo"], "proveedor": p["nombre"], "local": es_local(p),
            "ayuda": "" if ok else f"Falta la clave de API de {p['nombre']}: agréguela en Configuración → Proveedores de IA."}


def _elegido(raiz: Path, uso: str) -> tuple[dict, str]:
    cfg = leer(raiz)
    u = cfg["uso"][uso]
    p = cfg["proveedores"].get(u["proveedor"])
    if not p:
        raise ErrorIA("El proveedor asignado no existe. Revíselo en Configuración.")
    return p, u["modelo"]


# ------------------------------------------------------------------ Claude (SDK oficial)

def _cliente_anthropic(p: dict):
    import anthropic
    return anthropic.Anthropic(api_key=p["api_key"]) if p.get("api_key") else anthropic.Anthropic()


def _parametros_claude(modelo: str, effort: str) -> dict:
    """Pensamiento adaptativo y respaldo automático ante rechazos en los modelos que los admiten."""
    extra = {}
    if modelo.startswith(("claude-opus-5", "claude-sonnet-5", "claude-fable-5", "claude-haiku-5", "claude-opus-4-6",
                          "claude-opus-4-7", "claude-opus-4-8", "claude-sonnet-4-6")):
        extra["thinking"] = {"type": "adaptive"}
        extra["output_config"] = {"effort": effort}
    if modelo in ("claude-opus-5-5", "claude-opus-5", "claude-sonnet-5-5", "claude-fable-5-1"):
        extra["betas"] = ["server-side-fallback-2026-07-01"]
        extra["fallbacks"] = "default"
    return extra


def _errores_anthropic(fn):
    import anthropic
    try:
        return fn()
    except anthropic.AuthenticationError:
        raise ErrorIA("La clave de Claude no es válida. Revísela en Configuración.") from None
    except anthropic.RateLimitError:
        raise ErrorIA("Límite de uso de Claude alcanzado; intente en unos minutos.") from None
    except anthropic.NotFoundError:
        raise ErrorIA("El modelo de Claude configurado no existe. Revíselo en Configuración.") from None
    except anthropic.APIStatusError as e:
        raise ErrorIA(f"Error de la API de Claude ({e.status_code}).") from None
    except anthropic.APIConnectionError:
        raise ErrorIA("Sin conexión con la API de Claude.") from None
    except TypeError as e:                               # el SDK no encontró ninguna clave configurada
        if "authentication" in str(e).lower():
            raise ErrorIA("No hay clave de Claude configurada. Agréguela en Configuración → Proveedores de IA.") from None
        raise


# ------------------------------------------------------------------ APIs compatibles con OpenAI (HTTP)

def _post_openai(p: dict, ruta: str, cuerpo: dict, timeout: int = 300) -> dict:
    url = p["base_url"].rstrip("/") + ruta
    cab = {"Content-Type": "application/json"}
    if p.get("api_key"):
        cab["Authorization"] = f"Bearer {p['api_key']}"
    req = urllib.request.Request(url, data=json.dumps(cuerpo).encode(), headers=cab, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        detalle = ""
        try:
            detalle = json.loads(e.read().decode()).get("error", {}).get("message", "")[:200]
        except Exception:
            pass
        if e.code in (401, 403):
            raise ErrorIA(f"{p['nombre']}: la clave de API no es válida o no tiene permiso.") from None
        if e.code == 404:
            raise ErrorIA(f"{p['nombre']}: el modelo o la URL no existen ({detalle or 'revise Configuración'}).") from None
        if e.code == 400 and "tool" in detalle.lower():
            raise SinHerramientas(f"{p['nombre']}: el modelo no admite herramientas.") from None
        if e.code == 429:
            raise ErrorIA(f"{p['nombre']}: límite de uso alcanzado; intente en unos minutos.") from None
        raise ErrorIA(f"{p['nombre']}: error {e.code}. {detalle}") from None
    except (urllib.error.URLError, TimeoutError) as e:
        raise ErrorIA(f"Sin conexión con {p['nombre']} ({getattr(e, 'reason', e)}).") from None


# ------------------------------------------------------------------ operaciones

def json_estructurado(raiz: Path, uso: str, sistema: str, usuario: str, esquema: dict, effort: str = "high") -> tuple[dict, str]:
    """Pide una respuesta JSON que cumpla el esquema. Devuelve (datos, modelo que respondió)."""
    p, modelo = _elegido(raiz, uso)
    if p["tipo"] == "anthropic":
        cliente = _cliente_anthropic(p)

        extra = _parametros_claude(modelo, effort)
        extra["output_config"] = {**extra.get("output_config", {}), "format": {"type": "json_schema", "schema": esquema}}

        def llamar():
            with cliente.beta.messages.stream(model=modelo, max_tokens=32000, system=sistema,
                                              messages=[{"role": "user", "content": usuario}], **extra) as s:
                return s.get_final_message()
        resp = _errores_anthropic(llamar)
        if resp.stop_reason == "refusal":
            raise ErrorIA("El modelo no pudo responder esta solicitud.")
        if resp.stop_reason == "max_tokens":
            raise ErrorIA("La respuesta salió demasiado larga; acote la solicitud.")
        texto = next((b.text for b in resp.content if b.type == "text"), "")
        modelo_resp = resp.model
    else:
        if es_local(p) and len(usuario) > LIMITE_LOCAL_ENTRADA:   # contexto corto del modelo local
            usuario = usuario[:LIMITE_LOCAL_ENTRADA] + "\n\n[… fuentes recortadas por longitud]"
        base = _cuerpo(p, {"model": modelo, "messages": [{"role": "system", "content": sistema}, {"role": "user", "content": usuario}]})
        try:
            r = _post_openai(p, "/chat/completions", {**base, "response_format": {
                "type": "json_schema", "json_schema": {"name": "respuesta", "schema": esquema, "strict": True}}})
        except ErrorIA as e:
            if "error 400" not in str(e):
                raise
            # algunos proveedores no admiten json_schema: modo JSON simple con el esquema en las instrucciones
            base["messages"][0]["content"] += "\n\nResponde SOLO con un objeto JSON que cumpla este esquema:\n" + json.dumps(esquema, ensure_ascii=False)
            r = _post_openai(p, "/chat/completions", {**base, "response_format": {"type": "json_object"}})
        ch = r["choices"][0]
        if ch.get("finish_reason") == "length":
            raise ErrorIA("La respuesta salió demasiado larga; acote la solicitud.")
        texto = _limpiar(ch["message"].get("content") or "")
        modelo_resp = r.get("model", modelo)
    try:
        datos = json.loads(texto.strip().removeprefix("```json").removesuffix("```").strip())
    except json.JSONDecodeError:
        raise ErrorIA("La respuesta no tuvo el formato esperado; intente de nuevo.") from None
    for k, v in list(datos.items()):                     # los modelos locales a veces envuelven el Markdown en ```
        if isinstance(v, str):
            datos[k] = re.sub(r"^\s*```(?:markdown|md)?\s*\n(.*?)\n?```\s*$", r"\1", v, flags=re.S)
    faltan = [k for k in esquema.get("required", []) if k not in datos]
    if faltan:
        raise ErrorIA(f"La respuesta no trajo los campos {', '.join(faltan)}; intente de nuevo.")
    return datos, modelo_resp


def texto(raiz: Path, uso: str, sistema: str, usuario: str, effort: str = "medium") -> str:
    p, modelo = _elegido(raiz, uso)
    if p["tipo"] == "anthropic":
        cliente = _cliente_anthropic(p)

        def llamar():
            with cliente.beta.messages.stream(model=modelo, max_tokens=16000, system=sistema,
                                              messages=[{"role": "user", "content": usuario}],
                                              **_parametros_claude(modelo, effort)) as s:
                return s.get_final_message()
        resp = _errores_anthropic(llamar)
        return "".join(b.text for b in resp.content if b.type == "text")
    r = _post_openai(p, "/chat/completions", _cuerpo(p, {"model": modelo, "messages": [
        {"role": "system", "content": sistema}, {"role": "user", "content": usuario}]}))
    return _limpiar(r["choices"][0]["message"].get("content") or "")


def probar(raiz: Path, pid: str, modelo: str) -> str:
    cfg = leer(raiz)
    p = cfg["proveedores"].get(pid)
    if not p:
        raise ErrorIA("No existe ese proveedor.")
    if p["tipo"] == "anthropic":
        cliente = _cliente_anthropic(p)
        r = _errores_anthropic(lambda: cliente.messages.create(model=modelo, max_tokens=64,
                                                               messages=[{"role": "user", "content": "Responde solo: OK"}]))
        return f"Conexión correcta con {r.model}."
    r = _post_openai(p, "/chat/completions", {"model": modelo, "max_tokens": 16,
                                              "messages": [{"role": "user", "content": "Responde solo: OK"}]}, timeout=60)
    return f"Conexión correcta con {r.get('model', modelo)}."


def conversar(raiz: Path, uso: str, sistema: str, mensajes: list[dict], herramientas: dict, ejecutar, validar,
              max_vueltas: int = 12):
    """Conversación con herramientas. herramientas = {nombre: (descripción, esquema)}. Genera eventos:
    {"tipo": "texto"|"herramienta"|"fin"|"error", …}."""
    try:
        p, modelo = _elegido(raiz, uso)
    except ErrorIA as e:
        yield {"tipo": "error", "mensaje": str(e)}
        return
    try:
        if p["tipo"] == "anthropic":
            yield from _conversar_claude(p, modelo, sistema, mensajes, herramientas, ejecutar, validar, max_vueltas)
        else:
            yield from _conversar_openai(p, modelo, sistema, mensajes, herramientas, ejecutar, validar, max_vueltas)
    except SinHerramientas:
        raise                                            # el asistente pasa a modo de recuperación
    except ErrorIA as e:
        yield {"tipo": "error", "mensaje": str(e)}


def _resultado(nombre, args, ejecutar, validar):
    error = validar(nombre, args)
    if error:
        return f"Entrada inválida: {error}.", True
    try:
        return ejecutar(nombre, args), False
    except Exception as ex:                              # una lectura fallida no corta la conversación
        return f"Error al leer: {ex}", True


def _conversar_claude(p, modelo, sistema, mensajes, herramientas, ejecutar, validar, max_vueltas):
    cliente = _cliente_anthropic(p)
    tools = [{"name": n, "description": d, "input_schema": s, "eager_input_streaming": True} for n, (d, s) in herramientas.items()]
    historial, reintentos = list(mensajes), 0
    for _ in range(max_vueltas):
        eventos = []

        def llamar():
            with cliente.beta.messages.stream(model=modelo, max_tokens=64000, system=sistema, tools=tools,
                                              messages=historial, cache_control={"type": "ephemeral"},
                                              **_parametros_claude(modelo, "medium")) as stream:
                for ev in stream:
                    if ev.type == "content_block_delta" and ev.delta.type == "text_delta":
                        eventos.append(ev.delta.text)
                return stream.get_final_message()
        try:
            resp = _errores_anthropic(llamar)
            reintentos = 0
        except ValueError:                               # JSON de herramienta imposible de interpretar
            reintentos += 1
            if reintentos > 2:
                raise ErrorIA("El modelo devolvió una consulta ilegible; intente de nuevo.") from None
            continue
        for t in eventos:
            yield {"tipo": "texto", "t": t}
        if resp.stop_reason == "refusal":
            raise ErrorIA("El modelo no pudo responder esta consulta.")
        if resp.stop_reason == "pause_turn":
            historial.append({"role": "assistant", "content": resp.content})
            continue
        usos = [b for b in resp.content if b.type == "tool_use"]
        if not usos:
            yield {"tipo": "fin"}
            return
        if resp.stop_reason == "max_tokens":
            raise ErrorIA("La respuesta se cortó por longitud; haga una pregunta más acotada.")
        res = []
        for b in usos:
            args = b.input if isinstance(b.input, dict) else {}
            yield {"tipo": "herramienta", "nombre": b.name, "args": args}
            contenido, es_error = _resultado(b.name, b.input, ejecutar, validar)
            res.append({"type": "tool_result", "tool_use_id": b.id, "content": contenido, **({"is_error": True} if es_error else {})})
        historial.append({"role": "assistant", "content": resp.content})
        historial.append({"role": "user", "content": res})
        yield {"tipo": "texto", "t": "\n\n"}
    raise ErrorIA("La consulta necesitó demasiados pasos; intente una pregunta más concreta.")


def _conversar_openai(p, modelo, sistema, mensajes, herramientas, ejecutar, validar, max_vueltas):
    tools = [{"type": "function", "function": {"name": n, "description": d, "parameters": s}} for n, (d, s) in herramientas.items()]
    historial = [{"role": "system", "content": sistema}] + [{"role": m["role"], "content": m["content"]} for m in mensajes]
    for _ in range(max_vueltas):
        r = _post_openai(p, "/chat/completions", _cuerpo(p, {"model": modelo, "messages": historial, "tools": tools}))
        ch = r["choices"][0]
        msg = ch["message"]
        llamadas = msg.get("tool_calls") or []
        if _limpiar(msg.get("content")).strip():
            yield {"tipo": "texto", "t": _limpiar(msg["content"])}
        if not llamadas:
            if ch.get("finish_reason") == "length":
                raise ErrorIA("La respuesta se cortó por longitud; haga una pregunta más acotada.")
            yield {"tipo": "fin"}
            return
        historial.append({"role": "assistant", "content": msg.get("content") or "", "tool_calls": llamadas})
        for c in llamadas:
            nombre = c["function"]["name"]
            try:
                args = json.loads(c["function"].get("arguments") or "{}")
            except json.JSONDecodeError:
                args = None
            yield {"tipo": "herramienta", "nombre": nombre, "args": args if isinstance(args, dict) else {}}
            contenido, _ = _resultado(nombre, args, ejecutar, validar) if nombre in herramientas else ("Herramienta no disponible.", True)
            historial.append({"role": "tool", "tool_call_id": c["id"], "content": contenido})
        yield {"tipo": "texto", "t": "\n\n"}
    raise ErrorIA("La consulta necesitó demasiados pasos; intente una pregunta más concreta.")
