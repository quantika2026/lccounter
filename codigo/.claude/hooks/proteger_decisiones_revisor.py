#!/usr/bin/env python3
"""Hook PreToolUse de Claude Code: Claude nunca toma las decisiones del revisor fiscal (CLAUDE.md).

Bloquea (código de salida 2) los comandos y llamadas que aprueban, devuelven, revierten o cambian el NIT, y la
edición directa de los archivos que guardan esas decisiones. El mensaje vuelve a Claude para que deje el comando
listo al revisor en lugar de ejecutarlo.
"""
import json
import re
import sys

COMANDOS = (r"(aprobar|aprobar-receta|aprobar-propuesta|revertir|cambiar-nit|devolver|auditar|"
            r"rsi-aprobar|rsi-rechazar|rsi-restaurar)")


def _reglas():
    return [
        re.compile(r"-m\s+agente\b.*?\s" + COMANDOS + r"(\s|$|[\"';])"),                    # CLI del agente
        re.compile(r"\b(orquestador\.(aprobar|devolver)|mejora\.(aprobar_propuesta|revertir|publicar_plantilla)|"
                   r"estructurador\.aprobar|integridad\.cambiar_nit|auditoria\.registrar_conclusion|"
                   r"rsi\.(aprobar|rechazar|restaurar_estrategia|guardar_config))\s*\("),         # Python directo
        re.compile(r"/api/(aprobar|devolver|auditar|clientes/[^/\s]+/nit|propuestas/[^/\s]+/aprobar|"
                   r"versiones/[^/\s]+/revertir|receta/|mapa/[^/\s]+/[^/\s]+(\s|[\"']|$)|rsi/propuestas/|"
                   r"rsi/estrategias/|configuracion/|usuarios)"),                               # app web
        re.compile(r"\bCLAUDECODE\s*=|unset\s+CLAUDECODE|env\s+-u\s+CLAUDECODE"),            # saltarse la guarda
    ]


# archivos que registran decisiones o aprendizaje aprobado
ARCHIVOS = re.compile(r"espacio-trabajo/(salidas/[^/]+/[^/]+/estado\.json|conocimiento/|clientes/[^/]+/(mapa\.json|ficha\.yaml))")


def main():
    try:
        datos = json.load(sys.stdin)
    except Exception:
        return 0
    herramienta, entrada = datos.get("tool_name", ""), datos.get("tool_input") or {}
    motivo = None
    if herramienta == "Bash":
        cmd = entrada.get("command", "")
        if any(rx.search(cmd) for rx in _reglas()):
            motivo = "es una decisión del revisor fiscal"
        elif ARCHIVOS.search(cmd) and re.search(r"(>|\bsed\s+-i|\btee\b|\bmv\b|\bcp\b|\brm\b|write_text|open\([^)]*['\"]w)", cmd):
            motivo = "modifica un archivo que registra decisiones del revisor"
    elif herramienta in ("Write", "Edit", "MultiEdit", "NotebookEdit"):
        ruta = entrada.get("file_path") or entrada.get("notebook_path") or ""
        if ARCHIVOS.search(ruta.replace("\\", "/")):
            motivo = "modifica un archivo que registra decisiones del revisor"
    if motivo:
        print(f"Bloqueado: {motivo} (CLAUDE.md, «Reglas que no se negocian»). No lo ejecutes: deja el comando o el "
              "botón de la app listo y pide al revisor fiscal que lo haga.", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:                                   # si la guarda falla, se bloquea (no se deja pasar)
        print(f"Bloqueado: el hook de protección falló ({e}). Revise .claude/hooks/proteger_decisiones_revisor.py.",
              file=sys.stderr)
        sys.exit(2)
