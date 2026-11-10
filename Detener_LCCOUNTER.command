#!/bin/bash
# Detiene la aplicación LCCOUNTER que abrió Iniciar_LCCOUNTER.command.
PID=$(lsof -nP -tiTCP:8790 -sTCP:LISTEN 2>/dev/null)
if [ -n "$PID" ]; then kill $PID && echo "LCCOUNTER detenido."; else echo "LCCOUNTER no estaba corriendo."; fi
