#!/bin/bash
# Doble clic para abrir LCCOUNTER en este Mac. La aplicación solo escucha en 127.0.0.1:
# los datos de los clientes no salen del equipo.
cd "$(dirname "$0")/codigo" || exit 1
PY=".venv/bin/python"
[ -x "$PY" ] || PY=".venv.nosync/bin/python"
if [ ! -x "$PY" ]; then
  echo "No encuentro el entorno de Python. Instálelo con:"
  echo "  python3 -m venv .venv.nosync && ln -sfn .venv.nosync .venv && .venv/bin/pip install -r requirements.txt"
  read -r -p "Enter para cerrar"; exit 1
fi
"$PY" -c "import fastapi, uvicorn, openpyxl, pptx, yaml" 2>/dev/null || "$PY" -m pip install -q -r requirements.txt
if ! command -v soffice >/dev/null && [ ! -x /Applications/LibreOffice.app/Contents/MacOS/soffice ]; then
  echo "Aviso: LibreOffice no está instalado. Sin él no hay PDF, recálculo de fórmulas ni imágenes EMF."
fi
PUERTO=8790
if lsof -nP -iTCP:$PUERTO -sTCP:LISTEN >/dev/null 2>&1; then
  echo "LCCOUNTER ya está corriendo."
else
  nohup "$PY" -m app.servidor --raiz ../espacio-trabajo --puerto $PUERTO >> ../espacio-trabajo/servidor.log 2>&1 &
  for i in $(seq 1 30); do curl -s -o /dev/null "http://127.0.0.1:$PUERTO/" && break; sleep 1; done
fi
open "http://127.0.0.1:$PUERTO/"
echo "LCCOUNTER abierto en http://127.0.0.1:$PUERTO  (para detenerlo: doble clic en Detener_LCCOUNTER.command)"
