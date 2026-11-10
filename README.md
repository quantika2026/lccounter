# LCCOUNTER — versión mejorada

Esta carpeta es una evolución independiente de `01-MVP-codigo`. No modifica el código anterior ni usa su espacio de trabajo.

- Código: `codigo/`
- Espacio de trabajo aislado: `espacio-trabajo/`
- Objetivo inicial: asegurar identidad, periodo, trazabilidad de fuentes y aprobación por versión.

## Inicio

```bash
cd codigo
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m app.servidor --raiz ../espacio-trabajo
```

La migración de clientes o expedientes del sistema anterior debe hacerse de forma explícita después de validar esta versión.
