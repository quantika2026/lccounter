---
name: incorporar-cliente
description: Da de alta una copropiedad nueva en LCCOUNTER y adapta el flujo a SU Excel (cualquier estructura) para que salga el informe estándar — crea la ficha, carga el Excel, detecta el mapa de hojas, prepara recetas para lo que tenga otra estructura y, si se decide, la migra a la plantilla estándar. Úsala cuando digan «nuevo cliente», «crea la copropiedad X», «este Excel viene distinto», «incorpora a X».
argument-hint: <NIT> "<nombre de la copropiedad>" <AAAA-MM> [ruta del Excel]
---

# Incorporar un cliente nuevo

Desde `08-Version-Mejorada/codigo`, `A` = `.venv/bin/python -m agente --raiz ../espacio-trabajo`. Lee `CLAUDE.md`.
Se pueden crear todos los clientes que se quiera: cada uno tiene su carpeta `clientes/<NIT sin puntos>/`.

1. **Ficha**: `A nuevo-cliente <NIT> --copropiedad "<nombre>" [--software <perfil>]`. La ficha la completa una
   persona (app → Ficha del cliente): unidades, número de informe, revisor y, si el NIT aún no es el real,
   `nit_provisional: true`. Claude no edita `ficha.yaml` (el hook lo bloquea) ni inventa esos datos: pídelos.
2. **Excel**: cópialo a `../espacio-trabajo/entradas/<P>/<id>/` (uno solo; dos Excel bloquean el expediente).
3. **Diagnóstico** (subagente `recepcion-ph`): identidad, secciones ok / vacías / con otra estructura / faltantes.
   El mapa de hojas se detecta solo, por nombre y por contenido, la primera vez que corre `A correr`.
4. **Generar una primera vez**: `A correr --periodo P --clientes <id>`. Revisa `cobertura_esquema.json`:
   - «ok» → el lector estándar ya entiende esa sección.
   - «formato distinto», o «falta» con una hoja sin usar que la contiene → `estructurador-excel` (una sección por vez;
     verifica tú cada receta con `A probar-receta`).
   - vacía de verdad → faltante del cliente (T06): anótalo con `A memoria anotar <id> --tipo pendiente`.
5. **Revisor**: aprueba el mapa (app → Mapa del Excel) y las recetas (`aprobar-receta`). Deja los comandos listos;
   no los ejecutes. Al aprobar el mapa, sus nombres de hoja se aprenden para los siguientes clientes.
6. **Opcional, migración a la plantilla estándar** (una vez por cliente): subagente `migrador-ph`. Desde el mes
   siguiente el auditor trabaja sobre la plantilla y el flujo es idéntico para todos.
7. **Memoria**: registra las particularidades del Excel del cliente
   (`A memoria anotar <id> --tipo particularidad --texto "…"`) y lo aprendido para el flujo (`… flujo --tipo leccion`).

Entrega: ficha creada, mapa detectado (hoja por sección), recetas propuestas, faltantes para pedir al cliente,
diapositivas generadas y qué aprueba el revisor.
