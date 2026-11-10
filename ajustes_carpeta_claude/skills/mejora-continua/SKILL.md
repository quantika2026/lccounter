---
name: mejora-continua
description: Etapa E6 de LCCOUNTER — analiza las correcciones del revisor, corre la automejora acotada (RSI) de las plantillas de opinión, la regresión contra el golden set y la auditoría mensual, y deja las propuestas listas para que el revisor decida. Úsala cuando pidan «mejora continua», «qué aprendió el sistema», «corre el RSI», «automejora», o después de aprobar los informes del mes.
argument-hint: [AAAA-MM]
---

# Mejora continua y automejora (RSI)

Desde `08-Version-Mejorada/codigo`, `A` = `.venv/bin/python -m agente --raiz ../espacio-trabajo`. Lee `CLAUDE.md`.
Qué puede mejorar el sistema solo y qué no: mejora **textos** (plantillas de opinión) y propone reglas; **nunca**
cambia cálculos, extracción de cifras, reglas de validación ni código. Todo queda versionado y se puede revertir.

1. **Correcciones → propuestas**: delega en `mejora-continua-ph` → `A mejorar`. Propuestas `BIB-*` (opinión),
   `TXT-*`, `CIF-*` (posible error de cálculo: va a un desarrollador), `EST-*`, `REG-*` (borrador de regla por
   devoluciones repetidas), `AUD-*` (auditoría mensual). Lo que el revisor aceptó tal como lo redactó Claude
   (origen «claude») no cuenta como aprendizaje.
2. **RSI** (dos niveles: mejora las plantillas y mejora su propia estrategia de mejora):
   - `A rsi-estado`: modo, estrategia vigente y retadora, casos por plantilla, propuestas pendientes.
   - `A rsi-ciclo`: un ciclo. Desde Claude solo corre en modo «proponer» (en «automático» el CLI lo rechaza porque
     publicaría sin revisor). Usa el proveedor de IA configurado en la app para generar candidatos; Python los mide
     contra un caso reservado del revisor y contra el golden set.
   - Con menos de 2 casos de una plantilla no hay prueba independiente: dilo y no lo presentes como mejora.
3. **Regresión**: `A regresion`. Si falla o no hay casos, nada se publica; repórtalo.
4. **Auditoría mensual**: `A muestra --periodo P` genera la muestra para el segundo revisor (no puede ser quien aprobó).
5. **Al revisor** (no lo ejecutes; el hook lo bloquea): `aprobar-propuesta <id> --revisor`, `rsi-aprobar <id> --revisor`,
   `rsi-rechazar <id> --revisor`, `revertir <version>`, `rsi-restaurar <estrategia>`, `auditar …`.
6. **Memoria**: `A memoria anotar flujo --tipo leccion --texto "…"` con lo aprendido que no sea una propuesta
   (p. ej. «la plantilla de cartera necesita más casos antes de medir»).

Entrega: tabla de propuestas (id · tipo · evidencia · ganancia medida · riesgo · comando para el revisor) y la
tendencia por versión de la biblioteca solo si cada versión tiene varios informes.
