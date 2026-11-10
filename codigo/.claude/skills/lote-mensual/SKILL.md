---
name: lote-mensual
description: Orquesta el cierre mensual de TODOS los clientes de LCCOUNTER (proceso TO-BE E1→E6) — corre el lote, reparte cada cliente a su etapa y subagente, y entrega la cola de validación del revisor. Úsala cuando pidan «corre el mes», «procesa todos los clientes de septiembre», «cómo va el lote» o «qué falta del mes».
argument-hint: <AAAA-MM>
---

# Lote mensual (orquestador)

Eres el **orquestador** del flujo. Los subagentes de `.claude/agents/` no pueden lanzar otros subagentes, por eso la
orquestación vive aquí, en la sesión principal: tú decides qué etapa sigue para cada cliente y a qué subagente se la
delegas. Trabaja desde `08-Version-Mejorada/codigo` con `.venv/bin/python -m agente --raiz ../espacio-trabajo …`
(abreviado `A`). Lee primero `CLAUDE.md`. Nunca ejecutes decisiones del revisor (el hook las bloquea).

| Etapa | Quién | Herramienta |
|---|---|---|
| Cliente nuevo o Excel con otra estructura | Claude | skill `/incorporar-cliente` (recepcion-ph, estructurador-excel, migrador-ph) |
| E1 recepción, T03–T06 | Claude | subagente `recepcion-ph` |
| E4 T09 generar y validar | Python | `A correr --periodo P` |
| E4 T11 opinión · T12 re-auditoría | Claude | `redactor-ph` · `reauditor-ph` |
| E4 T10 · E5 T13–T14 | **Revisor** | comandos que dejas listos |
| E6 T16–T17 mejora y RSI | Claude + revisor | skill `/mejora-continua` (mejora-continua-ph) |

## 1. Retomar (memoria)
Para cada cliente de `A estado --periodo P` corre `A memoria contexto <id> --periodo P`. Anota qué quedó abierto:
pendientes pedidos al cliente, recetas sin aprobar, mapa sin aprobar, hallazgos abiertos del mes anterior y
lecciones del flujo. No repitas trabajo que la memoria dice que ya se hizo.

## 2. Generar (T09)
`A correr --periodo P` (todo el lote en paralelo; un cliente con problemas queda BLOQUEADO y el lote sigue).

## 3. Repartir por estado
Lee `salidas/P/<id>/estado.json` de cada cliente y clasifica:
- **BLOQUEADO**: lee la última línea de la bitácora. Excel ausente, duplicado, de otro cliente o de otro mes → lo
  resuelve el auditor (anótalo como pendiente). Mapa o estructura → `recepcion-ph` y, si hace falta,
  `/incorporar-cliente`.
- **PENDIENTE_VALIDACION**: delega en paralelo, un cliente por subagente, `redactor-ph` (T11) y luego `reauditor-ph` (T12).
  Cuenta los hallazgos ALTA desde `hallazgos.json`.
- **DEVUELTO**: el auditor corrige el Excel; cuando llegue la nueva versión, vuelve al paso 2.
- **APROBADO**: nada que hacer en E4; pasa a E6.

## 4. Cola del revisor
Presenta una tabla en el orden del calendario (`orden` de la ficha): cliente · estado · ALTA/MEDIA · qué decide el
revisor · comando listo (`aprobar … --borrador-final … [--justificacion …]` o `devolver … --motivo …`). Abre el panel:
`A panel --periodo P`.

## 5. Cerrar la sesión (memoria)
Por cada cliente, registra lo que la próxima sesión necesita y no está en el sistema:
`A memoria anotar <id> --periodo P --tipo pendiente|particularidad|acuerdo --texto "…"`; cierra los pendientes
resueltos con `A memoria resolver <id> --nota N00x --texto "…"`. Lo que sirva para todos los clientes va como
`A memoria anotar flujo --tipo leccion --texto "…"`. Nunca anotes cifras ni datos de propietarios: solo hechos
operativos (qué falta, dónde está, qué se acordó).

Resumen final: una tabla por cliente con etapa alcanzada, pendiente y responsable (auditor, revisor, despacho).
