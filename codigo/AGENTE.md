# Agente de Revisoría Fiscal PH (v0.5.0)

> v0.4: el agente genera el **informe completo** (estructura del modelo 006-2026) por defecto. Para solo la sección financiera, ponga `informe: financiero` en la ficha del cliente.

Convierte el MVP de un solo informe en un **sistema por lote**: uno solo, el mismo kit, para 15, 30 o 50 clientes.
Cada cliente pasa por los mismos agentes; el único paso que espera es la **validación del revisor fiscal**.

> **Python calcula, la IA organiza y redacta, el profesional decide.**

## Los agentes

| Agente | Módulo | Qué hace |
|---|---|---|
| Orquestador | `agente/orquestador.py` | Corre el lote en paralelo, lleva el estado de cada cliente y la cola de validación en el orden del calendario de visitas. Un cliente con problemas no detiene el lote. |
| Normalizador | `agente/normalizador.py` | Ubica el Excel estándar y revisa sus hojas. Lee el balance de prueba exportado del software contable con el **perfil** de ese software. Si el formato es nuevo, **propone** el mapeo y espera aprobación. Cuadra el Excel contra el balance (**V19**). |
| Validador | `revisoria/validador.py` | Las 18 reglas del MVP más la **V20** (total de la hoja ≠ suma de sus partidas). |
| Generador | `revisoria/generador_pptx.py` | Presentación institucional con tablas y gráficos nativos. |
| Redactor | `agente/redaccion.py` | Borrador de hallazgos, nivel de riesgo y recomendación desde la **biblioteca de opiniones** del revisor. Las cifras salen del modelo; la IA (opcional, `--llm`) solo pule la redacción, con candado de cifras. |
| Re-auditor | `agente/reauditor.py` | Revisa de forma independiente el PowerPoint ya generado: títulos frente al Excel (R01), cifras sin origen (R02), mes anterior (R03), secciones faltantes (R04), opiniones pendientes (R05), hallazgos ALTA por decidir (R06) y seguimiento de los abiertos del mes anterior (R07). |
| Mejora continua | `agente/mejora.py` | Captura cada corrección del revisor, la clasifica, propone ajustes, los prueba contra el **golden set** y publica versiones reversibles. |
| Auditoría periódica | `agente/auditoria.py` | Muestra mensual al azar de informes aprobados, con lista de chequeo para un segundo revisor. |
| Panel | `agente/panel.py` | `Panel_<periodo>.html`: estado del lote, cola de validación, hallazgos, cuadre, métricas y propuestas. |

## Mejora continua y automejora recursiva (RSI acotada)

**Mejora continua supervisada** (desde la v0.5): las correcciones del revisor, las devoluciones repetidas y la auditoría
mensual generan propuestas (`BIB-*`, `REG-*`, `AUD-*`…), y cada versión de la biblioteca se mide por correcciones por informe.

**Automejora recursiva** (`app/rsi.py`, desde la aplicación: ⚙ Configuración → Automejora recursiva):

| Nivel | Qué mejora | Cómo se decide |
|---|---|---|
| 1 | Plantillas de opinión de la biblioteca | La IA genera candidatos; Python mide su similitud con los textos finales del revisor en un caso **reservado** (no visto al generar) y verifica restricciones (solo marcadores permitidos, sin cifras escritas a mano, regresión contra el golden set). |
| 2 (recursivo) | La **estrategia** con la que el sistema genera candidatos | Con el historial de ganadores y perdedores, la IA reescribe su estrategia; la nueva compite contra la vigente y la reemplaza solo si gana la mayoría de las comparaciones. |

Modos: desactivado · proponer (el revisor aprueba) · automático (publica lo que mejora en la prueba reservada, como
versión reversible). **Límites fijos:** nunca modifica cálculos, extracción de cifras, reglas de validación ni código.

## Ciclo de vida de un informe

`RECIBIDO → (normalizar · validar · generar · redactar · re-auditar) → PENDIENTE_VALIDACION → APROBADO | DEVUELTO`

- **Nada sale sin aprobación.** El PDF solo se exporta al aprobar.
- Si hay hallazgos **ALTA**, el sistema no deja aprobar sin una `--justificacion` del revisor.
- Un informe aprobado no se regenera. Un informe devuelto se vuelve a correr cuando corrigen el Excel.
- Al aprobar se guardan: las correcciones (para la mejora continua), los hallazgos abiertos (para el seguimiento del mes siguiente) y el caso en el golden set.

## Espacio de trabajo

```
<raiz>/
  clientes/<id>/ficha.yaml        datos del cliente, software contable, revisor, notas del mes
  entradas/<periodo>/<id>/        Excel estándar + balance de prueba del software contable
  salidas/<periodo>/<id>/         informe, validación, borrador de opinión, re-auditoría, estado
  conocimiento/                   perfiles_software.yaml · biblioteca_opiniones.yaml · correcciones.jsonl
                                  propuestas.json · golden/ · versiones.json · metricas.json · auditorias/
  Panel_<periodo>.html
```

## Comandos

```bash
python3 -m agente --raiz <espacio> nuevo-cliente torres --copropiedad "Torres del Parque PH" --software generico
python3 -m agente --raiz <espacio> correr --periodo 2026-08 [--clientes a,b] [--paralelo 4] [--llm]
python3 -m agente --raiz <espacio> estado --periodo 2026-08
python3 -m agente --raiz <espacio> aprobar picasso --periodo 2026-08 --revisor "Nombre" \
        [--justificacion "..."] [--pptx-final editado.pptx] [--borrador-final editado.md]
python3 -m agente --raiz <espacio> devolver arboleda --periodo 2026-08 --revisor "Nombre" --motivo "..."
python3 -m agente --raiz <espacio> mejorar
python3 -m agente --raiz <espacio> aprobar-propuesta <id> --revisor "Nombre"
python3 -m agente --raiz <espacio> revertir bib-001
python3 -m agente --raiz <espacio> regresion
python3 -m agente --raiz <espacio> muestra --periodo 2026-08 --k 4
python3 -m agente --raiz <espacio> auditar picasso --periodo 2026-08 --auditor "Otro profesional" \
        --conclusion sin_observaciones|menores|reproceso [--observacion "..."]
```

## Demostración (10 minutos)

```bash
bash demo/demo.sh "/ruta/EXCEL PICASSO AGOSTO 2026 .xlsx"      # ENTER entre pasos
PAUSA=0 bash demo/demo.sh "/ruta/EXCEL PICASSO AGOSTO 2026 .xlsx"   # sin pausas
```

Cuatro clientes: **Picasso** (real) y tres **ficticios** derivados de él solo para mostrar el lote: Mirador (otro
software contable, con perfil aprobado), Arboleda (error sembrado en CXP y un software sin perfil) y San Rafael
(no entregó el Excel). Guion sugerido:

| Paso | Qué mostrar | Mensaje |
|---|---|---|
| 1 | El lote corre en segundos; abrir el panel | Un solo kit para todos los clientes |
| 2 | Aprobar sin justificar: el sistema lo impide | El profesional siempre decide |
| 3 | Aprobar Picasso con ediciones; abrir el PDF | Nada sale sin aprobación; se capturan las correcciones |
| 4 | Devolver Arboleda (regla V20) | El error digitado no llega al Consejo |
| 6–7 | `mejorar` y aprobar la propuesta | El sistema aprende la redacción de la revisora, probado contra el golden set y reversible |
| 8 | El borrador de Arboleda ya usa esa redacción | Automejora con control |
| 9 | Regresión y muestra mensual | Auditoría periódica del propio sistema |

## Lo que falta para producción

- **Perfiles reales** de los 4–5 programas contables (se necesitan exportes reales de cada uno).
- Secciones documentales y de control interno del informe (contratación, mantenimientos, actas).
- Integración con la carpeta compartida del despacho y la firma digital del PDF.
- Historial de correcciones de varios meses para que las métricas de mejora sean significativas.
