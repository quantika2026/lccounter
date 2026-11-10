# Auditoría completa del flujo — 18 de octubre de 2026

Se auditó la versión mejorada (`08-Version-Mejorada/codigo`) y se comparó contra el sistema anterior (`01-MVP-codigo`).
Las dos versiones se corrieron sobre **copias aisladas** de sus espacios de trabajo. No se modificaron datos reales.
Cada afirmación de este documento sale de una prueba ejecutada, no de leer la documentación.

## Resumen

| | 01-MVP-codigo | 08-Version-Mejorada |
|---|---|---|
| Pruebas automáticas | 32 de 32 pasan | 34 pasan, **8 no corren** (error en el script de demo) |
| Lote de agosto (Picasso y Aldea) | Ambos pendientes de validación | Ambos pendientes de validación |
| Excel de otro cliente cargado por error | **Genera el informe** con los datos equivocados | **Bloquea** el expediente |
| Aprobar después de cambiar la ficha | **Aprueba** | **Rechaza** y pide reprocesar |
| Estructura del informe de Aldea | 41 diapositivas (omite las secciones vacías) | 49 diapositivas (estructura completa, con «información pendiente») |
| Deudores varios de Picasso | $5.299.410 (correcto) | **Vacío (regresión)** |

**Conclusión:** 08 es claramente más segura. Cierra los dos riesgos más graves del sistema anterior: el informe de un cliente armado con el Excel de otro, y la aprobación de una versión distinta de la validada. Antes de reemplazar a 01 hay que corregir una regresión de lectura y el script de demo, que hoy deja 8 pruebas sin correr.

## Mejoras verificadas (08 frente a 01)

1. **Identidad del expediente: verificada.** Se cargó el Excel de Aldea en el cliente Picasso, el mismo incidente que ocurrió el 14 de octubre en el espacio real.
   - **08** lo bloqueó: «El archivo no identifica a Conjunto Residencial Picasso; declara Conjunto Aldea de San Marco PH en EGRESOS!D49».
   - **01** generó un informe «de Picasso» con cifras de Aldea y lo dejó pendiente de validación.
2. **Aprobación ligada a la revisión: verificada.** Se cambió la ficha después de validar y luego se intentó aprobar.
   - **08** respondió: «los archivos, la ficha o el mapa cambiaron después de validar. Vuelva a procesar antes de aprobar».
   - **01** aprobó.
3. **Historial de cargas: observado.** El espacio real de 08 conserva las revisiones archivadas en `historial/2026-08` e `historial/2026-10`.
4. **Fórmulas sin calcular (V21) y periodo automático: verificados** por `test_integridad_v2.py`, cuyas 10 pruebas pasan. El expediente de octubre del espacio real se bloqueó correctamente, porque el Excel cargado es el de agosto.
5. **Informe estándar: verificado.** Cuando falta una sección, 08 conserva la diapositiva con «Información pendiente en los papeles de trabajo» en vez de omitirla. Esto cumple el requisito de que el PowerPoint no varíe entre clientes.
6. **JSON atómico: verificado en el código** (`agente/comun.py`, escritura en un temporal y reemplazo).
7. **Se conserva lo de 01:**
   - mapa del Excel, esquema estándar y firma de formato;
   - reparación del mapa y aprendizaje entre clientes;
   - secciones opcionales que no bloquean;
   - varios proveedores de IA y RSI.

## Hallazgos

| # | Severidad | Hallazgo | Evidencia | Acción |
|---|---|---|---|---|
| A1 | **ALTA** | **Regresión: 08 no lee deudores varios cuando el título no trae el año.** El título es «DEUDORES VARIOS A AGOSTO CTA 1380». La regla nueva «mes y año» descarta la tabla correcta. Se pierde la cifra ($5.299.410, celda D172) y desaparece el hallazgo V15 (saldo antiguo de Soluciones Propiedad Horizontal SAS). | Mismo Excel: 01 lee D172 y 08 devuelve vacío. | Si el título no tiene año, aceptar el año del periodo cuando el bloque es el último del año en la hoja. Agregar una prueba con este título. |
| A2 | **ALTA** | **8 pruebas del agente no corren en 08.** `demo/preparar_demo.py` formatea `total_hoja = None` (efecto de A1 y de la mejora V21). | `TypeError ... NoneType.__format__` en `balance_desde_modelo`. | Hacer que la demo tolere `None` y, cuando A1 esté corregido, volver a correr las 42 pruebas. |
| A3 | **ALTA** (operativa) | **El proyecto está en el Escritorio sincronizado con iCloud.** macOS descarga y desaloja archivos de `.venv` sobre la marcha. Importar las librerías tardó 3 min 56 s con 0,4 s de CPU. | Pruebas: 18 s antes, entre 6 y 10 min hoy. Lote de 2 clientes: 3,5 min. | Mover el proyecto fuera de iCloud (p. ej. `~/Proyectos/LCCOUNTER`), o como mínimo renombrar `.venv` a `.venv.nosync`. |
| A4 | MEDIA | **El nombre de otro cliente aparece en los entregables.** Las diapositivas «pendientes» dicen «El ítem se conserva para mantener la estructura del informe **modelo Picasso**». En el informe de Aldea eso expone a otro cliente. | 8 diapositivas del informe de Aldea en 08. | Cambiar el texto a «…estructura estándar del informe». |
| A5 | MEDIA | **El NIT de Aldea en 08 es sintético pero parece real** (900.123.456-7), y además es el identificador del cliente (carpeta `9001234567`). Cuando llegue el NIT real habrá que migrar la carpeta y los expedientes. 08 no tiene el control F01 de 01, que bloquea un NIT provisional. | `clientes/9001234567/ficha.yaml`. | Traer F01 a 08 y definir el procedimiento de cambio de NIT (renombrar la carpeta y conservar el historial). |
| A6 | MEDIA | **Hay dos Excel en la misma entrada** (`…2026-08.xlsx` y `…2026-08 2.xlsx`, copia de Finder). Hoy son idénticos (misma huella SHA-256), pero el sistema toma el primero en orden alfabético sin avisar. | `espacio-trabajo/entradas/2026-08/9001234567/`. | Si hay más de un Excel candidato, bloquear y pedir cuál usar. Borrar la copia. |
| A7 | MEDIA | **Hay dos versiones que divergen.** 01 tiene F01 y 08 tiene la integridad. Cada corrección hecha en una sola versión se pierde en la otra. | Comparación de código. | Adoptar 08 como versión única, traer F01 y dejar 01 solo como respaldo. |
| A8 | BAJA | Con el número de informe vacío, el pie de página queda «Informe  de Revisoría Fiscal», con doble espacio. | Diapositivas de Aldea. | Omitir el espacio cuando no hay número. |
| A9 | INFO | Siguen pendientes en el espacio real de 01: el Excel de Picasso es en realidad el de Aldea (desde el 14 de octubre), y ningún mapa está aprobado, así que el aprendizaje sigue en cero. | `07-Espacio-trabajo`. | Recargar el Excel de Picasso y aprobar los mapas. |

## Hallazgos de negocio que confirman las dos versiones

- **Picasso:** 4 hallazgos ALTA del Excel.
  - **V04:** la cartera no coincide con su hoja.
  - **V05:** la disponibilidad no cuadra con los excedentes.
  - **V18:** los rubros sobreejecutados están subestimados.
  - **V08:** hay diferencia en la reserva del fondo de imprevistos.
- **Aldea San Marco:** 1 hallazgo ALTA, **V11** (diferencia no explicada en el banco AV Villas, cuenta 43367-4).
- **Aldea San Marco, secciones vacías en el Excel:** anticipos, deudores varios, actas de consejo, contratos y pólizas. Además, la matriz de riesgo existe con el nombre «Matriz riesgo 2» y aún no está asignada en el mapa.

## Orden recomendado

1. **A3:** mover el proyecto fuera de iCloud. Es lo más rápido y afecta todo lo demás.
2. **A1 y A2:** corregir la regresión y la demo, y volver a correr las 42 pruebas.
3. **A4, A5 y A7:** quitar «modelo Picasso» de los entregables, traer F01 a 08 y adoptar 08 como versión única.
4. **A6 y A9:** limpiar las entradas duplicadas, recargar el Excel de Picasso y aprobar los mapas para que empiece el aprendizaje.

## Correcciones aplicadas (18 de octubre de 2026, mismo día)

| # | Estado | Qué se hizo | Verificación |
|---|---|---|---|
| A1 | ✅ Corregido | `extractor._bloque_del_mes`: un título sin año («DEUDORES VARIOS A AGOSTO CTA 1380») se acepta solo si es el último bloque de su tipo y el año vigente de la hoja coincide con el periodo. Queda marcado con `anio_supuesto` para trazabilidad. | Picasso vuelve a leer $5.299.410 (D172) y recupera V15. Pedir agosto de 2025 o de 2027 no devuelve nada. Prueba nueva. |
| A2 | ✅ Corregido | La demo omite del balance simulado las cuentas sin cifra, en vez de inventarlas o fallar. | **45 de 45 pruebas pasan** (antes: 34 pasaban y 8 no corrían). |
| A3 | ✅ Mitigado | `.venv` → `.venv.nosync`, con enlace `.venv` para que los comandos no cambien. iCloud deja de quitar las librerías del disco. | Importar las librerías: de **4 min 21 s a 0,3 s**. Suite completa: de 6–10 min a **21 s**. Mover el proyecto fuera de iCloud sigue siendo lo ideal. |
| A4 | ✅ Corregido | Texto neutro: «…mantener la estructura estándar del informe». | El informe de Aldea ya no menciona a Picasso. |
| A5 | ✅ Corregido | La ficha admite `nit_provisional: true`, que genera el hallazgo **F01 ALTA**. Nuevo «Cambiar NIT…» (en la ficha, el endpoint `POST /api/clientes/{id}/nit` y el comando `cambiar-nit`): mueve la ficha, el mapa, las entradas, las salidas y el historial al NIT real, y deja registrado el NIT anterior. Se niega si ya hay informes aprobados. | Aldea marcada como provisional y F01 aparece. Tras cambiar el NIT, F01 desaparece. Prueba nueva. |
| A6 | ✅ Corregido | Con dos o más Excel en la entrada, el expediente se bloquea y se nombran los archivos. La copia «… 2.xlsx» de Aldea (idéntica) se movió a `historial/duplicados/`. | Prueba nueva. |
| A7 | ✅ Resuelto | F01 ya está en 08. 08 tiene todo lo de 01 más la integridad. | — |
| A8 | ✅ Corregido | Sin número de informe, el pie queda «Informe de Revisoría Fiscal», sin doble espacio. | Verificado en el PPTX de Aldea. |
| A9 | ⏳ Pendiente (le corresponde al revisor) | Recargar el Excel real de Picasso y aprobar los mapas. Picasso todavía no existe como cliente en 08. | — |
