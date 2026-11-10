# Mejoras incorporadas

La opción elegida fue evolucionar el sistema actual en una copia independiente. Así se conserva lo que ya funciona, se pueden comparar resultados y el proyecto anterior sigue disponible como respaldo.

## Cambios prioritarios resueltos

1. **Identidad del expediente.** Antes de generar, se contrastan la copropiedad, el NIT, el número de unidades y el periodo con el Excel recibido.
2. **Aprobación ligada a una revisión.** Excel, balance, ficha y mapa producen una huella conjunta. Si cualquiera cambia, la aprobación anterior deja de ser válida.
3. **Historial de cargas.** Al reemplazar entradas se archivan los archivos y salidas de la revisión anterior.
4. **Fórmulas sin calcular.** Ya no parecen valores cero. Se reportan como V21 y bloquean la aprobación hasta recalcular el libro.
5. **Selección correcta de periodos.** El periodo automático ignora columnas futuras vacías y los anticipos se buscan por mes y año.
6. **Redacción asistida protegida.** Una propuesta se descarta si cambia el orden, signo, cantidad, formato o porcentaje de alguna cifra.
7. **Presentación final reauditada.** La versión editada, incluidos sus gráficos, se revisa antes de convertirse en el archivo final aprobado.
8. **Persistencia segura.** Los archivos JSON se escriben de forma atómica y un JSON dañado genera un error visible.
9. **Aprendizaje reversible.** Las regresiones conservan el mapa del cliente y una reversión ya no puede borrar accidentalmente la biblioteca.
10. **Cliente identificado por NIT.** Al crear una copropiedad, su carpeta y todos sus expedientes usan los diez dígitos del NIT como identificador estable; el formato visible se conserva como `900.123.456-7`.

## Ubicación

- Sistema anterior: `01-MVP-codigo/`
- Versión mejorada: `08-Version-Mejorada/codigo/`
- Datos de la versión mejorada: `08-Version-Mejorada/espacio-trabajo/`

No se migraron clientes ni expedientes automáticamente. Esto evita mezclar datos reales con la prueba de la nueva versión.

## Siguiente fase recomendada

Probar un expediente real completo en la copia, comparar sus cifras y presentación contra el resultado anterior, y solo después migrar de forma explícita las fichas de clientes aprobadas. Para operación continua conviene añadir respaldo externo cifrado, recuperación probada y firma o sellado de los entregables emitidos.

## Excel con otra estructura → informe estándar (18 de octubre de 2026)

Cada cliente puede organizar su Excel a su manera. El informe sale siempre con la estructura del informe modelo (Picasso).

- **Recetas de lectura** (`revisoria/lectores.py`). Una receta es un JSON sin código: hoja, bloque del mes (`"titulo": "cuentas por pagar a {mes} {anio}"`) o fila de encabezado, y qué columna es cada campo estándar. Un lector genérico la ejecuta y entrega la misma forma que el lector nativo. Sirve hoy para cuentas por pagar, anticipos, deudores varios, retención, egresos, pólizas, matriz de riesgo, matriz de manuales y actas de consejo.
- **Entrenamiento por ejemplos aprobados** (no se reentrena ningún modelo). `config/lectores_modelo.json` guarda las 9 recetas del informe modelo con el vistazo de su hoja. `tests/test_lectores.py` exige que reproduzcan **exactamente** al lector nativo. Cada receta que un revisor aprueba se suma a `conocimiento/lectores_aprendidos.json` y el estructurador la usa como ejemplo con el siguiente cliente.
- **Subagente estructurador** (`agente/estructurador.py`). Mira la hoja con coordenadas, propone la receta con los ejemplos, el sistema la prueba y le devuelve los errores para que se corrija (hasta 3 intentos). Dice cuándo una hoja **no** contiene la sección. Nunca aprueba: lo hace el revisor.
  - **En la app:** Mapa del Excel → «Otra estructura…» en cada sección. Permite proponer, editar la receta a mano, probarla con vista de lo leído, aprobarla y quitarla.
  - **Por comandos:** `estructurar`, `probar-receta` y `aprobar-receta`.
  - **En Claude Code:** el subagente `.claude/agents/estructurador-excel.md`.
  - **Proveedor de IA:** el uso «estructurador» en Configuración. Por defecto hereda el del asistente.
- **Resultados verificados con DeepSeek:**
  - **Aldea, «Matriz procedimientos»** (columnas corridas): se corrigió solo en el 2.º intento y leyó el texto real de E+F.
  - **Aldea, «POLIZA»** (pagos de pólizas): respondió que no es la verificación de pólizas.
  - **Hoja de proveedores con otra estructura** (dos bloques mensuales, otro título, otro orden de columnas): la resolvió al 1.er intento. Tomó agosto, no julio, y el total cuadra exacto.
  - **Aldea, «Matriz riesgo 2»:** con la receta aprobada, las diapositivas estándar de la matriz de control interno salen con sus datos en vez de «Información pendiente».
- **Corrección encontrada al reproducir Picasso:** la fuente de cada anticipo y deudor vario apuntaba a la columna DETALLE en vez de VALOR.
- **Liquidez, cartera y presupuesto** (agregadas el mismo día). Estas recetas llevan, además de columnas, los textos de las filas clave (encabezado de meses, total, monetización, resumen). La columna del mes se busca sola por su encabezado, aunque venga como texto («ago-26», «agosto 2026»).
  - **Cartera** admite dos diseños: listas por mes con título, con columnas relativas al título («+1», «+2»), o una tabla con una columna por mes.
  - **Presupuesto** detecta si el total de la hoja tiene el signo invertido.
  - **Verificación:** las 3 recetas de Picasso reproducen idéntico al lector nativo. Con DeepSeek, las tres estructuras distintas de prueba se resolvieron al 1.er intento. Al cargar un Excel, el periodo se detecta con la receta de liquidez del cliente, si la tiene.
- **Pendientes:**
  - Patrimonio, deudores, reservas y conciliaciones aún no admiten receta.
  - La matriz de manuales se lee, pero el generador del informe todavía no tiene diapositiva para ella.

## Opción C: plantilla estándar y migración de clientes (18 de octubre de 2026)

Cada cliente se migra **una vez** a la plantilla estándar y desde entonces el auditor la diligencia cada mes. Así el flujo mensual es idéntico para todos.

- **Plantilla estándar v1** (`revisoria/plantilla.py`). Tiene 22 hojas con los nombres oficiales (la estructura del informe modelo), una hoja de control con instrucciones, columnas de meses de diciembre a diciembre, fórmulas en los totales y celdas amarillas para lo que diligencia el auditor. La plantilla vacía de 2026 está en `Plantilla_Estandar_LCCOUNTER_v1_2026.xlsx`, y el lector estándar la lee sin errores.
- **Escritores por sección.** Ponen cada sección en la disposición estándar, con un comentario de origen («Origen: Hoja!Celda») en cada dato. Conservan negrillas, rojos e imágenes de soporte. El ida y vuelta (escribir y volver a leer) da lo mismo en todas las secciones con escritor.
- **Migrador** (`agente/migrador.py`, comando `migrar` y subagente `migrador-ph`). Parte del Excel original:
  - renombra las hojas que el lector estándar ya entiende y corrige las fórmulas que las citan;
  - reconstruye las de otra estructura, usando recetas, y deja la original como «ANEXO»;
  - crea vacías las secciones que faltan;
  - recalcula con LibreOffice;
  - **verifica** con el lector estándar sin mapa ni recetas, sección por sección y hallazgo por hallazgo.

  El resultado es una propuesta en `clientes/<id>/migracion/…`, con su reporte.
- **Resultados:**
  - **Picasso** (prueba de identidad): ACEPTABLE, con 22 hallazgos iguales.
  - **Piloto Aldea San Marcos:** ACEPTABLE. 11 hojas renombradas, la matriz de riesgo reconstruida desde «Matriz riesgo 2» (receta propuesta, pendiente de aprobación), 2 hojas vacías y 4 anexos. Todas las secciones son idénticas y los 11 hallazgos coinciden.
- **Errores encontrados en el camino y corregidos:**
  - El validador reconocía las hojas financieras solo por el nombre exacto de Picasso, así que las reglas V01, V03 y V21 no revisaban, por ejemplo, la hoja de deudores de Aldea.
  - El lector de presupuesto tomaba el título «RUBROS SOBRE EJECUTADOS» como encabezado. En Picasso eso agregaba un rubro falso «Rubro» y perdía el encabezado del periodo en la diapositiva.
- **Pendiente:**
  - Las secciones que se migran renombrando conservan la estructura original del cliente cuando ya es compatible. Si el despacho quiere un formato visual uniforme, falta un paso de normalización estética.
  - La adopción de la plantilla por los auditores la decide el despacho.
- **Escritor de «may deudores»** (19 de octubre de 2026). Escribe la hoja de seguimiento de cobro con saldo anterior y actual, gestión del abogado, proceso y observación de la revisoría. Conserva las observaciones combinadas que aplican a varios inmuebles, y las negrillas y rojos. El ida y vuelta con los 13 deudores de Picasso da lo mismo. La migración **no** crea esta hoja vacía cuando el cliente no la tiene: en ese caso el informe toma los mayores deudores de la evolución mensual, y una hoja vacía los haría desaparecer.

## Ajustes del flujo con Claude (23 de octubre de 2026)

Resultado de la revisión del flujo agéntico. Pruebas: **85 de 85 pasan** (70 anteriores + 15 nuevas en `tests/test_flujo_claude.py`).

- **Las decisiones del revisor ya no dependen solo de las instrucciones.** El hook `.claude/hooks/proteger_decisiones_revisor.py`
  (registrado en `.claude/settings.json`) bloquea a Claude: `aprobar`, `devolver`, `aprobar-receta`, `aprobar-propuesta`, `revertir`,
  `cambiar-nit` y `auditar`; las mismas acciones llamadas desde Python o desde la API de la app; y la edición directa de
  `estado.json`, `conocimiento/`, `mapa.json` y `ficha.yaml`. Además, el CLI rechaza esos comandos cuando corre dentro de Claude Code.
  El revisor los sigue usando igual desde su terminal o la app.
- **La redacción de Claude no se aprende como si fuera del revisor.** `redactor-ph` escribe `Borrador_propuesta_claude.md`.
  Al aprobar con `--borrador-final`, las opiniones que quedaron exactamente como las propuso Claude se registran con origen
  «claude» y no generan propuestas `BIB-*` ni cuentan en las métricas de opiniones editadas.
- **Habilidad `/informe-mensual`:** el paso 8 deja listo `--borrador-final` (antes la opinión redactada se perdía) y el paso 5
  deja listo `devolver` cuando el revisor dice que no (antes la devolución no quedaba registrada).
- **Migración: imágenes en su comprobante.** LibreOffice movía los soportes de egresos a otras filas al recalcular. Ahora se
  devuelven a su celda original después del recálculo (`libreoffice.restaurar_anclas_imagenes`). Picasso vuelve a migrar ACEPTABLE.
- **Conteo de ALTA al aprobar:** R06 ya no se cuenta dos veces (Picasso: 5, antes decía 6).
- **`--raiz` obligatorio:** sin él (o sin `REVISORIA_RAIZ`) el comando falla, en vez de crear un espacio `trabajo` nuevo.
- **Nuevo comando `diagnosticar <id> --periodo P`** (solo lectura). `recepcion-ph` lo usa en lugar del script pegado en sus instrucciones.

## Excel de cada cliente con otra estructura (23 de octubre de 2026)

Se probó con 3 clientes creados desde cero (`nuevo-cliente`): Picasso, Aldea (Excel real con otra estructura) y una
variante de Picasso con hojas renombradas y columnas reordenadas. Los tres salen con el **mismo informe estándar**:
las mismas secciones en el mismo orden; solo cambian las cifras y las páginas de evidencias. Pruebas: **87 de 87**.

- **Detección por contenido.** Antes el mapa del cliente nuevo se armaba solo por el nombre de la hoja: «PROVEEDORES Y
  CUENTAS», «EJECUCION PPTO» o «COMPROBANTES MES» no se encontraban. Ahora, si el nombre no calza, se busca la única hoja
  libre cuyo contenido tiene la firma de la sección (p. ej. «cuentas por pagar» en sus encabezados). Con varias
  candidatas no adivina y la deja al revisor. La variante de prueba quedó mapeada completa sin intervención.
- **Falso positivo corregido.** «EGRESO» se asignaba a «FIRMAS E INF DE GESTIÓN» por la letra «e»; ya no.
- **Una hoja de cuentas por pagar ausente ya no bloquea al cliente.** El informe sale con «Información pendiente» en esa
  diapositiva y un hallazgo M02 **ALTA** (exige justificación al aprobar). Lo mismo para cualquier sección obligatoria
  vacía o faltante en el esquema.
- **Lo que sigue necesitando receta:** hojas con otra disposición de columnas que el lector estándar no entiende (p. ej. la
  «Matriz riesgo 2» de Aldea). Las prepara el subagente `estructurador-excel` y las aprueba el revisor.
- **Aún no soportado:** liquidez, deudores, patrimonio y conciliaciones siguen siendo indispensables (sin ellas el cliente
  queda bloqueado con el motivo), y patrimonio, deudores, reservas y conciliaciones no admiten receta.

## Memoria persistente, automejora (RSI) y orquestación con habilidades (23 de octubre de 2026)

Revisión: el flujo ya tenía la memoria del sistema (`conocimiento/`), una automejora RSI de dos niveles en la app
(`app/rsi.py`) y 6 subagentes, pero la sesión de Claude no recordaba nada entre sesiones, el RSI no se podía operar
desde Claude y solo había una habilidad. Pruebas: **90 de 90**.

- **Memoria de Claude entre sesiones** (`agente/memoria.py`, comando `memoria`): notas por cliente (pendientes T06,
  particularidades del Excel, acuerdos) y lecciones del flujo en `espacio-trabajo/memoria/`. `memoria contexto` junta
  esas notas con el estado del mes actual y anterior, el mapa, las recetas y los hallazgos abiertos.
- **RSI desde Claude**: `rsi-estado` y `rsi-ciclo` (solo en modo «proponer»; en «automático» el CLI se niega porque
  publicaría sin revisor). `rsi-aprobar`, `rsi-rechazar` y `rsi-restaurar` son del revisor y quedan bloqueados para Claude.
- **El RSI ya no aprende del texto de Claude**: los casos con origen «claude» se excluyen.
- **Habilidades**: `/lote-mensual` (orquestador del mes), `/incorporar-cliente` (cliente nuevo y Excel distinto) y
  `/mejora-continua` (E6 + RSI), además de `/informe-mensual`, que ahora empieza y termina con la memoria.
  El orquestador es la sesión principal porque los subagentes no pueden lanzar otros subagentes.
- **Hook ampliado**: bloquea también la configuración de la app (incluido el modo del RSI) y las aprobaciones del RSI;
  si el propio hook falla, bloquea en vez de dejar pasar.

## Prueba completa de punta a punta (23 de octubre de 2026)

Pruebas automáticas **91 de 91**, demo guiada de 11 pasos sin errores, y un ciclo real en un espacio aislado con 4
clientes nuevos: carga, detección, receta, aprobación con PDF, mejora continua, regresión, auditoría mensual,
migración ACEPTABLE, mes siguiente, cambio de NIT y la app web por su API (ingreso, validar, aprobar, sin sesión → 401).
Correcciones de esta prueba:
- `probar-receta` deja la receta como propuesta: antes solo `estructurar` (proveedor de IA de la app) la dejaba, y la
  receta escrita por Claude no se podía aprobar con `aprobar-receta`.
- El RSI sin clave de IA configurada mostraba un error técnico; ahora dice qué configurar.

## Ronda 2026-10-24 — imágenes EMF/WMF del Excel

- openpyxl descarta las imágenes EMF/WMF al abrir y al guardar. Ahora `revisoria/libreoffice.imagenes_vectoriales`
  las lee directo del zip, las convierte a PNG con LibreOffice (recortando el margen blanco) y `documental._imagenes_por_fila`
  las entrega con su celda de anclaje, igual que las demás imágenes (egresos, auditoría documental, recetas).
- La migración a la plantilla estándar las repone como PNG en su celda original (`migrador._reponer_vectoriales`).
- Picasso agosto 2026: 7 EMF. 1 en AUD CONT Y ADMON!C17 (tabla de retención en la fuente, ítem «Obligaciones
  tributarias»); 6 en «Hoja3» (fotos de liquidez, deudores, anticipos, CXP, resultado y excedentes), hoja que el informe
  no usa. Todas son copias de tablas que el informe ya construye en nativo desde sus hojas fuente (IMPUESTOS, etc.),
  así que la diapositiva de Retención sigue mostrando la tabla, no la foto (igual que el informe modelo 006-2026).
- Sin LibreOffice la recuperación devuelve [] y el informe sale igual. Prueba nueva: test_imagen_emf_del_excel_se_recupera.
  92/92 pruebas.

## Ronda 2026-11-01 — revalidación

- Código del Mac idéntico al probado (incluida la carpeta .claude ya copiada). 93/93 pruebas con LibreOffice.
- Corrida real del lote sobre una copia del espacio de trabajo: Aldea y Picasso quedan PENDIENTE_VALIDACION con el
  Excel correcto de cada uno; un cliente nuevo con hojas renombradas se procesa sin intervención.
- Corregido: si dos hojas libres mencionan la firma de una sección (p. ej. «Cuentas x pagar» y «Matriz procedimientos»
  con «cuentas por pagar»), la detección por contenido no asignaba ninguna. Ahora desempata por la hoja que ABRE con
  ese título (`mapa._abre_con_firma`); si sigue habiendo empate, no adivina.
- Pendiente en el espacio de trabajo real: las carpetas de Picasso 2026-08 y Aldea 2026-10 tienen el Excel de Aldea
  de agosto (el flujo las bloquea con razón); NIT de Aldea provisional; ningún mapa aprobado.

## Ronda 2026-11-04 — cliente nuevo Arboleda Park (formato distinto)

- Ninguna sección financiera bloquea ya el informe: si la hoja no está o tiene otro formato, la sección sale vacía,
  la validación la marca M02 ALTA y la diapositiva se conserva como «Información pendiente» con el motivo.
- Detección de hojas con variantes de encabezado («a|b» en FIRMAS): PATRIMONIO «resultado ejercicio», MAYORES DEUDORES
  «inmueble», CONTRATOS «vencimiento», matriz «nivel riesgo».
- Lectores más flexibles: egresos acepta «COMPROBANTES DE EGRESOS <MES>», deudores encuentra la tabla en cualquier
  columna, reservas acepta un cuadro con columna «Valor». Receta de presupuesto sin fila total.
- Concentración de cartera: el título usa el tramo más alto con inmuebles (no «0 inmuebles concentran el 0%»).
  Portada sin «None Aptos» cuando la ficha no tiene unidades. La verificación de identidad muestra el nombre corto que
  declara el archivo («EDIFICIO ARBOLEDA PARK»), no un párrafo.
- Recetas propuestas para Arboleda (liquidez, cartera, presupuesto) en recetas_arboleda/: las aprueba el revisor.
- Arboleda queda con pendientes reales del Excel: CXP sin bloque de agosto (el último es junio), PATRIMONIO sin
  encabezados de mes, y no hay hojas de anticipos, impuestos ni caja menor. 96/96 pruebas.
