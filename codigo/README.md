# LCCOUNTER — versión mejorada independiente

> Esta es la versión mejorada independiente. Usa `../espacio-trabajo` y no modifica la versión original.

## Controles de integridad añadidos

- El identificador permanente de cada cliente es su NIT completo, normalizado sin puntos ni guion.
- Cada carga crea una revisión identificada por las huellas de Excel, balance, ficha y mapa.
- La carga anterior y sus salidas se conservan en `espacio-trabajo/historial/`.
- Si cambian las fuentes después de validar, la aprobación se rechaza.
- Antes de generar se contrastan copropiedad, NIT, unidades y periodo con el archivo.
- Las fórmulas sin resultado guardado se reportan como V21; no se interpretan silenciosamente como cero.
- Anticipos y deudores varios se seleccionan por mes y año completos.
- El texto asistido solo se acepta si conserva exactamente signos, cifras, porcentajes y repeticiones.
- El PowerPoint final se reaudita, incluidos valores de gráficos, antes de aprobarse.

Genera la **sección financiera** del informe mensual de Revisoría Fiscal de una propiedad horizontal
a partir del Excel de papeles de trabajo, y produce un **reporte de validación** con las
inconsistencias del Excel y, opcionalmente, de un informe PPTX ya elaborado.

Caso de prueba: Conjunto Residencial Picasso PH, Informe 006-2026, agosto de 2026.

**Nuevo en v0.4.0:** informe COMPLETO con la estructura del modelo 006-2026 (auditoría regulatoria, financiera y de control interno), `python generar_informe.py ... --completo`; estructura editable en `config/estructura_informe.yaml`; reglas D01–D10.

**v0.3.0:** el agente por lote (orquestador, normalizador, re-auditor, redactor, mejora continua, auditoría periódica y panel). Ver [AGENTE.md](AGENTE.md) y la demostración en `demo/demo.sh`.

## Qué hace

| Paso | Módulo | Resultado |
|---|---|---|
| 1. Leer | `revisoria/extractor.py` | Modelo de datos con cada cifra y su celda fuente (`modelo_AAAA-MM.json`) |
| 2. Validar Excel | `revisoria/validador.py` | 18 reglas: #REF!, totales incompletos, cifras digitadas, cruces entre hojas, conciliaciones, retención, antigüedad de saldos… |
| 3. Auditar PPTX (opcional) | `revisoria/auditor_pptx.py` | Títulos que no cuadran, cifras del mes anterior, texto oculto tras imágenes, periodos mal digitados, citas normativas por verificar |
| 4. Generar | `revisoria/generador_pptx.py` | Diapositivas con la plantilla LC Counter, tablas y gráficos **nativos** y notas del orador con las fuentes |
| 5. Reportar | `revisoria/reporte.py` | `Validacion_<copropiedad>_<periodo>.html` con filtros por severidad |

Principio: **los números nunca los calcula un LLM**. Python calcula; la redacción con Claude es
opcional (`--llm`) y tiene un candado que descarta cualquier texto que cambie una cifra.

## Instalación de esta copia

```bash
cd "/Users/airm4/Desktop/01-PROYECTOS 🚀/LCCOUNTER/08-Version-Mejorada/codigo"
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python -m app.servidor
```

## Uso

```bash
# Sección financiera + validación del Excel
python generar_informe.py --excel "EXCEL PICASSO AGOSTO 2026 .xlsx" --config config/picasso.yaml

# Además, auditar el informe que ya se había armado a mano
python generar_informe.py --excel "EXCEL PICASSO AGOSTO 2026 .xlsx" --config config/picasso.yaml \
  --revisar-pptx "Inf 006-2026 PICASSO AGOSTO 2026.pptx"

# Otro periodo del mismo Excel, y pulir redacción con Claude
export ANTHROPIC_API_KEY=...
python generar_informe.py --excel ... --config ... --periodo 2026-07 --llm
```

La aplicación usa por defecto `../espacio-trabajo`, separado del entorno anterior.

## Diapositivas que genera

1. Portada de sección «Auditoría Financiera a <mes año>»
2. Monetización del fondo de imprevistos
3. Liquidez de efectivo (tabla + evolución mensual)
4. Cartera (título calculado, composición y evolución)
5. Concentración de cartera por rangos (título calculado)
6. Mayores deudores > $1 millón (se pagina solo)
7. Anticipos 1330 y deudores varios 1380, con antigüedad
8. Opinión de la Revisoría Fiscal sobre anticipos (si hay nota)
9. Cuentas por pagar
10. Resultado del ejercicio y rubros sobreejecutados (total recalculado, presupuesto anual y del periodo)
11. Soporte de rubros sobreejecutados: los extractos del auxiliar pegados en la hoja PRESUPUESTO, por rubro
12. Excedentes acumulados y su respaldo

Las observaciones de la Revisoría Fiscal que están en celdas combinadas del Excel (una observación para
varios inmuebles) se respetan: salen en una celda fusionada y el grupo no se separa entre páginas.

## El archivo de configuración

`config/<copropiedad>.yaml` guarda lo que **no** está en el Excel: nombre, NIT, unidades,
número de informe y las **notas del revisor** (opiniones, contexto de gestión, observaciones por
rubro). El sistema no inventa opiniones: si una nota está vacía, la diapositiva sale sin ella.
Para una copropiedad nueva, copie `config/picasso.yaml` y ajústelo.

## Cómo encuentra los datos

Busca por **etiquetas**, no por coordenadas: el texto de la fila («Total equivalente al efectivo
positivo»), el título del bloque («CUENTAS POR PAGAR A AGOSTO 2026») y el mes del encabezado.
Por eso tolera que las filas se muevan, pero sí necesita que se conserven los nombres de las hojas,
las etiquetas y los títulos de bloque. Lo más importante para el siguiente mes:

- En **CXP** y **ANTIC. E INGRESOS X C**, crear el bloque nuevo con el título del mes
  («CUENTAS POR PAGAR A SEPTIEMBRE 2026», «ANTICIPOS Y AVANCES A CONTRATISTAS A SEPTIEMBRE 2026»).
- En **Rangos de Cartera**, el listado del mes con el encabezado «CARTERA A SEPTIEMBRE 2026».
- En las hojas con columnas mensuales, la nueva columna con la fecha del mes en el encabezado.

## Pruebas

```bash
python -m pytest -q tests/test_integridad_v2.py
REVISORIA_DATOS=/ruta/a/la/carpeta/del/excel python -m pytest -q tests
```

Las pruebas leen el Excel por coordenada (camino independiente del extractor) y verifican que
las cifras clave aparezcan en el PPTX, que no queden cifras del mes anterior, que las tablas sean
nativas y que el candado del LLM funcione.

## Alcance heredado

- El alcance exacto depende de la ficha de cada cliente: puede generar la sección financiera
  o el informe completo disponible en el proyecto.
- La hoja **may deudores** no tiene columnas por mes: muestra los dos saldos que traiga.
- La regla de retención compara fecha de presentación y de pago; no conoce el calendario
  tributario. Señala para verificar, no concluye.
- Las citas normativas que marca el auditor son para **verificar** por el revisor fiscal.
