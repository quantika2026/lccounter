# LCCOUNTER 08 — instrucciones para Claude

Claude (Claude Code y los subagentes de `.claude/agents/`) es la herramienta con la que se configura, opera y
desarrolla este flujo. No se usa DeepSeek para estructurar ni validar.

## El flujo (proceso TO-BE, `../../02-Levantamiento-proceso/Plantilla_Proceso_Futuro_TO-BE.xlsx`)

| Etapa | Actividades | Quién | Herramienta de Claude |
|---|---|---|---|
| E1 Recepción y normalización | T03 carpeta del cliente · T04 normalizar · T05 ¿completa? · T06 documentar faltantes | Kit + auditor | subagente `recepcion-ph`; `estructurador-excel` si el Excel trae otra estructura |
| E2 Visita y auditoría documental | T07 | Auditor (presencial) | — |
| E3 Herramientas de trabajo | T08 completar el Excel estándar | Auditor | — |
| E4 Generación y validación | T09 generar y validar · **T10 ¿resultados correctos? (revisor)** · T11 redactar opinión · T12 re-auditar | Kit + revisor | `python -m agente correr`; `redactor-ph`; `reauditor-ph` |
| E5 Revisión, firma y presentación | **T13 ¿aprobado? · T14 firmar (revisor)** · T15 presentar | Revisor | — |
| E6 Seguimiento y mejora | T16 hallazgos abiertos · T17 incorporar correcciones | Kit + revisor | `mejora-continua-ph` |

### Orquestación: habilidades y subagentes
Los subagentes no pueden lanzar otros subagentes; por eso el **orquestador es la sesión principal** siguiendo una
habilidad, que reparte cada etapa al subagente que corresponde.

| Habilidad | Para qué | Subagentes que usa |
|---|---|---|
| `/lote-mensual <P>` | Orquestador del mes: todos los clientes, cola del revisor | todos, vía las otras habilidades |
| `/informe-mensual <cliente> <P>` | Un cliente de punta a punta (E1→E6) | recepcion-ph, estructurador-excel, redactor-ph, reauditor-ph, mejora-continua-ph |
| `/incorporar-cliente <NIT> "<nombre>" <P>` | Cliente nuevo y Excel con otra estructura | recepcion-ph, estructurador-excel, migrador-ph |
| `/mejora-continua [P]` | E6: correcciones, RSI, regresión, auditoría mensual | mejora-continua-ph |

### Memoria
- **Del sistema** (`../espacio-trabajo/conocimiento/`, la escribe Python, solo cambia con aprobación del revisor):
  correcciones, biblioteca de opiniones versionada, golden set, sinónimos de hojas, recetas aprobadas, RSI.
- **De Claude entre sesiones** (`../espacio-trabajo/memoria/`, comando `memoria`): al empezar
  `memoria contexto <id> --periodo P`; al terminar `memoria anotar` (pendientes, particularidades del Excel, acuerdos)
  y `memoria anotar flujo --tipo leccion`. Sin cifras ni datos de propietarios.

### Automejora (RSI)
`app/rsi.py`, acotada: mejora solo las plantillas de opinión y su propia estrategia de mejora, medida contra un caso
reservado del revisor y el golden set. Desde Claude solo en modo «proponer» (`rsi-estado`, `rsi-ciclo`); publicar
(`rsi-aprobar`) es del revisor.

**Incorporación de clientes (opción C, una vez por cliente):** cada cliente se migra a la plantilla estándar
(`revisoria/plantilla.py`, versión 1 = estructura del informe modelo) con el subagente `migrador-ph`
(`migrar <id> --periodo P [--receta r.json]`). La migración es una propuesta en `clientes/<id>/migracion/…` con su
reporte de verificación; la aprueba el revisor y el auditor trabaja sobre ella desde el mes siguiente.
Plantilla vacía para los auditores: `plantilla --periodo P --salida archivo.xlsx`.

## Reglas que no se negocian
- Python calcula; Claude organiza, explica y redacta. Ninguna cifra sale de Claude: todas vienen del Excel con su
  celda de origen.
- Las decisiones T10 y T13, la aprobación de mapas y recetas, y la firma son del **revisor fiscal**. Claude nunca
  corre `aprobar`, `aprobar-receta`, `aprobar-propuesta`, `revertir` ni `cambiar-nit` por su cuenta, ni aprueba en
  la app. Propone y deja el comando o el botón listo para el revisor. Esto está forzado: el hook
  `.claude/hooks/proteger_decisiones_revisor.py` (en `.claude/settings.json`) bloquea esos comandos, las llamadas a
  la app que aprueban y la edición de `estado.json`, `conocimiento/`, `mapa.json` y `ficha.yaml`; además el CLI los
  rechaza cuando corre dentro de Claude Code.
- No se escribe texto legal o normativo de memoria: solo lo que esté cargado en la base normativa.
- Si falta información el proceso no se detiene: se documenta (T06) y queda como hallazgo o nota.
- Datos de propietarios (nombres, apartamentos, deudas): no se envían a servicios externos fuera de Claude.
  Pregunta abierta del levantamiento (n.º 26): confirmar con el despacho qué puede salir a la nube.

## Comandos (desde esta carpeta)
```bash
.venv/bin/python -m agente --raiz ../espacio-trabajo <comando> ...
```
`correr --periodo P [--clientes C]` · `estado --periodo P` · `diagnosticar C --periodo P` · `estructurar C --periodo P --seccion S --hoja H` ·
`probar-receta C --periodo P --receta archivo.json` · `migrar C --periodo P [--receta r.json]` ·
`plantilla --periodo P --salida f.xlsx` · `mejorar` · `regresion` · `panel --periodo P` · `memoria …` · `rsi-estado` · `rsi-ciclo`.
Del revisor (no los corre Claude): `aprobar`, `devolver`, `aprobar-receta`, `aprobar-propuesta`, `revertir`, `cambiar-nit`, `auditar`, `rsi-aprobar`, `rsi-rechazar`, `rsi-restaurar`.
`--raiz` es obligatorio (o `REVISORIA_RAIZ`); ya no hay espacio `trabajo` por defecto.

Pruebas: `REVISORIA_DATOS=../../05-Datos-prueba-agosto-2026 PYTHONPATH=$PWD .venv/bin/python -m pytest -q tests`.

## Dónde está cada cosa
- Lectura del Excel: `revisoria/extractor.py`, `revisoria/documental.py`; recetas para otras estructuras:
  `revisoria/lectores.py` (ejemplos de entrenamiento en `config/lectores_modelo.json`).
- Mapa del cliente (qué hoja es cada sección): `revisoria/mapa.py`; esquema estándar: `revisoria/esquema.py`.
- Orquestador del mes: `agente/orquestador.py`; integridad (identidad, revisiones): `agente/integridad.py`.
- Espacio de trabajo: `../espacio-trabajo/` (`clientes/<NIT>/`, `entradas/<P>/<NIT>/`, `salidas/<P>/<NIT>/`).
