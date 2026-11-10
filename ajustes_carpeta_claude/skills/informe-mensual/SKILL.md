---
name: informe-mensual
description: Opera el informe mensual de Revisoría Fiscal de UN cliente de punta a punta siguiendo el proceso TO-BE (E1 recepción → E4 generación y validación → E5 aprobación del revisor → E6 mejora), con Claude como herramienta y deteniéndose en las decisiones del revisor fiscal. Úsala cuando pidan «procesa a <cliente> de <mes>», «informe de agosto de Aldea», «corre el flujo de este cliente» o carguen el Excel de un cliente nuevo.
argument-hint: <cliente (NIT o nombre)> <AAAA-MM>
---

# Informe mensual de un cliente (proceso TO-BE)

Trabaja desde `08-Version-Mejorada/codigo` con `.venv/bin/python -m agente --raiz ../espacio-trabajo …`.
Lee primero `CLAUDE.md` (reglas). El cliente se identifica por el campo **`id`** de su ficha (los 10 dígitos del NIT
sin puntos, que es también el nombre de su carpeta): `../espacio-trabajo/clientes/<id>/ficha.yaml`. Ese `id` es el
que usan `--clientes`, `probar-receta`, etc.; el campo `nit` es solo el formato visible (900.123.456-7).

Lleva un registro breve de cada etapa para el resumen final. No avances una etapa si la anterior dejó un
bloqueo que solo resuelve una persona: repórtalo y detente.

## 0 · Retomar (memoria persistente)
`memoria contexto <id> --periodo <P>`: estado actual y del mes anterior, mapa y recetas, hallazgos abiertos,
pendientes anotados por sesiones anteriores y lecciones del flujo. Parte de ahí. Si el cliente es nuevo o su Excel
viene con otra estructura, usa primero `/incorporar-cliente`.

## E1 · Recepción y normalización (T03–T06)
1. Delega en el subagente **recepcion-ph** (cliente, periodo). Devuelve: archivos recibidos, identidad
   (copropiedad/NIT/periodo), secciones del informe estándar completas, vacías o faltantes, y la lista de
   faltantes a documentar (T06).
2. Candidatas para **estructurador-excel** (una sección por vez): solo las que `recepcion-ph` marque como
   candidatas, es decir, secciones «formato distinto» o «falta», o «vacía» **cuando hay una hoja sin usar que parece
   contenerla**. Una sección «vacía» cuya hoja se leyó bien pero no trae filas es un **faltante del cliente** (T06),
   no un caso para el estructurador. Verifica tú mismo cada receta con `probar-receta` antes de presentarla.
   Si la única forma de que pase la prueba es tomar como dato una columna que no lo es (p. ej. un campo obligatorio
   que en el Excel está vacío), **no fuerces la receta**: es un faltante del cliente para T06.
   Deja la aprobación al revisor (botón «Aprobar receta» o `aprobar-receta <id> --periodo <P> --seccion <S> --revisor "<nombre>"`).
   **Orden:** una receta solo cambia el informe después de aprobada y de volver a correr T09. Si el revisor está
   presente, pídele que la apruebe antes de T09; si no, corre T09 igual y anota que habrá que repetirlo.
3. Si hay bloqueo de identidad (Excel de otro cliente o de otro mes), **detente**: lo resuelve el auditor.

## E4 · Generación y validación (T09–T12)
4. T09: `correr --periodo <P> --clientes <id>`. Lee `salidas/<P>/<id>/estado.json`, `hallazgos.json` y
   `cobertura_esquema.json`. Cuenta los hallazgos desde **`hallazgos.json`** (la línea «Validador: N» de la bitácora
   es el conteo antes de generar). La línea «Siguiente por validar» de `correr` es la cola de todo el lote, no
   solo de este cliente. Resume: diapositivas, hallazgos ALTA y MEDIA con su celda de origen, cuadre.
5. **T10 (revisor):** presenta los hallazgos ALTA uno por uno con evidencia (celda del Excel; o el archivo, si el
   hallazgo es de la ficha, p. ej. F01 → `ficha.yaml`, que se resuelve con «Cambiar NIT…») y pregunta si los
   resultados son correctos. Si el revisor dice que no, deja listo para él
   `devolver <id> --periodo <P> --revisor "<nombre>" --motivo "<motivo en sus palabras>"` (o el botón «Devolver» de
   la app): la devolución alimenta la mejora continua (dos con el mismo motivo proponen una regla). No lo ejecutes.
   Luego vuelve a E3/T08 (lo corrige el auditor en el Excel) y repite T09.
6. T11: delega en **redactor-ph** la mejora del borrador de opinión (sin tocar cifras). El revisor da la idea.
   El redactor escribe `Borrador_propuesta_claude.md`; el revisor la revisa y guarda **su** versión final
   (p. ej. `Borrador_final_revisor.md`), aunque la acepte sin cambios.
7. T12: delega en **reauditor-ph** la re-auditoría independiente del PPTX.

## E5 · Aprobación y firma (T13–T14) — solo el revisor
8. Presenta el resumen de la re-auditoría y deja listo el comando para el revisor:
   `aprobar <id> --periodo <P> --revisor "<nombre>" [--justificacion "…"] --borrador-final <ruta del borrador final del revisor> [--pptx-final <PPTX editado>]`
   (con hallazgos ALTA exige justificación), o el botón «Aprobar…» de la app. Sin `--borrador-final` la opinión
   redactada no queda registrada. Lo que el revisor acepta tal como lo propuso Claude se registra con origen
   «claude» y no entra al aprendizaje de la biblioteca. No lo ejecutes: un hook lo bloquea y el CLI lo rechaza.

## E6 · Seguimiento y mejora (T16–T17)
9. Después de que el revisor apruebe: delega en **mejora-continua-ph** para analizar sus correcciones y proponer
   mejoras (las aprueba el revisor).

## Cierre · memoria
Anota lo que la próxima sesión necesita y no queda en el sistema: `memoria anotar <id> --periodo <P> --tipo
pendiente|particularidad|acuerdo --texto "…"`, cierra los pendientes resueltos con `memoria resolver <id> --nota
N00x --texto "…"` y las lecciones generales con `memoria anotar flujo --tipo leccion`. Sin cifras ni datos de
propietarios.

## Resumen final
Una tabla por etapa: qué se hizo, resultado, y qué queda pendiente y de quién (auditor, revisor, despacho).
