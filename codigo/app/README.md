# LCCOUNTER · aplicación web

Interfaz local del agente de Revisoría Fiscal PH, conectada a Claude.

```bash
cd 08-Version-Mejorada/codigo
.venv/bin/pip install -r requirements.txt
export ANTHROPIC_API_KEY=...            # solo para el asistente
.venv/bin/python -m app.servidor --raiz ../espacio-trabajo
# abrir http://127.0.0.1:8790
```

| Vista | Qué hace |
|---|---|
| Panel del lote | Estado de cada cliente, cola de validación y hallazgos ALTA del periodo. |
| Cargar y validar | Sube el Excel de papeles de trabajo y el balance de prueba; corre el lote (normalizar, validar, generar, redactar, re-auditar). Crea clientes. |
| Hallazgos y decisión | Cinco pestañas. **Comunicaciones:** Claude redacta un borrador de oficio al Consejo de Administración, un requerimiento a la administración o la respuesta a una comunicación recibida, usando solo las fuentes elegidas (ficha, cifras del informe, hallazgos marcados, borrador de opinión y el texto recibido), con citas (F1, F2…), tabla de trazabilidad y observaciones; no inventa normas, fechas ni plazos (deja marcadores [entre corchetes]); señala toda cifra sin respaldo en las fuentes; se edita, un revisor lo marca como revisado y se descarga en Word con el logo (con la marca BORRADOR mientras no esté revisado). La aplicación nunca envía comunicaciones. **Diapositivas:** visor del informe (miniaturas, flechas del teclado) y edición de cada texto y celda de tabla; las ediciones se guardan en una copia (`Informe_Editado_…pptx`) que se envía como informe final al aprobar, así la mejora continua captura cada corrección (si se cambia una cifra avisa y genera una alerta para revisar el cálculo); permite comparar con la versión generada, descargar o descartar. Las imágenes se generan con LibreOffice + PyMuPDF y quedan en caché. **Hallazgos y decisión:** botón «Abrir el informe en el navegador» (PDF en una pestaña nueva), filtros y búsqueda, cuadre contra balance, descargas (PPTX, validación, PDF) y aprobar o devolver (con hallazgos ALTA exige justificación). **Borrador de opinión:** editor con vista previa; el texto guardado se envía como borrador final al aprobar y sus cambios alimentan la mejora continua. Si el informe se vuelve a validar y el borrador generado cambia, la edición queda desactualizada y no deja aprobar hasta revisarla. **Ficha del cliente:** datos de la copropiedad, revisor, número de informe, tipo de informe y notas del revisor del mes (cartera, anticipos y observaciones por rubro); guarda la versión anterior en `clientes/<nit>/historial/` y permite volver a validar. |
| Mejora continua | Métricas, propuestas (aprobar), resultado por versión de la biblioteca, regresión y muestra mensual. |
| Base normativa | Leyes, decretos, NIIF y manuales que el despacho carga (PDF con texto, Word, TXT, Markdown), generales o por copropiedad. Se dividen por artículo o numeral y se indexan; búsqueda con resaltado y visor por artículo. Cargan y eliminan revisores o administradores. **Nunca se completa de memoria**: el asistente y las comunicaciones solo citan lo que está cargado, con documento y artículo. |
| ⚙ Configuración | Fuera del menú, junto al nombre del usuario (administradores; los revisores la consultan). **Proveedores de IA:** Claude por su SDK oficial y cualquier API compatible con OpenAI (OpenAI, Gemini, DeepSeek, Mistral, Groq, Ollama local u otro); claves guardadas solo en el equipo (`configuracion_ia.json`, permisos 600) y siempre enmascaradas; botón «Probar». **Qué proveedor y modelo usa cada función** (asistente, comunicaciones, automejora). **Automejora recursiva (RSI):** modo, margen mínimo, candidatos, «Ejecutar ciclo ahora», propuestas con el texto actual y el propuesto lado a lado, estrategias (nivel recursivo) y ciclos recientes. |
| Asistente | Preguntas a las bases de conocimiento: hallazgos, cifras con celda fuente, Excel, mejora continua, biblioteca y bitácora. |

## Asistente

- Modelo `claude-opus-5-5` con pensamiento adaptativo, streaming y respaldo automático ante rechazos (`fallbacks: "default"`).
- Solo **lee** el espacio de trabajo con 6 herramientas (`app/asistente.py`). No aprueba, no edita, no calcula cifras:
  las cita de los resultados con su fuente.
- Sin `ANTHROPIC_API_KEY` el resto de la aplicación funciona; el asistente muestra cómo configurarla.

## Usuarios y roles

La primera vez la aplicación pide crear el **administrador** (queda con los roles administrador y revisor fiscal).
Desde la vista «Usuarios» crea a los demás con una contraseña temporal que deben cambiar al entrar.

| Rol | Puede |
|---|---|
| Asistente | Ver todo, cargar archivos, validar, crear clientes, usar el asistente Claude |
| Revisor fiscal | Lo anterior, más editar ficha, borrador y diapositivas; aprobar o devolver; aprobar mejoras |
| Auditor (segundo revisor) | Ver, generar la muestra mensual y registrar su conclusión; nunca sobre un informe que él aprobó |
| Administrador | Crear usuarios, asignar roles, desactivar, restablecer contraseñas y ver el registro de accesos |

- Quien aprueba, devuelve o audita queda registrado con el **nombre de su sesión**, no con un nombre escrito a mano.
- Contraseñas: mínimo 10 caracteres con mayúsculas, minúsculas y números; se guardan derivadas (PBKDF2-SHA256)
  en `<espacio>/usuarios.json` (permisos 600). 5 intentos fallidos bloquean el usuario 15 minutos.
- Sesión en cookie HttpOnly/SameSite=Strict, vence tras 12 horas sin uso; las acciones solo se aceptan desde la
  propia aplicación. Reiniciar el servidor cierra las sesiones.
- Todo queda en `<espacio>/registro_accesos.jsonl` (ingresos, cargas, validaciones, aprobaciones, consultas al asistente).
- Si se pierde el acceso del administrador:

```bash
.venv/bin/python -m app.seguridad --raiz ../espacio-trabajo clave <usuario>
```

## Seguridad

Escucha solo en `127.0.0.1`. Para abrirla a la red de la oficina hace falta HTTPS (un proxy como Caddy) y
revisar el respaldo del espacio de trabajo; no la exponga a internet tal como está.
