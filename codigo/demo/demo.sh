#!/usr/bin/env bash
# Demostración completa del agente de Revisoría Fiscal PH (lote de 4 clientes, agosto 2026).
# Uso:  bash demo/demo.sh "<ruta al EXCEL PICASSO AGOSTO 2026 .xlsx>"
# Pausa entre pasos (ENTER). Para correr sin pausas:  PAUSA=0 bash demo/demo.sh ...
set -e
cd "$(dirname "$0")/.."
EXCEL="${1:?Indique la ruta del Excel de Picasso}"
W=demo/trabajo; S=$W/salidas/2026-08; A="python3 -m agente --raiz $W"
paso() { echo; echo "════ $1"; [ "${PAUSA:-1}" = "0" ] || read -r -p "      (ENTER para continuar) " _; }

paso "0. Preparar el espacio: 4 clientes (Picasso real + 3 ficticios con distinto software)"
python3 demo/preparar_demo.py --excel "$EXCEL"

paso "1. Orquestador: corre el lote completo en paralelo (normalizar → validar → generar → redactar → re-auditar)"
$A correr --periodo 2026-08

paso "2. El revisor intenta aprobar Picasso sin resolver los hallazgos ALTA → el sistema lo impide"
$A aprobar picasso --periodo 2026-08 --revisor "Revisora fiscal A" || true

paso "3. La revisora edita el informe y el borrador de opinión, y aprueba con justificación (se exporta el PDF)"
python3 demo/simular_revisor.py $S/picasso
$A aprobar picasso --periodo 2026-08 --revisor "Revisora fiscal A" \
   --justificacion "Diferencias del Excel informadas al Consejo como hallazgo" \
   --pptx-final $S/picasso/Informe_editado.pptx --borrador-final $S/picasso/Borrador_editado.md

paso "4. Arboleda: el revisor lo devuelve (total de CXP digitado, detectado por la regla V20)"
$A devolver arboleda --periodo 2026-08 --revisor "Revisor fiscal B" --motivo "Total de CXP digitado; corregir la fórmula"

paso "5. Mirador: se aprueba con una corrección de texto"
python3 demo/simular_revisor.py $S/mirador
$A aprobar mirador --periodo 2026-08 --revisor "Revisora fiscal A" --justificacion "Hallazgos trasladados al informe" \
   --pptx-final $S/mirador/Informe_editado.pptx

paso "6. Mejora continua: las correcciones se convierten en propuestas y métricas"
$A mejorar

paso "7. La revisora aprueba la propuesta de opinión → se prueba contra el golden set y se publica una versión reversible"
$A aprobar-propuesta BIB-concentracion_cartera-1 --revisor "Revisora fiscal A"

paso "8. Corrigen el Excel de Arboleda y se vuelve a correr: el borrador ya usa la redacción de la revisora"
python3 demo/corregir_excel.py $W/entradas/2026-08/arboleda/Excel_arboleda_2026-08.xlsx
$A correr --periodo 2026-08 --clientes arboleda --paralelo 1
sed -n '/Concentración/,+4p' $S/arboleda/Borrador_Opinion_arboleda_2026-08.md

$A aprobar arboleda --periodo 2026-08 --revisor "Revisor fiscal B" \
   --justificacion "Hallazgos restantes informados al Consejo"

paso "9. Auditoría periódica: regresión del golden set y muestra mensual para un segundo revisor"
$A regresion
$A muestra --periodo 2026-08 --k 2 --semilla 1

paso "10. El segundo revisor registra su conclusión → alimenta la mejora continua"
CID=$(python3 -c "import json;print(json.load(open('$W/conocimiento/auditorias/muestra_2026-08.json'))['muestra'][0]['cliente'])")
$A auditar "$CID" --periodo 2026-08 --auditor "Revisor fiscal C" --conclusion menores \
   --observacion "La nota de cartera no menciona los acuerdos de pago vigentes"

paso "11. Mejora continua de nuevo: ¿sirvió la versión de la biblioteca? y lo que encontró la auditoría"
$A mejorar

echo; echo "Listo. Abra el panel: $W/Panel_2026-08.html"
