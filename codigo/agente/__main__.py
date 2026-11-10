"""Agente de Revisoría Fiscal PH — línea de comandos.

  python -m agente --raiz <espacio> correr --periodo 2026-08 [--clientes a,b] [--paralelo 4] [--llm]
  python -m agente --raiz <espacio> estado --periodo 2026-08
  python -m agente --raiz <espacio> aprobar <cliente> --periodo 2026-08 --revisor "Nombre"
        [--justificacion "..."] [--pptx-final informe_editado.pptx] [--borrador-final borrador_editado.md]
  python -m agente --raiz <espacio> devolver <cliente> --periodo 2026-08 --revisor "Nombre" --motivo "..."
  python -m agente --raiz <espacio> mejorar                 propuestas + métricas
  python -m agente --raiz <espacio> aprobar-propuesta <id> --revisor "Nombre"
  python -m agente --raiz <espacio> revertir <version>
  python -m agente --raiz <espacio> regresion
  python -m agente --raiz <espacio> muestra --periodo 2026-08 [--k 4]
  python -m agente --raiz <espacio> auditar <cliente> --periodo 2026-08 --auditor "Nombre"
        --conclusion sin_observaciones|menores|reproceso [--observacion "..."]
  python -m agente --raiz <espacio> panel --periodo 2026-08
  python -m agente --raiz <espacio> nuevo-cliente <nit> --copropiedad "..." [--software generico]
  python -m agente --raiz <espacio> memoria contexto|anotar|resolver <cliente|flujo> [--periodo P] [--tipo T] [--texto ..] [--nota N]
  python -m agente --raiz <espacio> rsi-estado | rsi-ciclo          automejora (RSI); aprobar: rsi-aprobar <id> --revisor
  python -m agente --raiz <espacio> diagnosticar <cliente> --periodo 2026-08   (solo lectura, etapa E1)
  python -m agente --raiz <espacio> cambiar-nit <nit actual> <nit real> --revisor "..."   (NIT provisional → real)

El espacio (--raiz o variable REVISORIA_RAIZ) tiene: clientes/, entradas/, salidas/, conocimiento/.
"""
from __future__ import annotations

import argparse
import os
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from . import auditoria, mejora, orquestador, panel  # noqa: E402
from .comun import Espacio, mostrar_nit, normalizar_nit  # noqa: E402

FICHA_NUEVA = """# Ficha del cliente: datos permanentes, cómo llega su información y notas del revisor del mes.
copropiedad: {copropiedad}
nit: "{nit}"
unidades:
firma: LC Counter Contadores y Auditores SAS
numero_informe: ""
revisor: ""
orden: 99                 # orden en el calendario de visitas
software: {software}      # perfil del software contable (conocimiento/perfiles_software.yaml)
entradas:
  excel: "*.xlsx"         # Excel estándar de papeles de trabajo
  balance:
    archivo: "balance*"   # balance de prueba exportado del software contable
    perfil: {software}
notas:
  concentracion_cartera: ""
  opinion_anticipos: ""
  presupuesto: {{}}
"""


# Decisiones del revisor fiscal (CLAUDE.md): Claude las deja listas pero nunca las ejecuta
COMANDOS_DEL_REVISOR = ("aprobar", "devolver", "aprobar-receta", "aprobar-propuesta", "revertir", "cambiar-nit", "auditar",
                        "rsi-aprobar", "rsi-rechazar", "rsi-restaurar")


def main(argv=None):
    ap = argparse.ArgumentParser(prog="agente", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--raiz", default=os.environ.get("REVISORIA_RAIZ"),
                    help="espacio de trabajo (obligatorio; o la variable REVISORIA_RAIZ)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("correr"); s.add_argument("--periodo", required=True); s.add_argument("--clientes")
    s.add_argument("--paralelo", type=int, default=4); s.add_argument("--llm", action="store_true")
    s = sub.add_parser("estado"); s.add_argument("--periodo", required=True)
    s = sub.add_parser("aprobar"); s.add_argument("cliente"); s.add_argument("--periodo", required=True)
    s.add_argument("--revisor", required=True); s.add_argument("--justificacion", default="")
    s.add_argument("--pptx-final"); s.add_argument("--borrador-final")
    s = sub.add_parser("devolver"); s.add_argument("cliente"); s.add_argument("--periodo", required=True)
    s.add_argument("--revisor", required=True); s.add_argument("--motivo", required=True)
    sub.add_parser("mejorar")
    s = sub.add_parser("aprobar-propuesta"); s.add_argument("id"); s.add_argument("--revisor", required=True)
    s = sub.add_parser("revertir"); s.add_argument("version")
    sub.add_parser("regresion")
    s = sub.add_parser("muestra"); s.add_argument("--periodo", required=True); s.add_argument("--k", type=int, default=4)
    s.add_argument("--semilla", type=int)
    s = sub.add_parser("auditar"); s.add_argument("cliente"); s.add_argument("--periodo", required=True)
    s.add_argument("--auditor", required=True); s.add_argument("--observacion", default="")
    s.add_argument("--conclusion", required=True, choices=list(auditoria.CONCLUSIONES))
    s = sub.add_parser("panel"); s.add_argument("--periodo", required=True)
    s = sub.add_parser("nuevo-cliente"); s.add_argument("nit"); s.add_argument("--copropiedad", required=True)
    s.add_argument("--software", default="generico")
    s = sub.add_parser("estructurar"); s.add_argument("cliente"); s.add_argument("--periodo", required=True)
    s.add_argument("--seccion", required=True); s.add_argument("--hoja", required=True)
    s = sub.add_parser("adaptar", help="agente adaptador: propone (sin aprobar) recetas para leer el Excel de un cliente")
    s.add_argument("cliente"); s.add_argument("--periodo", required=True)
    s.add_argument("--ia", action="store_true", help="si la receta automática no basta, pedirla al proveedor de IA de la app")
    s = sub.add_parser("probar-receta"); s.add_argument("cliente"); s.add_argument("--periodo", required=True)
    s.add_argument("--receta", required=True, help="archivo JSON con la receta")
    s = sub.add_parser("aprobar-receta"); s.add_argument("cliente"); s.add_argument("--periodo", required=True)
    s.add_argument("--seccion", required=True); s.add_argument("--revisor", required=True)
    s = sub.add_parser("plantilla"); s.add_argument("--periodo", required=True); s.add_argument("--salida", required=True)
    s = sub.add_parser("migrar"); s.add_argument("cliente"); s.add_argument("--periodo", required=True)
    s.add_argument("--receta", action="append", default=[], help="receta propuesta (JSON) aún no aprobada; se puede repetir")
    s = sub.add_parser("diagnosticar", help="solo lectura: identidad, cobertura y formato del Excel de un cliente")
    s.add_argument("cliente"); s.add_argument("--periodo", required=True)
    s = sub.add_parser("memoria", help="memoria persistente de Claude: contexto | anotar | resolver")
    s.add_argument("accion", choices=["contexto", "anotar", "resolver"]); s.add_argument("cliente", help="id del cliente o «flujo»")
    s.add_argument("--periodo", default=""); s.add_argument("--tipo", default="pendiente")
    s.add_argument("--texto", default=""); s.add_argument("--nota", help="id de la nota que se resuelve")
    sub.add_parser("rsi-estado", help="automejora (RSI): modo, estrategias, casos y propuestas pendientes")
    sub.add_parser("rsi-ciclo", help="un ciclo de automejora; en modo «proponer» solo deja propuestas")
    s = sub.add_parser("rsi-aprobar"); s.add_argument("id"); s.add_argument("--revisor", required=True)
    s = sub.add_parser("rsi-rechazar"); s.add_argument("id"); s.add_argument("--revisor", required=True)
    s = sub.add_parser("rsi-restaurar"); s.add_argument("estrategia")
    s = sub.add_parser("cambiar-nit"); s.add_argument("cliente"); s.add_argument("nit")
    s.add_argument("--revisor", required=True)
    a = ap.parse_args(argv)
    if not a.raiz:
        ap.error("indique el espacio de trabajo con --raiz (p. ej. --raiz ../espacio-trabajo) o REVISORIA_RAIZ")
    if a.cmd not in ("nuevo-cliente", "plantilla") and not Path(a.raiz).is_dir():
        ap.error(f"el espacio de trabajo «{a.raiz}» no existe")
    if a.cmd in COMANDOS_DEL_REVISOR and os.environ.get("CLAUDECODE"):
        raise SystemExit(f"«{a.cmd}» es una decisión del revisor fiscal y no se ejecuta desde Claude. "
                         "Deje el comando listo para que el revisor lo corra en su terminal o use la app.")
    esp = Espacio(Path(a.raiz))
    periodo = getattr(a, "periodo", None)

    if a.cmd == "adaptar":
        from . import adaptador
        try:
            r = adaptador.adaptar(esp, a.cliente, a.periodo, usar_ia=a.ia, log=print)
        except ValueError as e:
            raise SystemExit(str(e))
        R = r["resumen"]
        print(f"{a.cliente} {a.periodo}: {R['completas']} de {R['secciones']} secciones completas hoy; "
              f"{R['completas_con_propuestas']} con las {R['propuestas']} receta(s)"
              + (f" y {R.get('ajustes_mapa')} corrección(es) del mapa" if R.get("ajustes_mapa") else "") + " propuesta(s).")
        for clave, aj in (r.get("ajustes_mapa") or {}).items():
            print(f"  ↻ mapa «{clave}»: {aj['antes'] or '(sin hoja)'} → «{aj['ahora']}» ({aj['motivo']})")
        for f in r["secciones"]:
            if f["estado"] == "ok" and not f["propuesta"]:
                continue
            if f["propuesta"]:
                pr = f["prueba"] or {}
                print(f"  ✔ {f['seccion']}: receta {f['origen']} en «{f['propuesta']['hoja']}» · {pr.get('n')} fila(s)"
                      + (f" · avisos: {'; '.join(pr.get('avisos') or [])}" if pr.get("avisos") else ""))
            elif f.get("ajuste_mapa") and f["estado_con_propuestas"] == "ok":
                print(f"  ✔ {f['seccion']}: se lee con la corrección del mapa («{f['ajuste_mapa']['clave']}» → «{f['ajuste_mapa']['ahora']}»)")
            else:
                print(f"  · {f['seccion']} ({f['estado']}): {f['motivo'] or f['detalle']}")
        print("Propuestas en salidas/" + a.periodo + "/" + a.cliente + "/adaptacion.json. Las aprueba un revisor en la app "
              "(Mapa del Excel → Adaptación al Excel).")
        return
    if a.cmd in ("estructurar", "probar-receta", "aprobar-receta"):
        import json as _json
        from revisoria import lectores
        from revisoria.excel import Libro, Periodo
        from . import estructurador, normalizador
        p = Periodo.desde_texto(a.periodo)
        excel = normalizador.ubicar_excel(esp.entrada(a.periodo, a.cliente), esp.ficha(a.cliente))
        if not excel:
            raise SystemExit(f"No hay Excel cargado para {a.cliente} en {a.periodo}.")
        sal = esp.salida(a.periodo, a.cliente)
        if a.cmd == "estructurar":
            from app.proveedores import ErrorIA
            try:
                r = estructurador.proponer(esp.raiz, Libro(str(excel)), a.seccion, a.hoja, p, esp.conocimiento, log=print)
            except (ErrorIA, ValueError) as e:
                raise SystemExit(f"El estructurador no pudo proponer: {e}")
            sal.mkdir(parents=True, exist_ok=True)
            destino = sal / f"receta_propuesta_{a.seccion}.json"
            destino.write_text(_json.dumps(r, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
            if not r["corresponde"]:
                print(f"La hoja no contiene «{a.seccion}»: {r['motivo']}")
            else:
                print(_json.dumps({"receta": r["receta"], "ok": r["prueba"]["ok"], "errores": r["prueba"]["errores"],
                                   "avisos": r["prueba"]["avisos"], "filas_leidas": r["prueba"]["n"],
                                   "muestra": r["prueba"]["muestra"][:3]}, ensure_ascii=False, indent=1, default=str))
                print(f"Propuesta guardada en {destino}. La aprueba un revisor con «aprobar-receta».")
        elif a.cmd == "probar-receta":
            receta = _json.loads(Path(a.receta).read_text(encoding="utf-8"))
            prueba = lectores.probar(Libro(str(excel)), receta, p)
            print(_json.dumps(prueba, ensure_ascii=False, indent=1, default=str))
            if prueba["ok"] and receta.get("seccion"):
                # queda como propuesta para que el revisor la apruebe con «aprobar-receta» (antes solo «estructurar»,
                # que usa el proveedor de IA de la app, dejaba la propuesta y la receta de Claude no se podía aprobar)
                sal.mkdir(parents=True, exist_ok=True)
                destino = sal / f"receta_propuesta_{receta['seccion']}.json"
                destino.write_text(_json.dumps({"corresponde": True, "receta": receta, "prueba": prueba, "origen": "probar-receta"},
                                               ensure_ascii=False, indent=1, default=str), encoding="utf-8")
                print(f"Propuesta guardada en {destino}. La aprueba un revisor con «aprobar-receta».")
        else:
            r = _json.loads((sal / f"receta_propuesta_{a.seccion}.json").read_text(encoding="utf-8"))
            if not r.get("corresponde") or not r.get("receta"):
                raise SystemExit("La propuesta no tiene receta para aprobar.")
            try:
                estructurador.aprobar(esp, a.cliente, r["receta"], p, excel, a.revisor)
            except ValueError as e:
                raise SystemExit(f"No se aprobó: {e}")
            print(f"Receta de «{a.seccion}» aprobada por {a.revisor}. Vuelva a validar el cliente.")
        return
    if a.cmd == "memoria":
        import json as _json
        from . import memoria
        try:
            if a.accion == "contexto":
                if not a.periodo:
                    raise ValueError("indique --periodo AAAA-MM")
                r = memoria.contexto(esp, a.cliente, a.periodo)
            elif a.accion == "anotar":
                r = memoria.anotar(esp, a.cliente, a.tipo, a.texto, a.periodo)
            else:
                r = memoria.resolver(esp, a.cliente, a.nota or "", a.texto)
        except ValueError as e:
            raise SystemExit(str(e))
        print(_json.dumps(r, ensure_ascii=False, indent=1, default=str))
        return
    if a.cmd.startswith("rsi-"):
        import json as _json
        from app import rsi
        if a.cmd == "rsi-estado":
            e = rsi.estado(esp)
            print(_json.dumps({"config": rsi.config(esp), "vigente": e["vigente"], "retadora": e.get("retadora"),
                               "casos": {k: len(v) for k, v in rsi.casos(esp).items()},
                               "pendientes": [p for p in rsi.propuestas(esp) if p["estado"] == "pendiente"],
                               "ultimos_ciclos": [{"fecha": c["fecha"], "acciones": [p["accion"] for p in c["plantillas"]],
                                                   "meta": c.get("meta")} for c in rsi.ciclos(esp, 3)]},
                              ensure_ascii=False, indent=1, default=str))
        elif a.cmd == "rsi-ciclo":
            if os.environ.get("CLAUDECODE") and rsi.config(esp)["modo"] == "automatico":
                raise SystemExit("La automejora está en modo «automático» y publicaría sin el revisor. Desde Claude solo se "
                                 "corre en modo «proponer» (Configuración de la app).")
            from app.proveedores import ErrorIA
            try:
                r = rsi.ciclo(esp, "agente (CLI)")
            except (ValueError, ErrorIA) as e:
                raise SystemExit(f"No se corrió la automejora: {e}")
            for p in r["plantillas"]:
                print(f"  {p['clave']:<28} ganancia {p['ganancia']:+.3f} · {p['accion']}")
            if r.get("meta"):
                print("  Estrategia:", r["meta"].get("accion", ""))
        elif a.cmd == "rsi-aprobar":
            print(f"Plantilla publicada en {rsi.aprobar(esp, a.id, a.revisor)}.")
        elif a.cmd == "rsi-rechazar":
            rsi.rechazar(esp, a.id, a.revisor)
            print(f"Propuesta {a.id} rechazada.")
        else:
            rsi.restaurar_estrategia(esp, a.estrategia)
            print(f"Estrategia {a.estrategia} restaurada como vigente.")
        return
    if a.cmd == "diagnosticar":
        import json as _json
        from . import diagnostico
        print(_json.dumps(diagnostico.diagnosticar(esp, a.cliente, a.periodo), ensure_ascii=False, indent=1, default=str))
        return
    if a.cmd == "plantilla":
        from revisoria import libreoffice, plantilla
        from revisoria.excel import Periodo
        salida = Path(a.salida)
        bruto = salida.with_name(f"_{salida.name}")
        plantilla.plantilla_vacia(Periodo.desde_texto(a.periodo), bruto)
        out = libreoffice.convertir(bruto, "xlsx", salida.parent / "_lo")
        if out:
            out.replace(salida)
            (salida.parent / "_lo").rmdir()
            bruto.unlink()
        else:
            bruto.replace(salida)
        print(f"Plantilla estándar v{plantilla.VERSION} ({a.periodo[:4]}): {salida}")
        return
    if a.cmd == "migrar":
        import json as _json
        from . import migrador
        extra = {}
        for f in a.receta:
            r = _json.loads(Path(f).read_text(encoding="utf-8"))
            extra[r["seccion"]] = r
        r = migrador.migrar(esp, a.cliente, a.periodo, extra or None)
        print(("ACEPTABLE" if r["aceptable"] else "CON DIFERENCIAS") + f" · {r['carpeta']}")
        print((Path(r["carpeta"]) / "reporte_migracion.md").read_text(encoding="utf-8"))
        return
    if a.cmd == "cambiar-nit":
        from . import integridad
        try:
            nuevo = integridad.cambiar_nit(esp, a.cliente, a.nit, a.revisor)
        except ValueError as e:
            raise SystemExit(f"No se cambió el NIT: {e}")
        print(f"NIT cambiado: el cliente ahora es {nuevo}. Vuelva a validar sus expedientes.")
        return
    if a.cmd == "correr":
        clientes = a.clientes.split(",") if a.clientes else None
        orquestador.correr(esp, a.periodo, clientes, a.paralelo, a.llm)
        cola = orquestador.cola(esp, a.periodo)
        if cola:
            print(f"\nSiguiente por validar: {cola[0].get('copropiedad')} ({cola[0]['cliente']})")
    elif a.cmd == "estado":
        for cid in esp.ids_clientes():
            orquestador._log_cliente(print, esp.estado(a.periodo, cid) | {"cliente": cid})
    elif a.cmd == "aprobar":
        est = orquestador.aprobar(esp, a.periodo, a.cliente, a.revisor, a.justificacion, a.pptx_final, a.borrador_final)
        print(est["bitacora"][-1]["detalle"])
        cola = orquestador.cola(esp, a.periodo)
        print(f"Siguiente por validar: {cola[0].get('copropiedad')} ({cola[0]['cliente']})" if cola else "Cola vacía.")
    elif a.cmd == "devolver":
        print(orquestador.devolver(esp, a.periodo, a.cliente, a.revisor, a.motivo)["bitacora"][-1]["detalle"])
    elif a.cmd == "mejorar":
        props = mejora.proponer(esp)
        m = mejora.metricas(esp)
        print(f"{sum(1 for p in props if p['estado'] == 'pendiente')} propuesta(s) pendientes:")
        for p in props:
            if p["estado"] == "pendiente":
                print(f"  {p['id']:<26} {p['titulo']}")
        print("Métricas:", {k: v for k, v in m.items() if k not in ("fecha", "por_version", "auditoria_mensual")})
        for v in m["por_version"]:
            print(f"  Biblioteca {v['version']:<8} {v['informes']} informe(s) · {v['correcciones_por_informe']} correcciones "
                  f"y {v['opiniones_editadas_por_informe']} opiniones editadas por informe")
        if m["auditoria_mensual"]:
            print("  Auditoría mensual:", m["auditoria_mensual"])
    elif a.cmd == "aprobar-propuesta":
        r = mejora.aprobar_propuesta(esp, a.id, a.revisor)
        print(r["detalle"])
    elif a.cmd == "revertir":
        print(mejora.revertir(esp, a.version))
    elif a.cmd == "regresion":
        r = mejora.regresion(esp)
        print(f"Regresión {'OK' if r['ok'] else 'FALLÓ'}: {r['n']} caso(s)")
        for c in r["casos"]:
            print(f"  {c['caso']:<28} {'OK' if c['ok'] else 'FALLA'} {c.get('n_diferencias', '')} {c.get('error', '')}")
    elif a.cmd == "muestra":
        print(auditoria.muestra_mensual(esp, a.periodo, a.k, a.semilla))
    elif a.cmd == "auditar":
        f = auditoria.registrar_conclusion(esp, a.periodo, a.cliente, a.auditor, a.conclusion, a.observacion)
        print(f"Auditoría registrada: {f['copropiedad'] or a.cliente} · {auditoria.CONCLUSIONES[a.conclusion]} · {a.auditor}")
    elif a.cmd == "nuevo-cliente":
        try:
            cid = normalizar_nit(a.nit)
        except ValueError as e:
            ap.error(str(e))
        d = esp.clientes / cid
        if (d / "ficha.yaml").exists():
            ap.error(f"Ya existe el cliente con NIT {mostrar_nit(cid)}")
        d.mkdir(parents=True, exist_ok=True)
        (d / "ficha.yaml").write_text(FICHA_NUEVA.format(copropiedad=a.copropiedad, nit=mostrar_nit(cid), software=a.software), encoding="utf-8")
        print(f"Ficha creada: {d / 'ficha.yaml'}")

    # el panel se refresca después de cada acción con periodo
    if periodo:
        print(f"Panel: {panel.escribir(esp, periodo)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
