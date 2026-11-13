"use strict";
// LCCOUNTER · interfaz del agente de Revisoría Fiscal PH. Todo dato viene de la API local (app/servidor.py).

const $ = (s, r = document) => r.querySelector(s);
const esc = (t) => String(t ?? "").replace(/[&<>"']/g, (c) => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c]));
const pesos = (v) => v == null ? "—" : "$" + Math.round(v).toLocaleString("es-CO");
const SEVS = ["ALTA", "MEDIA", "BAJA", "INFO"];
const ESTADOS = {PENDIENTE_VALIDACION: "Pendiente de validación", APROBADO: "Aprobado", DEVUELTO: "Devuelto",
                 BLOQUEADO: "Bloqueado", RECIBIDO: "Recibido"};
const st = {resumen: null, lote: null, cliente: null, chat: [], base: "todo", filtro: {texto: "", estado: "", revisor: ""}};

async function api(ruta, opciones = {}) {
  const r = await fetch(ruta, opciones);
  if (r.status === 401 && !ruta.startsWith("/api/login")) { mostrarAcceso("login", "Su sesión venció. Ingrese de nuevo."); throw new Error("Sesión vencida."); }
  if (!r.ok) {
    let m = "";
    try { m = (await r.json()).detail || ""; } catch (_) {}
    if (!m) m = [502, 503, 504].includes(r.status)
      ? `El servidor no respondió (error ${r.status}). Es probable que se haya reiniciado en medio del proceso, por ejemplo por falta de memoria. Espere un minuto y vuelva a intentarlo; si se repite, revise los registros del servidor.`
      : `Error ${r.status}${r.statusText ? " " + r.statusText : ""} del servidor.`;
    throw new Error(typeof m === "string" ? m : JSON.stringify(m));
  }
  return r.json();
}
const post = (ruta, cuerpo) => api(ruta, {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(cuerpo)});

function aviso(t, err = false) {
  const a = $("#aviso");
  $("#aviso-txt").textContent = t; $("#aviso-ico").textContent = err ? "!" : "✓";
  a.className = "aviso" + (err ? " err" : ""); a.hidden = false;
  clearTimeout(aviso.t); aviso.t = setTimeout(() => (a.hidden = true), err ? 9000 : 5000);
}
$("#aviso-x").onclick = () => { clearTimeout(aviso.t); $("#aviso").hidden = true; };
// Indicador de carga para procesos largos (validar el lote o un cliente)
async function conCarga(texto, fn) {
  $("#cargando-txt").textContent = texto; $("#cargando").hidden = false;
  try { return await fn(); } finally { $("#cargando").hidden = true; }
}
async function ocupado(btn, fn) {
  btn = btn || {};                                   // envío sin botón (p. ej. requestSubmit): no hay nada que deshabilitar
  const txt = btn.textContent; btn.disabled = true; btn.textContent = "Procesando…";
  try { return await fn(); } catch (e) { aviso(e.message, true); } finally { btn.disabled = false; btn.textContent = txt; }
}
const periodo = () => $("#periodo").value;
const nombreCliente = (cid) => st.resumen?.nombres?.[cid] || cid;
const opcionCliente = (cid, sel) => `<option value="${esc(cid)}" ${cid === sel ? "selected" : ""}>${esc(nombreCliente(cid))}</option>`;
const chip = (e) => `<span class="chip e-${esc(e)}">${esc(ESTADOS[e] || e || "Sin datos")}</span>`;
const sev = (c) => c ? ["ALTA", "MEDIA"].map((s) => `<span class="sev ${s}">${c[s] || 0}</span>`).join(" / ") : "—";

// Markdown mínimo y seguro (escapa primero)
function md(texto) {
  const lineas = esc(texto).split("\n"); const out = []; let tabla = [], lista = null;
  const inline = (t) => t.replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").replace(/`([^`]+)`/g, "<code>$1</code>").replace(/(^|[^*])\*([^*]+)\*/g, "$1<i>$2</i>");
  const cerrar = () => {
    if (tabla.length) {
      const filas = tabla.filter((f) => !/^\|?\s*:?-{2,}/.test(f));
      out.push("<table>" + filas.map((f, i) => "<tr>" + f.replace(/^\||\|$/g, "").split("|")
        .map((c) => i ? `<td>${inline(c.trim())}</td>` : `<th>${inline(c.trim())}</th>`).join("") + "</tr>").join("") + "</table>");
      tabla = [];
    }
    if (lista) { out.push(`</${lista}>`); lista = null; }
  };
  for (const l of lineas) {
    if (/^\s*\|/.test(l)) { if (lista) { out.push(`</${lista}>`); lista = null; } tabla.push(l.trim()); continue; }
    if (tabla.length) cerrar();
    let m;
    if ((m = l.match(/^(#{1,4})\s+(.*)/))) { cerrar(); out.push(`<h3>${inline(m[2])}</h3>`); }
    else if ((m = l.match(/^\s*[-*]\s+(.*)/))) { if (lista !== "ul") { cerrar(); out.push("<ul>"); lista = "ul"; } out.push(`<li>${inline(m[1])}</li>`); }
    else if ((m = l.match(/^\s*\d+\.\s+(.*)/))) { if (lista !== "ol") { cerrar(); out.push("<ol>"); lista = "ol"; } out.push(`<li>${inline(m[1])}</li>`); }
    else if ((m = l.match(/^&gt;\s?(.*)/))) { cerrar(); out.push(`<blockquote>${inline(m[1])}</blockquote>`); }
    else if (!l.trim()) cerrar();
    else { cerrar(); out.push(`<p>${inline(l)}</p>`); }
  }
  cerrar();
  return out.join("");
}

// ------------------------------------------------------------------ navegación
function mostrar(vista) {
  if (!st.resumen) return;                                // aún cargando o sin sesión: iniciar() la mostrará al terminar
  if (vista !== "asistente") detenerVideoAuditin();      // el video (y su sonido) no sigue en otras pantallas
  if (vista === "asistente" && st.vista && st.vista !== "asistente") st.vistaPrevia = st.vistaHash;   // para «Cerrar»
  st.vista = vista; st.vistaHash = location.hash.slice(1) || vista;
  document.querySelectorAll(".vista").forEach((v) => v.classList.toggle("activa", v.id === "v-" + vista));
  document.querySelectorAll("nav a").forEach((a) => a.classList.toggle("activa", a.dataset.vista === vista));
  ({panel: cargarPanel, cargar: prepararCarga, cliente: cargarCliente, mejora: cargarMejora, asistente: prepararAsistente, usuarios: cargarUsuarios, tiempos: cargarTiempos, conceptos: cargarConceptos, normativa: cargarNormativa, configuracion: cargarConfiguracion}[vista] || cargarPanel)();
}
window.addEventListener("hashchange", () => mostrar(location.hash.slice(1).split("/")[0] || "panel"));

// ------------------------------------------------------------------ panel
async function cargarPanel() {
  if (!periodo()) return;
  const L = st.lote = await api(`/api/lote/${periodo()}`);
  const n = (e) => L.clientes.filter((c) => c.estado === e).length;
  const altas = (c) => (c.resumen?.hallazgos?.ALTA || 0) + (c.resumen?.reauditoria?.ALTA || 0);
  $("#panel-sub").textContent = `${L.clientes.length} cliente(s) · ${textoPeriodo(periodo())}`;
  // Cada tarjeta filtra la tabla de clientes con un clic (y se quita con otro)
  $("#kpis").innerHTML = [["Clientes", L.clientes.length, "", "k-total"], ["Pendientes", n("PENDIENTE_VALIDACION"), "PENDIENTE_VALIDACION", "k-pend"],
    ["Aprobados", n("APROBADO"), "APROBADO", "k-ok"], ["Devueltos", n("DEVUELTO"), "DEVUELTO", "k-dev"], ["Bloqueados", n("BLOQUEADO"), "BLOQUEADO", "k-bloq"],
    ["Hallazgos ALTA", L.clientes.reduce((s, c) => s + altas(c), 0), "ALTA", "k-alta"]]
    .map(([t, v, f, cls]) => `<button type="button" class="kpi ${cls} ${v ? "" : "cero"} ${st.filtro.estado === f && f ? "activo" : ""}" data-f="${f}"
      title="${f ? "Filtrar la tabla: " + t : "Ver todos los clientes"}"><b>${v}</b><small>${t}</small></button>`).join("");
  document.querySelectorAll("#kpis .kpi").forEach((k) => k.onclick = () => {
    st.filtro.estado = k.dataset.f && st.filtro.estado !== k.dataset.f ? k.dataset.f : "";
    if (!k.dataset.f) Object.assign(st.filtro, {texto: "", revisor: ""});
    cargarPanel();
  });
  $("#cola").innerHTML = L.cola.length ? L.cola.map((cid) => {
    const c = L.clientes.find((x) => x.cliente === cid);
    return `<li><a href="#cliente/${esc(cid)}">${esc(c?.copropiedad || cid)}</a></li>`;
  }).join("") : `<li class="vacio">No hay informes esperando validación.</li>`;
  const revisores = [...new Set(L.clientes.map((c) => c.revisor).filter(Boolean))].sort();
  $("#f-estado").innerHTML = `<option value="">Todos los estados</option><option value="ALTA">Con hallazgos ALTA</option>` +
    Object.entries(ESTADOS).map(([k, v]) => `<option value="${k}">${esc(v)}</option>`).join("") + `<option value="SIN_DATOS">Sin procesar</option>`;
  $("#f-revisor").innerHTML = `<option value="">Todos los revisores</option>` + revisores.map((r) => `<option>${esc(r)}</option>`).join("") +
    `<option value="SIN_REVISOR">Sin revisor asignado</option>`;
  $("#f-buscar").value = st.filtro.texto; $("#f-estado").value = st.filtro.estado; $("#f-revisor").value = st.filtro.revisor;
  pintarTablaClientes();
}
function pintarTablaClientes() {
  const L = st.lote, F = st.filtro, q = F.texto.trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const sinTilde = (t) => String(t || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const filas = L.clientes.filter((c) =>
    (!q || sinTilde(c.copropiedad).includes(q) || String(c.nit || c.cliente).replace(/\D/g, "").includes(q.replace(/\D/g, "") || "\u0000"))
    && (!F.estado || (F.estado === "ALTA" ? (c.resumen?.hallazgos?.ALTA || 0) + (c.resumen?.reauditoria?.ALTA || 0) > 0
                     : F.estado === "SIN_DATOS" ? !c.estado : c.estado === F.estado))
    && (!F.revisor || (F.revisor === "SIN_REVISOR" ? !c.revisor : c.revisor === F.revisor)));
  const filtrando = !!(F.texto || F.estado || F.revisor);
  $("#f-limpiar").hidden = !filtrando;
  $("#filtro-cuenta").textContent = filtrando ? `${filas.length} de ${L.clientes.length}` : "";
  const vacio = !L.clientes.length
    ? `<div class="vacio-caja"><b>Aún no hay clientes.</b><span>Cree el primero con su NIT y el nombre de la copropiedad.</span><a class="btn" href="#cargar">Crear el primer cliente</a></div>`
    : `<div class="vacio-caja"><b>Ningún cliente coincide con los filtros.</b><button type="button" class="btn sec" id="f-limpiar2">Quitar filtros</button></div>`;
  $("#tabla-clientes").innerHTML = `<tr><th>Copropiedad</th><th>Estado</th><th>Cartera</th><th>Resultado</th><th>Excel A/M</th><th>Re-auditor A/M</th><th>Último evento</th></tr>` +
    (filas.map((c) => `<tr class="clic" data-cid="${esc(c.cliente)}"><td><b>${esc(c.copropiedad)}</b><br><small class="nota">NIT ${esc(c.nit || c.cliente)} · ${esc(c.software)}${c.revisor ? " · " + esc(c.revisor) : ""}</small></td>
      <td>${chip(c.estado)}</td><td class="n">${pesos(c.resumen?.cifras?.cartera)}</td><td class="n">${pesos(c.resumen?.cifras?.resultado)}</td>
      <td>${sev(c.resumen?.hallazgos)}</td><td>${sev(c.resumen?.reauditoria)}</td><td><small class="nota">${esc((c.ultimo?.detalle || "Sin procesar en este periodo").slice(0, 110))}</small></td></tr>`).join("")
      || `<tr><td colspan="7">${vacio}</td></tr>`);
  document.querySelectorAll("#tabla-clientes tr.clic").forEach((tr) => tr.onclick = () => (location.hash = "cliente/" + tr.dataset.cid));
  if ($("#f-limpiar2")) $("#f-limpiar2").onclick = limpiarFiltros;
}
function limpiarFiltros() { Object.assign(st.filtro, {texto: "", estado: "", revisor: ""}); cargarPanel(); }
$("#f-buscar").oninput = () => { st.filtro.texto = $("#f-buscar").value; pintarTablaClientes(); };
$("#f-estado").onchange = () => { st.filtro.estado = $("#f-estado").value; cargarPanel(); };
$("#f-revisor").onchange = () => { st.filtro.revisor = $("#f-revisor").value; pintarTablaClientes(); };
$("#f-limpiar").onclick = limpiarFiltros;
$("#btn-validar-todo").onclick = (ev) => ocupado(ev.target, async () => {
  const r = await conCarga(`Validando todo el lote de ${textoPeriodo(periodo())}…`, () => post("/api/validar", {periodo: periodo()}));
  aviso(r.log.filter((l) => l.startsWith("  ·")).length + " cliente(s) procesados.");
  cargarPanel();
});

// ------------------------------------------------------------------ cargar
function prepararCarga() {
  const previo = $("#c-cliente").value;
  $("#c-cliente").innerHTML = st.resumen.clientes.map((c) => opcionCliente(c, previo)).join("") || `<option value="">Cree primero un cliente</option>`;
}
for (const tipo of ["excel", "balance"]) {
  const zona = $("#zona-" + tipo), input = $("#f-" + tipo);
  zona.onclick = () => input.click();
  input.onchange = () => ($("#n-" + tipo).textContent = input.files[0]?.name || "");
  zona.ondragover = (e) => { e.preventDefault(); zona.classList.add("sobre"); };
  zona.ondragleave = () => zona.classList.remove("sobre");
  zona.ondrop = (e) => { e.preventDefault(); zona.classList.remove("sobre"); input.files = e.dataTransfer.files; input.onchange(); };
}
$("#form-cargar").onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
  const cid = $("#c-cliente").value, log = $("#log-cargar");
  if (!cid) throw new Error("Seleccione un cliente.");
  const fd = new FormData(); fd.append("periodo", periodo()); fd.append("cliente", cid);
  if ($("#f-excel").files[0]) fd.append("excel", $("#f-excel").files[0]);
  if ($("#f-balance").files[0]) fd.append("balance", $("#f-balance").files[0]);
  log.hidden = false; log.textContent = "";
  if ($("#f-excel").files[0] || $("#f-balance").files[0]) {
    const r = await api("/api/cargar", {method: "POST", body: fd});
    log.textContent = "Guardado: " + r.guardados.join(", ") + "\nValidando…\n";
  }
  const v = await conCarga(`Validando ${nombreCliente(cid)}…`, () => post("/api/validar", {periodo: periodo(), clientes: [cid]}));
  log.textContent += v.log.join("\n");
  aviso("Validación terminada: " + (ESTADOS[v.estados[cid]] || v.estados[cid]));
  location.hash = "cliente/" + cid;
}); };
$("#form-cliente").onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
  const fd = new FormData(); fd.append("nit", $("#n-nit").value.trim()); fd.append("copropiedad", $("#n-copro").value.trim()); fd.append("software", $("#n-soft").value.trim() || "generico");
  if ($("#n-logo").files[0]) fd.append("logo", $("#n-logo").files[0]);
  const r = await api("/api/clientes", {method: "POST", body: fd});
  st.resumen = await api("/api/resumen"); prepararCarga(); $("#c-cliente").value = r.id;
  aviso("Cliente creado con NIT " + r.nit + ". Complete los demás datos en la ficha."); ev.target.reset();
}); };

// ------------------------------------------------------------------ cliente
async function cargarCliente() {
  const cids = st.resumen.clientes;
  const cid = location.hash.split("/")[1] || st.cliente || cids[0];
  $("#cl-select").innerHTML = cids.map((c) => opcionCliente(c, cid)).join("");
  if (!cid) { $("#cl-cuerpo").innerHTML = `<div class="caja vacio-caja"><b>Aún no hay clientes.</b><span>Cree el primero para ver aquí sus hallazgos y su informe.</span><a class="btn" href="#cargar">Crear el primer cliente</a></div>`; return; }
  st.cliente = cid;
  pintarLogoCliente(cid);
  const d = await api(`/api/cliente/${periodo()}/${cid}`);
  const e = d.estado, todos = [...d.hallazgos.map((h) => ({...h, origen: "Excel"})), ...d.reauditoria.map((h) => ({...h, origen: "Re-auditoría"}))];
  $("#cl-titulo").textContent = d.ficha.copropiedad || cid;
  $("#cl-sub").innerHTML = `${chip(e.estado)} · Periodo ${esc(textoPeriodo(periodo()))} · Informe ${esc(d.ficha.numero_informe || "—")} · Revisor ${esc(d.ficha.revisor || "sin asignar")}`;
  const arch = e.archivos || {};
  const enlaces = ["pptx", "validacion", "borrador", "final", "pdf"].filter((k) => arch[k])
    .map((k) => `<a href="/api/archivo/${periodo()}/${cid}/${encodeURIComponent(arch[k])}" target="_blank">${esc(arch[k])}</a>`).join("");
  const conteo = Object.fromEntries(SEVS.map((s) => [s, todos.filter((h) => h.severidad === s).length]));
  const pendiente = e.estado === "PENDIENTE_VALIDACION";
  const pest = st.pestana || "decision", ed = d.borrador_editado;
  const sinAprobar = e.propuestas_sin_aprobar || e.resumen?.propuestas_sin_aprobar || [];
  $("#cl-cuerpo").innerHTML = `
    ${sinAprobar.length ? `<div class="banda-propuesta" role="status"><span class="banda-ico">!</span><div><b>Este borrador usa propuestas del agente que aún no se aprueban.</b>
      <span>${sinAprobar.length} sección(es) se leyeron con recetas o correcciones del mapa propuestas${e.resumen?.diapositivas_con_propuesta ? ` (${e.resumen.diapositivas_con_propuesta} diapositiva(s) llevan la etiqueta «Leído con propuesta del agente»)` : ""}.
      El informe final no se puede aprobar hasta que un revisor las apruebe y se vuelva a validar.</span></div>
      <button type="button" class="btn sec chico" id="ir-adaptacion">Ir a Adaptación al Excel</button></div>` : ""}
    <div class="pestanas">${[["decision", "Hallazgos y decisión"], ["diapositivas", "Diapositivas"], ["comunicaciones", "Comunicaciones"], ["borrador", "Borrador de opinión" + (ed ? (ed.desactualizado ? " ⚠" : " ✎") : "")], ["ficha", "Ficha del cliente"], ["mapa", "Mapa del Excel"]]
      .map(([k, t]) => `<button data-p="${k}" class="${k === pest ? "activa" : ""}">${t}</button>`).join("")}</div>
    <div id="p-diapositivas" ${pest === "diapositivas" ? "" : "hidden"}></div>
    <div id="p-comunicaciones" ${pest === "comunicaciones" ? "" : "hidden"}></div>
    <div id="p-borrador" ${pest === "borrador" ? "" : "hidden"}></div>
    <div id="p-ficha" ${pest === "ficha" ? "" : "hidden"}></div>
    <div id="p-mapa" ${pest === "mapa" ? "" : "hidden"}></div>
    <div id="p-decision" ${pest === "decision" ? "" : "hidden"}>
    <div class="kpis">${SEVS.map((s) => `<div class="kpi"><b class="sev ${s}">${conteo[s]}</b><small>${s}</small></div>`).join("")}
      <div class="kpi"><b>${e.resumen?.diapositivas ?? "—"}</b><small>Diapositivas</small></div></div>
    ${pendiente && puede("revisor") ? `<div class="caja"><h2>Decisión del revisor fiscal</h2><p class="nota">Nada sale sin aprobación. Con hallazgos ALTA el sistema exige una justificación escrita por el revisor.</p>
      ${sinAprobar.length ? `<p class="m-warn">Primero apruebe las propuestas del agente en Mapa del Excel → Adaptación al Excel y vuelva a validar.</p>` : ""}
      <div class="grupo"><button class="btn" id="btn-aprobar" ${sinAprobar.length ? "disabled title='Hay propuestas del agente sin aprobar'" : ""}>Aprobar…</button><button class="btn vino" id="btn-devolver">Devolver…</button></div></div>`
      : e.estado ? `<div class="caja"><h2>Estado</h2><p>${esc((e.bitacora || []).slice(-1)[0]?.detalle || "")}</p></div>` : ""}
    <div class="caja"><h2>Archivos</h2><div class="archivos">${arch.pptx ? `<a class="destacado" href="/api/pdf/${periodo()}/${cid}" target="_blank" rel="noopener">▶ Abrir el informe en el navegador</a>` : ""}${enlaces || `<span class="vacio">Aún no se ha validado este cliente en el periodo.</span>`}</div>
      ${arch.pptx ? `<p class="nota">Se abre como PDF en una pestaña nueva (la primera vez tarda unos segundos). Si hay diapositivas editadas, muestra la versión editada.</p>` : ""}</div>
    <div class="caja"><h2>Hallazgos (${todos.length})</h2>
      <div class="filtros"><input id="h-buscar" placeholder="Buscar en hallazgos…">${["Todas", ...SEVS].map((s, i) => `<button class="btn chico ${i ? "sec" : ""}" data-sev="${i ? s : ""}">${s}</button>`).join("")}</div>
      <div id="h-lista"></div></div>
    ${d.cuadre.length ? `<div class="caja"><h2>Cuadre contra el balance de prueba</h2><table><tr><th>Concepto</th><th>Excel</th><th>Balance</th><th>Diferencia</th><th></th></tr>
      ${d.cuadre.map((c) => `<tr><td>${esc(c.concepto)}</td><td class="n">${pesos(c.excel)}</td><td class="n">${pesos(c.balance)}</td><td class="n">${pesos(c.diferencia)}</td><td>${c.ok === null ? "—" : c.ok ? "✔" : "✖"}</td></tr>`).join("")}</table></div>` : ""}
    </div>`;
  if ($("#ir-adaptacion")) $("#ir-adaptacion").onclick = () => $('.pestanas button[data-p="mapa"]').click();
  document.querySelectorAll(".pestanas button").forEach((b) => b.onclick = () => {
    st.pestana = b.dataset.p;
    document.querySelectorAll(".pestanas button").forEach((x) => x.classList.toggle("activa", x === b));
    ["decision", "diapositivas", "comunicaciones", "borrador", "ficha", "mapa"].forEach((k) => ($("#p-" + k).hidden = k !== b.dataset.p));
    if (b.dataset.p === "diapositivas") pintarDiapositivas(cid, d.ficha, "actual", !!arch.pptx);
    if (b.dataset.p === "mapa") pintarMapa(cid);
    if (b.dataset.p === "comunicaciones") pintarComunicaciones(cid);
  });
  if (pest === "diapositivas") pintarDiapositivas(cid, d.ficha, "actual", !!arch.pptx);
  if (pest === "comunicaciones") pintarComunicaciones(cid);
  if (pest === "mapa") pintarMapa(cid);
  pintarBorrador(d, cid, pendiente);
  pintarFicha(d, cid, e.estado);
  let filtro = "";
  const pintar = () => {
    const q = $("#h-buscar").value.toLowerCase();
    const lista = todos.filter((h) => (!filtro || h.severidad === filtro) && (!q || JSON.stringify(h).toLowerCase().includes(q)));
    $("#h-lista").innerHTML = lista.map((h) => `<div class="hallazgo ${esc(h.severidad)}"><div class="t"><span class="sev ${esc(h.severidad)}">${esc(h.severidad)}</span> · ${esc(h.regla)} · ${esc(h.titulo)}</div>
      <div class="d">${esc(h.detalle)}</div><div class="m">${esc(h.origen)} · ${esc(h.categoria || "")}${(h.ubicacion || []).length ? " · " + esc(h.ubicacion.slice(0, 6).join(", ")) : ""}${h.recomendacion ? " · " + esc(h.recomendacion) : ""}</div></div>`).join("")
      || `<p class="vacio">Sin hallazgos con ese filtro.</p>`;
  };
  $("#h-buscar").oninput = pintar;
  document.querySelectorAll(".filtros [data-sev]").forEach((b) => b.onclick = () => {
    filtro = b.dataset.sev; document.querySelectorAll(".filtros [data-sev]").forEach((x) => x.classList.toggle("sec", x !== b)); pintar();
  });
  pintar();
  if (pendiente && puede("revisor")) {
    const altas = conteo.ALTA;
    $("#btn-aprobar").onclick = () => dialogo("Aprobar informe", `
      ${ed && !ed.desactualizado ? `<p class="nota">Se enviará el <b>borrador de opinión editado</b> (${esc(ed.guardado || "")}); sus cambios alimentan la mejora continua.</p>` : ""}
      <p class="nota">Si editó diapositivas, se enviará el informe editado como versión final.</p>
      ${ed && ed.desactualizado ? `<p class="sev ALTA">El borrador editado quedó desactualizado: revíselo en la pestaña «Borrador de opinión» antes de aprobar.</p>` : ""}
      <p class="nota">Queda registrado como aprobado por <b>${esc(st.yo.nombre)}</b>.</p>
      <label>Justificación ${altas ? `<b class="sev ALTA">(obligatoria: ${altas} hallazgo(s) ALTA)</b>` : "(opcional)"}<textarea name="justificacion" rows="3" ${altas ? "required" : ""}></textarea></label>`,
      async (f) => { const r = await post("/api/aprobar", {periodo: periodo(), cliente: cid, justificacion: f.justificacion}); aviso(r.detalle); cargarCliente(); });
    $("#btn-devolver").onclick = () => dialogo("Devolver informe", `
      <p class="nota">Queda registrado como devuelto por <b>${esc(st.yo.nombre)}</b>.</p>
      <label>Motivo<textarea name="motivo" rows="3" required></textarea></label>`,
      async (f) => { const r = await post("/api/devolver", {periodo: periodo(), cliente: cid, motivo: f.motivo}); aviso(r.detalle); cargarCliente(); });
  }
}
// Logo del cliente: arriba a la derecha al elegirlo; clic para subirlo o cambiarlo
function pintarLogoCliente(cid) {
  const img = $("#cl-logo-img");
  img.onload = () => { img.hidden = false; $("#cl-logo-vacio").hidden = true; };
  img.onerror = () => { img.hidden = true; $("#cl-logo-vacio").hidden = !puede("asistente", "revisor"); };
  img.alt = "Logo de " + nombreCliente(cid);
  if ((st.resumen.logos || []).includes(cid) || pintarLogoCliente.subido === cid) img.src = `/api/clientes/${encodeURIComponent(cid)}/logo?v=${Date.now()}`;
  else img.onerror();                                     // sin logo: no se pide (evita errores 404 en la consola)
  $("#cl-logo").disabled = !puede("asistente", "revisor");
  $("#cl-logo").title = puede("asistente", "revisor") ? "Subir o cambiar el logo del cliente" : "Logo del cliente";
}
$("#cl-logo").onclick = () => $("#f-logo").click();
$("#f-logo").onchange = async () => {
  const f = $("#f-logo").files[0]; $("#f-logo").value = "";
  if (!f || !st.cliente) return;
  const fd = new FormData(); fd.append("logo", f);
  try { await api(`/api/clientes/${encodeURIComponent(st.cliente)}/logo`, {method: "POST", body: fd}); pintarLogoCliente.subido = st.cliente; pintarLogoCliente(st.cliente); aviso("Logo del cliente actualizado."); }
  catch (e) { aviso(e.message, true); }
};
$("#cl-select").onchange = () => (location.hash = "cliente/" + $("#cl-select").value);

// Visor y editor de diapositivas: las ediciones se guardan en una copia que se envía como informe final al aprobar.
const digitos = (t) => (String(t).match(/\d[\d.,]*/g) || []).sort().join("|");
async function pintarDiapositivas(cid, ficha, version = "actual", hayInforme = true) {
  const caja = $("#p-diapositivas");
  if (!hayInforme) {
    caja.innerHTML = `<div class="caja"><h2>Todavía no hay diapositivas</h2>
      <p>Este cliente aún no tiene informe en ${esc(textoPeriodo(periodo()))}: las diapositivas se generan al validar el Excel.</p>
      <p class="nota">Vaya a <b>Cargar y validar</b>, elija el cliente y pulse validar (o use «Validar todo el lote» en el Panel del lote). Si ya lo validó y el estado sigue en «Recibido», la validación no terminó: vuelva a intentarlo y fíjese en el mensaje de error.</p>
      <div class="grupo"><a class="btn" href="#cargar">Ir a Cargar y validar</a></div></div>`;
    return;
  }
  if (st.dCliente !== cid + periodo()) { st.dCliente = cid + periodo(); st.dActual = 1; st.cambios = {}; }
  caja.innerHTML = `<div class="caja"><p class="vacio">Preparando las diapositivas… la primera vez toma unos segundos.</p></div>`;
  let D;
  try { D = await api(`/api/diapositivas/${periodo()}/${cid}?version=${version}`); }
  catch (e) { caja.innerHTML = `<div class="caja"><p class="vacio">${esc(e.message)}</p></div>`; return; }
  const ed = D.edicion, editable = D.editable && D.version !== "generado" || (D.editable && !ed);
  const pendientes = () => Object.values(st.cambios).reduce((s, x) => s + Object.keys(x).length, 0);
  caja.innerHTML = `
    ${ed?.desactualizado ? `<div class="caja alerta"><b>El informe se volvió a validar después de estas ediciones.</b> Descártelas y vuelva a hacerlas sobre el informe nuevo.</div>` : ""}
    <div class="caja barra-diap">
      <div class="grupo">
        <span class="chip">${D.version === "editado" ? "Versión editada" : "Versión generada"}</span>
        ${ed ? `<span class="nota">${ed.cambios} cambio(s) frente al generado · guardado ${esc(ed.guardado || "")}</span>` : `<span class="nota">Sin ediciones</span>`}
      </div>
      <div class="grupo">
        ${ed ? `<button class="btn sec chico" id="d-ver">${D.version === "editado" ? "Ver el generado" : "Ver el editado"}</button>
          <a class="btn sec chico" href="/api/archivo/${periodo()}/${cid}/${encodeURIComponent("Informe_Editado_" + cid + "_" + periodo() + ".pptx")}">Descargar editado</a>
          ${D.editable ? `<button class="btn sec chico" id="d-descartar">Descartar ediciones</button>` : ""}` : ""}
        ${editable ? `<button class="btn chico" id="d-guardar" disabled>Guardar cambios</button>` : ""}
      </div>
    </div>
    <div class="visor">
      <ol class="miniaturas" id="d-minis">${D.diapositivas.map((x) => `<li data-n="${x.n}" class="${x.editadas.length ? "editada" : ""}">
        <img loading="lazy" src="${x.imagen}" alt="Diapositiva ${x.n}"><span>${x.n}. ${esc((x.titulo || "").slice(0, 46))}</span></li>`).join("")}</ol>
      <div class="escenario">
        <div class="nav-diap"><button class="btn sec chico" id="d-ant">◀</button><b id="d-num"></b><button class="btn sec chico" id="d-sig">▶</button></div>
        <img id="d-img" alt="">
        <div id="d-textos"></div>
      </div>
    </div>`;
  const botonGuardar = () => { const b = $("#d-guardar"); if (b) { const n = pendientes(); b.disabled = !n; b.textContent = n ? `Guardar ${n} cambio(s)` : "Guardar cambios"; } };
  const ir = (n) => {
    n = Math.max(1, Math.min(D.diapositivas.length, n)); st.dActual = n;
    const x = D.diapositivas[n - 1], mios = st.cambios[n] || {};
    $("#d-img").src = x.imagen; $("#d-num").textContent = `${n} / ${D.diapositivas.length}`;
    document.querySelectorAll("#d-minis li").forEach((li) => li.classList.toggle("activa", +li.dataset.n === n));
    $(`#d-minis li[data-n="${n}"]`)?.scrollIntoView({block: "nearest"});
    const claves = Object.keys(x.unidades);
    $("#d-textos").innerHTML = claves.length ? `<h3>Textos de la diapositiva</h3>
      <p class="nota">${editable ? "Corrija la redacción aquí. Las cifras salen del Excel: si una está mal, corrija el Excel y vuelva a validar; cambiarla aquí genera una alerta para revisar el cálculo." : "Solo lectura."}</p>
      ${claves.map((k) => { const v = mios[k] ?? x.unidades[k], tabla = k.includes(".t");
        const [, f, c] = tabla ? k.match(/t(\d+)\.(\d+)/) : [];
        return `<label class="unidad ${x.editadas.includes(k) ? "guardada" : ""} ${k in mios ? "pendiente" : ""}">
          <span>${tabla ? `Tabla · fila ${+f + 1}, col. ${+c + 1}` : "Texto"}${x.editadas.includes(k) ? " · editado" : ""}<em class="aviso-cifra" hidden>cambia cifras</em></span>
          <textarea data-k="${esc(k)}" rows="${Math.min(6, Math.ceil(v.length / 70) || 1)}" ${editable ? "" : "readonly"}>${esc(v)}</textarea></label>`; }).join("")}`
      : `<p class="vacio">Esta diapositiva no tiene textos editables (imagen o gráfico).</p>`;
    document.querySelectorAll("#d-textos textarea").forEach((ta) => {
      const marcar = () => { const orig = x.unidades[ta.dataset.k];
        ta.closest(".unidad").querySelector(".aviso-cifra").hidden = digitos(ta.value) === digitos(orig); };
      marcar();
      ta.oninput = () => {
        const orig = x.unidades[ta.dataset.k];
        st.cambios[n] = st.cambios[n] || {};
        if (ta.value === orig) delete st.cambios[n][ta.dataset.k]; else st.cambios[n][ta.dataset.k] = ta.value;
        if (!Object.keys(st.cambios[n]).length) delete st.cambios[n];
        ta.closest(".unidad").classList.toggle("pendiente", ta.dataset.k in (st.cambios[n] || {}));
        $(`#d-minis li[data-n="${n}"]`).classList.toggle("pendiente", !!st.cambios[n]);
        marcar(); botonGuardar();
      };
    });
  };
  document.querySelectorAll("#d-minis li").forEach((li) => { li.onclick = () => ir(+li.dataset.n); li.classList.toggle("pendiente", !!st.cambios[li.dataset.n]); });
  $("#d-ant").onclick = () => ir(st.dActual - 1); $("#d-sig").onclick = () => ir(st.dActual + 1);
  caja.onkeydown = (e) => { if (e.target.tagName === "TEXTAREA") return; if (e.key === "ArrowLeft") ir(st.dActual - 1); if (e.key === "ArrowRight") ir(st.dActual + 1); };
  caja.tabIndex = -1;
  if ($("#d-ver")) $("#d-ver").onclick = () => pintarDiapositivas(cid, ficha, D.version === "editado" ? "generado" : "editado");
  if ($("#d-descartar")) $("#d-descartar").onclick = (ev) => ocupado(ev.target, async () => {
    if (!confirm("¿Descartar todas las ediciones guardadas de las diapositivas?")) return;
    await api(`/api/diapositivas/${periodo()}/${cid}`, {method: "DELETE"}); st.cambios = {}; aviso("Ediciones descartadas."); pintarDiapositivas(cid, ficha);
  });
  if ($("#d-guardar")) $("#d-guardar").onclick = (ev) => ocupado(ev.target, async () => {
    const r = await api("/api/diapositivas", {method: "PUT", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({periodo: periodo(), cliente: cid, cambios: st.cambios})});
    st.cambios = {};
    aviso(`Guardado. ${r.diferencias} cambio(s) frente al informe generado${r.cifras ? ` · ${r.cifras} cambian cifras (quedará una alerta para revisar el cálculo)` : ""}.`, r.cifras > 0);
    pintarDiapositivas(cid, ficha);
  });
  botonGuardar(); ir(st.dActual || 1);
}

// Comunicaciones al Consejo o a la administración: la IA configurada redacta un BORRADOR solo con las fuentes elegidas.
async function pintarComunicaciones(cid, abrir = null) {
  const caja = $("#p-comunicaciones");
  const C = await api(`/api/comunicaciones/${periodo()}/${cid}`);
  const opera = puede("asistente", "revisor");
  caja.innerHTML = `
    <div class="dos">
      <form class="caja" id="c-nueva">
        <h2>Nueva comunicación</h2>
        ${C.asistente.ok ? "" : `<p class="sev ALTA">${esc(C.asistente.ayuda)}</p>`}
        <label>Tipo<select name="tipo" id="c-tipo">${Object.entries(C.tipos).map(([k, v]) => `<option value="${k}">${esc(v)}</option>`).join("")}</select></label>
        <div id="c-recibida-caja" hidden><label>Texto de la comunicación recibida<textarea name="recibida" rows="6" placeholder="Pegue aquí el oficio, correo o solicitud que se responde"></textarea></label></div>
        <div class="cab-mini"><b class="nota">Hallazgos que fundamentan la comunicación</b>
          <span class="grupo"><button type="button" class="btn sec chico" data-sel="ALTA">Solo ALTA</button><button type="button" class="btn sec chico" data-sel="">Ninguno</button></span></div>
        <div class="lista-check" id="c-hallazgos">${C.hallazgos.map((h) => `<label class="check"><input type="checkbox" name="h" value="${esc(h.id)}" ${h.severidad === "ALTA" ? "checked" : ""}>
          <span><span class="sev ${esc(h.severidad)}">${esc(h.severidad)}</span> ${esc(h.regla)} · ${esc(h.titulo)}</span></label>`).join("") || `<p class="vacio">Valide primero el cliente.</p>`}</div>
        <label class="check"><input type="checkbox" name="borrador" checked> <span>Incluir el borrador de opinión (el editado por el revisor si existe)</span></label>
        <label class="check"><input type="checkbox" name="normativa" checked> <span>Buscar fundamentos en la base normativa (solo se citan normas cargadas)</span></label>
        <label>Instrucciones para el enfoque (opcional)<textarea name="instrucciones" rows="2" placeholder="Ej.: énfasis en la cartera y en los acuerdos de pago; tono para presentar en la próxima sesión"></textarea></label>
        <p class="nota">La IA solo usa las fuentes elegidas y la ficha del cliente; no inventa normas, fechas ni plazos (deja marcadores [entre corchetes]). Siempre es un borrador: la aplicación nunca lo envía.</p>
        <div class="acciones">${opera ? `<button class="btn" ${C.asistente.ok ? "" : "disabled"}>Redactar borrador</button>` : `<span class="nota">Su rol no permite redactar comunicaciones.</span>`}</div>
      </form>
      <div class="caja"><h2>Comunicaciones del periodo</h2>
        ${C.lista.map((x) => `<div class="item-com" data-id="${esc(x.id)}"><b>${esc(x.asunto || x.tipo_nombre)}</b>
          <small class="nota">${esc(x.tipo_nombre)} · ${esc(x.creado)} · ${esc(x.creado_por || "")}</small>
          ${x.estado === "revisado" ? `<span class="chip e-APROBADO">Revisado por ${esc(x.revisado_por)}</span>` : `<span class="chip e-PENDIENTE_VALIDACION">Borrador</span>`}</div>`).join("")
          || `<p class="vacio">Aún no hay comunicaciones para este cliente y periodo.</p>`}
      </div>
    </div>
    <div id="c-doc"></div>`;
  const tipo = $("#c-tipo");
  tipo.onchange = () => ($("#c-recibida-caja").hidden = tipo.value !== "respuesta");
  caja.querySelectorAll("[data-sel]").forEach((b) => b.onclick = () => caja.querySelectorAll("#c-hallazgos input").forEach((i) => {
    i.checked = !!b.dataset.sel && i.closest("label").querySelector(".sev").textContent === b.dataset.sel; }));
  caja.querySelectorAll(".item-com").forEach((x) => x.onclick = () => abrirComunicacion(cid, x.dataset.id));
  $("#c-nueva").onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
    const f = new FormData(ev.target);
    ev.submitter.textContent = "Redactando… (puede tardar un minuto)";
    const doc = await post("/api/comunicaciones", {periodo: periodo(), cliente: cid, tipo: f.get("tipo"), hallazgos: f.getAll("h"),
      incluir_borrador: f.get("borrador") === "on", incluir_normativa: f.get("normativa") === "on", recibida: f.get("recibida") || "", instrucciones: f.get("instrucciones") || ""});
    aviso("Borrador redactado. Revíselo antes de firmar."); await pintarComunicaciones(cid, doc.id);
  }); };
  if (abrir) abrirComunicacion(cid, abrir);
}

async function abrirComunicacion(cid, id) {
  const d = await api(`/api/comunicaciones/${periodo()}/${cid}/${id}`), caja = $("#c-doc");
  const opera = puede("asistente", "revisor"), nombres = Object.fromEntries(d.fuentes.map((f) => [f.id, f.titulo]));
  caja.innerHTML = `
    <div class="caja">
      <div class="cab-mini"><h2>${esc(d.tipo_nombre)}</h2>
        <span>${d.estado === "revisado" ? `<span class="chip e-APROBADO">Revisado por ${esc(d.revisado_por)}</span>` : `<span class="chip e-PENDIENTE_VALIDACION">Borrador</span>`}</span></div>
      <p class="nota">${esc(d.resumen)}</p>
      ${d.cifras_sin_respaldo.length ? `<div class="caja alerta"><b>Cifras sin respaldo en las fuentes:</b> ${d.cifras_sin_respaldo.map(esc).join(", ")}. Verifíquelas o retírelas antes de firmar.</div>` : ""}
      ${(d.citas_invalidas || []).length ? `<div class="caja alerta"><b>Citas a fuentes que no existen:</b> ${d.citas_invalidas.map(esc).join(", ")} (solo se entregaron ${d.fuentes.map((f) => esc(f.id)).join(", ")}). Corrija o retire esas referencias antes de firmar.</div>` : ""}
      ${d.observaciones.length ? `<div class="obs"><b>Observaciones de la redacción</b><ul>${d.observaciones.map((o) => `<li>${esc(o)}</li>`).join("")}</ul></div>` : ""}
      <label>Asunto<input id="c-asunto" value="${esc(d.asunto)}" ${opera ? "" : "readonly"}></label>
      <div class="editor"><textarea id="c-texto" ${opera ? "" : "readonly"}>${esc(d.oficio)}</textarea><div class="md vista-previa" id="c-previa"></div></div>
      <div class="acciones">
        <a class="btn sec" href="/api/comunicaciones/${periodo()}/${cid}/${d.id}/word">Descargar Word</a>
        ${opera ? `<button class="btn sec" id="c-eliminar">Eliminar</button><button class="btn sec" id="c-guardar">Guardar cambios</button>` : ""}
        ${puede("revisor") ? `<button class="btn" id="c-revisar">Guardar y marcar revisado</button>` : ""}
      </div>
      <details class="mt"><summary><b>Trazabilidad de fuentes (${d.trazabilidad.length})</b></summary>
        <table><tr><th>Elemento</th><th>Fuente</th><th>Referencia</th></tr>${d.trazabilidad.map((x) => `<tr><td>${esc(x.elemento)}</td><td>${esc(x.fuente)} <small class="nota">${esc(nombres[x.fuente] || "")}</small></td><td>${esc(x.referencia)}</td></tr>`).join("")}</table>
      </details>
      <details><summary><b>Fuentes entregadas a la IA (${d.fuentes.length})</b></summary>
        ${(d.fuentes_texto || []).map((f) => `<div class="fuente"><b>${esc(f.id)} · ${esc(f.titulo)}</b><pre>${esc(f.texto)}</pre></div>`).join("")}</details>
    </div>`;
  const ta = $("#c-texto"), previa = () => ($("#c-previa").innerHTML = md(ta.value));
  ta.oninput = previa; previa();
  const guardar = (revisado) => async (ev) => ocupado(ev.target, async () => {
    if (revisado && !confirm("¿Marcar como revisada? Queda registrado con su nombre. La aplicación no envía la comunicación: fírmela y envíela por el canal habitual.")) return;
    const r = await api(`/api/comunicaciones/${periodo()}/${cid}/${d.id}`, {method: "PUT", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({oficio: ta.value, asunto: $("#c-asunto").value, revisado})});
    const alertas = (r.cifras_sin_respaldo.length ? " Hay cifras sin respaldo." : "") + ((r.citas_invalidas || []).length ? " Hay citas a fuentes inexistentes." : "");
    aviso((revisado ? "Comunicación marcada como revisada." : "Cambios guardados.") + alertas, !!alertas);
    pintarComunicaciones(cid, d.id);
  });
  if ($("#c-guardar")) $("#c-guardar").onclick = guardar(false);
  if ($("#c-revisar")) $("#c-revisar").onclick = guardar(true);
  if ($("#c-eliminar")) $("#c-eliminar").onclick = (ev) => ocupado(ev.target, async () => {
    if (!confirm("¿Eliminar este borrador?")) return;
    await api(`/api/comunicaciones/${periodo()}/${cid}/${d.id}`, {method: "DELETE"}); aviso("Borrador eliminado."); pintarComunicaciones(cid);
  });
  caja.scrollIntoView({behavior: "smooth", block: "start"});
}

// Configuración: proveedores de IA y RSI
async function cargarConfiguracion() {
  if (!puede("admin", "revisor")) { location.hash = "panel"; return; }
  const C = await api("/api/configuracion"), ia = C.ia, admin = C.puede_editar;
  document.querySelectorAll("#v-configuracion [data-admin]").forEach((x) => (x.hidden = !admin));
  $("#cfg-proveedores").innerHTML = `<tr><th>Proveedor</th><th>Tipo</th><th>Clave</th><th>Modelos</th><th></th></tr>` + Object.entries(ia.proveedores).map(([id, p]) => `
    <tr><td><b>${esc(p.nombre)}</b><br><small class="nota">${esc(id)}${p.base_url ? " · " + esc(p.base_url) : ""}</small></td>
    <td><small>${p.tipo === "anthropic" ? "Claude (SDK)" : "Compatible OpenAI"}</small></td>
    <td>${p.con_clave ? `<span class="chip e-APROBADO">${esc(p.api_key || "del entorno")}</span>` : `<span class="chip e-DEVUELTO">Sin clave</span>`}</td>
    <td><small>${p.modelos.map(esc).join(", ") || "—"}</small></td>
    <td class="grupo">${admin ? `<button class="btn sec chico" data-editar="${esc(id)}">Editar</button><button class="btn sec chico" data-probar="${esc(id)}">Probar</button>
      <button class="btn sec chico" data-borrar-p="${esc(id)}">Eliminar</button>` : ""}</td></tr>`).join("");
  $("#cfg-preset").innerHTML = `<option value="">— elegir —</option>` + Object.entries(ia.presets).map(([k, p]) => `<option value="${k}">${esc(p.nombre)}</option>`).join("");
  const F = $("#cfg-prov-form");
  const llenar = (id, p) => { F.id.value = id; F.nombre.value = p.nombre; F.tipo.value = p.tipo; F.base_url.value = p.base_url || "";
    F.modelos.value = (p.modelos || []).join(", "); F.api_key.value = ""; F.scrollIntoView({block: "center"}); };
  $("#cfg-preset").onchange = () => { const k = $("#cfg-preset").value; if (k) llenar(k, ia.presets[k]); };
  document.querySelectorAll("[data-editar]").forEach((b) => b.onclick = () => llenar(b.dataset.editar, ia.proveedores[b.dataset.editar]));
  document.querySelectorAll("[data-borrar-p]").forEach((b) => b.onclick = (ev) => ocupado(ev.target, async () => {
    if (!confirm("¿Eliminar este proveedor y su clave?")) return;
    await api(`/api/configuracion/proveedores/${encodeURIComponent(b.dataset.borrarP)}`, {method: "DELETE"}); aviso("Proveedor eliminado."); cargarConfiguracion(); }));
  document.querySelectorAll("[data-probar]").forEach((b) => b.onclick = (ev) => ocupado(ev.target, async () => {
    const p = ia.proveedores[b.dataset.probar], modelo = prompt("Modelo a probar:", p.modelos[0] || "");
    if (!modelo) return;
    const r = await post("/api/configuracion/probar", {proveedor: b.dataset.probar, modelo}); aviso(r.detalle, !r.ok); }));
  F.onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
    await api("/api/configuracion/proveedores", {method: "PUT", headers: {"Content-Type": "application/json"}, body: JSON.stringify({
      id: F.id.value.trim(), nombre: F.nombre.value.trim(), tipo: F.tipo.value, base_url: F.base_url.value.trim(),
      api_key: F.api_key.value || null, modelos: F.modelos.value.split(",").map((m) => m.trim()).filter(Boolean)})});
    aviso("Proveedor guardado."); F.reset(); cargarConfiguracion(); }); };
  $("#cfg-usos").innerHTML = Object.entries(ia.usos).map(([uso, nombre]) => {
    const a = ia.uso[uso];
    return `<form class="grupo uso" data-uso="${uso}"><b class="uso-nombre">${esc(nombre)}</b>
      <select name="proveedor" ${admin ? "" : "disabled"}>${Object.entries(ia.proveedores).map(([id, p]) => `<option value="${esc(id)}" ${id === a.proveedor ? "selected" : ""}>${esc(p.nombre)}</option>`).join("")}</select>
      <input name="modelo" list="dl-${uso}" value="${esc(a.modelo)}" ${admin ? "" : "readonly"}>
      <datalist id="dl-${uso}"></datalist>${admin ? `<button class="btn sec chico">Asignar</button>` : ""}</form>`; }).join("");
  document.querySelectorAll("#cfg-usos form").forEach((f) => {
    const dl = () => (f.querySelector("datalist").innerHTML = (ia.proveedores[f.proveedor.value]?.modelos || []).map((m) => `<option value="${esc(m)}">`).join(""));
    f.proveedor.onchange = dl; dl();
    f.onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
      await api("/api/configuracion/uso", {method: "PUT", headers: {"Content-Type": "application/json"},
        body: JSON.stringify({uso: f.dataset.uso, proveedor: f.proveedor.value, modelo: f.modelo.value})});
      aviso("Asignación guardada."); }); };
  });
  const R = $("#cfg-rsi-form");
  $("#cfg-rsi-modo").innerHTML = Object.entries(C.rsi.modos).map(([k, v]) => `<option value="${k}" ${k === C.rsi.config.modo ? "selected" : ""}>${esc(v)}</option>`).join("");
  R.margen.value = +(C.rsi.config.margen * 100).toFixed(1); R.candidatos.value = C.rsi.config.candidatos; R.hora.value = C.rsi.config.hora || "04:00";
  [...R.elements].forEach((x) => { if (x.name) x.disabled = !admin; });
  R.onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
    await api("/api/configuracion/rsi", {method: "PUT", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({modo: R.modo.value, margen: +R.margen.value / 100, candidatos: +R.candidatos.value, hora: R.hora.value})});
    aviso("Configuración de RSI guardada."); cargarRSI(); }); };
  $("#cfg-rsi-ciclo").onclick = (ev) => ocupado(ev.target, async () => {
    ev.target.textContent = "Ejecutando… (puede tardar unos minutos)";
    const r = await post("/api/rsi/ciclo", {});
    aviso(r.plantillas.map((p) => `${p.clave}: ${p.accion}`).join(" · ") + (r.meta?.accion ? ` · ${r.meta.accion}` : "") || "Ciclo terminado.");
    cargarRSI(); });
  cargarRSI();
}
async function cargarRSI() {
  const R = await api("/api/rsi"), e = R.estado, pct = (v) => (v * 100).toFixed(1) + "%";
  const casos = Object.entries(R.casos);
  const estr = (id) => e.estrategias[id];
  $("#cfg-rsi").innerHTML = `
    ${R.disponible.ok ? "" : `<p class="sev ALTA">${esc(R.disponible.ayuda)}</p>`}
    <div class="kpis"><div class="kpi"><b>${casos.reduce((s, [, n]) => s + n, 0)}</b><small>Casos para aprender (opiniones editadas)</small></div>
      <div class="kpi"><b>${casos.length}</b><small>Plantillas con casos</small></div>
      <div class="kpi"><b>${esc(e.vigente)}</b><small>Estrategia vigente</small></div>
      <div class="kpi"><b>${esc(e.retadora || "—")}</b><small>Estrategia retadora</small></div></div>
    ${casos.length ? "" : `<p class="nota">Aún no hay casos: la automejora aprende cuando el revisor edita el borrador de opinión y aprueba el informe.</p>`}
    <h3>Propuestas de la automejora</h3>
    ${R.propuestas.filter((p) => p.estado === "pendiente").map((p) => `<div class="propuesta"><div class="id">${esc(p.id)}</div>
      <b>${esc(p.clave)}</b> · similitud con el revisor ${pct(p.puntaje_actual)} → <b>${pct(p.puntaje_nuevo)}</b> (+${pct(p.ganancia)})${p.independiente ? "" : " · <span class='sev MEDIA'>sin caso de prueba independiente</span>"}
      <div class="comparar"><div><small class="nota">Actual</small><blockquote>${esc(p.actual)}</blockquote></div><div><small class="nota">Propuesta</small><blockquote>${esc(p.texto)}</blockquote></div></div>
      <p class="nota">${esc(p.razon)}</p>
      ${puede("revisor") ? `<div class="grupo"><button class="btn chico" data-rsi="${esc(p.id)}|aprobar">Aprobar y publicar</button><button class="btn sec chico" data-rsi="${esc(p.id)}|rechazar">Rechazar</button></div>` : ""}</div>`).join("") || `<p class="vacio">Sin propuestas pendientes.</p>`}
    <h3>Estrategias (nivel recursivo)</h3>
    ${Object.entries(e.estrategias).reverse().map(([id, s]) => `<details ${id === e.vigente || id === e.retadora ? "open" : ""}><summary><b>${esc(id)}</b>
      ${id === e.vigente ? `<span class="chip e-APROBADO">vigente</span>` : id === e.retadora ? `<span class="chip e-PENDIENTE_VALIDACION">retadora</span>` : s.descartada ? `<span class="chip e-BLOQUEADO">descartada</span>` : ""}
      <small class="nota">${esc(s.origen)} · ${s.ciclos} ciclo(s) · ganancia acumulada ${pct(s.ganancia_total)}</small></summary>
      <p class="extracto">${esc(s.texto)}</p>${s.aprendizajes ? `<p class="nota"><b>Aprendizajes:</b> ${esc(s.aprendizajes)}</p>` : ""}
      ${puede("admin") && id !== e.vigente ? `<button class="btn sec chico" data-restaurar="${esc(id)}">Usar como vigente</button>` : ""}</details>`).join("")}
    <h3 class="mt">Ciclos recientes</h3>
    ${R.ciclos.map((c) => `<div class="propuesta"><small class="nota">${esc(c.fecha)} · ${esc(c.autor)} · modo ${esc(c.modo)} · estrategias ${c.estrategias.map(esc).join(" vs ")}</small>
      <ul>${c.plantillas.map((p) => `<li><b>${esc(p.clave)}</b>: ${pct(p.puntaje_actual)} → ${p.ganador ? pct(p.ganador.prueba) : "—"} · ${esc(p.accion)}</li>`).join("")}</ul>
      ${c.meta?.accion ? `<p class="nota">Nivel recursivo: ${esc(c.meta.accion)}</p>` : ""}</div>`).join("") || `<p class="vacio">Aún no se ha ejecutado ningún ciclo.</p>`}`;
  document.querySelectorAll("[data-rsi]").forEach((b) => b.onclick = (ev) => ocupado(ev.target, async () => {
    const [id, acc] = b.dataset.rsi.split("|"); const r = await post(`/api/rsi/propuestas/${encodeURIComponent(id)}/${acc}`, {}); aviso(r.detalle); cargarRSI(); }));
  document.querySelectorAll("[data-restaurar]").forEach((b) => b.onclick = (ev) => ocupado(ev.target, async () => {
    await post(`/api/rsi/estrategias/${encodeURIComponent(b.dataset.restaurar)}/restaurar`, {}); aviso("Estrategia restaurada como vigente."); cargarRSI(); }));
}

// Base normativa
async function cargarNormativa() {
  const N = await api("/api/normativa");
  const ambitos = {general: "General (todas las copropiedades)", ...N.clientes};
  $("#n-cliente-b").innerHTML = `<option value="">Solo general</option>` + Object.entries(N.clientes).map(([k, v]) => `<option value="${esc(k)}">General + ${esc(v)}</option>`).join("");
  $("#n-tipo").innerHTML = Object.entries(N.tipos).map(([k, v]) => `<option value="${k}">${esc(v)}</option>`).join("");
  $("#n-ambito").innerHTML = Object.entries(ambitos).map(([k, v]) => `<option value="${esc(k)}">${esc(v)}</option>`).join("");
  const edita = puede("revisor", "admin");
  $("#n-docs").innerHTML = N.documentos.length ? `<table><tr><th>Documento</th><th>Ámbito</th><th>Estructura</th><th></th></tr>${N.documentos.map((d) => `
    <tr><td><b>${esc(d.titulo)}</b><br><small class="nota">${esc(N.tipos[d.tipo])}${d.version ? " · " + esc(d.version) : ""} · cargado por ${esc(d.cargado_por)} (${esc(d.cargado.slice(0, 10))})</small></td>
    <td><small>${esc(ambitos[d.ambito] || d.ambito)}</small></td><td><small>${d.fragmentos} ${esc(d.estructura.toLowerCase())}s${(d.archivos || []).length > 1 ? ` · ${d.archivos.length} partes` : ""}</small>
      ${(d.advertencias || []).map((a) => `<div class="sev MEDIA"><small>⚠ ${esc(a)}</small></div>`).join("")}</td>
    <td class="grupo"><button class="btn sec chico" data-ver="${esc(d.ambito)}|${esc(d.id)}">Ver</button>
      ${(d.archivos || [d.archivo]).map((n, i, a) => `<a class="btn sec chico" title="${esc(n)}" href="/api/normativa/${encodeURIComponent(d.ambito)}/${encodeURIComponent(d.id)}/original?parte=${i + 1}">${a.length > 1 ? "Parte " + (i + 1) : "Original"}</a>`).join("")}
      ${edita ? `<button class="btn sec chico" data-borrar="${esc(d.ambito)}|${esc(d.id)}">Eliminar</button>` : ""}</td></tr>`).join("")}</table>`
    : `<p class="vacio">La base está vacía. Cargue la Ley 675 de 2001, las NIIF para pymes y los manuales de cada copropiedad desde su fuente oficial.</p>`;
  document.querySelectorAll("[data-ver]").forEach((b) => b.onclick = () => verNorma(...b.dataset.ver.split("|")));
  document.querySelectorAll("[data-borrar]").forEach((b) => b.onclick = (ev) => ocupado(ev.target, async () => {
    if (!confirm("¿Eliminar este documento de la base normativa?")) return;
    const [a, id] = b.dataset.borrar.split("|"); await api(`/api/normativa/${encodeURIComponent(a)}/${encodeURIComponent(id)}`, {method: "DELETE"});
    aviso("Documento eliminado."); cargarNormativa();
  }));
}
async function verNorma(ambito, id, fragmento = null) {
  const D = await api(`/api/normativa/${encodeURIComponent(ambito)}/${encodeURIComponent(id)}`);
  $("#n-visor").innerHTML = `<div class="caja"><div class="cab-mini"><h2>${esc(D.meta.titulo)}</h2><button class="btn sec chico" id="n-cerrar">Cerrar</button></div>
    <p class="nota">${D.fragmentos.length} fragmentos · ${esc(D.meta.archivo)}${D.meta.origen ? " · origen: " + esc(D.meta.origen) : ""}</p>
    ${D.fragmentos.map((f) => `<details class="frag" id="frag-${f.id}" ${f.id === fragmento ? "open" : ""}><summary><b>${esc(f.ref)}</b>${f.seccion ? ` <small class="nota">${esc(f.seccion)}</small>` : ""} <small class="nota">pág. ${f.pagina}</small></summary><pre>${esc(f.texto)}</pre></details>`).join("")}</div>`;
  $("#n-cerrar").onclick = () => ($("#n-visor").innerHTML = "");
  (fragmento !== null ? $(`#frag-${fragmento}`) : $("#n-visor")).scrollIntoView({behavior: "smooth", block: "start"});
}
$("#n-buscar").onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
  const f = new FormData(ev.target), q = f.get("q");
  const R = await api(`/api/normativa/buscar?q=${encodeURIComponent(q)}` + (f.get("cliente") ? `&cliente=${encodeURIComponent(f.get("cliente"))}` : ""));
  const palabras = q.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").split(/\W+/).filter((w) => w.length > 3);
  const resaltar = (t) => esc(t).replace(/\S+/g, (w) => palabras.some((p) => w.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").includes(p)) ? `<mark>${w}</mark>` : w);
  $("#n-resultados").innerHTML = R.resultados.length ? R.resultados.map((r) => `<div class="propuesta"><b>${esc(r.cita)}</b> <small class="nota">${esc(r.tipo)}${r.seccion ? " · " + esc(r.seccion) : ""} · pág. ${r.pagina}</small>
    <p class="extracto">${resaltar(r.texto.slice(0, 700))}${r.texto.length > 700 ? "…" : ""}</p>
    <button class="btn sec chico" data-abrir="${esc(r.ambito)}|${esc(r.doc_id)}|${r.fragmento}">Ver completo</button></div>`).join("")
    : `<p class="vacio">Sin resultados en la base normativa cargada.</p>`;
  document.querySelectorAll("[data-abrir]").forEach((b) => b.onclick = () => { const [a, id, fr] = b.dataset.abrir.split("|"); verNorma(a, id, +fr); });
}); };
// Partes del documento normativo: se unen en el orden de la lista (por defecto, alfabético por nombre de archivo)
st.partes = [];
function pintarPartes() {
  $("#n-partes").innerHTML = st.partes.map((f, i) => `<li><span>${esc(f.name)}</span>
    <button type="button" class="btn sec chico" data-mov="${i}|-1" ${i ? "" : "disabled"}>↑</button>
    <button type="button" class="btn sec chico" data-mov="${i}|1" ${i < st.partes.length - 1 ? "" : "disabled"}>↓</button>
    <button type="button" class="btn sec chico" data-quitar="${i}">✕</button></li>`).join("");
  document.querySelectorAll("#n-partes [data-mov]").forEach((b) => b.onclick = () => {
    const [i, d] = b.dataset.mov.split("|").map(Number); [st.partes[i], st.partes[i + d]] = [st.partes[i + d], st.partes[i]]; pintarPartes(); });
  document.querySelectorAll("#n-partes [data-quitar]").forEach((b) => b.onclick = () => { st.partes.splice(+b.dataset.quitar, 1); pintarPartes(); });
}
{ const zona = $("#zona-norma"), input = $("#f-norma");
  const agregar = (lista) => { st.partes.push(...lista); st.partes.sort((a, b) => a.name.localeCompare(b.name, "es", {numeric: true})); pintarPartes(); };
  zona.onclick = () => input.click(); input.onchange = () => { agregar([...input.files]); input.value = ""; };
  zona.ondragover = (e) => { e.preventDefault(); zona.classList.add("sobre"); }; zona.ondragleave = () => zona.classList.remove("sobre");
  zona.ondrop = (e) => { e.preventDefault(); zona.classList.remove("sobre"); agregar([...e.dataTransfer.files]); }; }
$("#n-form").onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
  if (!st.partes.length) throw new Error("Seleccione el documento.");
  const fd = new FormData(ev.target);
  st.partes.forEach((f) => fd.append("archivos", f, f.name));
  const m = await api("/api/normativa", {method: "POST", body: fd});
  const adv = m.advertencias || [];
  aviso(`${m.titulo}: ${m.fragmentos} fragmentos (${m.estructura.toLowerCase()}s) de ${m.archivos.length} archivo(s).` + (adv.length ? " ⚠ " + adv.join(" ") : ""), adv.length > 0);
  ev.target.reset(); st.partes = []; pintarPartes(); cargarNormativa();
}); };

// Borrador de opinión: el revisor lo edita aquí; al aprobar se envía como borrador final.
function pintarBorrador(d, cid, pendiente) {
  const caja = $("#p-borrador"), ed = d.borrador_editado;
  if (!d.borrador) { caja.innerHTML = `<div class="caja"><p class="vacio">Aún no hay borrador: valide primero el cliente en este periodo.</p></div>`; return; }
  const texto = ed ? ed.texto : d.borrador;
  pendiente = pendiente && puede("revisor");
  caja.innerHTML = `
    ${ed?.desactualizado ? `<div class="caja alerta"><b>El informe se volvió a validar después de esta edición.</b> Las cifras del borrador generado pueden haber cambiado.
      Compare con el borrador actual y guarde de nuevo, o descarte la edición.</div>` : ""}
    <div class="caja">
      <div class="cab-mini"><h2>${ed ? "Borrador editado" : "Borrador generado por el agente redactor"}</h2>
        <span class="nota">${ed ? `Guardado ${esc(ed.guardado || "")}${ed.por ? " por " + esc(ed.por) : ""}` : "Sin ediciones"}</span></div>
      <p class="nota">Edite la redacción, el nivel de riesgo o la recomendación. <b>Conserve los títulos «## N. …»</b>: la mejora continua compara cada sección con la generada.
        Borrar una sección completa la registra como descartada. Las cifras salen del Excel: si una no cuadra, corrija el Excel y vuelva a validar.</p>
      <div class="editor">
        <textarea id="b-texto" spellcheck="true" ${pendiente ? "" : "readonly"}>${esc(texto)}</textarea>
        <div class="md vista-previa" id="b-previa"></div>
      </div>
      <div class="acciones">
        ${ed ? `<button class="btn sec" id="b-original" type="button">Ver el generado</button><button class="btn sec" id="b-descartar" type="button">Descartar edición</button>` : ""}
        ${pendiente ? `<button class="btn" id="b-guardar" type="button">Guardar borrador</button>` : `<span class="nota">Solo se edita un informe pendiente de validación.</span>`}
      </div>
    </div>`;
  const ta = $("#b-texto"), previa = () => ($("#b-previa").innerHTML = md(ta.value));
  ta.oninput = previa; previa();
  const titulos = (t) => (t.match(/^##\s+\d+\.\s+.*$/gm) || []).map((x) => x.replace(/^##\s+\d+\.\s+/, "").trim());
  if ($("#b-guardar")) $("#b-guardar").onclick = (ev) => ocupado(ev.target, async () => {
    const faltan = titulos(d.borrador).filter((t) => !titulos(ta.value).includes(t));
    if (faltan.length && !confirm(`Estas secciones no aparecen y quedarán como descartadas:\n\n• ${faltan.join("\n• ")}\n\n¿Guardar de todos modos?`)) return;
    await api("/api/borrador", {method: "PUT", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({periodo: periodo(), cliente: cid, texto: ta.value})});
    aviso("Borrador guardado. Se enviará al aprobar el informe."); cargarCliente();
  });
  if ($("#b-descartar")) $("#b-descartar").onclick = (ev) => ocupado(ev.target, async () => {
    if (!confirm("¿Descartar la edición y volver al borrador generado?")) return;
    await api(`/api/borrador/${periodo()}/${cid}`, {method: "DELETE"}); aviso("Edición descartada."); cargarCliente();
  });
  if ($("#b-original")) $("#b-original").onclick = () => dialogo("Borrador generado por el agente", `<div class="md dlg-alto">${md(d.borrador)}</div>`, async () => {});
}

// Mapa del Excel: qué hoja del cliente alimenta cada sección del informe estándar. La IA propone, el revisor aprueba.
const ESTADO_SECCION = {ok: ["✔ completa", "m-ok"], vacia: ["vacía este mes", "m-gris"], falta: ["sin hoja", "m-gris"],
                        incompleta: ["incompleta", "m-warn"]};
// Agente adaptador: propone, para cada sección que no se pudo leer, la receta que lee ESTE Excel. El revisor aprueba.
const ESTADO_ADAPT = {ok: "completa", falta: "no se pudo leer", vacia: "vacía", incompleta: "incompleta"};
async function pintarAdaptacion(cid, caja) {
  let ad;
  try { ad = await api(`/api/adaptacion/${periodo()}/${cid}`); } catch (e) { caja.innerHTML = ""; return; }
  const R = ad.resumen;
  const filas = (ad.secciones || []).filter((f) => f.estado !== "ok" || f.propuesta || f.ajuste_mapa);
  const pendiente = (f) => (f.propuesta && !f.receta_aprobada) || (f.ajuste_mapa && !f.mapa_aprobado);
  const ajuste = (f) => f.ajuste_mapa ? `<div><span class="m-ok">Corregir el mapa:</span> «${esc(f.ajuste_mapa.clave)}» → hoja «${esc(f.ajuste_mapa.ahora)}»
    <div class="m-gris">${esc(f.ajuste_mapa.motivo)}</div></div>` : "";
  const muestra = (pr) => {
    const m = (pr && pr.muestra) || [];
    if (!m.length) return "";
    const cols = Object.keys(m[0]).filter((k) => k !== "fuente").slice(0, 5);
    return `<details><summary>Ver lo que leyó (${esc(pr.n)} fila(s))</summary><table class="mini"><tr>${cols.map((c) => `<th>${esc(c)}</th>`).join("")}<th>celda</th></tr>
      ${m.map((x) => `<tr>${cols.map((c) => `<td>${esc(typeof x[c] === "number" ? x[c].toLocaleString("es-CO") : x[c] ?? "")}</td>`).join("")}<td class="m-gris">${esc(x.fuente || "")}</td></tr>`).join("")}</table></details>`;
  };
  caja.innerHTML = `<div class="caja"><h2>Adaptación al Excel del cliente</h2>
    <p class="nota">El informe en PowerPoint es siempre el mismo. El <b>agente adaptador</b> revisa las secciones que no se pudieron leer
      y propone una <b>receta</b>: dónde está cada dato en <i>este</i> Excel. Python lee y calcula las cifras; usted revisa lo leído y aprueba.
      Lo que no tenga receta sale en el informe como «información pendiente», en su lugar.</p>
    ${R ? `<p><b>${R.completas} de ${R.secciones}</b> secciones completas con lo aprobado hoy${R.propuestas || R.ajustes_mapa ? ` · <b>${R.completas_con_propuestas}</b> si aprueba ${[R.propuestas ? `las ${R.propuestas} receta(s)` : "", R.ajustes_mapa ? `las ${R.ajustes_mapa} corrección(es) del mapa` : ""].filter(Boolean).join(" y ")} propuesta(s)` : ""}.
      <span class="m-gris">Revisado ${esc((ad.fecha || "").replace("T", " "))}${ad.ia ? " · con IA" : ""}.</span></p>` : `<p class="vacio">Aún no se ha revisado este Excel.</p>`}
    ${filas.length ? `<table class="mapa"><tr><th>Sección</th><th>Hoy</th><th>Propuesta del agente</th></tr>
      ${filas.map((f) => `<tr><td>${esc(f.descripcion)}${f.obligatoria ? " <b title='Obligatoria'>*</b>" : ""}</td>
        <td>${esc(ESTADO_ADAPT[f.estado] || f.estado)}</td>
        <td>${!pendiente(f) && (f.propuesta || f.ajuste_mapa) ? `<span class="m-ok">✔ ${f.propuesta ? "receta" : "corrección del mapa"} aprobada${f.aprobada_por ? " por " + esc(f.aprobada_por) : ""}</span>`
          : f.ajuste_mapa && !f.propuesta ? `<label class="ad-sel">${puede("revisor") ? `<input type="checkbox" data-adapt="${esc(f.seccion)}" checked>` : ""}${ajuste(f)}</label>
            ${f.estado_con_propuestas !== "ok" && f.motivo ? `<div class="m-warn">⚠ Aun corregida: ${esc(f.motivo)}</div>` : ""}`
          : f.propuesta ? `${ajuste(f)}<label class="ad-sel">${puede("revisor") ? `<input type="checkbox" data-adapt="${esc(f.seccion)}" checked>` : ""}<span><span class="m-ok">Receta ${esc(f.origen)}</span> en «${esc(f.propuesta.hoja)}»</span></label>
            ${(f.prueba && f.prueba.avisos || []).map((a) => `<div class="m-warn">⚠ ${esc(a)}</div>`).join("")}${muestra(f.prueba)}`
          : `<span class="m-gris">${esc(f.motivo || f.detalle || "")}</span>`}</td></tr>`).join("")}</table>
      <p class="nota">* obligatoria.</p>` : ""}
    <div class="acciones">
      ${puede("asistente", "revisor") ? `<button class="btn sec" id="ad-buscar">Revisar el Excel</button><button class="btn sec" id="ad-ia" title="Usa el proveedor de IA configurado para las secciones que la receta automática no resuelve">Revisar con IA</button>` : ""}
      ${puede("revisor") && filas.some(pendiente) ? `<button class="btn" id="ad-aprobar">Aprobar seleccionadas y volver a validar</button>` : ""}
      ${!puede("revisor") && filas.some(pendiente) ? `<span class="nota">Solo un revisor aprueba las recetas y las correcciones del mapa.</span>` : ""}
    </div></div>`;
  const correr = (ia) => async (ev) => ocupado(ev.target, async () => {
    const r = await post(`/api/adaptacion/${periodo()}/${cid}`, {ia});
    const n = r.resumen.propuestas + (r.resumen.ajustes_mapa || 0);
    aviso(n ? `${n} propuesta(s) (recetas o correcciones del mapa): revise lo leído y apruébelas.` : "No hay propuestas nuevas para este Excel.");
    pintarAdaptacion(cid, caja);
  });
  if ($("#ad-buscar")) $("#ad-buscar").onclick = correr(false);
  if ($("#ad-ia")) $("#ad-ia").onclick = correr(true);
  if ($("#ad-aprobar")) $("#ad-aprobar").onclick = (ev) => ocupado(ev.target, async () => {
    const secciones = [...caja.querySelectorAll("input[data-adapt]:checked")].map((i) => i.dataset.adapt);
    if (!secciones.length) return aviso("Marque al menos una propuesta.", true);
    await api(`/api/adaptacion/${periodo()}/${cid}`, {method: "PUT", headers: {"Content-Type": "application/json"}, body: JSON.stringify({secciones})});
    const r = await post("/api/validar", {periodo: periodo(), clientes: [cid]});
    aviso(`${secciones.length} propuesta(s) aprobada(s) y cliente validado: ` + (ESTADOS[r.estados[cid]] || r.estados[cid]) + ".");
    st.pestana = "mapa"; return cargarCliente();
  });
}

async function pintarMapa(cid) {
  const raiz = $("#p-mapa");
  raiz.innerHTML = `<div id="p-adapt"></div><div id="p-mapa-in"></div>`;
  pintarAdaptacion(cid, $("#p-adapt"));
  const caja = $("#p-mapa-in");
  caja.innerHTML = `<p class="vacio">Leyendo el Excel del cliente…</p>`;
  let v;
  try { v = await api(`/api/mapa/${periodo()}/${cid}`); } catch (e) { caja.innerHTML = `<p class="vacio">${esc(e.message)}</p>`; return; }
  const titulos = v.hojas_excel.map((h) => h.titulo);
  let prueba = v.prueba, propuestas = [], editor = null;   // editor: {seccion, hoja, resultado, receta, prueba}
  const opciones = (actual, lista) => (lista ? "" : `<option value="">— el cliente no tiene esta hoja —</option>`)
    + titulos.map((t) => `<option value="${esc(t)}" ${(lista ? (actual || []).includes(t) : actual === t) ? "selected" : ""}>${esc(t)}</option>`).join("");
  const estadoClave = (c) => {
    const fmt = (prueba.formato || []).filter((f) => f.clave === c.clave);
    if (fmt.length) return `<span class="m-warn" title="${esc(fmt.map((f) => f.detalle).join(" "))}">⚠ formato distinto</span>`;
    const est = (prueba.cobertura || []).filter((x) => c.secciones.some((s) => s.clave === x.seccion));
    if (!est.length) return `<span class="m-gris">—</span>`;
    return est.map((x) => { const [t, k] = ESTADO_SECCION[x.estado] || [x.estado, ""]; return `<span class="${k}" title="${esc(x.detalle)}">${t}</span>`; }).join("<br>");
  };
  const leer = () => Object.fromEntries([
    ...[...caja.querySelectorAll("select[data-clave]")].map((s) => [s.dataset.clave, s.value || null]),
    ...[...caja.querySelectorAll("[data-lista]")].map((d) => [d.dataset.lista, [...d.querySelectorAll("input:checked")].map((i) => i.value)])]);
  // ---- subagente estructurador: receta de lectura para una sección con otra estructura
  const letras = (x) => (Array.isArray(x) ? x.join("+") : x || "");
  const cajaEstructura = () => {
    const def = v.campos_receta[editor.seccion], rec = editor.receta || {}, cols = rec.columnas || {}, pr = editor.prueba, res = editor.resultado;
    const claves = pr && pr.muestra && pr.muestra.length ? Object.keys(pr.muestra[0]).filter((k) => k !== "fuente") : [];
    return `<div class="caja" id="e-caja"><h2>Otra estructura · ${esc(def.descripcion)}</h2>
      <p class="nota">El subagente <b>estructurador</b> mira la hoja y escribe una <b>receta</b>: dónde está cada campo estándar. El sistema la ejecuta,
        la prueba y, si falla, le devuelve los errores para que se corrija. No inventa datos; usted aprueba lo que se lee.</p>
      <div class="rejilla"><label>Hoja del Excel<select id="e-hoja">${titulos.map((t) => `<option value="${esc(t)}" ${t === editor.hoja ? "selected" : ""}>${esc(t)}</option>`).join("")}</select></label></div>
      <div class="acciones">${puede("asistente", "revisor") ? `<button type="button" class="btn sec" id="e-proponer">Proponer con el estructurador</button>` : ""}
        ${(v.lectores || {})[editor.seccion] && puede("revisor") ? `<button type="button" class="btn vino" id="e-quitar">Quitar receta</button>` : ""}
        <button type="button" class="btn sec" id="e-cerrar">Cerrar</button></div>
      ${res && !res.corresponde ? `<p class="m-warn">La hoja no contiene esta sección: ${esc(res.motivo)}</p>` : ""}
      ${res && res.intentos ? `<p class="nota">${res.intentos.map((i) => `Intento ${i.intento}: ${i.corresponde === false ? "no corresponde" : i.errores && i.errores.length ? "✖ " + esc(i.errores.join("; ")) : "✔ sin errores"}`).join(" · ")}${res.modelo ? ` · ${esc(res.modelo)}` : ""}</p>` : ""}
      ${res && res.motivo && res.corresponde ? `<p class="nota">${esc(res.motivo)}</p>` : ""}
      ${editor.receta ? `<h2 class="mt">Receta</h2><div class="rejilla">
          <label title="Plantilla del título del bloque del mes, con {mes} y {anio}">Título del bloque del mes<input id="e-titulo" value="${esc(rec.titulo || "")}" placeholder="cuentas por pagar a {mes} {anio}"></label>
          <label title="Solo si la hoja tiene una tabla única">Fila de encabezado<input id="e-fila" type="number" min="1" value="${esc(rec.titulo ? "" : rec.fila_encabezado ?? "")}"></label>
          ${def.campos.map((c) => `<label title="${esc(c.que_es)}">${esc(c.campo)}${c.obligatorio ? " *" : ""}<input data-campo="${esc(c.campo)}" value="${esc(letras(cols[c.campo]))}" placeholder="letra (p. ej. D o E+F)"></label>`).join("")}
          ${(def.parametros || []).map((x) => `<label title="${esc(x.que_es)}">${esc(x.nombre.replace("_", " "))}${x.obligatorio ? " *" : ""}<input data-param="${esc(x.nombre)}" value="${esc(rec[x.nombre] ?? "")}" placeholder="${esc(x.que_es.slice(0, 40))}"></label>`).join("")}
          ${def.matriz ? `<label title="Columnas de nivel de cumplimiento por periodo">Niveles<input id="e-niveles" value="${esc((rec.niveles || []).join("+"))}" placeholder="E+F"></label>` : ""}
        </div>` : ""}
      ${pr ? `${(pr.errores || []).map((x) => `<p class="m-warn">✖ ${esc(x)}</p>`).join("")}${(pr.avisos || []).map((x) => `<p class="nota">⚠ ${esc(x)}</p>`).join("")}
        <p class="nota">${pr.n || 0} fila(s) leída(s). Primeras filas:</p>
        ${claves.length ? `<div class="tabla-scroll"><table><tr>${claves.map((k) => `<th>${esc(k)}</th>`).join("")}<th>fuente</th></tr>
          ${pr.muestra.map((f) => `<tr>${claves.map((k) => `<td>${esc(typeof f[k] === "object" ? JSON.stringify(f[k]) : f[k]).slice(0, 140)}</td>`).join("")}<td class="m-gris">${esc(f.fuente || "")}</td></tr>`).join("")}</table></div>` : ""}` : ""}
      ${editor.receta ? `<div class="acciones">${puede("asistente", "revisor") ? `<button type="button" class="btn sec" id="e-probar">Probar receta</button>` : ""}
        ${puede("revisor") ? `<button type="button" class="btn" id="e-aprobar" ${pr && pr.ok ? "" : "disabled title='Primero pruebe la receta sin errores'"}>Aprobar receta</button>` : `<span class="nota">Solo un revisor aprueba la receta.</span>`}</div>` : ""}
    </div>`;
  };
  const recetaDelFormulario = () => {
    const r = {seccion: editor.seccion, hoja: $("#e-hoja").value, columnas: {}};
    const t = $("#e-titulo").value.trim(), f = $("#e-fila").value.trim();
    if (t) r.titulo = t; else if (f) r.fila_encabezado = +f;
    caja.querySelectorAll("[data-campo]").forEach((i) => {
      const val = i.value.trim();
      if (/^\+\d+$/.test(val)) { r.columnas[i.dataset.campo] = val; return; }   // columna relativa al título (+2)
      const xs = val.split(/[+,\s]+/).map((x) => x.trim().toUpperCase()).filter(Boolean);
      if (xs.length) r.columnas[i.dataset.campo] = xs.length === 1 ? xs[0] : xs;
    });
    caja.querySelectorAll("[data-param]").forEach((i) => {
      const val = i.value.trim();
      if (val) r[i.dataset.param] = i.dataset.param === "signo_total" ? +val : val;
    });
    if ($("#e-niveles")) { const n = $("#e-niveles").value.split(/[+,\s]+/).map((x) => x.trim().toUpperCase()).filter(Boolean); if (n.length) r.niveles = n; }
    return r;
  };
  const enlazarEstructura = () => {
    caja.querySelectorAll("[data-estructura]").forEach((b) => b.onclick = async () => {
      const sec = b.dataset.estructura, existente = (v.lectores || {})[sec];
      const clave = v.claves.find((c) => c.secciones.some((x) => x.clave === sec));
      const sel = caja.querySelector(`select[data-clave="${CSS.escape(clave.clave)}"]`);
      const hoja = existente ? existente.hoja : (sel && sel.value) || titulos[0];
      editor = {seccion: sec, hoja, receta: existente || {seccion: sec, hoja, columnas: {}}, prueba: null, resultado: null};
      if (existente) editor.prueba = await post(`/api/receta/${periodo()}/${cid}/probar`, {receta: existente});
      pintar(); $("#e-caja").scrollIntoView({behavior: "smooth", block: "start"});
    });
    if (!editor) return;
    $("#e-cerrar").onclick = () => { editor = null; pintar(); };
    $("#e-hoja").onchange = (ev) => { editor.hoja = ev.target.value; };
    if ($("#e-proponer")) $("#e-proponer").onclick = (ev) => ocupado(ev.target, async () => {
      editor.hoja = $("#e-hoja").value;
      const r = await post(`/api/receta/${periodo()}/${cid}/proponer`, {seccion: editor.seccion, hoja: editor.hoja});
      editor.resultado = r; editor.receta = r.receta || null; editor.prueba = r.prueba || null; pintar();
      aviso(!r.corresponde ? "El estructurador indica que esa hoja no contiene la sección." : r.prueba.ok ? `Receta lista en ${r.intentos.length} intento(s). Revise lo leído y apruébela.` : "La receta aún tiene errores: corríjala a mano o pruebe otra hoja.", !r.corresponde || !r.prueba.ok);
    });
    if ($("#e-probar")) $("#e-probar").onclick = (ev) => ocupado(ev.target, async () => {
      editor.receta = recetaDelFormulario(); editor.hoja = editor.receta.hoja;
      editor.prueba = await post(`/api/receta/${periodo()}/${cid}/probar`, {receta: editor.receta}); pintar();
    });
    if ($("#e-aprobar")) $("#e-aprobar").onclick = (ev) => ocupado(ev.target, async () => {
      const receta = recetaDelFormulario();
      await api(`/api/receta/${periodo()}/${cid}`, {method: "PUT", headers: {"Content-Type": "application/json"}, body: JSON.stringify({receta})});
      aviso("Receta aprobada: queda en el mapa del cliente y como ejemplo para el estructurador. Vuelva a validar el cliente.");
      editor = null; v = await api(`/api/mapa/${periodo()}/${cid}`); prueba = v.prueba; caja.innerHTML = ""; pintar();
    });
    if ($("#e-quitar")) $("#e-quitar").onclick = (ev) => ocupado(ev.target, async () => {
      await api(`/api/receta/${periodo()}/${cid}/${editor.seccion}`, {method: "DELETE"});
      aviso("Receta quitada: la sección vuelve al lector estándar.");
      editor = null; v = await api(`/api/mapa/${periodo()}/${cid}`); prueba = v.prueba; caja.innerHTML = ""; pintar();
    });
  };
  const pintar = () => {
    const actual = caja.querySelector("select[data-clave]") ? leer() : null;
    const valor = (c) => actual ? actual[c.clave] : c.hoja;
    caja.innerHTML = `
      <div class="caja"><h2>Mapa del Excel · ${esc(v.excel)}</h2>
        <p class="nota">El informe es siempre el mismo; lo único que cambia entre clientes es <b>en qué hoja</b> está cada dato.
          ${v.guardado && v.aprobado_por ? `Mapa aprobado por <b>${esc(v.aprobado_por)}</b> (${esc((v.fecha || "").replace("T", " "))}).`
            : `<b>Mapa propuesto automáticamente, pendiente de aprobación del revisor.</b>`}</p>
        ${Object.entries(v.ajustes || {}).map(([k, a]) => a.ahora
          ? `<p class="m-warn">↻ «${esc(k)}»: la hoja «${esc([].concat(a.antes).flat().filter((x) => x && typeof x === "string").join(", "))}» ya no está; se propone «${esc(a.ahora)}» (${esc(a.motivo)}). Confírmelo aprobando el mapa.</p>`
          : `<p class="m-warn">⚠ «${esc(k)}»: ${esc(a.motivo)}. Elija la hoja a mano o deje «no tiene».</p>`).join("")}
        ${(v.hojas_nuevas || []).length ? `<p class="m-warn">＋ Hojas nuevas desde la última aprobación: ${v.hojas_nuevas.map(esc).join(", ")}.</p>` : ""}
        ${prueba.ok ? "" : `<p class="m-warn">Con este mapa la extracción falla: ${esc(prueba.error)}</p>`}
        ${(prueba.desajustes || []).map((d) => `<p class="m-warn">⚠ ${esc(d)}</p>`).join("")}
        <table class="mapa"><tr><th>Sección del informe</th><th>Hoja del Excel del cliente</th><th>Resultado</th></tr>
        ${v.claves.map((c) => `<tr><td>${c.secciones.map((s) => esc(s.descripcion) + (s.obligatoria ? " <b title='Obligatoria'>*</b>" : "")).join("<br>") || esc(c.ayuda)}
            <div class="m-gris">${esc(c.clave)}</div></td>
          <td>${c.lista ? `<div class="m-lista" data-lista="${esc(c.clave)}">${titulos.map((t) => `<label><input type="checkbox" value="${esc(t)}" ${(valor(c) || []).includes(t) ? "checked" : ""}> ${esc(t)}</label>`).join("")}</div>`
            : `<select data-clave="${esc(c.clave)}">${opciones(valor(c), false)}</select>`}</td>
          <td>${estadoClave(c)}${c.secciones.filter((x) => v.campos_receta[x.clave]).map((x) => (v.lectores || {})[x.clave]
              ? `<div><span class="m-ok" title="Se lee con una receta aprobada">📐 receta${c.secciones.length > 1 ? " · " + esc(x.descripcion) : ""}</span> <button type="button" class="btn sec chico" data-estructura="${esc(x.clave)}">Ver</button></div>`
              : puede("asistente", "revisor") ? `<div><button type="button" class="btn sec chico" data-estructura="${esc(x.clave)}" title="El Excel trae esta sección con otra organización">Otra estructura…${c.secciones.length > 1 ? " (" + esc(x.descripcion.split(" ")[0]) + ")" : ""}</button></div>` : "").join("")}</td></tr>`).join("")}</table>
        <p class="nota">* obligatoria. «Vacía este mes» significa que la hoja existe pero la tabla no trae filas: es un dato que falta en el Excel, no un error del mapa.</p>
        <div class="acciones">
          ${puede("asistente", "revisor") ? `<button class="btn sec" id="m-proponer">Proponer con IA</button><button class="btn sec" id="m-probar">Probar mapa</button>` : ""}
          ${puede("revisor") ? `<button class="btn" id="m-guardar">Aprobar mapa</button><button class="btn" id="m-validar">Aprobar y volver a validar</button>`
            : `<span class="nota">Solo un revisor aprueba el mapa.</span>`}
        </div></div>
      ${editor ? cajaEstructura() : ""}
      ${propuestas.length ? `<div class="caja"><h2>Propuestas</h2>${propuestas.map((p, i) => `<div class="hallazgo ${p.formato_ok ? "INFO" : "MEDIA"}">
          <div class="t">${esc(p.clave)} → «${esc(p.hoja)}» · confianza ${esc(p.confianza)}</div><div class="d">${esc(p.motivo)}</div>
          ${p.formato_ok ? "" : `<div class="m-warn">⚠ ${esc(p.formato)} Aplicarla dejaría la sección sin leer o mal leída.</div>`}
          <button class="btn chico ${p.formato_ok ? "" : "sec"}" data-aplicar="${i}">Aplicar al mapa</button></div>`).join("")}</div>` : ""}
      <div class="caja"><h2>Hojas del Excel</h2>${v.hojas_excel.map((h) => `<details><summary>${esc(h.titulo)} ${h.usada ? "" : "<span class='m-warn'>· sin usar</span>"}</summary>
        <pre class="vistazo">${esc(h.vistazo.join("\n")) || "(vacía)"}</pre></details>`).join("")}</div>`;
    const probar = async () => { prueba = await post(`/api/mapa/${periodo()}/${cid}/probar`, {hojas: leer()}); pintar(); };
    enlazarEstructura();
    if ($("#m-probar")) $("#m-probar").onclick = (ev) => ocupado(ev.target, async () => { await probar(); aviso("Mapa probado con el Excel actual."); });
    if ($("#m-proponer")) $("#m-proponer").onclick = (ev) => ocupado(ev.target, async () => {
      const r = await post(`/api/mapa/${periodo()}/${cid}/proponer`, {});
      propuestas = r.propuestas; pintar();
      aviso((r.propuestas.length ? `${r.propuestas.length} propuesta(s)` : "Sin propuestas") + (r.modelo ? ` de ${r.modelo}` : "") + (r.nota ? ". " + r.nota : ""));
    });
    caja.querySelectorAll("[data-aplicar]").forEach((b) => b.onclick = () => ocupado(b, async () => {
      const p = propuestas[+b.dataset.aplicar], s = caja.querySelector(`select[data-clave="${CSS.escape(p.clave)}"]`);
      if (s) s.value = p.hoja;
      else caja.querySelectorAll(`[data-lista="${CSS.escape(p.clave)}"] input`).forEach((i) => { if (i.value === p.hoja) i.checked = true; });
      propuestas.splice(+b.dataset.aplicar, 1);
      await probar();
    }));
    const guardar = async (validar) => {
      const g = await api(`/api/mapa/${periodo()}/${cid}`, {method: "PUT", headers: {"Content-Type": "application/json"}, body: JSON.stringify({hojas: leer()})});
      const aprendio = (g.mapa.aprendido || []).length ? " Aprendió para otros clientes: " + g.mapa.aprendido.join("; ") + "." : "";
      if (validar) {
        const r = await post("/api/validar", {periodo: periodo(), clientes: [cid]});
        aviso("Mapa aprobado y cliente validado: " + (ESTADOS[r.estados[cid]] || r.estados[cid]) + "." + aprendio);
        st.pestana = "mapa"; return cargarCliente();
      }
      aviso("Mapa aprobado. La versión anterior quedó en clientes/" + cid + "/historial/." + aprendio);
      v = await api(`/api/mapa/${periodo()}/${cid}`); prueba = v.prueba; caja.innerHTML = ""; pintar();
    };
    if ($("#m-guardar")) $("#m-guardar").onclick = (ev) => ocupado(ev.target, () => guardar(false));
    if ($("#m-validar")) $("#m-validar").onclick = (ev) => ocupado(ev.target, () => guardar(true));
  };
  caja.innerHTML = ""; pintar();
}

// Ficha del cliente: datos permanentes y notas del revisor del mes.
function pintarFicha(d, cid, estado) {
  const f = d.ficha, n = f.notas || {}, pres = Object.entries(n.presupuesto || {});
  const fila = (k = "", v = "") => `<div class="fila-pres"><input class="p-rubro" placeholder="Rubro tal como está en la hoja PRESUPUESTO" value="${esc(k)}">
    <textarea class="p-texto" rows="2" placeholder="Observación del revisor">${esc(v)}</textarea><button type="button" class="btn sec chico p-quitar" title="Quitar">✕</button></div>`;
  $("#p-ficha").innerHTML = `
    <form class="caja" id="f-ficha">
      <h2>Datos de la copropiedad</h2>
      <div class="rejilla">
        <label class="ancho">Copropiedad<input name="copropiedad" required value="${esc(f.copropiedad || "")}"></label>
        <label>NIT${f.nit_provisional ? ` <b class="sev ALTA">PROVISIONAL</b>` : ""}<input name="nit" value="${esc(f.nit || "")}" placeholder="900.000.000-0" ${/^\d{10}$/.test(cid) ? "readonly" : ""}>
          ${/^\d{10}$/.test(cid) && puede("revisor") ? `<button type="button" class="btn sec chico" id="f-cambiar-nit">Cambiar NIT…</button>` : ""}</label>
        <label>Unidades<input name="unidades" type="number" min="0" value="${esc(f.unidades ?? "")}"></label>
        <label>Número de informe<input name="numero_informe" value="${esc(f.numero_informe || "")}" placeholder="007-2026"></label>
        <label>Revisor fiscal<input name="revisor" value="${esc(f.revisor || "")}"></label>
        <label>Firma<input name="firma" value="${esc(f.firma || "")}"></label>
        <label>Ciudad<input name="ciudad" value="${esc(f.ciudad || "")}" placeholder="Para la fecha de los oficios"></label>
        <label>Orden en el calendario<input name="orden" type="number" min="0" max="999" value="${esc(f.orden ?? 99)}"></label>
        <label>Software contable<input name="software" pattern="[a-z0-9_]+" value="${esc(f.software || "generico")}"></label>
        <label>Tipo de informe<select name="informe">${["completo", "financiero"].map((t) => `<option ${((f.informe || "completo") === t) ? "selected" : ""}>${t}</option>`).join("")}</select></label>
      </div>
      <h2 class="mt">Notas del revisor fiscal para el mes</h2>
      <p class="nota">Son palabras del revisor: el generador las lleva tal cual al informe y nunca las inventa. Si una nota queda vacía, la diapositiva sale sin ella.</p>
      <label>Gestión de cobro de cartera (diapositiva de concentración de cartera)<textarea name="concentracion_cartera" rows="4">${esc(n.concentracion_cartera || "")}</textarea></label>
      <label>Opinión sobre anticipos<textarea name="opinion_anticipos" rows="5">${esc(n.opinion_anticipos || "")}</textarea></label>
      <div class="cab-mini"><h2 class="mt">Observaciones a rubros sobreejecutados</h2><button type="button" class="btn sec chico" id="p-agregar">+ Rubro</button></div>
      <div id="p-filas">${pres.map(([k, v]) => fila(k, v)).join("") || ""}</div>
      <div class="acciones">
        <span class="nota" id="f-estado">${estado === "APROBADO" ? "El informe de este periodo ya está aprobado: los cambios aplican desde el próximo." : "Los cambios se reflejan en el informe al volver a validar."}</span>
        ${puede("revisor") ? `<button class="btn sec" type="submit" value="guardar">Guardar ficha</button>` : `<span class="nota">Solo un revisor puede editar la ficha.</span>`}
        ${estado !== "APROBADO" && puede("asistente", "revisor") ? `<button class="btn" type="submit" value="validar">Guardar y volver a validar</button>` : ""}
      </div>
    </form>
    <div class="dos">
      <div class="caja"><div class="cab-mini"><h2>Documentos del cliente</h2><span class="nota" id="d-cuenta"></span></div>
        <p class="nota">Reglamento, actas, contratos, pólizas, manuales… Quedan en la carpeta del cliente y Auditin los ve.</p>
        ${puede("asistente", "revisor") ? `<div class="zona" id="zona-docs"><input type="file" id="f-docs" multiple hidden><b>Subir documentos</b><span>Arrastre aquí o haga clic · PDF, Word, Excel, imágenes · hasta 60 MB</span></div>` : ""}
        <ul class="docs" id="d-lista"><li class="vacio">Cargando…</li></ul></div>
      <form class="caja" id="f-norma-cli"><h2>Ficha de normatividad</h2>
        <p class="nota">Lo que rige a esta copropiedad. La escribe el revisor; el sistema no redacta normas de memoria.</p>
        <label>Reglamento de propiedad horizontal<input name="reglamento" placeholder="Escritura, fecha y notaría del reglamento vigente"></label>
        <label>Escritura de constitución<input name="escritura"></label>
        <label>Marco contable<input name="marco_contable" placeholder="Ej.: NIIF para microempresas (grupo 3)"></label>
        <label>Normas aplicables (una por línea)<textarea name="normas" rows="4" placeholder="Ley 675 de 2001&#10;…"></textarea></label>
        <label>Manuales internos (uno por línea)<textarea name="manuales" rows="3" placeholder="Manual de contratación&#10;Manual de convivencia"></textarea></label>
        <label>Obligaciones especiales<textarea name="obligaciones" rows="2"></textarea></label>
        <label>Observaciones<textarea name="observaciones" rows="2"></textarea></label>
        <div class="acciones"><span class="nota" id="nc-estado"></span>${puede("revisor") ? `<button class="btn sec">Guardar ficha de normatividad</button>` : ""}</div>
      </form>
    </div>`;
  pintarDocumentos(cid); pintarNormatividad(cid);
  const filas = $("#p-filas");
  const enlazar = () => filas.querySelectorAll(".p-quitar").forEach((b) => b.onclick = () => b.parentElement.remove());
  $("#p-agregar").onclick = () => { filas.insertAdjacentHTML("beforeend", fila()); enlazar(); };
  enlazar();
  if ($("#f-cambiar-nit")) $("#f-cambiar-nit").onclick = () => dialogo("Cambiar NIT", `
      <p class="nota">El NIT identifica al cliente: se moverán su ficha, su mapa y todos sus expedientes al NIT nuevo, conservando el historial.
        No se permite si ya hay informes aprobados con el NIT actual. Después hay que volver a validar.</p>
      <label>NIT real (RUT de la copropiedad)<input name="nit" required placeholder="900.000.000-0"></label>`,
    async (datos) => {
      const r = await post(`/api/clientes/${cid}/nit`, {nit: datos.nit});
      aviso("NIT cambiado a " + r.nit + ". Vuelva a validar el cliente.");
      st.resumen = await api("/api/resumen");
      st.cliente = r.id; location.hash = "#cliente/" + r.id; cargarCliente();
    });
  $("#f-ficha").onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
    const F = Object.fromEntries(new FormData(ev.target));
    const presupuesto = {};
    filas.querySelectorAll(".fila-pres").forEach((r) => { const k = r.querySelector(".p-rubro").value.trim(), v = r.querySelector(".p-texto").value.trim(); if (k && v) presupuesto[k] = v; });
    const cuerpo = {copropiedad: F.copropiedad, nit: F.nit, unidades: F.unidades === "" ? null : +F.unidades, numero_informe: F.numero_informe,
      revisor: F.revisor, firma: F.firma, ciudad: F.ciudad, orden: +F.orden || 0, software: F.software || "generico", informe: F.informe,
      notas: {...n, concentracion_cartera: F.concentracion_cartera, opinion_anticipos: F.opinion_anticipos, presupuesto}};
    await api(`/api/ficha/${cid}`, {method: "PUT", headers: {"Content-Type": "application/json"}, body: JSON.stringify(cuerpo)});
    if (ev.submitter?.value === "validar") {
      const v = await post("/api/validar", {periodo: periodo(), clientes: [cid]});
      aviso("Ficha guardada y cliente validado: " + (ESTADOS[v.estados[cid]] || v.estados[cid]));
    } else aviso("Ficha guardada. La versión anterior quedó en clientes/" + cid + "/historial/.");
    cargarCliente();
  }); };
}

// Documentos y ficha de normatividad del cliente
const tamano = (b) => b > 1048576 ? (b / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(b / 1024)) + " KB";
async function pintarDocumentos(cid) {
  const D = await api(`/api/clientes/${encodeURIComponent(cid)}/documentos`);
  $("#d-cuenta").textContent = D.documentos.length ? `${D.documentos.length} documento(s)` : "";
  $("#d-lista").innerHTML = D.documentos.map((d) => `<li><a href="/api/clientes/${encodeURIComponent(cid)}/documentos/${encodeURIComponent(d.nombre)}">${esc(d.nombre)}</a>
      <small class="nota">${tamano(d.bytes)} · ${esc(d.fecha.replace("T", " "))}</small>
      ${puede("revisor") ? `<button type="button" class="btn sec chico" data-quitar="${esc(d.nombre)}" title="Pasa al historial del cliente">Retirar</button>` : ""}</li>`).join("")
    || `<li class="vacio">Aún no hay documentos para este cliente.</li>`;
  $("#d-lista").querySelectorAll("[data-quitar]").forEach((b) => b.onclick = async () => {
    await api(`/api/clientes/${encodeURIComponent(cid)}/documentos/${encodeURIComponent(b.dataset.quitar)}`, {method: "DELETE"});
    aviso("Documento retirado (quedó en el historial del cliente)."); pintarDocumentos(cid);
  });
  const zona = $("#zona-docs"), input = $("#f-docs");
  if (!zona) return;
  const subir = async (files) => {
    if (!files.length) return;
    const fd = new FormData(); [...files].forEach((f) => fd.append("archivos", f));
    try { const r = await api(`/api/clientes/${encodeURIComponent(cid)}/documentos`, {method: "POST", body: fd}); aviso(`${r.guardados.length} documento(s) guardado(s).`); }
    catch (e) { aviso(e.message, true); }
    pintarDocumentos(cid);
  };
  zona.onclick = () => input.click();
  input.onchange = () => { subir(input.files); input.value = ""; };
  zona.ondragover = (e) => { e.preventDefault(); zona.classList.add("sobre"); };
  zona.ondragleave = () => zona.classList.remove("sobre");
  zona.ondrop = (e) => { e.preventDefault(); zona.classList.remove("sobre"); subir(e.dataTransfer.files); };
}
async function pintarNormatividad(cid) {
  const N = await api(`/api/clientes/${encodeURIComponent(cid)}/normatividad`), f = $("#f-norma-cli");
  for (const k of ["reglamento", "escritura", "marco_contable", "obligaciones", "observaciones"]) f[k].value = N[k] || "";
  f.normas.value = (N.normas || []).join("\n"); f.manuales.value = (N.manuales || []).join("\n");
  [...f.elements].forEach((x) => { if (x.name) x.readOnly = !puede("revisor"); });
  $("#nc-estado").textContent = N.actualizada_por ? `Actualizada por ${N.actualizada_por} · ${(N.fecha || "").replace("T", " ")}` : "";
  f.onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
    const F = Object.fromEntries(new FormData(f));
    const lista = (t) => t.split("\n").map((x) => x.trim()).filter(Boolean);
    await api(`/api/clientes/${encodeURIComponent(cid)}/normatividad`, {method: "PUT", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({...F, normas: lista(F.normas), manuales: lista(F.manuales)})});
    aviso("Ficha de normatividad guardada."); pintarNormatividad(cid);
  }); };
}

function dialogo(titulo, html, alConfirmar) {
  const dlg = $("#dlg"), form = $("#dlg-form");
  $("#dlg-titulo").textContent = titulo; $("#dlg-cuerpo").innerHTML = html;
  form.onsubmit = async (ev) => {
    if (ev.submitter?.value !== "ok") return;
    ev.preventDefault();
    const datos = Object.fromEntries(new FormData(form));
    try { await alConfirmar(datos); dlg.close(); } catch (e) { aviso(e.message, true); }
  };
  dlg.showModal();
}

// ------------------------------------------------------------------ mejora
async function cargarMejora() {
  const M = await api("/api/mejora"), met = M.metricas || {};
  const fmt = (v, p) => v == null ? "—" : p ? Math.round(v * 100) + "%" : v;
  $("#m-kpis").innerHTML = [["Diapositivas sin corrección", fmt(met.pct_diapositivas_sin_correccion, 1)], ["Correcciones por informe", fmt(met.correcciones_por_informe)],
    ["Informes aprobados", fmt(met.aprobados)], ["Horas de validación", fmt(met.horas_validacion_promedio)],
    ["Regresión", M.regresion ? (M.regresion.ok ? `OK · ${M.regresion.n}` : "FALLÓ") : "—"],
    ["Mapas aprobados", met.mapas ? `${met.mapas.aprobados}/${met.mapas.clientes}` : "—"],
    ["Hojas aprendidas", fmt(met.mapas?.sinonimos)]]
    .map(([t, v]) => `<div class="kpi"><b>${esc(v)}</b><small>${t}</small></div>`).join("");
  const pend = M.propuestas.filter((p) => p.estado === "pendiente");
  $("#m-propuestas").innerHTML = pend.map((p) => `<div class="propuesta"><div class="id">${esc(p.id)}</div><b>${esc(p.titulo)}</b>
    <p class="nota">${esc(p.evidencia || "")} · riesgo ${esc(p.riesgo || "")}</p>
    ${p.plantilla ? `<div class="md"><blockquote>${esc(p.plantilla)}</blockquote></div>` : ""}
    ${p.archivo ? `<p class="nota">Borrador: <code>${esc(p.archivo)}</code></p>` : ""}
    <button class="btn chico" data-pid="${esc(p.id)}">Aprobar…</button></div>`).join("")
    || `<p class="vacio">Sin propuestas pendientes. Pulse «Analizar correcciones» después de aprobar informes.</p>`;
  document.querySelectorAll("#m-propuestas [data-pid]").forEach((b) => b.onclick = () => dialogo("Aprobar propuesta " + b.dataset.pid,
    `<p class="nota">Las plantillas de opinión se prueban contra los informes aprobados antes de publicarse; los demás tipos quedan registrados para aplicación manual.</p>
     <p class="nota">Queda registrada como aprobada por <b>${esc(st.yo.nombre)}</b>.</p>`,
    async () => { const r = await post(`/api/propuestas/${encodeURIComponent(b.dataset.pid)}/aprobar`, {}); aviso(r.detalle, !r.ok); cargarMejora(); }));
  const pv = met.por_version || [];
  $("#m-versiones").innerHTML = pv.length ? `<table><tr><th>Versión</th><th>Informes</th><th>Correcciones / informe</th><th>Opiniones editadas</th></tr>
    ${pv.map((v) => `<tr><td>${esc(v.version)}</td><td class="n">${v.informes}</td><td class="n">${v.correcciones_por_informe}</td><td class="n">${v.opiniones_editadas_por_informe}</td></tr>`).join("")}</table>
    <p class="nota">Con pocos informes por versión la comparación aún no es concluyente.</p>` : `<p class="vacio">Aún no hay informes aprobados.</p>`;
  $("#m-borradores").innerHTML = M.borradores.length ? M.borradores.map((b) => `<div><code>${esc(b)}</code></div>`).join("") : `<p class="vacio">Sin borradores.</p>`;
  pintarMuestra((await api(`/api/muestra/${periodo()}`)).muestra);
}
$("#btn-analizar").onclick = (ev) => ocupado(ev.target, async () => { await post("/api/mejora/analizar", {}); aviso("Propuestas y métricas actualizadas."); cargarMejora(); });
$("#btn-regresion").onclick = (ev) => ocupado(ev.target, async () => { const r = await post("/api/regresion", {}); aviso(`Regresión ${r.ok ? "OK" : "FALLÓ"}: ${r.n} caso(s).`, !r.ok); cargarMejora(); });
$("#btn-muestra").onclick = (ev) => ocupado(ev.target, async () => {
  const r = await post("/api/muestra", {periodo: periodo(), k: +$("#m-k").value || 4});
  pintarMuestra(r.muestra);
});

// ------------------------------------------------------------------ asistente
const BASES = {todo: "Todo", clientes: "Clientes", informes: "Informes", hallazgos: "Hallazgos", cifras: "Cifras del informe",
               excel: "Excel de papeles", acuerdos: "Acuerdos de reunión", mejora: "Mejora continua", biblioteca: "Biblioteca de opiniones",
               auditoria: "Bitácora y auditoría", normativa: "Base normativa"};
const SUGERENCIAS = {
  todo: ["Resume el estado del lote y qué debe decidir hoy la revisora", "¿Qué clientes tienen hallazgos ALTA y por qué?"],
  clientes: ["¿Qué sé de Arboleda Park: ficha, documentos y hojas de su Excel?", "¿Qué clientes no tienen revisor asignado?"],
  informes: ["¿Qué salió en el informe de Arboleda Park y qué quedó pendiente?", "¿Qué diapositivas usan propuestas sin aprobar?"],
  acuerdos: ["¿Qué se definió en la última reunión?", "¿Qué quedó pendiente por definir con el despacho?"],
  hallazgos: ["Explica los hallazgos ALTA de Picasso y qué celdas revisar", "¿Qué hallazgos se repiten entre clientes?"],
  cifras: ["¿Cuál es la cartera de Picasso y de qué celda sale?", "Muestra la concentración de cartera por rangos"],
  excel: ["¿Qué deudores tienen acuerdo de pago y cuál es su estado?", "Busca las notas sobre radicación de demandas"],
  mejora: ["¿Qué aprendió el sistema este mes?", "¿Qué propuestas están pendientes y qué riesgo tienen?"],
  biblioteca: ["¿Qué opiniones estándar tiene la biblioteca y cuándo se usan?"],
  auditoria: ["¿Quién aprobó cada informe y con qué justificación?"],
  normativa: ["¿Qué dice la base normativa sobre el fondo de imprevistos?", "¿Qué norma o manual respalda el hallazgo de cartera de Picasso?"],
};
function prepararAsistente() {
  const A = st.resumen.asistente;
  $("#as-sub").innerHTML = A.ok ? `Pregúntele todo sobre los clientes. Solo lectura: cita las cifras y normas con su fuente y no aprueba ni modifica nada.
    <span class="modelo-chip ${A.local ? "local" : ""}">Responde: ${esc(A.modelo)} · ${esc(A.proveedor)}${A.local ? " · sin internet, los datos no salen de este equipo" : " · por internet"}</span>` : esc(A.ayuda);
  $("#bases").innerHTML = Object.entries(BASES).map(([k, v]) => `<button data-b="${k}" class="${k === st.base ? "activa" : ""}">${v}</button>`).join("");
  document.querySelectorAll("#bases button").forEach((b) => b.onclick = () => { st.base = b.dataset.b; prepararAsistente(); });
  $("#sugerencias").innerHTML = (SUGERENCIAS[st.base] || []).map((s) => `<button>${esc(s)}</button>`).join("");
  document.querySelectorAll("#sugerencias button").forEach((b) => b.onclick = () => { $("#pregunta").value = b.textContent; $("#form-chat").requestSubmit(); });
  if (!st.chat.length) $("#mensajes").innerHTML = `<div class="auditin-hola">
    <div class="auditin-video"><video id="au-video" src="auditin.mp4" autoplay muted loop playsinline aria-label="Video de presentación de Auditin"></video>
      <div class="au-sonido" role="group" aria-label="Sonido del video">
        <button type="button" id="au-mute" aria-label="Activar sonido" title="Activar sonido">🔇</button>
        <button type="button" id="au-menos" aria-label="Bajar volumen" title="Bajar volumen">−</button>
        <input type="range" id="au-vol" min="0" max="100" step="10" aria-label="Volumen">
        <button type="button" id="au-mas" aria-label="Subir volumen" title="Subir volumen">+</button>
        <span id="au-nivel" aria-live="polite"></span>
      </div>
      <p class="au-pista" id="au-pista" hidden>El navegador inició el video sin sonido: pulse 🔇 para activarlo.</p></div>
    <div><b>¡Hola${st.yo ? ", " + esc(st.yo.nombre.split(" ")[0]) : ""}! Soy Auditin.</b>
      <p>Pregúnteme por cualquier cliente: su ficha y documentos, lo que trae su Excel, lo que salió en su informe y lo que se acordó en las reuniones. Cito cada cifra con su hoja y celda, y no apruebo ni cambio nada.</p>
      ${A.ok ? "" : `<p class="m-warn">${esc(A.ayuda)}</p>`}</div></div>`;
  sonidoAuditin();
}
function detenerVideoAuditin() {
  const v = $("#au-video");
  if (v) { v.pause(); v.removeAttribute("src"); v.load(); v.closest(".auditin-video")?.remove(); }
}
$("#au-cerrar").onclick = () => { detenerVideoAuditin(); location.hash = st.vistaPrevia || "panel"; };
// Sonido del video de Auditin: arranca en silencio (los navegadores lo exigen); se recuerda la elección y el volumen
function sonidoAuditin() {
  const v = $("#au-video");
  if (!v) return;
  const leer = (k, d) => { try { return localStorage.getItem(k) ?? d; } catch (_) { return d; } };
  const guardar = (k, x) => { try { localStorage.setItem(k, x); } catch (_) {} };
  let vol = Math.min(100, Math.max(0, +leer("auditin-volumen", 70)));
  const pintar = () => {
    const mudo = v.muted || vol === 0;
    $("#au-mute").textContent = mudo ? "🔇" : "🔊";
    $("#au-mute").ariaLabel = $("#au-mute").title = mudo ? "Activar sonido" : "Silenciar";
    $("#au-vol").value = vol; $("#au-nivel").textContent = mudo ? "Silencio" : `${vol} %`;
    $("#au-menos").disabled = vol <= 0; $("#au-mas").disabled = vol >= 100;
  };
  const poner = (n) => { vol = Math.min(100, Math.max(0, n)); v.volume = vol / 100; guardar("auditin-volumen", vol);
    if (vol > 0 && v.muted) { v.muted = false; guardar("auditin-sonido", "1"); } pintar(); };
  v.volume = vol / 100;
  $("#au-mute").onclick = () => {
    v.muted = !v.muted;
    if (!v.muted && vol === 0) poner(50);
    guardar("auditin-sonido", v.muted ? "0" : "1"); $("#au-pista").hidden = true;
    v.play().catch(() => {}); pintar();
  };
  $("#au-menos").onclick = () => poner(vol - 10);
  $("#au-mas").onclick = () => poner(vol + 10);
  $("#au-vol").oninput = (e) => poner(+e.target.value);
  if (leer("auditin-sonido", "0") === "1") {            // lo activó antes: se intenta con sonido
    v.muted = false;
    v.play().catch(() => { v.muted = true; v.play().catch(() => {}); $("#au-pista").hidden = false; pintar(); });
  }
  pintar();
}
$("#pregunta").onkeydown = (e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); $("#form-chat").requestSubmit(); } };
$("#form-chat").onsubmit = async (ev) => {
  ev.preventDefault();
  const q = $("#pregunta").value.trim(); if (!q) return;
  const caja = $("#mensajes"), btn = $("#btn-enviar");
  if (!st.chat.length) caja.innerHTML = "";
  st.chat.push({role: "user", content: q}); $("#pregunta").value = "";
  caja.insertAdjacentHTML("beforeend", `<div class="msg usuario">${esc(q)}</div>`);
  const fila = document.createElement("div"); fila.className = "msg-fila"; fila.innerHTML = `<img src="auditin.png" alt="Auditin" class="msg-cara">`;
  const div = document.createElement("div"); div.className = "msg claude"; div.innerHTML = `<span class="vacio">Auditin está consultando…</span>`; fila.append(div); caja.append(fila);
  btn.disabled = true;
  let texto = "", herr = [], terminado = false;
  const pintar = () => { div.innerHTML = herr.map((h) => `<div class="herr">${esc(h)}</div>`).join("") + md(texto.trim()); caja.scrollTop = caja.scrollHeight; };
  try {
    const r = await fetch("/api/preguntar", {method: "POST", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({mensajes: st.chat, base: st.base, periodo: periodo(), cliente: st.cliente})});
    if (!r.ok) throw new Error((await r.json()).detail || r.statusText);
    const lector = r.body.getReader(), dec = new TextDecoder(); let buf = "";
    for (;;) {
      const {value, done} = await lector.read(); if (done) break;
      buf += dec.decode(value, {stream: true});
      let i;
      while ((i = buf.indexOf("\n\n")) >= 0) {
        const linea = buf.slice(0, i); buf = buf.slice(i + 2);
        if (!linea.startsWith("data: ")) continue;
        const ev = JSON.parse(linea.slice(6));
        if (ev.tipo === "texto") texto += ev.t;
        else if (ev.tipo === "herramienta") herr.push(`${ev.nombre}(${Object.values(ev.args || {}).filter(Boolean).join(", ")})`);
        else if (ev.tipo === "error") { div.className = "msg error"; texto += (texto ? "\n\n" : "") + ev.mensaje; }
        else if (ev.tipo === "fin") terminado = true;
        pintar();
      }
    }
    if (!terminado && !div.className.includes("error"))     // el flujo se cortó (app reiniciada, Ollama cerrado…)
      throw new Error(texto ? "La respuesta se interrumpió antes de terminar." : "No llegó respuesta.");
    if (div.className.includes("error")) st.chat.pop(); else st.chat.push({role: "assistant", content: texto.trim()});
  } catch (e) {
    div.className = "msg error";
    div.innerHTML = (texto ? md(texto.trim()) + "<hr>" : "") + `<b>${esc(e.message.includes("fetch") || e.message.includes("network") ? "Se perdió la conexión con la aplicación." : e.message)}</b>
      <br>Si la aplicación se reinició o Ollama está cerrado (ícono de la llama en la barra de menú), ábralo y vuelva a preguntar.`;
    st.chat.pop();
  }
  finally { btn.disabled = false; }
};

// ------------------------------------------------------------------ sesión y roles
const puede = (...roles) => !!st.yo && roles.some((r) => st.yo.roles.includes(r));
const NOMBRES_ROL = {asistente: "Asistente", revisor: "Revisor fiscal", auditor: "Auditor (segundo revisor)", admin: "Administrador"};

function mostrarAcceso(modo, mensaje = "") {
  st.acceso = modo;
  const campos = {
    login: `<label>Usuario<input name="usuario" required autocomplete="username" autofocus></label>
            <label>Contraseña<input name="clave" type="password" required autocomplete="current-password"></label>`,
    configurar: `<label>Nombre completo<input name="nombre" required placeholder="Como aparecerá en las aprobaciones"></label>
            <label>Usuario<input name="usuario" class="campo-usuario" required autocomplete="username" pattern="[a-z0-9._\\-]{3,32}" title="3 a 32 caracteres: minúsculas, números, punto, guion o guion bajo (sin espacios ni tildes)" placeholder="javier.galindo"><small class="ayuda">Minúsculas, números, punto o guion; sin espacios ni tildes. Ej.: javier.galindo</small></label>
            <label>Contraseña<input name="clave" type="password" required minlength="10" autocomplete="new-password"></label>
            <label>Repita la contraseña<input name="clave2" type="password" required minlength="10" autocomplete="new-password"></label>`,
    clave: `<label>Contraseña actual<input name="actual" type="password" required autocomplete="current-password"></label>
            <label>Nueva contraseña<input name="nueva" type="password" required minlength="10" autocomplete="new-password"></label>
            <label>Repita la nueva contraseña<input name="nueva2" type="password" required minlength="10" autocomplete="new-password"></label>`,
  }[modo];
  $("#acceso-titulo").textContent = {login: "Ingresar", configurar: "Configuración inicial", clave: "Cambiar contraseña"}[modo];
  $("#acceso-sub").textContent = mensaje || {login: "LCCOUNTER · Revisoría Fiscal PH",
    configurar: "Cree el usuario administrador. Tendrá los roles de administrador y revisor fiscal.",
    clave: "Use al menos 10 caracteres con mayúsculas, minúsculas y números."}[modo];
  $("#acceso-campos").innerHTML = campos;
  $("#acceso-btn").textContent = {login: "Ingresar", configurar: "Crear y entrar", clave: "Guardar contraseña"}[modo];
  $("#acceso-error").hidden = true;
  $("#acceso").hidden = false;
  $("#acceso-campos input")?.focus();
}
$("#acceso-form").onsubmit = async (ev) => {
  ev.preventDefault();
  const F = Object.fromEntries(new FormData(ev.target)), err = $("#acceso-error"), btn = $("#acceso-btn");
  err.hidden = true; btn.disabled = true;
  try {
    if (st.acceso === "configurar") {
      if (F.clave !== F.clave2) throw new Error("Las contraseñas no coinciden.");
      await post("/api/configurar", {usuario: F.usuario, nombre: F.nombre, clave: F.clave});
    } else if (st.acceso === "clave") {
      if (F.nueva !== F.nueva2) throw new Error("Las contraseñas no coinciden.");
      await post("/api/clave", {actual: F.actual, nueva: F.nueva});
    } else {
      await post("/api/login", {usuario: F.usuario, clave: F.clave});
    }
    $("#acceso").hidden = true;
    iniciar();
  } catch (e) { err.textContent = e.message; err.hidden = false; }
  finally { btn.disabled = false; }
};

function pintarUsuario() {
  const u = st.yo;
  $("#perfil").hidden = false;
  $("#perfil-nombre").textContent = u.nombre; $("#perfil-nombre").title = u.nombre;
  $("#perfil-rol").textContent = u.roles.map((r) => NOMBRES_ROL[r] || r).join(" · ");
  pintarAvatar();
  $("#usuario-caja").innerHTML = `${puede("admin", "revisor") ? `<a class="btn sec boton-config" href="#configuracion" data-vista="configuracion">⚙ Configuración</a>` : ""}
    <div class="grupo"><button class="btn sec chico" id="u-clave">Contraseña</button><button class="btn sec chico" id="u-salir">Salir</button></div>`;
  $("#u-salir").onclick = async () => { await post("/api/logout", {}); location.reload(); };
  $("#u-clave").onclick = () => mostrarAcceso("clave");
  document.querySelectorAll("[data-rol]").forEach((x) => (x.hidden = !puede(...x.dataset.rol.split(","))));
  $("#btn-validar-todo").hidden = !puede("asistente", "revisor");
  $("#btn-analizar").hidden = !puede("revisor");
  $("#btn-regresion").hidden = $("#btn-muestra").hidden = !puede("auditor", "revisor");
  document.querySelectorAll("#form-cargar button, #form-cliente button").forEach((b) => (b.disabled = !puede("asistente", "revisor")));
}

// Usuarios (administrador)
async function cargarUsuarios() {
  if (!puede("admin")) { location.hash = "panel"; return; }
  const U = await api("/api/usuarios");
  $("#u-roles").innerHTML = `<legend>Roles</legend>` + Object.entries(U.roles).map(([r, d]) =>
    `<label class="check"><input type="checkbox" name="roles" value="${r}" ${r === "asistente" ? "checked" : ""}> <b>${esc(NOMBRES_ROL[r])}</b> <small>${esc(d)}</small></label>`).join("");
  $("#u-tabla").innerHTML = `<tr><th>Nombre</th><th>Roles</th><th>Estado</th><th>Último acceso</th><th></th></tr>` + U.usuarios.map((x) => `
    <tr><td><b>${esc(x.nombre)}</b><br><small class="nota">${esc(x.usuario)}</small></td>
    <td>${Object.keys(U.roles).map((r) => `<label class="check en-fila"><input type="checkbox" data-u="${esc(x.usuario)}" data-r="${r}" ${x.roles.includes(r) ? "checked" : ""}> ${esc(NOMBRES_ROL[r].split(" ")[0])}</label>`).join("")}</td>
    <td>${x.activo ? `<span class="chip e-APROBADO">Activo</span>` : `<span class="chip e-BLOQUEADO">Inactivo</span>`}</td>
    <td><small class="nota">${esc(x.ultimo_acceso || "—")}</small></td>
    <td class="grupo"><button class="btn sec chico" data-act="${x.activo ? "desactivar" : "activar"}" data-u="${esc(x.usuario)}">${x.activo ? "Desactivar" : "Activar"}</button>
      <button class="btn sec chico" data-act="clave" data-u="${esc(x.usuario)}">Restablecer contraseña</button></td></tr>`).join("");
  const cambiar = async (usuario, cuerpo, msg) => {
    try { await api(`/api/usuarios/${encodeURIComponent(usuario)}`, {method: "PATCH", headers: {"Content-Type": "application/json"}, body: JSON.stringify(cuerpo)}); aviso(msg); }
    catch (e) { aviso(e.message, true); }
    cargarUsuarios();
  };
  document.querySelectorAll("#u-tabla input[data-r]").forEach((c) => c.onchange = () => {
    const roles = [...document.querySelectorAll(`#u-tabla input[data-u="${c.dataset.u}"]:checked`)].map((x) => x.dataset.r);
    cambiar(c.dataset.u, {roles}, "Roles actualizados.");
  });
  document.querySelectorAll("#u-tabla button[data-act]").forEach((b) => b.onclick = () => {
    const u = b.dataset.u;
    if (b.dataset.act === "clave") dialogo(`Restablecer contraseña de ${u}`, `<label>Contraseña temporal<input name="clave" type="password" required minlength="10" autocomplete="new-password"></label>
      <p class="nota">Se cerrarán sus sesiones y deberá cambiarla al entrar.</p>`, async (f) => cambiar(u, {clave: f.clave}, "Contraseña restablecida."));
    else cambiar(u, {activo: b.dataset.act === "activar"}, b.dataset.act === "activar" ? "Usuario activado." : "Usuario desactivado y sesiones cerradas.");
  });
  $("#u-permisos").innerHTML = `<tr><th>Acción</th>${Object.keys(U.roles).map((r) => `<th>${esc(NOMBRES_ROL[r] || r)}</th>`).join("")}</tr>` +
    (U.permisos || []).map(([accion, roles]) => `<tr><td>${esc(accion)}</td>${Object.keys(U.roles).map((r) => `<td class="permiso">${roles.includes(r) ? "✔" : "—"}</td>`).join("")}</tr>`).join("");
  $("#u-registro").innerHTML = `<tr><th>Fecha</th><th>Usuario</th><th>Acción</th><th>Detalle</th></tr>` +
    (U.registro.map((r) => `<tr><td><small>${esc(r.fecha)}</small></td><td>${esc(r.usuario)}</td><td>${esc(r.accion)}</td><td><small class="nota">${esc(r.detalle || "")}</small></td></tr>`).join("")
     || `<tr><td colspan="4" class="vacio">Sin registros.</td></tr>`);
}
$("#u-form").onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
  const fd = new FormData(ev.target);
  const roles = fd.getAll("roles");
  if (!roles.length) throw new Error("Asigne al menos un rol.");
  await post("/api/usuarios", {usuario: fd.get("usuario"), nombre: fd.get("nombre"), roles, clave: fd.get("clave")});
  aviso("Usuario creado. Entréguele la contraseña temporal por un canal privado."); ev.target.reset(); cargarUsuarios();
}); };

// Conceptos del revisor fiscal (a todas las copropiedades)
async function cargarConceptos() {
  const C = await api("/api/conceptos");
  $("#c-ia").textContent = C.disponible.ok ? `Redacta: ${C.disponible.modelo}${C.disponible.local ? " (sin internet)" : ""}` : C.disponible.ayuda;
  $("#c-clientes").innerHTML = `<legend>Copropiedades (${st.resumen.clientes.length})</legend>` + st.resumen.clientes.map((c) =>
    `<label class="check"><input type="checkbox" name="cli" value="${esc(c)}" checked> ${esc(nombreCliente(c))}</label>`).join("");
  $("#c-lista").innerHTML = C.conceptos.map((x) => `<div class="propuesta clic" data-c="${esc(x.id)}"><b>${esc(x.asunto || x.tema)}</b>
      <div class="nota">${x.estado === "emitido" ? `<span class="chip e-APROBADO">Emitido</span> ${esc(x.emitido?.fecha?.replace("T", " ") || "")}` : `<span class="chip e-PENDIENTE_VALIDACION">Borrador</span>`}
      · ${x.clientes?.length || 0} copropiedad(es) · ${esc(x.por || "")}</div></div>`).join("") || `<p class="vacio">Aún no hay conceptos.</p>`;
  $("#c-lista").querySelectorAll("[data-c]").forEach((d) => d.onclick = () => verConcepto(d.dataset.c));
}
$("#c-buscar-btn").onclick = (ev) => ocupado(ev.target, async () => {
  const r = await post("/api/conceptos/fragmentos", {consulta: $("#c-buscar").value});
  $("#c-fragmentos").innerHTML = r.fragmentos.map((f, i) => `<label class="fragmento"><input type="checkbox" data-f='${esc(JSON.stringify({ambito: f.ambito, doc_id: f.doc_id, fragmento: f.fragmento}))}' ${i < 2 ? "checked" : ""}>
      <span><b>${esc(f.documento)} · ${esc(f.referencia)}</b><small>${esc(f.texto.slice(0, 260))}${f.texto.length > 260 ? "…" : ""}</small></span></label>`).join("")
    || `<p class="vacio">No hay resultados. Cargue la norma en «Base normativa» primero.</p>`;
});
$("#c-buscar").onkeydown = (e) => { if (e.key === "Enter") { e.preventDefault(); $("#c-buscar-btn").click(); } };
$("#c-nuevo").onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
  const F = new FormData(ev.target);
  const fragmentos = [...document.querySelectorAll("#c-fragmentos input:checked")].map((x) => JSON.parse(x.dataset.f));
  if (!fragmentos.length) throw new Error("Busque la norma y marque al menos un fragmento.");
  const d = await conCarga("Redactando el concepto…", () => post("/api/conceptos", {tema: F.get("tema"), opinion: F.get("opinion"), fragmentos, clientes: F.getAll("cli")}));
  aviso("Borrador listo: revíselo, ajústelo y emítalo."); ev.target.reset(); $("#c-fragmentos").innerHTML = "";
  await cargarConceptos(); verConcepto(d.id);
}); };
async function verConcepto(id) {
  const d = await api(`/api/conceptos/${id}`), ig = d.infografia || {}, emitido = d.estado === "emitido", ed = puede("revisor") && !emitido;
  const caja = $("#c-detalle");
  caja.innerHTML = `<form class="caja" id="c-form"><div class="cab-mini"><h2>${esc(d.tema)}</h2>${emitido ? `<span class="chip e-APROBADO">Emitido por ${esc(d.emitido.por)}</span>` : `<span class="chip e-PENDIENTE_VALIDACION">Borrador</span>`}</div>
    ${(d.citas_invalidas || []).length ? `<p class="m-warn">⚠ El texto cita fuentes que no existen: ${d.citas_invalidas.map(esc).join(", ")}. Las fuentes válidas son F0 (su opinión) y F1–F${d.fragmentos.length} (la norma). Corríjalas antes de emitir.</p>` : ""}
    <label>Asunto<input name="asunto" value="${esc(d.asunto)}" ${ed ? "" : "readonly"}></label>
    <label>Texto explicativo (Markdown)<textarea name="texto" rows="14" ${ed ? "" : "readonly"}>${esc(d.texto)}</textarea></label>
    <div class="dos"><div>
      <h2 class="mt">Infografía</h2>
      <label>Título<input name="ig_titulo" value="${esc(ig.titulo || "")}" ${ed ? "" : "readonly"}></label>
      <label>Lo que dice la norma (una idea por línea)<textarea name="ig_norma" rows="5" ${ed ? "" : "readonly"}>${esc((ig.norma || []).join("\n"))}</textarea></label>
      <label>Opinión del revisor (una idea por línea)<textarea name="ig_opinion" rows="5" ${ed ? "" : "readonly"}>${esc((ig.opinion || []).join("\n"))}</textarea></label>
      <label>Recomendación<textarea name="ig_rec" rows="2" ${ed ? "" : "readonly"}>${esc(ig.recomendacion || "")}</textarea></label>
    </div><div><h2 class="mt">Vista previa</h2><img class="infografia" src="/api/conceptos/${id}/infografia.png?v=${Date.now()}" alt="Infografía del concepto">
      <h2 class="mt">Anexo: norma</h2><ul class="docs">${d.fragmentos.map((f) => `<li><span><b>${esc(f.documento)}</b> · ${esc(f.referencia)}</span></li>`).join("")}</ul></div></div>
    <div class="acciones">
      ${emitido ? `<a class="btn" href="/api/conceptos/${id}/archivo/${encodeURIComponent(d.emitido.zip)}">Descargar todos (ZIP, ${d.emitido.archivos.length} Word)</a>` : ""}
      ${ed ? `<button class="btn sec" value="guardar">Guardar cambios</button><button class="btn" value="emitir">Emitir para ${d.clientes.length} copropiedad(es)</button>` : ""}
    </div>
    ${emitido ? `<details><summary>Ver cada documento</summary><ul class="docs">${d.emitido.archivos.map((a) => `<li><a href="/api/conceptos/${id}/archivo/${encodeURIComponent(a)}">${esc(a)}</a></li>`).join("")}</ul></details>` : ""}
    <p class="nota">La aplicación no envía nada: descargue los documentos y compártalos. Un concepto emitido queda congelado.</p></form>`;
  caja.scrollIntoView({behavior: "smooth", block: "start"});
  $("#c-form").onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
    const F = Object.fromEntries(new FormData(ev.target)), lineas = (t) => t.split("\n").map((x) => x.trim()).filter(Boolean);
    await api(`/api/conceptos/${id}`, {method: "PUT", headers: {"Content-Type": "application/json"}, body: JSON.stringify({asunto: F.asunto, texto: F.texto,
      infografia: {titulo: F.ig_titulo, norma: lineas(F.ig_norma), opinion: lineas(F.ig_opinion), recomendacion: F.ig_rec}})});
    if (ev.submitter?.value === "emitir") {
      await conCarga("Generando un Word por copropiedad…", () => post(`/api/conceptos/${id}/emitir`, {}));
      aviso("Concepto emitido: descargue el ZIP con los documentos.");
    } else aviso("Concepto guardado.");
    await cargarConceptos(); verConcepto(id);
  }); };
}

// Tiempos: en línea por usuario y de procesamiento por cliente; registro de acciones
const duracion = (s) => { s = Math.round(s || 0); const h = Math.floor(s / 3600), m = Math.round((s % 3600) / 60); return h ? `${h} h ${m} min` : s < 60 && s ? `${s} s` : `${m} min`; };
async function cargarTiempos() {
  if (!puede("admin", "revisor")) { location.hash = "panel"; return; }
  const [T, R] = await Promise.all([api("/api/tiempos"), api("/api/registro")]);
  $("#t-linea").innerHTML = `<tr><th>Usuario</th><th>Hoy</th><th>Últimos 7 días</th><th>Últimos 30 días</th></tr>` +
    (T.en_linea.sort((a, b) => b.mes - a.mes).map((x) => `<tr><td><b>${esc(x.nombre)}</b> ${x.en_linea ? `<span class="chip e-APROBADO">en línea</span>` : ""}</td>
      <td class="n">${duracion(x.hoy)}</td><td class="n">${duracion(x.semana)}</td><td class="n">${duracion(x.mes)}</td></tr>`).join("")
     || `<tr><td colspan="4" class="vacio">Aún no hay tiempos registrados.</td></tr>`);
  $("#t-proc").innerHTML = `<tr><th>Cliente</th><th>Periodo</th><th>Tiempo</th><th>Estado</th></tr>` +
    (T.procesamiento.sort((a, b) => b.periodo.localeCompare(a.periodo) || b.segundos - a.segundos).map((x) => `<tr><td><b>${esc(x.copropiedad)}</b></td>
      <td>${esc(textoPeriodo(x.periodo))}</td><td class="n">${duracion(x.segundos)}</td><td>${chip(x.estado)}</td></tr>`).join("")
     || `<tr><td colspan="4" class="vacio">Valide un cliente para medir su tiempo de procesamiento.</td></tr>`);
  $("#t-registro").innerHTML = `<tr><th>Fecha</th><th>Usuario</th><th>Acción</th><th>Detalle</th></tr>` +
    (R.registro.map((r) => `<tr><td><small>${esc(r.fecha)}</small></td><td>${esc(r.usuario)}</td><td>${esc(r.accion)}</td><td><small class="nota">${esc(r.detalle || "")}</small></td></tr>`).join("")
     || `<tr><td colspan="4" class="vacio">Sin registros.</td></tr>`);
}
// latido: suma el tiempo en línea solo con la app abierta y visible
function latir() {
  if (!st.yo || latir.parado) return;
  fetch("/api/latido", {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify({visible: !document.hidden})})
    .then((r) => { if (r.status === 401) latir.parado = true; }).catch(() => {});   // sin sesión no se insiste
}
setInterval(latir, 60_000);
document.addEventListener("visibilitychange", latir);

// Muestra mensual con la conclusión del segundo revisor
function pintarMuestra(M) {
  const m = M?.muestra || [];
  $("#m-muestra").innerHTML = m.length ? m.map((x) => {
    const rev = x.revision, propio = x.aprobado_por === st.yo.nombre;
    return `<div class="propuesta"><b>${esc(x.copropiedad || x.cliente)}</b> · aprobó ${esc(x.aprobado_por)} · ${x.hallazgos.length} observación(es) automática(s)
      ${x.hallazgos.map((h) => `<div class="nota">${esc(h.regla)} ${esc(h.titulo)}</div>`).join("")}
      ${rev ? `<p><span class="chip e-APROBADO">${esc({sin_observaciones: "Sin observaciones", menores: "Observaciones menores", reproceso: "Requiere reproceso"}[rev.conclusion])}</span> ${esc(rev.auditor)} · ${esc(rev.observacion || "")}</p>`
        : !puede("auditor") ? `<p class="nota">Pendiente de conclusión de un auditor.</p>`
        : propio ? `<p class="nota">Usted aprobó este informe: la auditoría la hace otro profesional.</p>`
        : `<form class="grupo concluir" data-c="${esc(x.cliente)}"><select name="conclusion"><option value="sin_observaciones">Sin observaciones</option><option value="menores">Observaciones menores</option><option value="reproceso">Requiere reproceso</option></select>
            <input name="observacion" placeholder="Observación (obligatoria si hay hallazgos)"><button class="btn chico">Registrar conclusión</button></form>`}</div>`;
  }).join("") + `<p class="nota">Lista de chequeo completa en <code>Auditoria_Muestra_${esc(periodo())}.html</code> (raíz del espacio de trabajo).</p>`
    : `<p class="vacio">Sin muestra para ${esc(periodo())}.</p>`;
  document.querySelectorAll("form.concluir").forEach((f) => f.onsubmit = (ev) => { ev.preventDefault(); ocupado(ev.submitter, async () => {
    const F = Object.fromEntries(new FormData(f));
    await post("/api/auditar", {periodo: periodo(), cliente: f.dataset.c, ...F});
    aviso("Conclusión registrada: alimenta la mejora continua.");
    pintarMuestra((await api(`/api/muestra/${periodo()}`)).muestra);
  }); });
}

// Avatar del usuario: foto propia o, si no tiene, sus iniciales
function pintarAvatar() {
  const u = st.yo, img = $("#avatar-img");
  $("#avatar-ini").textContent = u.nombre.split(/\s+/).filter(Boolean).slice(0, 2).map((p) => p[0].toUpperCase()).join("");
  img.onload = () => { img.hidden = false; $("#avatar-ini").hidden = true; };
  img.onerror = () => { img.hidden = true; $("#avatar-ini").hidden = false; };
  if (u.avatar) img.src = `/api/avatar/${encodeURIComponent(u.usuario)}?v=${Date.now()}`;
  else img.onerror();                                     // sin foto: iniciales, sin pedir la imagen
}
$("#avatar-btn").onclick = () => {
  if ($("#avatar-img").hidden) return $("#f-avatar").click();
  dialogo("Mi imagen", `<p class="nota">Elija una foto nueva o quite la actual para mostrar sus iniciales.</p>
    <div class="grupo"><button type="button" class="btn sec" id="av-cambiar">Elegir otra imagen</button><button type="button" class="btn sec" id="av-quitar">Quitar imagen</button></div>`, async () => {});
  $("#dlg-ok").hidden = true; $("#dlg").addEventListener("close", () => ($("#dlg-ok").hidden = false), {once: true});
  $("#av-cambiar").onclick = () => { $("#dlg").close(); $("#f-avatar").click(); };
  $("#av-quitar").onclick = async () => { $("#dlg").close(); await api("/api/avatar", {method: "DELETE"}); st.yo.avatar = false; pintarAvatar(); aviso("Imagen quitada."); };
};
$("#f-avatar").onchange = async () => {
  const f = $("#f-avatar").files[0]; $("#f-avatar").value = "";
  if (!f) return;
  if (f.size > 2 * 1024 * 1024) return aviso("La imagen pesa más de 2 MB.", true);
  const fd = new FormData(); fd.append("imagen", f);
  try { await api("/api/avatar", {method: "POST", body: fd}); st.yo.avatar = true; pintarAvatar(); aviso("Imagen actualizada."); } catch (e) { aviso(e.message, true); }
};

// Periodo: calendario de meses (funciona igual en Safari, Chrome y Edge)
const MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
const textoPeriodo = (p) => { const [a, m] = (p || "").split("-"); return m ? `${MESES[+m - 1]} ${a}` : "—"; };
const cal = {anio: null};
function ponerPeriodo(p) {
  $("#periodo").value = p; $("#periodo-txt").textContent = textoPeriodo(p);
  $("#periodo").onchange?.();
}
function pintarCalendario() {
  const sel = periodo(), hoy = new Date(), hoyP = `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, "0")}`;
  const conDatos = new Set(st.resumen?.periodos || []);
  $("#cal-anio").textContent = cal.anio;
  $("#cal-meses").innerHTML = MESES.map((m, i) => {
    const p = `${cal.anio}-${String(i + 1).padStart(2, "0")}`;
    const cls = [p === sel && "sel", p === hoyP && "hoy", conDatos.has(p) && "con-datos"].filter(Boolean).join(" ");
    return `<button type="button" data-p="${p}" class="${cls}" title="${m} ${cal.anio}${conDatos.has(p) ? " · con datos" : ""}">${m.slice(0, 3)}</button>`;
  }).join("");
  $("#cal-meses").querySelectorAll("button").forEach((b) => b.onclick = () => { cerrarCalendario(); if (b.dataset.p !== periodo()) ponerPeriodo(b.dataset.p); });
}
function abrirCalendario() {
  cal.anio = +(periodo() || String(new Date().getFullYear())).slice(0, 4);
  pintarCalendario(); $("#calendario").hidden = false; $("#periodo-btn").setAttribute("aria-expanded", "true");
}
function cerrarCalendario() { $("#calendario").hidden = true; $("#periodo-btn").setAttribute("aria-expanded", "false"); }
$("#periodo-btn").onclick = () => ($("#calendario").hidden ? abrirCalendario() : cerrarCalendario());
$("#cal-ant").onclick = () => { cal.anio--; pintarCalendario(); };
$("#cal-sig").onclick = () => { cal.anio++; pintarCalendario(); };
$("#cal-hoy").onclick = () => { const h = new Date(); cerrarCalendario(); ponerPeriodo(`${h.getFullYear()}-${String(h.getMonth() + 1).padStart(2, "0")}`); };
document.addEventListener("click", (e) => { if (!$("#calendario").hidden && !e.target.closest(".periodo-caja")) cerrarCalendario(); });
document.addEventListener("keydown", (e) => { if (e.key === "Escape" && !$("#calendario").hidden) { cerrarCalendario(); $("#periodo-btn").focus(); } });

// Modo día / noche (se recuerda en este equipo; si no se elige, sigue al sistema)
function temaActual() {
  return document.documentElement.dataset.theme || (matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
}
function pintarTema() {
  const noche = temaActual() === "dark";
  $("#tema-ico").textContent = noche ? "☀" : "☾"; $("#tema-txt").textContent = noche ? "Modo día" : "Modo noche";
}
$("#tema-btn").onclick = () => {
  const t = temaActual() === "dark" ? "light" : "dark";
  document.documentElement.dataset.theme = t;
  try { localStorage.setItem("lccounter-tema", t); } catch (_) {}
  pintarTema();
};
matchMedia("(prefers-color-scheme: dark)").addEventListener?.("change", pintarTema);
pintarTema();

// Barra lateral: plegable en escritorio y menú ☰ en celular
const lateral = $(".lateral");
function plegar(si) {
  document.body.classList.toggle("plegado", si);
  $("#plegar-btn").textContent = si ? "»" : "«";
  $("#plegar-btn").title = $("#plegar-btn").ariaLabel = si ? "Desplegar la barra lateral" : "Plegar la barra lateral";
  try { localStorage.setItem("lccounter-plegado", si ? "1" : ""); } catch (_) {}
}
$("#plegar-btn").onclick = () => plegar(!document.body.classList.contains("plegado"));
document.querySelectorAll("nav a").forEach((a) => (a.title = a.textContent.trim()));   // nombre visible al pasar el mouse con la barra plegada
try { if (localStorage.getItem("lccounter-plegado")) plegar(true); } catch (_) {}
function menu(abierto) { lateral.classList.toggle("abierto", abierto); $("#menu-btn").setAttribute("aria-expanded", String(abierto)); $("#menu-btn").textContent = abierto ? "✕" : "☰"; }
$("#menu-btn").onclick = () => menu(!lateral.classList.contains("abierto"));
window.addEventListener("hashchange", () => menu(false));

// ------------------------------------------------------------------ inicio
async function iniciar() {
  const S = await api("/api/sesion");
  if (!S.configurado) return mostrarAcceso("configurar");
  if (!S.usuario) return mostrarAcceso("login");
  st.yo = S.usuario;
  if (S.cambiar_clave) return mostrarAcceso("clave", "Su contraseña es temporal: cámbiela para continuar.");
  st.resumen = await api("/api/resumen");
  pintarUsuario(); latir();
  const hoy = new Date(); const def = st.resumen.periodos[0] || `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, "0")}`;
  if (!$("#periodo").value) $("#periodo").value = def;
  $("#periodo-txt").textContent = textoPeriodo(periodo());
  $("#periodo").onchange = () => mostrar(location.hash.slice(1).split("/")[0] || "panel");
  mostrar(location.hash.slice(1).split("/")[0] || "panel");
}
iniciar().catch((e) => aviso("No se pudo conectar con el servidor local: " + e.message, true));
window.addEventListener("beforeunload", (e) => {
  if (st.cambios && Object.keys(st.cambios).length) { e.preventDefault(); e.returnValue = ""; }
});

// El usuario se escribe siempre en minúsculas y sin tildes ni espacios (así lo exige el servidor).
document.addEventListener("input", (e) => {
  if (!e.target.classList?.contains("campo-usuario")) return;
  const limpio = e.target.value.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ".").replace(/[^a-z0-9._-]/g, "");
  if (limpio !== e.target.value) e.target.value = limpio;
});
