// Utilidades comunes: llamadas a la API, barra superior y cambio de contraseña.
async function api(url, body) {
  const op = body === undefined ? {} : {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(body)};
  let r;
  try { r = await fetch(url, op); } catch (e) { throw new Error("No se pudo conectar con el servidor. Intenta de nuevo."); }
  const j = await r.json().catch(() => ({}));
  if (r.status === 401) { location.href = "/login"; throw new Error(j.error || "Sesión terminada"); }
  if (!r.ok) throw new Error(j.error || "Error " + r.status);
  return j;
}
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c]));
const fmtFecha = iso => iso ? iso.split("-").reverse().join("-") : "";
const DIAS = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
const diaSemana = iso => DIAS[new Date(iso + "T12:00:00").getDay()];

async function barra(actual) {
  const yo = await api("/api/yo");
  const links = [["/", "Reporte diario"], ["/dashboard", "Dashboard y descarga"]];
  if (yo.rol === "admin") links.push(["/admin", "Usuarios y buses"]);
  const el = document.createElement("header");
  el.className = "top";
  el.innerHTML = `<div class="in">
    <div class="brand"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><rect x="3" y="4" width="18" height="13" rx="2"/><path d="M3 11h18M7 17v3M17 17v3"/><circle cx="7.5" cy="14" r=".8" fill="currentColor"/><circle cx="16.5" cy="14" r=".8" fill="currentColor"/></svg>Control de Flota</div>
    <nav class="nav">${links.map(([h, t]) => `<a href="${h}" ${h === actual ? 'aria-current="page"' : ""}>${t}</a>`).join("")}</nav>
    <div class="who"><span>${esc(yo.nombre)} · ${yo.rol === "admin" ? "Administrador" : "Operador"}</span>
      <button class="linkbtn" id="b-clave">Cambiar contraseña</button><button class="linkbtn" id="b-salir">Salir</button></div></div>
  <dialog id="dlg-clave"><form method="dialog" id="f-clave" style="display:flex;flex-direction:column;gap:10px">
    <h3>Cambiar mi contraseña</h3>
    <label class="f">Contraseña actual<input type="password" id="c-act" required autocomplete="current-password"></label>
    <label class="f">Nueva contraseña (mínimo 8)<input type="password" id="c-new" required minlength="8" autocomplete="new-password"></label>
    <p class="note" id="c-msg"></p>
    <div class="row"><button class="btn" value="ok">Guardar</button><button class="btn sec" value="cancel" formnovalidate>Cancelar</button></div>
  </form></dialog>`;
  document.body.prepend(el);
  el.querySelector("#b-salir").onclick = async () => { await api("/api/logout", {}); location.href = "/login"; };
  const dlg = el.querySelector("#dlg-clave");
  el.querySelector("#b-clave").onclick = () => { el.querySelector("#c-msg").textContent = ""; dlg.showModal(); };
  el.querySelector("#f-clave").addEventListener("submit", async e => {
    if (e.submitter && e.submitter.value === "cancel") return;
    e.preventDefault();
    try {
      await api("/api/mi-clave", {actual: el.querySelector("#c-act").value, nueva: el.querySelector("#c-new").value});
      dlg.close(); alert("Contraseña actualizada.");
    } catch (x) { el.querySelector("#c-msg").textContent = x.message; }
  });
  return yo;
}
function guardarPref(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
function leerPref(k, d) { try { return localStorage.getItem(k) || d; } catch (e) { return d; } }
