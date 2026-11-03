/* PortCo Learning Hub – secure edition.
   Portcos pick their company and enter a PIN; each sees only its own data.
   Admin sees everything. Data files are built and encrypted by build_secure.py. */

const CONFIG = {
  autoRefreshSeconds: 60,
  dueSoonDays: 7,
  statusOrder: ["Script Received", "In Development", "In Review", "Portco Feedback", "Delivered", "On Hold"]
};
/* Colour sets for each theme (charts can't read CSS variables, so they live here too) */
const THEMES = {
  dark:  { gold: "#E3C88B", gold2: "#B8964A", gold3: "#7A6332", red: "#E26D5F", amber: "#E7A94F", green: "#7CC49A", grey: "#5E5E5E",
           track: "#262626", muted: "#8C8C8C", line: "#2B2B2B", card: "#1B1B1B", tip: "#111111",
           status: ["#F3E2B8", "#E3C88B", "#C9A660", "#9C7E43", "#6E5A30", "#4A4A4A"] },
  light: { gold: "#2E5A8A", gold2: "#1F3A5F", gold3: "#7FA7D9", red: "#D64545", amber: "#E8A33D", green: "#2E9E6B", grey: "#9AA3B2",
           track: "#EEF2F8", muted: "#6B7385", line: "#E3E7EF", card: "#FFFFFF", tip: "#1F3A5F",
           status: ["#9AA3B2", "#3B7DD8", "#E8A33D", "#8E6CD8", "#2E9E6B", "#C9CED8"] }
};
let theme = document.documentElement.getAttribute("data-theme") === "light" ? "light" : "dark";
let C = THEMES[theme];
const HEALTH_COLOR = { gold: C.gold, green: C.green, amber: C.amber, red: C.red, grey: C.grey };
const HEALTH_LABEL = { gold: "Delivered", green: "On track", amber: "At risk", red: "Overdue", grey: "On hold" };

let DATA = { portcos: [], projects: [], comms: [], files: [] };
let charts = {}, lastSignature = "", usingLocalFile = false;

const $ = (id) => document.getElementById(id);
const today = () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; };

/* ---------- Theme ---------- */
function applyChartDefaults() {
  if (!window.Chart) return;
  Chart.defaults.color = C.muted;
  Chart.defaults.borderColor = C.line;
  Chart.defaults.font.family = "Manrope, Segoe UI, sans-serif";
  Chart.defaults.plugins.tooltip.backgroundColor = C.tip;
  Chart.defaults.plugins.tooltip.borderColor = C.gold3;
  Chart.defaults.plugins.tooltip.borderWidth = 1;
}
function setTheme(t) {
  theme = t; C = THEMES[t];
  document.documentElement.setAttribute("data-theme", t);
  try { localStorage.setItem("plh-theme", t); } catch (e) {}
  document.querySelectorAll(".theme-toggle span").forEach(s => s.classList.toggle("on", s.dataset.t === t));
  applyChartDefaults();
  const s = $("syncStatus").textContent;
  if (s && !s.startsWith("Loading")) setSync(s.replace(/^[●⚠] /, ""), !s.startsWith("⚠"));
  if (DATA.projects.length) render();
}

/* ---------- Secure access (PortCo PIN / Admin) ---------- */
/* The published site holds NO readable data. Each portco's data is encrypted with
   its own PIN (AES-256-GCM, key from PBKDF2-SHA256). The admin file holds everything. */
let SESSION = null;           // { role: "admin" | "portco", id, name, file, key }
let PORTCO_LIST = [];
const enc = new TextEncoder(), dec = new TextDecoder();
const b64 = (s) => Uint8Array.from(atob(s), c => c.charCodeAt(0));
const toB64 = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf)));

async function deriveKey(secret, saltB64, iterations) {
  const base = await crypto.subtle.importKey("raw", enc.encode(secret), "PBKDF2", false, ["deriveKey"]);
  return crypto.subtle.deriveKey({ name: "PBKDF2", salt: b64(saltB64), iterations, hash: "SHA-256" },
    base, { name: "AES-GCM", length: 256 }, true, ["decrypt"]);
}
async function fetchEnc(file) {
  const res = await fetch(`data/${file}?v=${Date.now()}`, { cache: "no-store" });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}
async function decryptWith(key, blob) {
  const plain = await crypto.subtle.decrypt({ name: "AES-GCM", iv: b64(blob.iv) }, key, b64(blob.ct));
  return JSON.parse(dec.decode(plain));
}

async function loadSecure(force = false) {
  if (!SESSION || (usingLocalFile && !force)) return;
  try {
    const blob = await fetchEnc(SESSION.file);
    if (blob.ct === lastSignature && !force) { return; }
    const data = await decryptWith(SESSION.key, blob);
    lastSignature = blob.ct; usingLocalFile = false;
    ingest(data, "Secure data");
  } catch (err) {
    console.warn(err);
    if (String(err.name) === "OperationError") { logout("Your PIN has changed. Please sign in again."); return; }
    setSync("Can't load data – try Refresh", false);
  }
}
const loadFromRepo = loadSecure;   // used by the Refresh button

function loadFromFile(file) {      // admin only: preview a local Excel file
  const r = new FileReader();
  r.onload = (e) => { usingLocalFile = true; parseWorkbook(e.target.result, file.name); };
  r.readAsArrayBuffer(file);
}
function parseWorkbook(buf, name) {
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheets = {};
  ["Portcos", "Projects", "Communications", "Files", "SyncInfo"].forEach(n => {
    const key = wb.SheetNames.find(s => s.toLowerCase() === n.toLowerCase());
    sheets[n] = key ? XLSX.utils.sheet_to_json(wb.Sheets[key], { defval: "" }) : [];
  });
  ingest(sheets, name);
}

function ingest(sheets, name) {
  const sheet = (n) => sheets[n] || [];
  DATA.portcos = sheet("Portcos").filter(r => r.PortcoName);
  DATA.projects = sheet("Projects").filter(r => r.ProjectID || r.ProjectName).map(r => ({
    ...r, ScriptReceived: toDate(r.ScriptReceived), StartDate: toDate(r.StartDate), DueDate: toDate(r.DueDate),
    Progress: Math.max(0, Math.min(100, Number(r.Progress) || 0)), FeedbackRounds: Number(r.FeedbackRounds) || 0
  }));
  DATA.comms = sheet("Communications").filter(r => r.Date || r.Subject).map(r => ({ ...r, Date: toDate(r.Date) }));
  DATA.files = sheet("Files").filter(r => r.FileName).map(r => ({ ...r, UploadedOn: toDate(r.UploadedOn) }));
  const known = new Set(DATA.portcos.map(p => p.PortcoName));
  DATA.projects.forEach(p => { if (p.PortcoName && !known.has(p.PortcoName)) { DATA.portcos.push({ PortcoName: p.PortcoName }); known.add(p.PortcoName); } });

  const info = {};
  sheet("SyncInfo").forEach(r => { if (r.Key) info[r.Key] = r.Value; });
  const fromMonday = !usingLocalFile && info.Source === "Monday.com";
  const syncedAt = info.LastSync ? new Date(info.LastSync) : null;
  const syncedTxt = syncedAt && !isNaN(syncedAt) ? syncedAt.toLocaleString([], { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }) : "";

  $("sourceName").textContent = usingLocalFile ? name : fromMonday ? "Monday.com" : "Project tracker";
  $("lastRefresh").textContent = fromMonday && syncedTxt ? syncedTxt : new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  setSync(usingLocalFile ? "Showing uploaded file" : fromMonday ? `Live · synced from Monday${syncedTxt ? " · " + syncedTxt : ""}` : "Live", true);
  buildFilterOptions();
  if (SESSION && SESSION.role === "portco") $("fPortco").value = "";
  render();
}
function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === "number") return new Date(Math.round((v - 25569) * 86400 * 1000));
  const d = new Date(v); return isNaN(d) ? null : d;
}
const fmt = (d) => d ? d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "–";
function setSync(text, ok) { const el = $("syncStatus"); el.textContent = (ok ? "● " : "⚠ ") + text; el.style.color = ok ? C.gold : C.amber; }

/* ---------- Health ---------- */
const isClosed = (p) => ["delivered", "on hold"].includes(String(p.Status).toLowerCase());
function health(p) {
  const s = String(p.Status).toLowerCase();
  if (s === "delivered") return "gold";
  if (s === "on hold") return "grey";
  if (!p.DueDate) return "green";
  const days = (p.DueDate - today()) / 86400000;
  if (days < 0) return "red";
  if (days <= CONFIG.dueSoonDays && p.Progress < 70) return "amber";
  return "green";
}

/* ---------- Filters ---------- */
const unique = (a) => [...new Set(a.filter(Boolean))];
function buildFilterOptions() {
  fill("fPortco", unique(DATA.portcos.map(p => p.PortcoName)), "ALL PORTCOS");
  fill("fStatus", unique([...CONFIG.statusOrder.filter(s => DATA.projects.some(p => p.Status === s)), ...DATA.projects.map(p => p.Status)]), "All statuses");
  fill("fOwner", unique(DATA.projects.map(p => p.IDOwner)), "All owners");
}
function fill(id, values, allLabel) {
  const el = $(id), cur = el.value;
  el.innerHTML = `<option value="">${allLabel}</option>` + values.map(v => `<option value="${esc(v)}">${esc(id === "fPortco" ? String(v).toUpperCase() : v)}</option>`).join("");
  if (values.includes(cur)) el.value = cur;
}
const F = () => ({ portco: $("fPortco").value, status: $("fStatus").value, owner: $("fOwner").value, q: $("fSearch").value.trim().toLowerCase() });
function filteredProjects() {
  const f = F();
  return DATA.projects.filter(p => (!f.portco || p.PortcoName === f.portco) && (!f.status || p.Status === f.status) &&
    (!f.owner || p.IDOwner === f.owner) && (!f.q || `${p.ProjectName} ${p.ProjectID}`.toLowerCase().includes(f.q)));
}

/* ---------- Render ---------- */
function render() {
  const projects = filteredProjects();
  const ids = new Set(projects.map(p => p.ProjectID));
  const f = F(), narrowed = f.status || f.owner || f.q;
  const comms = DATA.comms.filter(c => (!f.portco || c.PortcoName === f.portco) && (!narrowed || ids.has(c.ProjectID)));
  const files = DATA.files.filter(x => (!f.portco || x.PortcoName === f.portco) && (!narrowed || ids.has(x.ProjectID)));
  renderKPIs(projects, comms, files);
  renderStatus(projects);
  renderPortcoBars(projects);
  renderDelivery(projects);
  renderFunnel(projects);
  renderComms(comms);
  renderOwner(projects);
  renderPortcoCards();
  renderProjectTable(projects);
  renderCommsTable(comms);
  renderFilesTable(files);
}

function renderKPIs(projects, comms, files) {
  const active = projects.filter(p => !isClosed(p)).length;
  const overdue = projects.filter(p => health(p) === "red").length;
  const atRisk = projects.filter(p => health(p) === "amber").length;
  const avgFb = projects.length ? (projects.reduce((a, p) => a + p.FeedbackRounds, 0) / projects.length).toFixed(1) : "0";
  const sent = comms.filter(c => c.Direction === "Sent").length;
  const k = [
    ["PORTCOS", unique(projects.map(p => p.PortcoName)).length, "with projects", "gold"],
    ["TOTAL PROJECTS", projects.length, `${active} active`, ""],
    ["AT RISK", atRisk, `due ≤ ${CONFIG.dueSoonDays} days`, "amber"],
    ["OVERDUE", overdue, "past due date", "red"],
    ["COMMUNICATIONS", comms.length, `${sent} sent · ${comms.length - sent} received`, ""],
    ["FILES STORED", files.length, "scripts → final content", ""],
    ["AVG FEEDBACK ROUNDS", avgFb, "per project", "gold"]
  ];
  $("kpis").innerHTML = k.map(([l, v, n, c]) => `<div class="kpi ${c}"><div class="l">${l}</div><div class="v">${v}</div><div class="n">${n}</div></div>`).join("");
}

function renderStatus(projects) {
  const st = unique([...CONFIG.statusOrder, ...projects.map(p => p.Status)]).filter(s => projects.some(p => p.Status === s));
  const colors = st.map(s => C.status[CONFIG.statusOrder.indexOf(s)] || C.gold3);
  const counts = st.map(s => projects.filter(p => p.Status === s).length);
  drawChart("statusChart", "doughnut", { labels: st, datasets: [{ data: counts, backgroundColor: colors, borderColor: C.card, borderWidth: 4, spacing: 2 }] },
    { cutout: "80%", plugins: { legend: { display: false } } });
  $("statusCenter").textContent = projects.length;
  $("statusLegend").innerHTML = st.map((s, i) => `<li><span class="sw" style="background:${colors[i]}"></span>${esc(s)} – ${counts[i]}</li>`).join("") || `<li class="empty">No projects</li>`;
}

function renderPortcoBars(projects) {
  const pcs = unique(projects.map(p => p.PortcoName));
  const active = pcs.map(pc => projects.filter(p => p.PortcoName === pc && !isClosed(p)).length);
  const done = pcs.map(pc => projects.filter(p => p.PortcoName === pc && isClosed(p)).length);
  drawChart("portcoChart", "bar", {
    labels: pcs.map(shorten),
    datasets: [
      { label: "Active", data: active, backgroundColor: C.gold, borderRadius: 20, barThickness: 10 },
      { label: "Delivered / on hold", data: done, backgroundColor: C.gold3, borderRadius: 20, barThickness: 10 }
    ]
  }, { indexAxis: "y", scales: { x: { stacked: true, ticks: { precision: 0 }, grid: { color: C.line } }, y: { stacked: true, grid: { display: false } } },
       plugins: { legend: { position: "bottom", labels: { boxWidth: 10, boxHeight: 10, useBorderRadius: true, borderRadius: 3 } } } });
  $("activeTotal").textContent = active.reduce((a, b) => a + b, 0);
}

function renderDelivery(projects) {
  const delivered = projects.filter(p => health(p) === "gold").length;
  const pct = projects.length ? Math.round(delivered / projects.length * 100) : 0;
  drawChart("deliveryChart", "doughnut", {
    datasets: [{ data: [pct, 100 - pct], backgroundColor: [C.gold, C.track], borderWidth: 0, borderRadius: [20, 0] }]
  }, { cutout: "82%", rotation: 0, plugins: { legend: { display: false }, tooltip: { enabled: false } } });
  $("deliveryCenter").textContent = pct + "%";
  $("deliveredOf").textContent = `${delivered} of ${projects.length}`;
}

function renderFunnel(projects) {
  const stages = CONFIG.statusOrder.filter(s => s !== "On Hold");
  const total = projects.length || 1;
  const counts = stages.map(s => projects.filter(p => p.Status === s).length);
  const max = Math.max(1, ...counts);
  const lead = counts.indexOf(max);
  $("funnel").innerHTML = stages.map((s, i) => {
    const w = Math.round(counts[i] / max * 100);
    const inStage = projects.filter(p => p.Status === s);
    const isDone = s === "Delivered";
    const avgDays = isDone ? avgDuration(inStage) : avgAge(inStage);
    return `<div class="f-row">
      <div class="f-stage ${i === lead && counts[i] ? "lead" : ""}"><div class="fill" style="width:${i === lead ? 100 : 0}%;"></div>
        ${i !== lead ? `<div class="fill" style="width:${w}%;opacity:.18"></div>` : ""}
        <div class="t">${esc(s)}</div><div class="s">${avgDays !== null ? `avg ${avgDays} days ${isDone ? "to deliver" : "since script"}` : "–"}</div></div>
      <div class="f-cell">${counts[i]}</div>
      <div class="f-cell">${Math.round(counts[i] / total * 100)}%</div>
    </div>`;
  }).join("");
}
/* Delivered projects: days from script received to completion (latest end date) */
function avgDuration(list) {
  const ds = list.filter(p => p.ScriptReceived && p.DueDate).map(p => (p.DueDate - p.ScriptReceived) / 86400000);
  if (!ds.length) return null;
  return Math.round(ds.reduce((a, d) => a + d, 0) / ds.length);
}
function avgAge(list) {
  const ds = list.map(p => p.ScriptReceived).filter(Boolean);
  if (!ds.length) return null;
  return Math.round(ds.reduce((a, d) => a + (today() - d) / 86400000, 0) / ds.length);
}

function renderComms(comms) {
  const by = {};
  comms.forEach(c => { if (!c.Date) return; const k = `${c.Date.getFullYear()}-${String(c.Date.getMonth() + 1).padStart(2, "0")}`; by[k] = by[k] || { s: 0, r: 0 }; c.Direction === "Received" ? by[k].r++ : by[k].s++; });
  const m = Object.keys(by).sort();
  const lbl = (k) => new Date(k + "-01").toLocaleDateString("en-GB", { month: "short", year: "2-digit" });
  drawChart("commsChart", "bar", {
    labels: m.map(lbl),
    datasets: [
      { label: "Sent", data: m.map(k => by[k].s), backgroundColor: C.gold, borderRadius: 6, maxBarThickness: 34 },
      { label: "Received", data: m.map(k => by[k].r), backgroundColor: C.gold3, borderRadius: 6, maxBarThickness: 34 }
    ]
  }, { scales: { y: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: C.line } }, x: { grid: { display: false } } },
       plugins: { legend: { position: "bottom", labels: { boxWidth: 10, boxHeight: 10 } } } });
}

function renderOwner(projects) {
  const act = projects.filter(p => !isClosed(p));
  const o = unique(act.map(p => p.IDOwner));
  drawChart("ownerChart", "bar", {
    labels: o, datasets: [{ label: "Active projects", data: o.map(x => act.filter(p => p.IDOwner === x).length), backgroundColor: C.gold, borderRadius: 8, maxBarThickness: 60 }]
  }, { scales: { y: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: C.line } }, x: { grid: { display: false } } }, plugins: { legend: { display: false } } });
}

function drawChart(id, type, data, options) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart($(id), { type, data, options: { responsive: true, maintainAspectRatio: false, animation: { duration: 600 }, ...options } });
}

function renderPortcoCards() {
  const sel = $("fPortco").value;
  $("portcoCards").innerHTML = DATA.portcos.map(pc => {
    const ps = DATA.projects.filter(p => p.PortcoName === pc.PortcoName);
    const active = ps.filter(p => !isClosed(p)).length;
    const overdue = ps.filter(p => health(p) === "red").length;
    const cm = DATA.comms.filter(c => c.PortcoName === pc.PortcoName);
    const last = cm.map(c => c.Date).filter(Boolean).sort((a, b) => b - a)[0];
    const avg = ps.length ? Math.round(ps.reduce((a, p) => a + p.Progress, 0) / ps.length) : 0;
    return `<div class="pc ${sel === pc.PortcoName ? "active" : ""}" data-pc="${esc(pc.PortcoName)}">
      <h3>${esc(pc.PortcoName)}</h3><div class="ind">${esc(pc.Industry || "")}</div>
      <div class="stats"><div><b>${ps.length}</b>Projects</div><div><b>${active}</b>Active</div>
        <div><b class="${overdue ? "alert" : ""}">${overdue}</b>Overdue</div><div><b>${cm.length}</b>Comms</div></div>
      <div class="bar"><span style="width:${avg}%"></span></div>
      <div class="meta">Avg progress ${avg}% · SPOC ${esc(pc.SPOC || "–")}<br>Last contact ${fmt(last)}</div>
    </div>`;
  }).join("") || `<div class="empty">No portcos found in the Excel file.</div>`;
  if (SESSION && SESSION.role === "portco") return;
  document.querySelectorAll(".pc").forEach(el => el.onclick = () => { $("fPortco").value = $("fPortco").value === el.dataset.pc ? "" : el.dataset.pc; render(); });
}

function renderProjectTable(projects) {
  const rank = { red: 0, amber: 1, green: 2, grey: 3, gold: 4 };
  const rows = [...projects].sort((a, b) => rank[health(a)] - rank[health(b)] || (a.DueDate || 0) - (b.DueDate || 0));
  $("projectTable").querySelector("tbody").innerHTML = rows.map(p => {
    const h = health(p);
    return `<tr><td><i class="dot ${h}"></i>${HEALTH_LABEL[h]}</td>
      <td><strong>${esc(p.ProjectName)}</strong><br><small>${esc(p.ProjectID)}</small></td>
      <td>${esc(p.PortcoName)}</td><td>${esc(p.ContentType)}</td><td>${esc(p.IDOwner)}</td>
      <td><span class="pill st">${esc(p.Status)}</span></td>
      <td><div class="prog"><div class="bar"><span style="width:${p.Progress}%"></span></div>${p.Progress}%</div></td>
      <td>${fmt(p.DueDate)}</td><td style="text-align:center">${p.FeedbackRounds}</td>
      <td><span class="pill ${esc(p.Priority)}">${esc(p.Priority || "–")}</span></td></tr>`;
  }).join("") || `<tr><td colspan="10" class="empty">No projects match the filters.</td></tr>`;
}
function renderCommsTable(comms) {
  $("commsTable").querySelector("tbody").innerHTML = [...comms].sort((a, b) => (b.Date || 0) - (a.Date || 0)).map(c => `<tr>
    <td>${fmt(c.Date)}</td><td>${esc(c.PortcoName)}</td><td>${esc(c.ProjectID)}</td><td>${esc(c.Type)}</td>
    <td><span class="pill ${c.Direction === "Sent" ? "st" : ""}">${esc(c.Direction)}</span></td><td>${esc(c.Subject)}</td></tr>`).join("")
    || `<tr><td colspan="6" class="empty">No communications logged.</td></tr>`;
}
function renderFilesTable(files) {
  $("filesTable").querySelector("tbody").innerHTML = [...files].sort((a, b) => (b.UploadedOn || 0) - (a.UploadedOn || 0)).map(f => `<tr>
    <td>${f.Link ? `<a href="${esc(f.Link)}" target="_blank" rel="noopener">${esc(f.FileName)}</a>` : esc(f.FileName)}</td>
    <td>${esc(f.PortcoName)}</td><td>${esc(f.FileType)}</td><td>${esc(f.Version)}</td><td>${fmt(f.UploadedOn)}</td></tr>`).join("")
    || `<tr><td colspan="5" class="empty">No files logged.</td></tr>`;
}

const shorten = (s) => String(s).length > 20 ? String(s).slice(0, 19) + "…" : s;
function esc(v) { return String(v ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

/* ---------- Wire up ---------- */
["fPortco", "fStatus", "fOwner"].forEach(id => $(id).addEventListener("change", render));
$("fSearch").addEventListener("input", render);
$("clearBtn").onclick = () => { ["fPortco", "fStatus", "fOwner", "fSearch"].forEach(id => $(id).value = ""); render(); };
$("refreshBtn").onclick = () => { usingLocalFile = false; loadFromRepo(true); };
["fileInput", "fileInputTop"].forEach(id => $(id).onchange = (e) => { if (e.target.files[0]) loadFromFile(e.target.files[0]); e.target.value = ""; });
document.querySelectorAll("nav a").forEach(a => a.addEventListener("click", () => {
  document.querySelectorAll("nav a").forEach(x => x.classList.remove("active")); a.classList.add("active");
}));

$("themeToggle").onclick = () => setTheme(theme === "dark" ? "light" : "dark");
setTheme(theme);

/* ---------- Login screen ---------- */
const LOCK = { max: 5, seconds: 30 };
let pending = null;   // { role, id, name, file }

async function initLogin() {
  try {
    const res = await fetch(`portcos.json?v=${Date.now()}`, { cache: "no-store" });
    if (!res.ok) throw new Error(res.status);
    PORTCO_LIST = await res.json();
  } catch (e) {
    $("pcList").innerHTML = `<div class="empty">The dashboard hasn't been published with access settings yet.</div>`;
    return;
  }
  renderPortcoList();
  // Resume a session in this browser tab
  try {
    const s = JSON.parse(sessionStorage.getItem("plh-session") || "null");
    if (s) {
      const key = await crypto.subtle.importKey("raw", b64(s.key), { name: "AES-GCM" }, true, ["decrypt"]);
      await enterDashboard({ ...s, key }, null);
    }
  } catch (e) { sessionStorage.removeItem("plh-session"); }
}

function renderPortcoList() {
  const q = $("pcSearch").value.trim().toLowerCase();
  const list = PORTCO_LIST.filter(p => !q || p.name.toLowerCase().includes(q));
  $("pcList").innerHTML = list.map(p => `
    <button class="login-item" data-id="${esc(p.id)}">
      <span class="li-icon">${esc(initials(p.name))}</span>
      <span class="li-text"><b>${esc(p.name)}</b><small>${esc(p.industry || "Portfolio company")}</small></span>
      <span class="li-go">&#8250;</span>
    </button>`).join("") || `<div class="empty">No portco matches “${esc(q)}”.</div>`;
  document.querySelectorAll("#pcList .login-item").forEach(el => el.onclick = () => {
    const p = PORTCO_LIST.find(x => x.id === el.dataset.id);
    openModal({ role: "portco", id: p.id, name: p.name, file: p.file });
  });
}
const initials = (s) => String(s).split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join("").toUpperCase();

function openModal(target) {
  pending = target;
  const admin = target.role === "admin";
  $("mTitle").textContent = admin ? "Admin login" : "Enter your PIN";
  $("mSub").textContent = admin ? "REVANTAGE — ALL PORTCOS" : `PORTCO — ${target.name.toUpperCase()}`;
  $("mLabel").textContent = admin ? "Admin password" : "PIN (digits only, up to 4)";
  $("pinInput").hidden = admin; $("adminInput").hidden = !admin;
  $("pinInput").value = ""; $("adminInput").value = ""; $("mError").textContent = "";
  $("pinModal").classList.remove("hidden");
  getBlob(target.file).catch(() => {});   // load early so the check is instant
  setTimeout(() => (admin ? $("adminInput") : $("pinInput")).focus(), 50);
  checkLock();
}
function closeModal() { $("pinModal").classList.add("hidden"); pending = null; clearTimeout(autoTimer); setChecking(false); }

function lockState() { try { return JSON.parse(sessionStorage.getItem("plh-lock") || "{}"); } catch (e) { return {}; } }
function checkLock() {
  const st = lockState()[pending ? pending.id : ""] || {};
  const left = Math.ceil(((st.until || 0) - Date.now()) / 1000);
  if (left > 0) {
    $("mGo").disabled = true;
    $("mError").textContent = `Too many attempts. Try again in ${left}s.`;
    setTimeout(checkLock, 1000);
    return true;
  }
  $("mGo").disabled = false;
  if ($("mError").textContent.startsWith("Too many")) $("mError").textContent = "";
  return false;
}
function recordFail() {
  const all = lockState(), st = all[pending.id] || { n: 0 };
  st.n = (st.n || 0) + 1;
  if (st.n >= LOCK.max) { st.until = Date.now() + LOCK.seconds * 1000; st.n = 0; }
  all[pending.id] = st; sessionStorage.setItem("plh-lock", JSON.stringify(all));
  return LOCK.max - st.n;
}

/* Unlocks as soon as the right code is typed – no need to press "Unlock & view".
   While typing, each code is checked quietly (no error, no attempt used). A wrong PIN only
   counts once all 4 digits are typed, or when Enter / the button is pressed. */
const blobCache = {};
let checkSeq = 0, autoTimer = null;

async function getBlob(file) {
  if (!blobCache[file]) blobCache[file] = fetchEnc(file).catch(e => { delete blobCache[file]; throw e; });
  return blobCache[file];
}

function scheduleAutoCheck() {
  clearTimeout(autoTimer);
  if (!pending) return;
  const admin = pending.role === "admin";
  const value = admin ? $("adminInput").value : $("pinInput").value;
  $("mError").textContent = "";
  if (!value) { setChecking(false); return; }
  if (!admin && value.length === 4) { submitLogin(); return; }   // full PIN: real attempt
  autoTimer = setTimeout(() => submitLogin({ silent: true }), admin ? 450 : 250);
}

function setChecking(on) {
  $("mCheck").textContent = on ? "Checking…" : "";
}

async function submitLogin(opts = {}) {
  const silent = !!opts.silent;
  if (!pending || checkLock()) return;
  const target = pending;
  const admin = target.role === "admin";
  const secret = admin ? $("adminInput").value : $("pinInput").value.trim();
  if (!admin && !/^\d{1,4}$/.test(secret)) { if (!silent) $("mError").textContent = "PIN must be 1–4 digits."; return; }
  if (admin && !secret) { if (!silent) $("mError").textContent = "Enter the admin password."; return; }
  const seq = ++checkSeq;
  setChecking(true);
  if (!silent) { $("mGo").disabled = true; $("mGo").textContent = "Verifying…"; $("mError").textContent = ""; }
  try {
    const blob = await getBlob(target.file);
    const key = await deriveKey(secret, blob.salt, blob.iter);
    const data = await decryptWith(key, blob);
    if (pending !== target) return;                      // pop-up was closed meanwhile
    const all = lockState(); delete all[target.id]; sessionStorage.setItem("plh-lock", JSON.stringify(all));
    await enterDashboard({ ...target, key }, { data, blob });
    closeModal();
  } catch (e) {
    if (seq !== checkSeq || pending !== target) return;  // a newer check is running
    if (String(e.name) === "OperationError") {
      if (silent) return;                                // quiet check: keep typing
      const left = recordFail();
      $("mError").textContent = admin ? "Incorrect password." : `Incorrect PIN.${left < LOCK.max && left > 0 ? ` ${left} attempt${left === 1 ? "" : "s"} left.` : ""}`;
      if (!admin) { $("pinInput").value = ""; $("pinInput").focus(); }
      checkLock();
    } else {
      console.warn(e);
      $("mError").textContent = "Couldn't reach the dashboard data. Check your connection and try again.";
    }
  } finally {
    if (seq === checkSeq) setChecking(false);
    if (!silent) { $("mGo").textContent = "Unlock & view"; if (!checkLock()) $("mGo").disabled = false; }
  }
}

async function enterDashboard(session, first) {
  SESSION = session;
  const raw = await crypto.subtle.exportKey("raw", session.key);
  sessionStorage.setItem("plh-session", JSON.stringify({ role: session.role, id: session.id, name: session.name, file: session.file, key: toB64(raw) }));
  document.body.classList.remove("locked");
  document.body.classList.toggle("portco-mode", session.role === "portco");
  $("viewBadge").textContent = session.role === "admin" ? "ADMIN · ALL PORTCOS" : session.name.toUpperCase();
  $("pageTitle").textContent = session.role === "admin" ? "DASHBOARD" : `${session.name.toUpperCase()} DASHBOARD`;
  if (first) { lastSignature = first.blob.ct; usingLocalFile = false; ingest(first.data, "Secure data"); }
  else await loadSecure(true);
}

function logout(message) {
  SESSION = null; lastSignature = ""; usingLocalFile = false;
  DATA = { portcos: [], projects: [], comms: [], files: [] };
  sessionStorage.removeItem("plh-session");
  ["fPortco", "fStatus", "fOwner", "fSearch"].forEach(id => $(id).value = "");
  document.body.classList.add("locked"); document.body.classList.remove("portco-mode");
  $("loginNote").textContent = message || "";
  window.scrollTo(0, 0);
}

$("pcSearch").addEventListener("input", renderPortcoList);
$("adminItem").onclick = () => openModal({ role: "admin", id: "admin", name: "Admin", file: "admin.json" });
$("mClose").onclick = closeModal; $("mCancel").onclick = closeModal;
$("mGo").onclick = () => submitLogin();
$("pinInput").addEventListener("input", (e) => { e.target.value = e.target.value.replace(/\D/g, "").slice(0, 4); scheduleAutoCheck(); });
$("adminInput").addEventListener("input", scheduleAutoCheck);
["pinInput", "adminInput"].forEach(id => $(id).addEventListener("keydown", (e) => { if (e.key === "Enter") { clearTimeout(autoTimer); submitLogin(); } if (e.key === "Escape") closeModal(); }));
$("pinModal").addEventListener("click", (e) => { if (e.target.id === "pinModal") closeModal(); });
$("logoutBtn").onclick = () => logout();
$("loginTheme").onclick = () => setTheme(theme === "dark" ? "light" : "dark");

initLogin();
setInterval(() => loadSecure(false), CONFIG.autoRefreshSeconds * 1000);
