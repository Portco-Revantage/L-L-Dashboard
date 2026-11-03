/* PortCo Learning Hub – reads data.xlsx and renders the dashboard.
   Update the Excel file, commit it to the repo, and the dashboard refreshes automatically. */

const CONFIG = {
  dataFile: "data.xlsx",        // Excel file in the repo root (or a direct-download URL)
  autoRefreshSeconds: 60,       // how often to check for a new version of the Excel file
  dueSoonDays: 7,               // "amber" window
  statusOrder: ["Script Received", "In Development", "In Review", "Portco Feedback", "Delivered", "On Hold"],
  statusColors: ["#9AA3B2", "#3B7DD8", "#E8A33D", "#8E6CD8", "#2E9E6B", "#C9CED8"]
};

let DATA = { portcos: [], projects: [], comms: [], files: [] };
let charts = {};
let lastSignature = "";
let usingLocalFile = false;

const $ = (id) => document.getElementById(id);
const today = () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; };

/* ---------- Load Excel ---------- */
async function loadFromRepo(force = false) {
  if (usingLocalFile && !force) return;
  try {
    const res = await fetch(`${CONFIG.dataFile}?v=${Date.now()}`, { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buf = await res.arrayBuffer();
    const sig = await signature(buf);
    if (sig === lastSignature && !force) { setSync("Up to date", true); return; }
    lastSignature = sig;
    usingLocalFile = false;
    parseWorkbook(buf, CONFIG.dataFile);
  } catch (err) {
    console.warn(err);
    setSync(location.protocol === "file:"
      ? "Open via GitHub Pages or use 'Load Excel'"
      : "Couldn't read data.xlsx – use 'Load Excel'", false);
  }
}

async function signature(buf) {
  if (crypto && crypto.subtle) {
    const h = await crypto.subtle.digest("SHA-1", buf);
    return Array.from(new Uint8Array(h)).map(b => b.toString(16).padStart(2, "0")).join("");
  }
  return String(buf.byteLength);
}

function loadFromFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => { usingLocalFile = true; parseWorkbook(e.target.result, file.name); };
  reader.readAsArrayBuffer(file);
}

function parseWorkbook(buf, name) {
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheet = (n) => {
    const key = wb.SheetNames.find(s => s.toLowerCase() === n.toLowerCase());
    return key ? XLSX.utils.sheet_to_json(wb.Sheets[key], { defval: "" }) : [];
  };
  DATA.portcos = sheet("Portcos");
  DATA.projects = sheet("Projects").filter(r => r.ProjectID || r.ProjectName).map(r => ({
    ...r,
    ScriptReceived: toDate(r.ScriptReceived), StartDate: toDate(r.StartDate), DueDate: toDate(r.DueDate),
    Progress: Math.max(0, Math.min(100, Number(r.Progress) || 0)),
    FeedbackRounds: Number(r.FeedbackRounds) || 0
  }));
  DATA.comms = sheet("Communications").filter(r => r.Date || r.Subject).map(r => ({ ...r, Date: toDate(r.Date) }));
  DATA.files = sheet("Files").filter(r => r.FileName).map(r => ({ ...r, UploadedOn: toDate(r.UploadedOn) }));

  // Portcos that appear only in Projects still get a card
  const known = new Set(DATA.portcos.map(p => p.PortcoName));
  DATA.projects.forEach(p => { if (p.PortcoName && !known.has(p.PortcoName)) { DATA.portcos.push({ PortcoName: p.PortcoName }); known.add(p.PortcoName); } });

  $("sourceName").textContent = name;
  $("lastRefresh").textContent = new Date().toLocaleString();
  setSync(usingLocalFile ? "Showing uploaded file" : "Live · synced with Excel", true);
  buildFilterOptions();
  render();
}

function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === "number") return new Date(Math.round((v - 25569) * 86400 * 1000));
  const d = new Date(v);
  return isNaN(d) ? null : d;
}
const fmt = (d) => d ? d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "–";

function setSync(text, ok) {
  const el = $("syncStatus");
  el.textContent = (ok ? "● " : "⚠ ") + text;
  el.style.color = ok ? "#BFF5D8" : "#FFE0A3";
}

/* ---------- Health (RAG) ---------- */
function health(p) {
  const s = String(p.Status).toLowerCase();
  if (s === "delivered") return "blue";
  if (s === "on hold") return "grey";
  if (!p.DueDate) return "green";
  const days = (p.DueDate - today()) / 86400000;
  if (days < 0) return "red";
  if (days <= CONFIG.dueSoonDays && p.Progress < 70) return "amber";
  return "green";
}
const healthLabel = { green: "On track", amber: "At risk", red: "Overdue", blue: "Delivered", grey: "On hold" };

/* ---------- Filters ---------- */
function buildFilterOptions() {
  fill("fPortco", unique(DATA.portcos.map(p => p.PortcoName)), "All portcos");
  const statuses = unique([...CONFIG.statusOrder.filter(s => DATA.projects.some(p => p.Status === s)), ...DATA.projects.map(p => p.Status)]);
  fill("fStatus", statuses, "All statuses");
  fill("fOwner", unique(DATA.projects.map(p => p.IDOwner)), "All owners");
}
function fill(id, values, allLabel) {
  const el = $(id), cur = el.value;
  el.innerHTML = `<option value="">${allLabel}</option>` + values.map(v => `<option>${esc(v)}</option>`).join("");
  if (values.includes(cur)) el.value = cur;
}
const unique = (arr) => [...new Set(arr.filter(Boolean))];
function filters() {
  return { portco: $("fPortco").value, status: $("fStatus").value, owner: $("fOwner").value, q: $("fSearch").value.trim().toLowerCase() };
}
function filteredProjects() {
  const f = filters();
  return DATA.projects.filter(p =>
    (!f.portco || p.PortcoName === f.portco) &&
    (!f.status || p.Status === f.status) &&
    (!f.owner || p.IDOwner === f.owner) &&
    (!f.q || `${p.ProjectName} ${p.ProjectID}`.toLowerCase().includes(f.q)));
}

/* ---------- Render ---------- */
function render() {
  const projects = filteredProjects();
  const ids = new Set(projects.map(p => p.ProjectID));
  const f = filters();
  const narrowed = f.status || f.owner || f.q;
  const comms = DATA.comms.filter(c => (!f.portco || c.PortcoName === f.portco) && (!narrowed || ids.has(c.ProjectID)));
  const files = DATA.files.filter(x => (!f.portco || x.PortcoName === f.portco) && (!narrowed || ids.has(x.ProjectID)));

  renderKPIs(projects, comms, files);
  renderCharts(projects, comms);
  renderPortcoCards();
  renderProjectTable(projects);
  renderComms(comms);
  renderFiles(files);
}

function renderKPIs(projects, comms, files) {
  const active = projects.filter(p => !["delivered", "on hold"].includes(String(p.Status).toLowerCase()));
  const delivered = projects.filter(p => health(p) === "blue").length;
  const overdue = projects.filter(p => health(p) === "red").length;
  const atRisk = projects.filter(p => health(p) === "amber").length;
  const portcoCount = unique(projects.map(p => p.PortcoName)).length;
  const avgFb = projects.length ? (projects.reduce((a, p) => a + p.FeedbackRounds, 0) / projects.length).toFixed(1) : "0";
  const kpis = [
    { label: "Portcos", value: portcoCount, note: "with projects", cls: "" },
    { label: "Total projects", value: projects.length, note: `${active.length} active`, cls: "blue" },
    { label: "Delivered", value: delivered, note: projects.length ? `${Math.round(delivered / projects.length * 100)}% of total` : "", cls: "green" },
    { label: "At risk", value: atRisk, note: `due ≤ ${CONFIG.dueSoonDays} days`, cls: "amber" },
    { label: "Overdue", value: overdue, note: "past due date", cls: "red" },
    { label: "Communications", value: comms.length, note: `${comms.filter(c => c.Direction === "Sent").length} sent · ${comms.filter(c => c.Direction === "Received").length} received`, cls: "" },
    { label: "Files stored", value: files.length, note: "scripts, storyboards, finals", cls: "" },
    { label: "Avg feedback rounds", value: avgFb, note: "per project", cls: "" }
  ];
  $("kpis").innerHTML = kpis.map(k => `<div class="kpi ${k.cls}"><div class="label">${k.label}</div><div class="value">${k.value}</div><div class="note">${k.note}</div></div>`).join("");
}

function renderCharts(projects, comms) {
  // Status doughnut
  const statuses = unique([...CONFIG.statusOrder, ...projects.map(p => p.Status)]).filter(s => projects.some(p => p.Status === s));
  const sColors = statuses.map(s => CONFIG.statusColors[CONFIG.statusOrder.indexOf(s)] || "#6b7385");
  drawChart("statusChart", "doughnut", {
    labels: statuses,
    datasets: [{ data: statuses.map(s => projects.filter(p => p.Status === s).length), backgroundColor: sColors, borderWidth: 2 }]
  }, { plugins: { legend: { position: "right" } }, cutout: "60%" });

  // Projects per portco – stacked by health
  const pcs = unique(projects.map(p => p.PortcoName));
  const groups = [["Delivered", "blue", "#3B7DD8"], ["On track", "green", "#2E9E6B"], ["At risk", "amber", "#E8A33D"], ["Overdue", "red", "#D64545"], ["On hold", "grey", "#9AA3B2"]];
  drawChart("portcoChart", "bar", {
    labels: pcs,
    datasets: groups.map(([label, h, color]) => ({ label, backgroundColor: color, data: pcs.map(pc => projects.filter(p => p.PortcoName === pc && health(p) === h).length) }))
  }, { indexAxis: "y", scales: { x: { stacked: true, ticks: { precision: 0 } }, y: { stacked: true } }, plugins: { legend: { position: "bottom" } } });

  // Comms per month
  const byMonth = {};
  comms.forEach(c => { if (!c.Date) return; const k = `${c.Date.getFullYear()}-${String(c.Date.getMonth() + 1).padStart(2, "0")}`; byMonth[k] = byMonth[k] || { Sent: 0, Received: 0 }; byMonth[k][c.Direction === "Received" ? "Received" : "Sent"]++; });
  const months = Object.keys(byMonth).sort();
  const label = (k) => new Date(k + "-01").toLocaleDateString("en-GB", { month: "short", year: "2-digit" });
  drawChart("commsChart", "bar", {
    labels: months.map(label),
    datasets: [
      { label: "Sent", data: months.map(m => byMonth[m].Sent), backgroundColor: "#1F3A5F" },
      { label: "Received", data: months.map(m => byMonth[m].Received), backgroundColor: "#7FA7D9" }
    ]
  }, { scales: { y: { beginAtZero: true, ticks: { precision: 0 } } }, plugins: { legend: { position: "bottom" } } });

  // Owner workload
  const active = projects.filter(p => !["delivered", "on hold"].includes(String(p.Status).toLowerCase()));
  const owners = unique(active.map(p => p.IDOwner));
  drawChart("ownerChart", "bar", {
    labels: owners,
    datasets: [{ label: "Active projects", data: owners.map(o => active.filter(p => p.IDOwner === o).length), backgroundColor: "#2E5A8A", borderRadius: 6 }]
  }, { scales: { y: { beginAtZero: true, ticks: { precision: 0 } } }, plugins: { legend: { display: false } } });
}

function drawChart(id, type, data, options) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart($(id), { type, data, options: { responsive: true, maintainAspectRatio: false, ...options } });
}

function renderPortcoCards() {
  const sel = $("fPortco").value;
  $("portcoCards").innerHTML = DATA.portcos.map(pc => {
    const ps = DATA.projects.filter(p => p.PortcoName === pc.PortcoName);
    const active = ps.filter(p => !["delivered", "on hold"].includes(String(p.Status).toLowerCase())).length;
    const overdue = ps.filter(p => health(p) === "red").length;
    const comms = DATA.comms.filter(c => c.PortcoName === pc.PortcoName);
    const last = comms.map(c => c.Date).filter(Boolean).sort((a, b) => b - a)[0];
    const avg = ps.length ? Math.round(ps.reduce((a, p) => a + p.Progress, 0) / ps.length) : 0;
    return `<div class="pc ${sel === pc.PortcoName ? "active" : ""}" data-pc="${esc(pc.PortcoName)}">
      <h3>${esc(pc.PortcoName)}</h3><div class="ind">${esc(pc.Industry || "")}</div>
      <div class="stats">
        <div><b>${ps.length}</b>Projects</div><div><b>${active}</b>Active</div>
        <div><b style="color:${overdue ? "var(--red)" : "inherit"}">${overdue}</b>Overdue</div><div><b>${comms.length}</b>Comms</div>
      </div>
      <div class="bar"><span style="width:${avg}%"></span></div>
      <div class="spoc">Avg progress ${avg}% · SPOC: ${esc(pc.SPOC || "–")} · Last contact: ${fmt(last)}</div>
    </div>`;
  }).join("") || `<div class="empty">No portcos found in the Excel file.</div>`;
  document.querySelectorAll(".pc").forEach(el => el.onclick = () => {
    $("fPortco").value = $("fPortco").value === el.dataset.pc ? "" : el.dataset.pc; render();
  });
}

function renderProjectTable(projects) {
  const rank = { red: 0, amber: 1, green: 2, grey: 3, blue: 4 };
  const rows = [...projects].sort((a, b) => rank[health(a)] - rank[health(b)] || (a.DueDate || 0) - (b.DueDate || 0));
  $("projectTable").querySelector("tbody").innerHTML = rows.map(p => {
    const h = health(p);
    return `<tr>
      <td><span class="dot ${h}" title="${healthLabel[h]}"></span> ${healthLabel[h]}</td>
      <td><strong>${esc(p.ProjectName)}</strong><br><small style="color:var(--muted)">${esc(p.ProjectID)}</small></td>
      <td>${esc(p.PortcoName)}</td><td>${esc(p.ContentType)}</td><td>${esc(p.IDOwner)}</td>
      <td><span class="pill">${esc(p.Status)}</span></td>
      <td><div class="prog"><div class="bar"><span style="width:${p.Progress}%"></span></div>${p.Progress}%</div></td>
      <td>${fmt(p.DueDate)}</td><td style="text-align:center">${p.FeedbackRounds}</td>
      <td><span class="pill ${esc(p.Priority)}">${esc(p.Priority || "–")}</span></td>
    </tr>`;
  }).join("") || `<tr><td colspan="10" class="empty">No projects match the filters.</td></tr>`;
}

function renderComms(comms) {
  const rows = [...comms].sort((a, b) => (b.Date || 0) - (a.Date || 0));
  $("commsTable").querySelector("tbody").innerHTML = rows.map(c => `<tr>
    <td>${fmt(c.Date)}</td><td>${esc(c.PortcoName)}</td><td>${esc(c.ProjectID)}</td><td>${esc(c.Type)}</td>
    <td><span class="pill">${esc(c.Direction)}</span></td><td>${esc(c.Subject)}</td></tr>`).join("")
    || `<tr><td colspan="6" class="empty">No communications logged.</td></tr>`;
}

function renderFiles(files) {
  const rows = [...files].sort((a, b) => (b.UploadedOn || 0) - (a.UploadedOn || 0));
  $("filesTable").querySelector("tbody").innerHTML = rows.map(f => `<tr>
    <td>${f.Link ? `<a href="${esc(f.Link)}" target="_blank" rel="noopener">${esc(f.FileName)}</a>` : esc(f.FileName)}</td>
    <td>${esc(f.PortcoName)}</td><td>${esc(f.FileType)}</td><td>${esc(f.Version)}</td><td>${fmt(f.UploadedOn)}</td></tr>`).join("")
    || `<tr><td colspan="5" class="empty">No files logged.</td></tr>`;
}

function esc(v) {
  return String(v ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

/* ---------- Wire up ---------- */
["fPortco", "fStatus", "fOwner"].forEach(id => $(id).addEventListener("change", render));
$("fSearch").addEventListener("input", render);
$("clearBtn").onclick = () => { ["fPortco", "fStatus", "fOwner", "fSearch"].forEach(id => $(id).value = ""); render(); };
$("refreshBtn").onclick = () => { usingLocalFile = false; loadFromRepo(true); };
$("fileInput").onchange = (e) => { if (e.target.files[0]) loadFromFile(e.target.files[0]); e.target.value = ""; };

loadFromRepo(true);
setInterval(() => loadFromRepo(false), CONFIG.autoRefreshSeconds * 1000);
