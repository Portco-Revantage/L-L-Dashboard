/* PortCo Learning Hub – motion layer (loads after app.js).
   The menu highlight moves like a liquid pill: the leading edge goes first,
   the trailing edge catches up. Plus count-up KPIs, growing bars and
   scroll-in cards. Nothing here changes data or security. */
(function () {
  "use strict";
  const html = document.documentElement;
  const reduce = window.matchMedia && matchMedia("(prefers-reduced-motion: reduce)").matches;
  if (reduce) html.classList.add("motion-off");
  const off = () => html.classList.contains("motion-off");
  const raf = window.requestAnimationFrame.bind(window);

  /* ---------- Tiny tween engine (GSAP-style eases, no library needed) ---------- */
  const EASE = {
    "power3.out": (t) => 1 - Math.pow(1 - t, 3),
    "power3.inOut": (t) => t < .5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2,
    "expo.out": (t) => t === 1 ? 1 : 1 - Math.pow(2, -10 * t)
  };
  const running = new WeakMap();           // target -> { key: tweenId }
  let tweenSeq = 0;
  function to(target, props, { duration = .5, ease = "power3.out", delay = 0, onUpdate } = {}) {
    const ids = running.get(target) || {}; running.set(target, ids);
    const id = ++tweenSeq, from = {};
    Object.keys(props).forEach(k => { from[k] = target[k]; ids[k] = id; });
    if (off() || duration <= 0) { Object.assign(target, props); onUpdate && onUpdate(); return; }
    const fn = EASE[ease] || EASE["power3.out"];
    let start = null;
    const step = (now) => {
      if (start === null) start = now + delay * 1000;
      const p = Math.min(1, Math.max(0, (now - start) / (duration * 1000)));
      const e = fn(p);
      let alive = false;
      Object.keys(props).forEach(k => { if (ids[k] === id) { target[k] = from[k] + (props[k] - from[k]) * e; alive = true; } });
      if (alive) onUpdate && onUpdate();
      if (alive && p < 1) raf(step);
    };
    raf(step);
  }

  /* ---------- Liquid pill ---------- */
  function makePill(container, itemSel, activeSel) {
    container.classList.add("has-pill");
    const pill = document.createElement("span");
    pill.className = "m-pill"; pill.setAttribute("aria-hidden", "true");
    container.insertBefore(pill, container.firstChild);
    let st = null;                                  // current edges {t, b, l, r}
    const paint = () => {
      pill.style.top = st.t + "px"; pill.style.left = st.l + "px";
      pill.style.height = Math.max(0, st.b - st.t) + "px"; pill.style.width = Math.max(0, st.r - st.l) + "px";
    };
    function move(instant) {
      const el = container.querySelector(activeSel);
      if (!el || !el.offsetParent || !container.offsetParent) { if (pill.classList.contains("ready")) pill.classList.remove("ready"); return; }
      const to_ = { t: el.offsetTop, l: el.offsetLeft, b: el.offsetTop + el.offsetHeight, r: el.offsetLeft + el.offsetWidth };
      if (!st || instant || !pill.classList.contains("ready")) { st = { ...to_ }; paint(); if (!pill.classList.contains("ready")) pill.classList.add("ready"); return; }
      const prev = { ...st };
      if (prev.t === to_.t && prev.l === to_.l && prev.b === to_.b && prev.r === to_.r) return;
      // the leading edge goes first, the trailing edge catches up
      const vertical = Math.abs(to_.t - prev.t) > Math.abs(to_.l - prev.l);
      const fwd = vertical ? to_.t > prev.t : to_.l > prev.l;
      const [lead, trail] = vertical ? (fwd ? ["b", "t"] : ["t", "b"]) : (fwd ? ["r", "l"] : ["l", "r"]);
      const cross = vertical ? ["l", "r"] : ["t", "b"];
      to(st, { [lead]: to_[lead] }, { duration: .36, ease: "power3.out", onUpdate: paint });
      to(st, { [trail]: to_[trail] }, { duration: .55, ease: "power3.inOut", onUpdate: paint });
      to(st, { [cross[0]]: to_[cross[0]], [cross[1]]: to_[cross[1]] }, { duration: .45, ease: "power3.inOut", onUpdate: paint });
    }
    // follow the app whenever it changes which item is active
    new MutationObserver((recs) => { if (recs.some(r => r.target !== pill)) move(false); }).observe(container, { subtree: true, attributes: true, attributeFilter: ["class"] });
    window.addEventListener("resize", () => move(true));
    raf(() => move(true));
    return move;
  }

  const nav = document.querySelector(".sidebar nav");
  const navLinks = nav ? [...nav.querySelectorAll("a")] : [];
  navLinks.forEach((a, i) => a.style.setProperty("--i", i));
  const moveNavPill = nav ? makePill(nav, "a", "a.active") : () => {};
  const toggleMovers = [...document.querySelectorAll(".theme-toggle")].map(t => makePill(t, "span", "span.on"));

  // re-measure when the dashboard opens (it is hidden behind the login screen)
  new MutationObserver(() => raf(() => { moveNavPill(true); toggleMovers.forEach(m => m(true)); }))
    .observe(document.body, { attributes: true, attributeFilter: ["class"] });
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(() => { moveNavPill(true); toggleMovers.forEach(m => m(true)); });

  /* ---------- Scroll-spy: the pill follows you down the page ---------- */
  let spyLockUntil = 0;
  navLinks.forEach(a => a.addEventListener("click", () => { spyLockUntil = Date.now() + 1000; }));
  function spy() {
    if (Date.now() < spyLockUntil || document.body.classList.contains("locked")) return;
    const visible = navLinks.filter(a => a.offsetParent);
    let current = visible.find(a => a.getAttribute("href") === "#top");
    let best = -Infinity;
    const atBottom = window.innerHeight + window.scrollY >= document.documentElement.scrollHeight - 4;
    visible.forEach(a => {
      const id = a.getAttribute("href").slice(1);
      if (id === "top") return;
      const sec = document.getElementById(id);
      if (!sec || !sec.offsetParent) return;
      const top = sec.getBoundingClientRect().top;
      if ((top <= 140 || atBottom && top < window.innerHeight) && top > best) { best = top; current = a; }
    });
    if (atBottom) {   // at the very end, pick the lowest section on screen
      let low = -Infinity;
      visible.forEach(a => { const s = document.getElementById(a.getAttribute("href").slice(1)); if (s && s.id !== "top" && s.offsetParent) { const t = s.getBoundingClientRect().top; if (t < window.innerHeight && t > low) { low = t; current = a; } } });
    }
    if (current && !current.classList.contains("active")) {
      navLinks.forEach(x => x.classList.toggle("active", x === current));
    }
  }
  let spyQueued = false;
  window.addEventListener("scroll", () => { if (!spyQueued) { spyQueued = true; raf(() => { spyQueued = false; spy(); }); } }, { passive: true });

  /* ---------- Count-up numbers ---------- */
  const lastVal = {};
  function countUp(el, key, { suffix = "", duration = .9 } = {}) {
    const m = String(el.textContent).trim().match(/^(-?\d+(?:\.\d+)?)(\D*)$/);
    if (!m) { lastVal[key] = undefined; return; }
    const target = parseFloat(m[1]); suffix = m[2] || suffix;
    const from = lastVal[key] !== undefined ? lastVal[key] : 0;
    lastVal[key] = target;
    if (from === target || off()) return;
    const o = { v: from };
    el.textContent = Math.round(from) + suffix;
    to(o, { v: target }, { duration, ease: "expo.out", onUpdate: () => { el.textContent = Math.round(o.v) + suffix; } });
  }

  /* ---------- Hook into the dashboard's drawing functions ---------- */
  function wrap(name, after) {
    const orig = window[name];
    if (typeof orig !== "function") return;
    window[name] = function () { const r = orig.apply(this, arguments); try { after.apply(this, arguments); } catch (e) { console.warn("motion:", e); } return r; };
  }

  wrap("renderKPIs", () => {
    document.querySelectorAll("#kpis .kpi").forEach((k, i) => {
      k.style.setProperty("--i", i);
      const label = (k.querySelector(".l") || {}).textContent || i;
      const v = k.querySelector(".v"); if (v) countUp(v, "kpi:" + label);
    });
  });

  wrap("renderHealth", () => {
    const segs = [...document.querySelectorAll("#healthBar i")];
    const widths = segs.map(s => s.style.width);
    if (!off()) {
      segs.forEach((s, i) => { s.style.setProperty("--d", (i * 90) + "ms"); s.style.width = "0%"; });
      void document.body.offsetWidth;
      raf(() => segs.forEach((s, i) => { s.style.width = widths[i]; }));
    }
    const score = document.getElementById("healthScore"); if (score) countUp(score, "health");
  });

  wrap("renderTimeline", () => {
    const box = document.getElementById("timeline"); if (!box) return;
    box.querySelectorAll(".tl-row").forEach((r, i) => {
      const n = Math.min(i, 20);
      r.style.setProperty("--i", n);
      const bar = r.querySelector(".tl-bar"); if (bar) bar.style.setProperty("--i", n);
    });
    box.classList.remove("tl-anim"); void box.offsetWidth; box.classList.add("tl-anim");
  });

  wrap("renderProjectTable", () => {
    const t = document.getElementById("projectTable"); if (!t) return;
    t.classList.add("m-rows");
    const rows = [...t.querySelectorAll("tbody tr")];
    rows.forEach((r, i) => r.style.setProperty("--i", Math.min(i, 15)));
    const bars = [...t.querySelectorAll(".prog .bar span")];
    const w = bars.map(b => b.style.width);
    if (!off()) {
      bars.forEach((b, i) => { b.style.setProperty("--d", (Math.min(i, 15) * 30 + 200) + "ms"); b.style.width = "0%"; });
      void t.offsetWidth;
      raf(() => bars.forEach((b, i) => { b.style.width = w[i]; }));
    }
  });
  ["renderCommsTable", "renderFilesTable"].forEach(fn => wrap(fn, () => {
    const id = fn === "renderCommsTable" ? "commsTable" : "filesTable";
    const t = document.getElementById(id); if (!t) return;
    t.classList.add("m-rows");
    t.querySelectorAll("tbody tr").forEach((r, i) => r.style.setProperty("--i", Math.min(i, 12)));
  }));

  // Charts: slower, smoother draw; bars grow one after another
  const origDraw = window.drawChart;
  if (typeof origDraw === "function") {
    window.drawChart = function (id, type, data, options) {
      const anim = off() ? false : {
        duration: 1100, easing: "easeOutQuart",
        delay: (ctx) => (type !== "doughnut" && ctx.type === "data" && ctx.mode === "default") ? ctx.dataIndex * 70 + ctx.datasetIndex * 40 : 0
      };
      const extra = type === "doughnut" ? { animation: anim && { ...anim, animateRotate: true, animateScale: true } } : { animation: anim };
      return origDraw.call(this, id, type, data, { ...extra, ...(options || {}) });
    };
  }

  /* ---------- Cards fade up as they come into view (and charts replay) ---------- */
  const replayed = new WeakSet();
  function replayCharts(card) {
    if (replayed.has(card)) return;
    replayed.add(card);
    card.querySelectorAll("canvas").forEach(cv => {
      const ch = window.Chart && Chart.getChart ? Chart.getChart(cv) : null;
      if (ch) { ch.reset(); ch.update(); }
    });
  }
  const cards = [...document.querySelectorAll(".main .card")];
  if ("IntersectionObserver" in window && !off()) {
    const io = new IntersectionObserver((entries) => entries.forEach(en => {
      if (!en.isIntersecting) return;
      en.target.classList.add("m-in");
      replayCharts(en.target);
      io.unobserve(en.target);
    }), { threshold: .12, rootMargin: "0px 0px -40px 0px" });
    cards.forEach(c => {
      const row = c.parentElement && c.parentElement.classList.contains("row") ? c.parentElement : null;
      const idx = row ? [...row.children].indexOf(c) : 0;
      c.style.setProperty("--d", (idx * 110) + "ms");
      c.classList.add("m-reveal");
      io.observe(c);
    });
    // cards already on screen when the dashboard opens
    new MutationObserver(() => {
      if (document.body.classList.contains("locked")) {      // logged out: replay next time
        cards.forEach(c => { c.classList.remove("m-in"); io.observe(c); });
        Object.keys(lastVal).forEach(k => delete lastVal[k]);
      }
    }).observe(document.body, { attributes: true, attributeFilter: ["class"] });
  }

  /* ---------- Login list, PIN errors, refresh spin ---------- */
  const pcList = document.getElementById("pcList");
  if (pcList) {
    pcList.classList.add("m-intro");
    new MutationObserver(() => pcList.querySelectorAll(".login-item").forEach((b, i) => b.style.setProperty("--i", Math.min(i, 12)))).observe(pcList, { childList: true });
    document.getElementById("pcSearch") && document.getElementById("pcSearch").addEventListener("input", () => pcList.classList.remove("m-intro"), { once: true });
  }
  const mErr = document.getElementById("mError");
  if (mErr) new MutationObserver(() => {
    if (!mErr.textContent.trim()) return;
    const card = document.querySelector("#pinModal .modal-card");
    card.classList.remove("m-shake"); void card.offsetWidth; card.classList.add("m-shake");
    setTimeout(() => card.classList.remove("m-shake"), 450);
  }).observe(mErr, { childList: true, characterData: true, subtree: true });
  const rb = document.getElementById("refreshBtn");
  if (rb) rb.addEventListener("click", () => { rb.classList.remove("m-spin"); void rb.offsetWidth; rb.classList.add("m-spin"); });
  if (rb) rb.addEventListener("animationend", () => rb.classList.remove("m-spin"));

  /* ---------- Snapshot export: switch motion off so nothing is caught mid-animation ---------- */
  const origSnap = window.snapshotBlob;
  if (typeof origSnap === "function") {
    window.snapshotBlob = async function () {
      const had = html.classList.contains("motion-off");
      html.classList.add("motion-off");
      document.querySelectorAll(".main .m-reveal").forEach(c => c.classList.add("m-in"));
      try { return await origSnap.apply(this, arguments); }
      finally { if (!had) html.classList.remove("motion-off"); }
    };
  }

})();
