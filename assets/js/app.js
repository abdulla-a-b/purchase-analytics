/* Panda Procurement Control Room — UI */
(function () {
  "use strict";
  const CFG = window.DASHBOARD_CONFIG || {};
  const E = window.PurchaseEngine;
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const esc = (s) => String(s === null || s === undefined ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const N = (v) => (typeof v === "number" && isFinite(v) ? v : 0);

  /* ---------------------------------------------------------------- formatting (৳, lakh, crore) */
  const inr = (x, d = 0) => N(x).toLocaleString("en-IN", { minimumFractionDigits: d, maximumFractionDigits: d });
  function tk(x) {
    const v = N(x), a = Math.abs(v);
    if (a >= 1e7) return "৳" + (v / 1e7).toFixed(2) + " Cr";
    if (a >= 1e5) return "৳" + (v / 1e5).toFixed(2) + " L";
    return "৳" + inr(v);
  }
  const lakh = (x) => Math.round(N(x) / 1e3) / 100;
  const pct = (x, d = 0) => (x === null || x === undefined ? "—" : (x * 100).toFixed(d) + "%");
  const days = (x, d = 1) => (x === null || x === undefined ? "—" : x.toFixed(d) + " d");
  const dfmt = (day) => (day === null || day === undefined ? "—" : E.dayToDate(day).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "2-digit", timeZone: "UTC" }));
  const icon = (st) => `<svg aria-hidden="true"><use href="#i-${st}"/></svg>`;
  const STATE = { good: "On track", warning: "Watch", serious: "At risk", critical: "Critical", none: "No data" };
  const pill = (st, text) => `<span class="pill ${st}">${icon(st === "neutral" ? "none" : st === "info" ? "info" : st)}${esc(text)}</span>`;
  const prioState = { High: "critical", Medium: "serious", Low: "warning" };
  const stageState = (s) => ({ "Awaiting Approval": "warning", "Approved - PO Pending": "warning", "PO Issued - Awaiting Delivery": "info",
    "Receiving In-house": "info", "In-house Closed - Payment Due": "info", "Closed": "good", "Rejected": "neutral" }[s] || "neutral");
  const finState = (s) => ({ "Fully Paid": "good", "Partially Paid": "info", "Initial / Advance Paid": "info", "Not Paid": "warning", "No PO Yet": "neutral" }[s] || "neutral");
  const pendState = (s) => !s ? "neutral" : s.startsWith("All In-house") ? "good" : s.startsWith("Short") ? "critical" : s.startsWith("Pending") ? "warning" : s.startsWith("Not Started") ? "info" : "neutral";

  /* ---------------------------------------------------------------- state */
  const S = { data: null, all: [], lines: [], A: null, R: null, tab: "overview", prio: "", sort: { key: "_row", dir: 1 }, charts: {}, source: "", drawn: {} };
  const store = {
    get(k) { try { return window.localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { window.localStorage.setItem(k, v); } catch (e) { /* storage unavailable */ } },
  };
  function toast(msg) {
    document.querySelectorAll(".toast").forEach((x) => x.remove());
    const t = document.createElement("div"); t.className = "toast"; t.setAttribute("role", "status"); t.textContent = msg;
    document.body.appendChild(t); setTimeout(() => t.remove(), 3800);
  }
  function todayDay() { const d = new Date(); return Math.round(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000); }

  /* ---------------------------------------------------------------- data loading */
  async function fetchJSON(url) {
    const r = await fetch(url, { cache: "no-store" });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return r.json();
  }
  async function loadData() {
    if (CFG.dataSource === "apps_script" && CFG.appsScriptUrl) {
      try { const d = await fetchJSON(CFG.appsScriptUrl); if (d && d.lines) return [d, "Live Google Sheet"]; }
      catch (e) { toast("Live Google Sheet not reachable — showing the last published data."); }
    }
    const d = await fetchJSON(CFG.jsonUrl || "data/purchase_data.json");
    return [d, "Published data"];
  }
  function setData(d, source) {
    S.data = d; S.source = source;
    $("#company-name").textContent = (d.meta && d.meta.company) || CFG.company || "";
    $("#app-title").textContent = CFG.title || "Procurement Control Room";
    const gen = d.meta && d.meta.generated_at ? new Date(d.meta.generated_at) : null;
    $("#src-chip").textContent = `${source}${gen ? " · updated " + gen.toLocaleString("en-GB", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }) : ""}`;
    if ($("#foot-source")) $("#foot-source").textContent = `Source: ${(d.meta && d.meta.source) || "—"} · ${(d.meta && d.meta.rows) || 0} purchase lines${d.meta && d.meta.warnings && d.meta.warnings.length ? " · " + d.meta.warnings.length + " data warning(s)" : ""}`;
    if ($("#foot-prepared")) $("#foot-prepared").textContent = CFG.preparedBy || "";
    if (d.meta && d.meta.warnings && d.meta.warnings.length) toast(`${d.meta.warnings.length} cell(s) could not be read — see the footer.`);
    buildFilters();
    recompute();
    if (window.PPCR_ENTRY && !document.querySelector("#single-form, #ed-form")) { window.PPCR_ENTRY.renderEntry(); window.PPCR_ENTRY.renderAdmin(); }
  }

  /* ---------------------------------------------------------------- filters */
  const FILTERS = [["f-dept", "dept", "All departments"], ["f-category", "category", "All categories"], ["f-supplier", "supplier", "All suppliers"],
    ["f-spend", "spend", "OPEX + CAPEX"], ["f-origin", "origin", "Local + import"]];
  function buildFilters() {
    const raw = S.data.lines || [];
    FILTERS.forEach(([id, key, label]) => {
      const el = $("#" + id), cur = el.value;
      const vals = Array.from(new Set(raw.map((l) => l[key]).filter(Boolean))).sort();
      el.innerHTML = `<option value="">${label}</option>` + vals.map((v) => `<option${v === cur ? " selected" : ""}>${esc(v)}</option>`).join("");
    });
    const st = $("#reg-stage"); st.innerHTML = `<option value="">All stages</option>` + E.STAGES.map((s) => `<option>${s}</option>`).join("");
    const fs = $("#reg-fin"); fs.innerHTML = `<option value="">All payment states</option>` + E.FIN.map((s) => `<option>${s}</option>`).join("");
  }
  function activeFilters() { const f = {}; FILTERS.forEach(([id, key]) => { const v = $("#" + id).value; if (v) f[key] = v; }); return f; }

  function recompute() {
    S.R = E.toDay($("#asof").value) ?? todayDay();
    S.all = E.enrich(S.data, S.R);
    E.analyze(S.all, S.data, S.R, CFG.targets);   // attaches sourcing / invoice / delivery controls to every line
    const f = activeFilters();
    S.lines = S.all.filter((l) => Object.entries(f).every(([k, v]) => l[k] === v));
    const data = f.dept ? Object.assign({}, S.data, { budgets: (S.data.budgets || []).filter((b) => b.dept === f.dept) }) : S.data;
    S.A = E.analyze(S.lines, data, S.R, CFG.targets);
    const n = Object.keys(f).length;
    $("#f-count").textContent = n ? `${S.lines.length} of ${S.all.length} lines shown` : `${S.all.length} purchase lines`;
    const hi = S.A.kpis.high; const b = $("#badge-high"); b.hidden = !hi; b.textContent = hi;
    $("#loading").hidden = true;
    S.drawn = {};
    renderAll();
    showTab(S.tab, false);
  }

  /* ---------------------------------------------------------------- tabs */
  function showTab(t, push = true) {
    S.tab = t;
    $$(".tab-btn").forEach((b) => b.setAttribute("aria-selected", b.dataset.tab === t ? "true" : "false"));
    $$(".tab-panel").forEach((p) => { p.hidden = p.dataset.panel !== t; });
    if (push) { try { history.replaceState(null, "", "#" + t); } catch (e) { /* sandboxed */ } }
    if (!S.drawn[t]) { drawCharts(t); S.drawn[t] = true; }
  }

  /* ---------------------------------------------------------------- render: overview */
  function renderAll() {
    renderOverview(); renderLeadtime(); renderActions(); renderSuppliers(); renderSpend(); renderFinance(); renderControls(); renderRegister();
  }
  const RANK = { critical: 4, serious: 3, warning: 2, good: 1, none: 0 };

  function renderOverview() {
    const A = S.A, k = A.kpis;
    const sig = (area) => { const x = A.signals.find((s) => s.area === area); return x && x.status !== "none" ? x.status : null; };
    const worst = A.signals.reduce((w, s) => (RANK[s.status] > RANK[w] ? s.status : w), "good");
    const bad = A.signals.filter((s) => s.status === "critical" || s.status === "serious");
    const crit = A.signals.filter((s) => s.status === "critical").length;
    const head = crit ? `${crit} critical risk signal${crit > 1 ? "s" : ""} need management attention`
      : bad.length ? `${bad.length} area${bad.length > 1 ? "s" : ""} at risk` : worst === "warning" ? "On track, with items to watch" : "All procurement signals on track";
    $("#alert-strip").innerHTML = `<div class="alert-strip ${worst}">${pill(worst, STATE[worst])}<strong>${esc(head)}</strong>
      <div class="items">${bad.sort((a, b) => RANK[b.status] - RANK[a.status]).map((s) => pill(s.status, s.area + ": " + s.metric)).join("")}</div></div>`;

    // health gauge
    const h = A.health, hs = h >= 80 ? "good" : h >= 65 ? "warning" : h >= 50 ? "serious" : "critical";
    const hl = { good: "Healthy", warning: "Watch", serious: "At risk", critical: "Critical" }[hs];
    const cx = 130, cy = 128, r = 100;
    const pt = (v, rr) => { const a = Math.PI * (1 - v / 100); return [cx + rr * Math.cos(a), cy - rr * Math.sin(a)]; };
    const arc = (a, b, rr) => { const p = pt(a, rr), q = pt(b, rr); return `M${p[0].toFixed(1)} ${p[1].toFixed(1)} A${rr} ${rr} 0 0 1 ${q[0].toFixed(1)} ${q[1].toFixed(1)}`; };
    const bands = [[0, 50, "critical"], [50, 65, "serious"], [65, 80, "warning"], [80, 100, "good"]];
    const tickL = [0, 50, 65, 80, 100].map((v) => { const p = pt(v, 124); return `<text x="${p[0].toFixed(1)}" y="${(p[1] + 4).toFixed(1)}" font-size="10" text-anchor="middle" fill="var(--muted)">${v}</text>`; }).join("");
    $("#health").innerHTML = `<div class="gauge-wrap">
      <svg class="gauge" viewBox="-12 0 284 150" role="img" aria-label="Procurement health ${h} out of 100, ${hl}">
        ${bands.map(([a, b, st]) => `<path d="${arc(a + 0.6, b - 0.6, 114)}" stroke="var(--${st})" stroke-width="5" fill="none" stroke-linecap="butt" opacity=".85"/>`).join("")}
        <path d="${arc(0, 100, r)}" stroke="var(--grid)" stroke-width="18" fill="none" stroke-linecap="round"/>
        <path d="${arc(0, Math.max(1, h), r)}" stroke="var(--${hs})" stroke-width="18" fill="none" stroke-linecap="round"/>
        ${tickL}
        <text x="${cx}" y="${cy - 16}" text-anchor="middle" font-size="52" font-weight="800" fill="var(--ink)">${h}</text>
        <text x="${cx}" y="${cy + 6}" text-anchor="middle" font-size="12" fill="var(--muted)">out of 100</text>
      </svg>
      <div class="health-label">${pill(hs, hl)}</div></div>
      <div class="parts">${A.healthParts.map((c) => {
        const st = c.s >= 0.9 ? "good" : c.s >= 0.75 ? "warning" : c.s >= 0.5 ? "serious" : "critical";
        return `<div class="part"><span>${esc(c.name)}</span><span class="bar"><i style="width:${(c.s * 100).toFixed(0)}%;background:var(--${st})"></i></span><span class="pts">${(c.w * c.s).toFixed(1)}/${c.w}</span></div>`;
      }).join("")}</div>`;

    // pipeline
    const st = A.settings;
    const ageState = (stage, age) => {
      if (age === null) return null;
      if (stage === "Awaiting Approval") return age > 2 * st.sla_appr ? "serious" : age > st.sla_appr ? "warning" : "good";
      if (stage === "Approved - PO Pending") return age > 2 * st.sla_po ? "serious" : age > st.sla_po ? "warning" : "good";
      return "info";
    };
    const short = { "Awaiting Approval": "Awaiting OPEX/CAPEX approval", "Approved - PO Pending": "Approved, PO not issued", "PO Issued - Awaiting Delivery": "PO issued, awaiting goods",
      "Receiving In-house": "Receiving in-house", "In-house Closed - Payment Due": "In-house, payment due", "Closed": "Closed & paid" };
    const flow = A.pipeline.filter((p) => p.stage !== "Rejected");
    const rej = A.pipeline.find((p) => p.stage === "Rejected");
    $("#pipeline").innerHTML = `<div class="flow">${flow.map((p) => {
      const as = ageState(p.stage, p.avgAge);
      return `<div class="stage" role="button" tabindex="0" data-stage="${esc(p.stage)}" aria-label="${esc(p.stage)}: ${p.count} lines">
        <div class="st-name">${esc(short[p.stage])}</div><div class="st-count">${p.count}</div><div class="st-val">${tk(p.value)}</div>
        <div class="st-age">${p.count && as ? pill(as, "avg " + Math.round(p.avgAge) + " d here") : "&nbsp;"}</div></div>`;
    }).join("")}</div>
      <div class="flow-foot"><span><b>${S.lines.length}</b> PR lines · <b>${k.pos}</b> POs</span><span>Committed <b>${tk(k.committed)}</b></span>
      <span>Still to come in-house <b>${tk(k.pendingVal)}</b></span><span>Rejected <b>${rej ? rej.count : 0}</b></span>
      <span>SLA: approval ${st.sla_appr} d · PO ${st.sla_po} d</span></div>
      <h3 style="font-size:14px;margin:18px 0 8px">Follow up this week</h3>${followUps()}`;
    $$("#pipeline .stage").forEach((el) => {
      const go = () => { $("#reg-stage").value = el.dataset.stage; renderRegister(); showTab("register"); };
      el.addEventListener("click", go); el.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go(); } });
    });

    // signals
    $("#signals").innerHTML = A.signals.map((s) => `<div class="signal ${s.status}">
      <div class="sig-ico">${icon(s.status)}</div>
      <div class="sig-area"><span>${esc(s.area)}</span><span class="sig-state">${STATE[s.status]}</span></div>
      <div class="sig-metric">${esc(s.metric)}</div><div class="sig-note">${esc(s.note)}</div></div>`).join("");

    // KPIs
    const T = Object.assign({ otif: 0.9, infull: 0.95 }, CFG.targets);
    const kpi = (lbl, val, sub, st, bar) => `<div class="kpi"><div class="k-lbl"><span>${esc(lbl)}</span>${st ? `<span class="ico-${st}" title="${STATE[st]}">${icon(st).replace("<svg", '<svg width="14" height="14"')}</span>` : ""}</div>
      <div class="k-val">${val}</div><div class="k-sub">${sub}</div>${bar ? `<div class="k-bar"><i style="width:${Math.min(100, bar.v * 100).toFixed(1)}%;background:var(--${bar.c || "brand"})"></i>${bar.t !== undefined ? `<b style="left:${Math.min(100, bar.t * 100)}%"></b>` : ""}</div>` : ""}</div>`;
    const lt = (v, tgt) => (v === null ? null : v >= tgt ? "good" : v >= tgt - 0.15 ? "warning" : "critical");
    const local = A.origin.find((o) => o.origin === "Local"), imp = A.origin.find((o) => o.origin === "Import");
    $("#kpis").innerHTML = [
      kpi("Committed spend", tk(k.committed), `${k.pos} POs · OPEX ${tk(k.opex)} · CAPEX ${tk(k.capex)}`),
      kpi("Paid to suppliers", pct(k.pctPaid), `${tk(k.paid)} paid`, null, { v: k.pctPaid, c: "good" }),
      kpi("Outstanding payable", tk(k.outstanding), `${tk(k.overdue)} overdue`, sig("Supplier payments")),
      kpi("Still to come in-house", tk(k.pendingVal), `${k.pendingLines} open line(s) · ${k.shortClosed} short-closed`, sig("Open PO lead time")),
      kpi("On-time in-full (OTIF)", pct(k.otif), `Target ${pct(T.otif)}`, sig("On-time delivery"), { v: N(k.otif), c: sig("On-time delivery") || "brand", t: T.otif }),
      kpi("In-full delivery", pct(k.infull), `Target ${pct(T.infull)}`, sig("In-full delivery"), { v: N(k.infull), c: sig("In-full delivery") || "brand", t: T.infull }),
      kpi("Lead time PO → in-house", days(k.dPoClose), `Local ${days(local && local.lead, 0)} · Import ${days(imp && imp.lead, 0)}`),
      kpi("Full cycle PR → close", days(k.dCycle), `PR→approval ${days(k.dPrAppr)} · →PO ${days(k.dApprPo)}`),
      kpi("Budget used (all depts)", pct(k.budgetUtil), `${tk(k.budgetCommitted)} of ${tk(k.budget)}${k.overBudget ? ` · <b>${k.overBudget} dept over budget</b>` : ""}`, sig("Budget control"),
        { v: N(k.budgetUtil), c: "brand", t: st.watch }),
      kpi("Price vs budget", (k.netPVpct >= 0 ? "+" : "") + pct(k.netPVpct, 1), `Saved ${tk(k.savings)} · over ${tk(k.overspend)}`, sig("Price discipline")),
      kpi("Paid ahead of receipt", tk(k.paidAhead), k.atRisk > 0 ? `${tk(k.atRisk)} on late POs` : "Advances on POs still on schedule", sig("Cash at risk")),
      kpi("High-priority actions", String(k.high), `${k.medium} medium · ${k.low} low · ${k.dataIssues} data issue(s)`, k.high ? "critical" : k.medium ? "warning" : "good"),
    ].join("");

    // top actions
    const top = A.actions.slice(0, 7);
    $("#top-actions").innerHTML = top.length ? `<div class="tbl-wrap"><table><thead><tr><th>Priority</th><th>PR</th><th>Supplier</th><th>What to do</th></tr></thead><tbody>
      ${top.map((l) => `<tr class="click prio-${l.all_priority}" data-pr="${esc(l.pr_no)}"><td>${pill(prioState[l.all_priority], l.all_priority)}</td><td class="mono">${esc(l.pr_no)}</td>
      <td>${esc(l.supplier || "—")}</td><td>${esc(l.all_action.replace(/; $/, "").split("; ")[0])}${l.all_action.split("; ").filter(Boolean).length > 1 ? ` <span class="cell-sub">+${l.all_action.split("; ").filter(Boolean).length - 1} more</span>` : ""}</td></tr>`).join("")}
      </tbody></table></div>` : `<div class="empty">No open actions. Every line is on schedule.</div>`;

    // highlights
    const best = A.suppliers.filter((s) => s.score !== null).sort((a, b) => b.score - a.score);
    const topCat = A.categories[0];
    const aCats = A.categories.filter((c) => c.cls === "A");
    const H = [
      [sig("Supplier payments") || "good", `Spend: ${tk(k.committed)} committed on ${k.pos} POs — ${pct(k.pctPaid)} paid; ${tk(k.outstanding)} outstanding, of which ${tk(k.overdue)} is overdue.`],
      [RANK[sig("On-time delivery")] >= RANK[sig("Open PO lead time")] ? sig("On-time delivery") : sig("Open PO lead time"), `Delivery: OTIF ${pct(k.otif)}, in-full ${pct(k.infull)}, average lead time ${days(k.dPoClose, 0)}. ${k.lateOpen} open PO(s) late; ${tk(k.pendingVal)} of material still to come in-house; ${k.shortClosed} line(s) short-closed.`],
      [sig("Supplier concentration") || "good", `Concentration: ${k.topSupplier ? esc(k.topSupplier.supplier) + " holds " + pct(k.topSupplier.share) + " of spend" : "—"} (HHI ${Math.round(k.hhi).toLocaleString("en-US")}). ${topCat ? `Top category ${esc(topCat.category)} at ${pct(topCat.share)}; ${aCats.length} A-class categories carry ${pct(aCats.reduce((s, c) => s + c.share, 0))} of spend.` : ""}`],
      [best.length && best[best.length - 1].score < 60 ? "warning" : "good", best.length ? `Suppliers: ${A.suppliers.filter((s) => s.rating.startsWith("A")).length} rated A, ${A.suppliers.filter((s) => s.rating.startsWith("B")).length} B, ${A.suppliers.filter((s) => s.rating.startsWith("C")).length} C. Best ${esc(best[0].supplier)} (${best[0].score}); weakest ${esc(best[best.length - 1].supplier)} (${best[best.length - 1].score}).` : "Suppliers: no closed deliveries scored yet."],
      [RANK[sig("Budget control")] >= RANK[sig("Price discipline")] ? sig("Budget control") : sig("Price discipline"), `Budget & price: ${k.overBudget} department(s) over budget, ${k.watchBudget} on watch (${pct(k.budgetUtil)} used). Bought ${tk(k.savings)} under and ${tk(k.overspend)} over budget price; ${k.dearer} line(s) dearer than the previous PO.`],
      [RANK[sig("Invoice matching")] >= RANK[sig("Sourcing compliance")] ? (sig("Invoice matching") || "good") : (sig("Sourcing compliance") || "good"),
        A.controls.has.quotations || A.controls.has.invoices ? `Sourcing & invoices: ${pct(A.controls.quoteCompliance)} of PO value bought with enough quotes; ${A.controls.notL1.length} line(s) not at the lowest quote (${tk(A.controls.notL1Extra)} above it). ${A.controls.mismatches.length} invoice mismatch(es); ${tk(A.controls.paidOverInvoiced)} paid beyond invoices; ${pct(A.controls.receiptsOnTime)} of receipts on time.`
        : "Sourcing & invoices: start recording quotations and invoices to see sourcing compliance and three-way matching."],
      [k.breach ? "critical" : k.high ? "serious" : "good", `Control: ${k.breach} PO(s) without approval; ${k.high} high, ${k.medium} medium and ${k.low} low priority action line(s); ${tk(k.paidAhead)} paid ahead of goods received.`],
    ];
    $("#highlights").innerHTML = H.map(([s, t]) => `<li><span class="ico-${s}">${icon(s)}</span><span>${t}</span></li>`).join("");
  }

  function followUps() {
    const R = S.R, st = S.A.settings, items = [];
    S.lines.forEach((l) => {
      if (l.is_open && N(l.delay) > 0) items.push([0, "critical", `Late ${l.delay} d`, `Chase ${l.supplier || "supplier"} — ${l.item}`, l]);
      else if (l.is_open && l.days_to_commit !== null && l.days_to_commit <= 7) items.push([1, "warning", l.days_to_commit === 0 ? "Due today" : `Due in ${l.days_to_commit} d`, `Confirm delivery — ${l.item} (${l.supplier || "—"})`, l]);
      if (l.stage === "Awaiting Approval" && l.wait_days !== null && l.wait_days > st.sla_appr) items.push([2, l.wait_days > 2 * st.sla_appr ? "serious" : "warning", `${l.wait_days} d waiting`, `Get OPEX/CAPEX approval — ${l.item}`, l]);
      if (l.stage === "Approved - PO Pending" && l.wait_days !== null && l.wait_days > st.sla_po) items.push([2, l.wait_days > 2 * st.sla_po ? "serious" : "warning", `${l.wait_days} d since approval`, `Issue the PO — ${l.item}`, l]);
      if (N(l.outstanding) > 0 && l.due_date !== null && l.due_date >= R && l.due_date <= R + 7) items.push([3, "info", l.due_date === R ? "Pay today" : `Pay in ${l.due_date - R} d`, `${tk(l.outstanding)} to ${l.supplier || "supplier"}`, l]);
    });
    items.sort((a, b) => a[0] - b[0]);
    if (!items.length) return `<div class="empty">Nothing falls due in the next 7 days.</div>`;
    return `<div class="tbl-wrap"><table><tbody>${items.slice(0, 7).map(([, st2, when, what, l]) => `<tr class="click" data-pr="${esc(l.pr_no)}">
      <td style="width:150px">${pill(st2, when)}</td><td>${esc(what)}</td><td class="mono cell-sub" style="white-space:nowrap">${esc(l.po_no || l.pr_no)}</td></tr>`).join("")}</tbody></table></div>
      ${items.length > 7 ? `<p class="note">+${items.length - 7} more in the Lead Time and Action Center tabs.</p>` : ""}`;
  }

  /* ---------------------------------------------------------------- render: lead time */
  function renderLeadtime() {
    const A = S.A, k = A.kpis, st = A.settings, R = S.R;
    const local = A.origin.find((o) => o.origin === "Local"), imp = A.origin.find((o) => o.origin === "Import");
    const stat = (l, v, s) => `<div class="stat"><div class="s-l">${l}</div><div class="s-v">${v}</div><div class="s-s">${s}</div></div>`;
    $("#lt-stats").innerHTML = [
      stat("Open POs", k.openPOs, `${tk(k.pendingVal)} still to arrive`),
      stat("Late open POs", k.lateOpen, k.lateOpen ? `worst ${k.maxLate} days` : "none past committed date"),
      stat("Due within 3 days", k.dueSoon, "follow up with supplier"),
      stat("Avg lead time · local", days(local && local.lead, 0), `OTIF ${pct(local && local.otif)}`),
      stat("Avg lead time · import", days(imp && imp.lead, 0), `OTIF ${pct(imp && imp.otif)}`),
      stat("Avg PR → close", days(k.dCycle, 0), "full purchase cycle"),
      stat("Receipts on time", pct(A.controls.receiptsOnTime), `${A.controls.receipts} receipt(s) vs committed date`),
      stat("Split deliveries", A.controls.splitPOs, A.controls.avgReceipts === null ? "POs received in 2+ lots" : `avg ${A.controls.avgReceipts.toFixed(1)} receipts per closed PO`),
    ].join("");

    // countdown timeline
    const open = A.open.filter((l) => l.po_date !== null);
    if (!open.length) $("#timeline").innerHTML = `<div class="empty">No open purchase orders. Everything ordered is in-house.</div>`;
    else {
      const lo = Math.min(...open.map((l) => l.po_date)) - 2;
      const hi = Math.max(R, ...open.map((l) => (l.commit_date === null ? l.po_date : l.commit_date))) + 4;
      const span = Math.max(1, hi - lo), x = (d) => ((d - lo) / span) * 100;
      const ticks = Array.from({ length: 7 }, (_, i) => lo + Math.round((span * (i + 0.5)) / 7));
      $("#timeline").innerHTML = `<div class="tl-axis"><span>PO · supplier</span><div class="tl-ticks">${ticks.map((t) => `<span style="left:${x(t)}%">${dfmt(t)}</span>`).join("")}</div><span style="text-align:right">Status</span></div>
        ${open.map((l, i) => {
          const end = l.commit_date === null ? l.po_date : l.commit_date;
          const late = N(l.delay) > 0;
          const s = l.commit_date === null ? pill("neutral", "No committed date") : late ? pill("critical", `Late ${l.delay} d`)
            : l.days_to_commit <= 3 ? pill("warning", l.days_to_commit === 0 ? "Due today" : `Due in ${l.days_to_commit} d`) : pill("good", `${l.days_to_commit} d left`);
          return `<div class="tl-row click${i === 0 ? " first" : ""}" data-pr="${esc(l.pr_no)}" style="cursor:pointer">
            <div class="tl-label"><div class="t1">${esc(l.po_no)} · ${esc(l.supplier || "—")}</div><div class="t2">${esc(l.item)} · ${inr(l.qty)} ${esc(l.uom || "")} · ${pct(l.recv_pct)} received${l.n_deliveries ? ` in ${l.n_deliveries} lot${l.n_deliveries > 1 ? "s" : ""}` : ""}</div></div>
            <div class="tl-track" title="PO ${dfmt(l.po_date)} → committed ${dfmt(l.commit_date)}">
              <div class="tl-bar" style="left:${x(l.po_date)}%;width:${Math.max(0.8, x(end) - x(l.po_date))}%"><i style="width:${Math.min(100, N(l.recv_pct) * 100)}%"></i></div>
              ${late ? `<div class="tl-late" style="left:${x(end)}%;width:${x(R) - x(end)}%"></div>` : ""}
              <div class="tl-today" style="left:${x(R)}%"></div></div>
            <div class="tl-status">${s}</div></div>`;
        }).join("")}`;
    }

    // SLA bars
    const committedLead = (() => { const v = S.lines.filter((l) => l.po_date !== null && l.commit_date !== null).map((l) => l.commit_date - l.po_date); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; })();
    const rows = [
      { n: "PR → OPEX/CAPEX approval", v: k.dPrAppr, t: st.sla_appr, tl: `SLA ${st.sla_appr} d`, share: k.apprInSla },
      { n: "Approval → PO issued", v: k.dApprPo, t: st.sla_po, tl: `SLA ${st.sla_po} d`, share: k.poInSla },
      { n: "PO → in-house close (actual vs agreed)", v: k.dPoClose, t: committedLead, tl: `agreed ${days(committedLead, 0)}` },
      { n: "Full cycle PR → close", v: k.dCycle, t: null },
    ];
    const mx = Math.max(1, ...rows.map((r) => Math.max(N(r.v), N(r.t)))) * 1.15;
    $("#sla").innerHTML = rows.map((r) => {
      const s = r.v === null ? "none" : r.t === null ? "info" : r.v <= r.t ? "good" : r.v <= r.t * 1.5 ? "warning" : "serious";
      const col = s === "info" || s === "none" ? "var(--brand)" : `var(--${s})`;
      return `<div class="sla-row"><div class="sla-top"><span>${esc(r.n)}${r.share !== undefined && r.share !== null ? ` <span class="cell-sub">· ${pct(r.share)} within SLA</span>` : ""}</span><b>${days(r.v)}</b></div>
        <div class="sla-track"><i style="width:${(N(r.v) / mx * 100).toFixed(1)}%;background:${col}"></i>${r.t !== null && r.t !== undefined ? `<b style="left:${(r.t / mx * 100).toFixed(1)}%" data-l="${esc(r.tl)}"></b>` : ""}</div></div>`;
    }).join("") + `<p class="note" style="margin-top:14px">Green = within the service level, amber = up to 1.5× it, orange = beyond. SLAs are set on the Settings sheet.</p>`;

    // waiting queue
    const W = A.waiting;
    $("#waiting").innerHTML = W.length ? `<div class="tbl-wrap"><table><thead><tr><th>PR</th><th>Item</th><th>Step</th><th class="n">Waiting</th><th>Status</th></tr></thead><tbody>
      ${W.map((l) => {
        const sla = l.stage === "Awaiting Approval" ? st.sla_appr : st.sla_po;
        const s = l.wait_days === null ? "neutral" : l.wait_days > 2 * sla ? "serious" : l.wait_days > sla ? "warning" : "good";
        return `<tr class="click" data-pr="${esc(l.pr_no)}"><td class="mono">${esc(l.pr_no)}</td><td>${esc(l.item || "—")}<div class="cell-sub">${esc(l.dept || "")} · ${tk(l.total_cost)}</div></td>
          <td>${l.stage === "Awaiting Approval" ? "Approval" : "PO issue"}</td><td class="n">${l.wait_days === null ? "—" : l.wait_days + " d"}</td>
          <td>${pill(s, s === "good" ? `Within ${sla} d` : s === "neutral" ? "No date" : `${l.wait_days - sla} d over SLA`)}</td></tr>`;
      }).join("")}</tbody></table></div>` : `<div class="empty">Nothing is waiting for approval or a PO.</div>`;

    // origin
    $("#origin").innerHTML = `<div class="stats" style="grid-template-columns:1fr">${A.origin.map((o) =>
      stat(o.origin, days(o.lead, 0), `${o.lines} POs · ${tk(o.value)} · OTIF ${pct(o.otif)}`)).join("")}</div>
      <p class="note">Lead time = PO date to final in-house closing date, closed lines only.</p>`;
  }

  /* ---------------------------------------------------------------- render: actions */
  function renderActions() {
    const q = ($("#act-search").value || "").toLowerCase();
    const L = S.A.actions.filter((l) => (!S.prio || l.all_priority === S.prio) &&
      (!q || [l.pr_no, l.po_no, l.item, l.supplier, l.all_action, l.dept].join(" ").toLowerCase().includes(q)));
    $$("#prio-seg button").forEach((b) => {
      const p = b.dataset.p; b.setAttribute("aria-pressed", p === S.prio ? "true" : "false");
      const n = p ? S.A.actions.filter((l) => l.all_priority === p).length : S.A.actions.length;
      b.textContent = (p || "All") + " · " + n;
    });
    $("#act-table").innerHTML = `<thead><tr><th>Priority</th><th>PR / department</th><th>Item · supplier</th><th>Stage</th><th>In-house</th>
      <th class="n">Outstanding</th><th class="n">Paid ahead</th><th>What to do</th></tr></thead><tbody>
      ${L.length ? L.map((l) => `<tr class="click prio-${l.all_priority}" data-pr="${esc(l.pr_no)}"><td>${pill(prioState[l.all_priority], l.all_priority)}</td>
        <td><span class="mono cell-strong">${esc(l.pr_no)}</span><div class="cell-sub">${esc(l.dept || "—")}</div></td>
        <td>${esc(l.item || "—")}<div class="cell-sub">${esc(l.supplier || "—")}</div></td><td>${pill(stageState(l.stage), l.stage)}</td>
        <td>${pill(pendState(l.pend_status), l.pend_status || "—")}</td><td class="n">${l.outstanding ? tk(l.outstanding) : "—"}</td><td class="n">${l.paid_ahead ? tk(l.paid_ahead) : "—"}</td>
        <td>${l.all_action.split("; ").filter((x) => x.trim()).map((a) => `<div>• ${esc(a.replace(/;\s*$/, ""))}</div>`).join("")}</td></tr>`).join("")
        : `<tr><td colspan="8"><div class="empty">No actions match.</div></td></tr>`}</tbody>`;
  }

  /* ---------------------------------------------------------------- render: suppliers */
  const ratingState = (r) => (r.startsWith("A") ? "good" : r.startsWith("B") ? "warning" : r.startsWith("C") ? "critical" : "neutral");
  function renderSuppliers() {
    const A = S.A, k = A.kpis;
    $("#hhi-sub").textContent = `HHI ${Math.round(k.hhi).toLocaleString("en-US")} · ${k.hhi > 2500 ? "highly concentrated" : k.hhi > 1500 ? "moderately concentrated" : "diversified"}`;
    const L = A.suppliers.slice().sort((a, b) => b.spend - a.spend);
    const mb = (v, col) => v === null ? "—" : `<div class="minibar"><span><i style="width:${Math.min(100, v * 100)}%;background:${col}"></i></span><em>${pct(v)}</em></div>`;
    $("#sup-table").innerHTML = `<thead><tr><th>Supplier</th><th>Rating</th><th>Score</th><th class="n">POs</th><th class="n">Spend</th><th class="n">Share</th>
      <th>OTIF</th><th>In-full</th><th>QC accepted</th><th class="n">Lead time</th><th class="n">Avg delay</th><th class="n">Late open</th><th class="n">Short closed</th>
      <th class="n">Pending in-house</th><th class="n">Outstanding</th></tr></thead><tbody>
      ${L.map((s) => `<tr><td class="cell-strong" style="white-space:nowrap">${esc(s.supplier)}</td><td>${pill(ratingState(s.rating), s.rating)}</td>
        <td>${s.score === null ? "—" : `<div class="minibar"><span><i style="width:${s.score}%;background:var(--${ratingState(s.rating)})"></i></span><em>${s.score}</em></div>`}</td>
        <td class="n">${s.pos}</td><td class="n">${tk(s.spend)}</td><td class="n">${pct(s.share, 1)}</td>
        <td>${mb(s.otif, "var(--s1)")}</td><td>${mb(s.infull, "var(--s3)")}</td><td>${mb(s.qc, "var(--s7)")}</td>
        <td class="n">${days(s.lead, 0)}</td><td class="n">${days(s.delay, 1)}</td><td class="n">${s.lateOpen || "—"}</td><td class="n">${s.shortClosed || "—"}</td>
        <td class="n">${s.pending ? tk(s.pending) : "—"}</td><td class="n">${s.outstanding ? tk(s.outstanding) : "—"}</td></tr>`).join("")}</tbody>`;
  }

  /* ---------------------------------------------------------------- render: spend */
  function renderSpend() {
    const A = S.A, k = A.kpis;
    const stat = (l, v, s) => `<div class="stat"><div class="s-l">${l}</div><div class="s-v">${v}</div><div class="s-s">${s}</div></div>`;
    $("#spend-stats").innerHTML = [
      stat("Committed incl. tax", tk(k.committed), `${k.pos} POs`), stat("Saved vs budget price", tk(k.savings), "lines bought below budget"),
      stat("Overspend vs budget price", tk(k.overspend), "lines bought above budget"), stat("Net price variance", (k.netPVpct >= 0 ? "+" : "") + pct(k.netPVpct, 1), "of budgeted value"),
      stat("Dearer than last PO", k.dearer, "same item, earlier PO"), stat("A-class categories", A.categories.filter((c) => c.cls === "A").length, `of ${A.categories.length} carry ~80% of spend`),
    ].join("");
    const B = A.budgets.slice().sort((a, b) => N(b.util) - N(a.util));
    const watch = A.settings.watch;
    const mx = Math.max(1.2, ...B.map((b) => N(b.util) * 1.05));
    const bs = (s) => ({ "Over Budget": "critical", "No Budget - Spending": "serious", "Watch": "warning", "Within Budget": "good" }[s] || "neutral");
    $("#budget").innerHTML = B.length ? B.map((b) => `<div class="bud-row"><span class="cell-strong">${esc(b.dept)}</span>
      <div class="bud-track"><i style="width:${Math.min(100, N(b.util) / mx * 100).toFixed(1)}%;background:var(--${bs(b.status) === "neutral" ? "brand" : bs(b.status)})"></i>
      <div class="mk" style="left:${(watch / mx * 100).toFixed(1)}%" title="Watch ${pct(watch)}"></div><div class="mk" style="left:${(1 / mx * 100).toFixed(1)}%" title="100%"></div>
      <span>${b.util === null ? "no budget" : pct(b.util)}</span></div>
      <span>${pill(bs(b.status), b.status)}<div class="cell-sub">${tk(b.committed)} of ${tk(b.budget)}</div></span></div>`).join("")
      : `<div class="empty">No budget lines for this filter.</div>`;
    $("#pv-table").innerHTML = `<thead><tr><th>Category</th><th class="n">Saved</th><th class="n">Over</th><th class="n">Avg var.</th></tr></thead><tbody>
      ${A.categories.map((c) => `<tr><td>${esc(c.category)} <span class="cell-sub">${c.cls}</span></td><td class="n">${c.savings ? tk(c.savings) : "—"}</td>
      <td class="n">${c.overspend ? tk(c.overspend) : "—"}</td><td class="n" style="color:${N(c.apv) > A.settings.pv_alert ? "var(--critical-ink)" : "inherit"}">${c.apv === null ? "—" : (c.apv >= 0 ? "+" : "") + pct(c.apv, 1)}</td></tr>`).join("")}</tbody>`;
  }

  /* ---------------------------------------------------------------- render: finance */
  function renderFinance() {
    const A = S.A, k = A.kpis;
    const stat = (l, v, s) => `<div class="stat"><div class="s-l">${l}</div><div class="s-v">${v}</div><div class="s-s">${s}</div></div>`;
    $("#fin-stats").innerHTML = [
      stat("Outstanding payable", tk(k.outstanding), `${pct(k.pctPaid)} of committed paid`),
      stat("Overdue", tk(k.overdue), k.maxOverdue ? `oldest ${k.maxOverdue} days` : "nothing past due"),
      ...A.cashDue.map((c) => stat(`Due next ${c.days} days`, tk(c.value), "cash to plan")),
      stat("Paid ahead of receipt", tk(k.paidAhead), k.atRisk ? `${tk(k.atRisk)} on late POs` : "advances on schedule"),
    ].join("");
    const plan = A.controls.plan.filter((x) => x.day <= S.R + 60);
    $("#sched-table").innerHTML = plan.length ? `<thead><tr><th>Due</th><th>Supplier</th><th>PO · item</th><th>Basis</th><th class="n">Amount</th><th>Status</th></tr></thead><tbody>
      ${plan.map((x) => { const dd = x.day - S.R; return `<tr class="click" data-pr="${esc(x.line.pr_no)}"><td class="tnum">${dfmt(x.day)}</td><td>${esc(x.line.supplier || "—")}</td>
        <td><span class="mono">${esc(x.line.po_no || "")}</span><div class="cell-sub">${esc(x.line.item || "")}</div></td><td class="cell-sub">${esc(x.source)}</td><td class="n">${tk(x.amount)}</td>
        <td>${dd < 0 ? pill(-dd > 30 ? "critical" : "serious", `${-dd} d overdue`) : dd === 0 ? pill("warning", "Due today") : dd <= 7 ? pill("warning", `In ${dd} d`) : pill("info", `In ${dd} d`)}</td></tr>`; }).join("")}</tbody>`
      : `<tbody><tr><td><div class="empty">No payments due in the next 60 days.</div></td></tr></tbody>`;
    const T = A.payTerms, tot = T.reduce((s, t) => s + t.value, 0);
    $("#terms-table").innerHTML = `<thead><tr><th>Payment terms</th><th class="n">POs</th><th class="n">Value</th><th>Share</th><th class="n">Paid ahead</th></tr></thead><tbody>
      ${T.map((t) => `<tr><td>${esc(t.terms)}</td><td class="n">${t.lines}</td><td class="n">${tk(t.value)}</td>
      <td><div class="minibar"><span><i style="width:${tot ? (t.value / tot * 100) : 0}%;background:var(--s1)"></i></span><em>${pct(tot ? t.value / tot : 0)}</em></div></td>
      <td class="n">${t.paidAhead ? tk(t.paidAhead) : "—"}</td></tr>`).join("")}</tbody>`;
  }

  /* ---------------------------------------------------------------- render: sourcing & controls */
  function renderControls() {
    const A = S.A, C = A.controls, st = A.settings;
    const stat = (l, v, s, cls) => `<div class="stat${cls ? " " + cls : ""}"><div class="s-l">${l}</div><div class="s-v">${v}</div><div class="s-s">${s}</div></div>`;
    $("#ctl-stats").innerHTML = [
      stat("Sourcing compliance", pct(C.quoteCompliance), `of PO value ≥ ${tk(st.quote_threshold)} with ${st.min_quotes}+ quotes at the lowest price`),
      stat("Lines with a sourcing issue", C.quoteIssues.length, `${C.notL1.length} not at the lowest quote`),
      stat("Paid above the lowest quote", tk(C.notL1Extra), "on lines not bought at L1"),
      stat("Saved vs average quote", tk(C.quoteSaving), "competition value captured"),
      stat("Invoice mismatches", C.mismatches.length, `${tk(C.mismatches.reduce((s2, l) => s2 + N(l.total_cost), 0))} of PO value`),
      stat("Paid beyond invoices", tk(C.paidOverInvoiced), "recover or collect invoices"),
      stat("Awaiting invoice", (C.match.find((m) => m.status === "Awaiting invoice") || {}).lines || 0, "goods received, no invoice yet"),
      stat("Quotes per PO line", C.quotesPerPO === null ? "—" : C.quotesPerPO.toFixed(1), "average suppliers asked"),
    ].join("");
    $("#ctl-note").hidden = C.has.quotations && C.has.invoices;
    const show = S.qAll ? "all" : "issues";
    $$("#q-seg button").forEach((b) => b.setAttribute("aria-pressed", b.dataset.q === show ? "true" : "false"));
    const Q = S.lines.filter((l) => l.q_status && (show === "all" || l.q_issue || /^Quotes needed/.test(l.q_status)))
      .sort((a, b) => (b.q_issue - a.q_issue) || N(b.total_cost) - N(a.total_cost));
    const qs = (l) => l.q_issue ? (l.q_status === "Not lowest quote" ? "serious" : "critical") : /needed/.test(l.q_status) ? "warning" : "good";
    $("#q-table").innerHTML = `<thead><tr><th>PR / PO</th><th>Item · supplier</th><th class="n">PO value</th><th class="n">Quotes</th><th>Lowest quote (৳/unit)</th><th class="n">Chosen (৳/unit)</th><th>Status</th></tr></thead><tbody>
      ${Q.length ? Q.map((l) => `<tr class="click" data-pr="${esc(l.pr_no)}"><td><span class="mono cell-strong">${esc(l.pr_no)}</span><div class="cell-sub mono">${esc(l.po_no || "no PO yet")}</div></td>
        <td>${esc(l.item)}<div class="cell-sub">${esc(l.supplier || "—")}</div></td><td class="n">${l.total_cost === null ? "—" : tk(l.total_cost)}</td>
        <td class="n">${l.q_count}${N(l.total_cost) >= st.quote_threshold ? ` / ${st.min_quotes}` : ""}</td>
        <td>${l.q_l1 === null ? "—" : `${inr(l.q_l1, 2)}<div class="cell-sub">${esc(l.q_l1_supplier)}</div>`}</td>
        <td class="n">${l.q_sel === null ? "—" : inr(l.q_sel, 2) + (l.q_l1 && l.q_sel > l.q_l1 * 1.000001 ? `<div class="cell-sub" style="color:var(--serious-ink)">+${pct(l.q_sel / l.q_l1 - 1, 1)}</div>` : "")}</td>
        <td>${pill(qs(l), l.q_status)}</td></tr>`).join("") : `<tr><td colspan="7"><div class="empty">${C.has.quotations ? "No sourcing issues — every line above the threshold has enough quotes at the lowest price." : "No quotations recorded yet."}</div></td></tr>`}</tbody>`;
    const M = S.lines.filter((l) => l.match_status && l.match_status !== "Not due")
      .sort((a, b) => ({ Mismatch: 0, "Awaiting invoice": 1, "Partly invoiced": 2, "Invoiced, goods pending": 3, Matched: 4 }[a.match_status] - { Mismatch: 0, "Awaiting invoice": 1, "Partly invoiced": 2, "Invoiced, goods pending": 3, Matched: 4 }[b.match_status]) || N(b.total_cost) - N(a.total_cost));
    const ms = (s2) => ({ Matched: "good", "Partly invoiced": "info", "Invoiced, goods pending": "info", "Awaiting invoice": "warning", Mismatch: "critical" }[s2] || "neutral");
    $("#m-table").innerHTML = `<thead><tr><th>PO · supplier</th><th class="n">Ordered</th><th class="n">Received</th><th class="n">Accepted</th><th class="n">Invoiced</th><th class="n">PO value</th><th class="n">Invoiced ৳</th><th class="n">Paid</th><th>Match</th></tr></thead><tbody>
      ${M.length ? M.map((l) => `<tr class="click" data-pr="${esc(l.pr_no)}"><td><span class="mono cell-strong">${esc(l.po_no)}</span><div class="cell-sub">${esc(l.supplier || "—")} · ${esc(l.item)}</div></td>
        <td class="n">${inr(l.qty)}</td><td class="n">${inr(l.qty_recv)}</td><td class="n">${inr(l.accepted)}</td>
        <td class="n"${l.inv_qty > N(l.accepted) + 1e-9 ? ' style="color:var(--critical-ink);font-weight:700"' : ""}>${inr(l.inv_qty)}</td>
        <td class="n">${tk(l.total_cost)}</td><td class="n">${l.inv_amt ? tk(l.inv_amt) : "—"}</td><td class="n">${tk(l.total_paid)}</td>
        <td>${pill(ms(l.match_status), l.match_status)}${l.match_issues.map((t) => `<div class="t-err">${esc(t)}</div>`).join("")}</td></tr>`).join("")
        : `<tr><td colspan="9"><div class="empty">${C.has.invoices ? "No received goods to match yet." : "No invoices recorded yet."}</div></td></tr>`}</tbody>`;
  }

  /* ---------------------------------------------------------------- render: register */
  const REG_COLS = [["pr_no", "PR No"], ["pr_date", "PR date"], ["dept", "Department"], ["item", "Item"], ["supplier", "Supplier"], ["stage", "Stage"],
    ["total_cost", "Value", "n"], ["pct_paid", "Paid", "n"], ["fin_status", "Payment"], ["pend_status", "In-house"], ["delay", "Delay", "n"], ["all_priority", "Action"]];
  function renderRegister() {
    const q = ($("#reg-search").value || "").toLowerCase(), stg = $("#reg-stage").value, fin = $("#reg-fin").value;
    let L = S.lines.filter((l) => (!stg || l.stage === stg) && (!fin || l.fin_status === fin) &&
      (!q || [l.pr_no, l.po_no, l.item, l.supplier, l.dept, l.category, l.spec].join(" ").toLowerCase().includes(q)));
    const { key, dir } = S.sort;
    const pr = { High: 1, Medium: 2, Low: 3 };
    L = L.slice().sort((a, b) => {
      let x = a[key], y = b[key];
      if (key === "all_priority") { x = pr[x] || 9; y = pr[y] || 9; }
      if (x === null || x === undefined) return 1; if (y === null || y === undefined) return -1;
      return (x > y ? 1 : x < y ? -1 : 0) * dir;
    });
    $("#reg-count").textContent = `${L.length} line(s)`;
    $("#reg-table").innerHTML = `<thead><tr>${REG_COLS.map(([k, h, c]) => `<th class="sortable ${c || ""}" data-k="${k}" aria-sort="${S.sort.key === k ? (dir > 0 ? "ascending" : "descending") : "none"}">${h}${S.sort.key === k ? (dir > 0 ? " ▲" : " ▼") : ""}</th>`).join("")}</tr></thead><tbody>
      ${L.map((l) => `<tr class="click${l.all_priority ? " prio-" + l.all_priority : ""}" data-pr="${esc(l.pr_no)}"><td class="mono cell-strong">${esc(l.pr_no)}</td><td class="tnum">${dfmt(l.pr_date)}</td>
        <td>${esc(l.dept || "—")}</td><td>${esc(l.item || "—")}</td><td>${esc(l.supplier || "—")}</td><td>${pill(stageState(l.stage), l.stage)}</td>
        <td class="n">${l.total_cost === null ? "—" : tk(l.total_cost)}</td><td class="n">${pct(l.pct_paid)}</td><td>${pill(finState(l.fin_status), l.fin_status)}</td>
        <td>${pill(pendState(l.pend_status), l.pend_status || "—")}</td><td class="n">${N(l.delay) ? l.delay + " d" : "—"}</td>
        <td>${l.all_priority ? pill(prioState[l.all_priority], l.all_priority) : pill("good", "None")}</td></tr>`).join("") || `<tr><td colspan="12"><div class="empty">No lines match.</div></td></tr>`}</tbody>`;
    $$("#reg-table th.sortable").forEach((th) => th.addEventListener("click", () => {
      const k = th.dataset.k; S.sort = { key: k, dir: S.sort.key === k ? -S.sort.dir : 1 }; renderRegister();
    }));
  }

  /* ---------------------------------------------------------------- drawer (one line's journey) */
  function openLine(prNo) {
    const l = S.all.find((x) => x.pr_no === prNo); if (!l) return;
    const R = S.R, st = S.A.settings;
    const gap = (a, b) => (a !== null && b !== null ? `+${b - a} d` : "");
    const steps = [
      ["PR received", l.pr_date, "", l.pr_date !== null],
      ["OPEX/CAPEX approved" + (l.appr_by ? " · " + l.appr_by : ""), l.appr_date, gap(l.pr_date, l.appr_date), l.appr_date !== null, l.d_pr_appr !== null && l.d_pr_appr > st.sla_appr],
      ["PO issued" + (l.po_no ? " · " + l.po_no : ""), l.po_date, gap(l.appr_date, l.po_date), l.po_date !== null, l.d_appr_po !== null && l.d_appr_po > st.sla_po],
      ["Supplier committed delivery", l.commit_date, gap(l.po_date, l.commit_date), l.commit_date !== null && l.commit_date <= R && l.close_date !== null && l.close_date <= l.commit_date, N(l.delay) > 0],
      ["Started coming in-house", l.inhouse_start, gap(l.po_date, l.inhouse_start), l.inhouse_start !== null],
      ["Final in-house closing", l.close_date, gap(l.po_date, l.close_date), l.close_date !== null, l.close_date !== null && l.commit_date !== null && l.close_date > l.commit_date],
      ["Last payment", l.last_pay_date, "", l.last_pay_date !== null],
      ["Next payment due", l.due_date, N(l.days_overdue) > 0 ? `${l.days_overdue} d overdue` : "", N(l.outstanding) === 0 && l.due_date !== null, N(l.days_overdue) > 0],
    ];
    const kv = [["Department", l.dept], ["Requested by", l.req_by], ["Category", l.category], ["Spec", l.spec], ["Spend type", l.spend],
      ["Approval ref", l.appr_ref], ["Supplier", l.supplier], ["Local / import", l.origin], ["Order quantity", l.qty === null ? null : inr(l.qty) + " " + (l.uom || "")],
      ["Received", l.qty_recv === null ? null : inr(l.qty_recv) + " " + (l.uom || "") + " (" + pct(l.recv_pct) + ")"], ["In-house status", l.pend_status],
      ["Unit price", l.unit_price === null ? null : inr(l.unit_price, 2) + " " + (l.ccy || "BDT")], ["Unit price (BDT)", l.unit_bdt === null ? null : "৳" + inr(l.unit_bdt, 2)],
      ["Previous PO price", l.prev_price === null ? null : "৳" + inr(l.prev_price, 2) + " (" + (l.price_chg >= 0 ? "+" : "") + pct(l.price_chg, 1) + ")"],
      ["Budget unit price", l.budget_price === null ? null : "৳" + inr(l.budget_price, 2)], ["Variance vs budget", l.pv_pct === null ? null : (l.pv_pct >= 0 ? "+" : "") + pct(l.pv_pct, 1) + " · " + tk(l.pv_bdt)],
      ["Total incl. tax", l.total_cost === null ? null : tk(l.total_cost)], ["Paid", tk(l.total_paid) + " (" + pct(l.pct_paid) + ")"],
      ["Outstanding", l.outstanding === null ? null : tk(l.outstanding)], ["Paid ahead of receipt", l.paid_ahead ? tk(l.paid_ahead) : null],
      ["Payment terms", l.pay_terms], ["QC / material approval", l.qc_status], ["Approved sample photo", l.photo_link], ["Remarks", l.remarks],
      ["Data check", l.data_check]];
    $("#drawer-root").innerHTML = `<div class="drawer-bg" data-close></div>
      <aside class="drawer" role="dialog" aria-modal="true" aria-labelledby="dr-title">
        <button class="close" type="button" data-close aria-label="Close">✕</button>
        <div class="cell-sub mono">${esc(l.pr_no)}${l.po_no ? " · " + esc(l.po_no) : ""}</div>
        <h2 id="dr-title">${esc(l.item || "—")}</h2>
        <div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:8px">${pill(stageState(l.stage), l.stage)}${pill(finState(l.fin_status), l.fin_status)}${l.all_priority ? pill(prioState[l.all_priority], l.all_priority + " priority") : ""}${l.otif ? pill(l.otif === "OTIF" ? "good" : "serious", l.otif) : ""}</div>
        ${l.all_action ? `<div class="action-box"><b>What to do:</b>${l.all_action.split("; ").filter((x) => x.trim()).map((a) => `<div>• ${esc(a.replace(/;\s*$/, ""))}</div>`).join("")}</div>` : ""}
        <ol class="steps">${steps.map(([n, d, g, done, late]) => `<li class="${late ? "late" : done ? "done" : ""}"><span class="dot">${late ? icon("critical") : done ? icon("good") : ""}</span>
          <div><div class="s-name">${esc(n)}</div><div class="s-date">${dfmt(d)}</div></div><div class="s-gap">${esc(g)}</div></li>`).join("")}</ol>
        <div class="kv">${kv.filter(([, v]) => v !== null && v !== undefined && v !== "").map(([k, v]) => `<div><span>${esc(k)}</span><b>${esc(v)}</b></div>`).join("")}</div>
        ${drawerLogs(l)}
      </aside>`;
    const close = () => { $("#drawer-root").innerHTML = ""; document.removeEventListener("keydown", onKey); };
    const onKey = (e) => { if (e.key === "Escape") close(); };
    $$("#drawer-root [data-close]").forEach((b) => b.addEventListener("click", close));
    document.addEventListener("keydown", onKey);
    $("#drawer-root .close").focus();
  }

  function drawerLogs(l) {
    const g = l.logs || {}, out = [];
    const mini = (title, head, rows) => rows.length ? `<h3 class="subh">${title}</h3><div class="tbl-wrap"><table><thead><tr>${head.map((h) => `<th${h[1] ? ' class="n"' : ""}>${h[0]}</th>`).join("")}</tr></thead><tbody>${rows.join("")}</tbody></table></div>` : "";
    if (l.q_status) out.push(`<h3 class="subh">Sourcing ${pill(l.q_issue ? "serious" : /needed/.test(l.q_status) ? "warning" : "good", l.q_status)}</h3>` +
      ((g.quotations || []).length ? `<div class="tbl-wrap"><table><thead><tr><th>Supplier</th><th>Date</th><th class="n">Unit price</th><th class="n">Lead time</th></tr></thead><tbody>${g.quotations.slice().sort((a, b) => N(a.unit_price) - N(b.unit_price)).map((q) =>
        `<tr${String(q.supplier).toLowerCase() === String(l.supplier || "").toLowerCase() ? ' style="font-weight:700"' : ""}><td>${esc(q.supplier)}${String(q.supplier).toLowerCase() === String(l.supplier || "").toLowerCase() ? " ✓" : ""}</td><td class="tnum">${dfmt(E.toDay(q.date))}</td>
        <td class="n">${inr(q.unit_price, 2)} ${esc(q.ccy || l.ccy || "")}</td><td class="n">${q.lead_days === null || q.lead_days === undefined ? "—" : q.lead_days + " d"}</td></tr>`).join("")}</tbody></table></div>` : `<p class="cell-sub">No quotations recorded.</p>`));
    if (l.match_status) out.push(`<h3 class="subh">Invoice match ${pill({ Matched: "good", "Partly invoiced": "info", "Invoiced, goods pending": "info", "Awaiting invoice": "warning", Mismatch: "critical" }[l.match_status] || "neutral", l.match_status)}</h3>` +
      l.match_issues.map((t) => `<div class="t-err">${esc(t)}</div>`).join("") +
      `<p class="cell-sub">Ordered ${inr(l.qty)} · accepted ${inr(l.accepted)} · invoiced ${inr(l.inv_qty)} ${esc(l.uom || "")} — invoiced ${tk(l.inv_amt)} of ${tk(l.total_cost)}</p>`);
    out.push(mini("Deliveries", [["ID"], ["Date"], ["Qty", 1], ["Accepted", 1], ["On time"]], (g.deliveries || []).map((d) => `<tr><td class="mono">${esc(d.id)}</td><td class="tnum">${dfmt(E.toDay(d.date))}</td><td class="n">${inr(d.qty)}</td><td class="n">${d.qty_ok === null || d.qty_ok === undefined ? "—" : inr(d.qty_ok)}</td>
      <td>${l.commit_date === null ? "—" : E.toDay(d.date) <= l.commit_date ? pill("good", "On time") : pill("serious", "Late " + (E.toDay(d.date) - l.commit_date) + " d")}</td></tr>`)));
    out.push(mini("Payments", [["ID"], ["Entry"], ["Date"], ["Amount", 1], ["Ref"]], (g.payments || []).map((x) => `<tr><td class="mono">${esc(x.id)}</td><td>${esc(x.entry)} · ${esc(x.ptype)}${x.settles ? `<div class="cell-sub">settles ${esc(x.settles)}</div>` : ""}</td><td class="tnum">${dfmt(E.toDay(x.date))}</td><td class="n">${tk(x.amount)}</td><td class="cell-sub">${esc(x.invoice_no || x.ref || "")}</td></tr>`)));
    out.push(mini("Invoices", [["Invoice"], ["Date"], ["Qty", 1], ["Unit price", 1], ["Amount", 1]], (g.invoices || []).map((x) => `<tr><td class="mono">${esc(x.invoice_no)}</td><td class="tnum">${dfmt(E.toDay(x.date))}</td><td class="n">${inr(x.qty)}</td><td class="n">${inr(x.unit_price, 2)}</td><td class="n">${tk(x.amount)}</td></tr>`)));
    return out.join("");
  }

  /* ---------------------------------------------------------------- charts */
  function css(v) { return getComputedStyle(document.documentElement).getPropertyValue(v).trim(); }
  function chart(id, cfg) {
    if (S.charts[id]) { S.charts[id].destroy(); delete S.charts[id]; }
    const el = document.getElementById(id); if (!el || !window.Chart) return;
    S.charts[id] = new window.Chart(el, cfg);
  }
  function baseOpts(horizontal, unit) {
    const ink = css("--ink-2"), grid = css("--grid");
    return {
      responsive: true, maintainAspectRatio: false, indexAxis: horizontal ? "y" : "x", animation: { duration: 350 },
      layout: { padding: { top: horizontal ? 0 : 22, right: horizontal ? 56 : 8 } },
      plugins: {
        legend: { display: false, labels: { color: ink, boxWidth: 12, boxHeight: 12, usePointStyle: true, pointStyle: "rectRounded" } },
        tooltip: { backgroundColor: css("--ink"), titleColor: css("--surface"), bodyColor: css("--surface"), padding: 10, cornerRadius: 6 },
        datalabels: { display: true, color: css("--ink"), font: { size: 11, weight: "600", family: css("--f-body") || undefined },
          anchor: "end", align: horizontal ? "right" : "top", offset: 2, clamp: true },
      },
      scales: {
        x: { grid: { display: horizontal, color: grid, drawTicks: false }, border: { color: css("--line") }, ticks: { color: ink, font: { size: 11 } }, beginAtZero: true },
        y: { grid: { display: !horizontal, color: grid, drawTicks: false }, border: { color: css("--line") }, ticks: { color: ink, font: { size: 11 } }, beginAtZero: true,
          title: unit ? { display: true, text: unit, color: css("--muted"), font: { size: 11 } } : undefined },
      },
    };
  }
  const barDS = (label, data, color, extra) => Object.assign({ label, data, backgroundColor: color, borderRadius: 4, borderSkipped: "start", maxBarThickness: 34, categoryPercentage: .78, barPercentage: .92 }, extra || {});
  const wrapLabel = (s, n = 18) => { const w = String(s).split(" "); const out = []; let cur = ""; w.forEach((x) => { if ((cur + " " + x).trim().length > n) { if (cur) out.push(cur); cur = x; } else cur = (cur + " " + x).trim(); }); if (cur) out.push(cur); return out; };

  function drawCharts(tab) {
    if (!window.Chart) return;
    const A = S.A;
    if (tab === "leadtime") {
      const L = A.suppliers.filter((s) => s.lead !== null).sort((a, b) => b.lead - a.lead);
      const o = baseOpts(true, ""); o.plugins.legend.display = true; o.plugins.legend.position = "bottom";
      o.plugins.datalabels.formatter = (v) => (v === null ? "" : v.toFixed(1) + " d");
      o.scales.x.title = { display: true, text: "days", color: css("--muted") };
      chart("ch-lead", { type: "bar", data: { labels: L.map((s) => s.supplier), datasets: [
        barDS("Avg lead time PO → in-house", L.map((s) => s.lead), css("--s1")), barDS("Avg delay vs committed date", L.map((s) => s.delay), css("--s2"))] }, options: o });
    }
    if (tab === "suppliers") {
      const L = A.suppliers.filter((s) => s.score !== null).sort((a, b) => b.score - a.score);
      const o = baseOpts(true); o.scales.x.max = 100;
      o.plugins.datalabels.formatter = (v, c) => `${v} · ${L[c.dataIndex].rating.charAt(0)}`;
      o.plugins.tooltip.callbacks = { label: (c) => ` Score ${c.raw} — ${L[c.dataIndex].rating}` };
      chart("ch-score", { type: "bar", data: { labels: L.map((s) => s.supplier), datasets: [barDS("Score", L.map((s) => s.score), L.map((s) => css("--" + ratingState(s.rating))))] }, options: o });
      const P = A.suppliers.slice().sort((a, b) => b.spend - a.spend);
      const o2 = baseOpts(true); o2.plugins.datalabels.formatter = (v, c) => `${v.toFixed(1)}% · ${tk(P[c.dataIndex].spend)}`;
      o2.scales.x.ticks.callback = (v) => v + "%";
      chart("ch-share", { type: "bar", data: { labels: P.map((s) => s.supplier), datasets: [barDS("Share of spend", P.map((s) => +(N(s.share) * 100).toFixed(1)), css("--s1"))] }, options: o2 });
    }
    if (tab === "spend") {
      const C = A.categories;
      const o = baseOpts(false, "৳ lakh"); o.layout.padding.top = 34;
      o.plugins.datalabels.formatter = (v, c) => [`${v.toFixed(1)}`, `${pct(C[c.dataIndex].cum)} · ${C[c.dataIndex].cls}`];
      o.plugins.datalabels.textAlign = "center";
      o.plugins.tooltip.callbacks = { label: (c) => ` ${tk(C[c.dataIndex].spend)} · share ${pct(C[c.dataIndex].share, 1)} · cumulative ${pct(C[c.dataIndex].cum, 1)} · class ${C[c.dataIndex].cls}` };
      const clsCol = { A: css("--s1"), B: css("--s3"), C: css("--s7"), "-": css("--muted") };
      chart("ch-pareto", { type: "bar", data: { labels: C.map((c) => wrapLabel(c.category, 12)), datasets: [barDS("Spend", C.map((c) => lakh(c.spend)), C.map((c) => clsCol[c.cls]))] }, options: o });
      const k = A.kpis;
      chart("ch-opex", { type: "doughnut", data: { labels: ["OPEX", "CAPEX"], datasets: [{ data: [lakh(k.opex), lakh(k.capex)], backgroundColor: [css("--s1"), css("--s2")], borderColor: css("--surface"), borderWidth: 2 }] },
        options: { responsive: true, maintainAspectRatio: false, cutout: "58%", plugins: {
          legend: { position: "bottom", labels: { color: css("--ink-2"), usePointStyle: true, pointStyle: "rectRounded" } },
          datalabels: { color: "#fff", font: { weight: "700", size: 12 }, textAlign: "center",
            formatter: (v, c) => { const t = c.dataset.data.reduce((a, b) => a + b, 0); return t && v ? [`${c.chart.data.labels[c.dataIndex]}`, `৳${v.toFixed(1)} L`, pct(v / t)] : ""; } } } } });
      const M = A.monthly;
      const o3 = baseOpts(false, "৳ lakh"); o3.plugins.legend.display = true; o3.plugins.legend.position = "bottom";
      o3.plugins.datalabels.formatter = (v) => (v >= 1 ? v.toFixed(0) : ""); o3.plugins.datalabels.font.size = 10;
      chart("ch-monthly", { type: "bar", data: { labels: M.map((m) => m.label), datasets: [
        barDS("PO value", M.map((m) => lakh(m.poValue)), css("--s1")), barDS("Paid", M.map((m) => lakh(m.paid)), css("--s3")), barDS("Outstanding due", M.map((m) => lakh(m.due)), css("--s2"))] }, options: o3 });
    }
    if (tab === "controls") {
      const Mt = A.controls.match.filter((m) => m.status !== "Not due");
      const col = { Matched: css("--good"), "Partly invoiced": css("--s1"), "Invoiced, goods pending": css("--s7"), "Awaiting invoice": css("--warning"), Mismatch: css("--critical") };
      const o = baseOpts(true); o.plugins.datalabels.formatter = (v, c) => (v ? `${v.toFixed(1)} L · ${Mt[c.dataIndex].lines} line(s)` : ""); o.layout.padding.right = 110;
      o.scales.x.title = { display: true, text: "৳ lakh of PO value", color: css("--muted") };
      chart("ch-match", { type: "bar", data: { labels: Mt.map((m) => m.status), datasets: [barDS("PO value", Mt.map((m) => lakh(m.value)), Mt.map((m) => col[m.status]))] }, options: o });
    }
    if (tab === "finance") {
      const G = A.aging;
      const col = { "Not Yet Due": css("--s1"), "1-30 Days": css("--warning"), "31-60 Days": css("--serious"), "61-90 Days": css("--critical"), "90+ Days": css("--critical"), "No Due Date": css("--muted") };
      const o = baseOpts(false, "৳ lakh"); o.plugins.datalabels.formatter = (v, c) => (v ? [`${v.toFixed(1)}`, `${G[c.dataIndex].lines} line(s)`] : ""); o.plugins.datalabels.textAlign = "center"; o.layout.padding.top = 30;
      chart("ch-aging", { type: "bar", data: { labels: G.map((g) => g.bucket), datasets: [barDS("Outstanding", G.map((g) => lakh(g.value)), G.map((g) => col[g.bucket]))] }, options: o });
      const F = A.finance.filter((f) => f.status !== "No PO Yet");
      const o2 = baseOpts(false, "৳ lakh"); o2.plugins.datalabels.formatter = (v, c) => [`${v.toFixed(1)}`, `${F[c.dataIndex].lines} PO(s)`]; o2.plugins.datalabels.textAlign = "center"; o2.layout.padding.top = 30;
      chart("ch-finstatus", { type: "bar", data: { labels: F.map((f) => f.status), datasets: [barDS("Value", F.map((f) => lakh(f.value)), css("--s1"))] }, options: o2 });
      const P = A.suppliers.filter((s) => s.outstanding > 0).sort((a, b) => b.outstanding - a.outstanding);
      const o3 = baseOpts(true); o3.plugins.legend.display = true; o3.plugins.legend.position = "bottom";
      o3.plugins.datalabels.formatter = (v) => (v ? v.toFixed(1) : "");
      chart("ch-supplier-pay", { type: "bar", data: { labels: P.map((s) => s.supplier), datasets: [
        barDS("Outstanding", P.map((s) => lakh(s.outstanding)), css("--s1")), barDS("Of which overdue", P.map((s) => lakh(s.overdue)), css("--s2"))] }, options: o3 });
    }
  }

  /* ---------------------------------------------------------------- theme */
  function applyTheme(t) { if (t) document.documentElement.setAttribute("data-theme", t); else document.documentElement.removeAttribute("data-theme"); }
  function toggleTheme() {
    const cur = document.documentElement.getAttribute("data-theme") || (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
    const nxt = cur === "dark" ? "light" : "dark"; applyTheme(nxt); store.set("ppcr-theme", nxt);
    S.drawn = {}; showTab(S.tab, false);
  }

  /* ---------------------------------------------------------------- events */
  function bind() {
    $$(".tab-btn").forEach((b) => b.addEventListener("click", () => showTab(b.dataset.tab)));
    $$("[data-goto]").forEach((b) => b.addEventListener("click", () => showTab(b.dataset.goto)));
    FILTERS.forEach(([id]) => $("#" + id).addEventListener("change", recompute));
    $("#f-reset").addEventListener("click", () => { FILTERS.forEach(([id]) => { $("#" + id).value = ""; }); recompute(); });
    $("#asof").addEventListener("change", recompute);
    $("#btn-theme").addEventListener("click", toggleTheme);
    $("#btn-refresh").addEventListener("click", async () => { try { const [d, src] = await loadData(); setData(d, src); toast("Data refreshed."); } catch (e) { toast("Could not refresh: " + e.message); } });
    $("#xlsx-file").addEventListener("change", async (e) => {
      const f = e.target.files && e.target.files[0]; if (!f) return;
      toast("Reading " + f.name + "…");
      try { const d = await window.PurchaseXlsx.parse(f); setData(d, "Loaded file: " + f.name); toast(`${d.lines.length} purchase lines loaded from ${f.name}.`); }
      catch (err) { toast(err.message); }
      e.target.value = "";
    });
    if (window.self === window.top) { const p = $("#btn-print"); p.hidden = false; p.addEventListener("click", () => window.print()); }
    $("#prio-seg").addEventListener("click", (e) => { const b = e.target.closest("button"); if (!b) return; S.prio = b.dataset.p; renderActions(); });
    $("#act-search").addEventListener("input", renderActions);
    $("#q-seg").addEventListener("click", (e) => { const b = e.target.closest("button"); if (!b) return; S.qAll = b.dataset.q === "all"; renderControls(); });
    ["#reg-search", "#reg-stage", "#reg-fin"].forEach((s) => $(s).addEventListener(s === "#reg-search" ? "input" : "change", renderRegister));
    document.addEventListener("click", (e) => { const tr = e.target.closest("[data-pr]"); if (tr && !e.target.closest(".drawer")) openLine(tr.dataset.pr); });
    if (window.matchMedia) window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => { if (!document.documentElement.getAttribute("data-theme")) { S.drawn = {}; showTab(S.tab, false); } });
  }

  /* API for the Data Entry and Admin tabs (assets/js/entry-admin.js) */
  window.PPCR = {
    get data() { return S.data; }, get all() { return S.all; }, get R() { return S.R; }, get source() { return S.source; },
    setData: (d, src) => setData(d, src || S.source),
    reload: async () => { const [d, src] = await loadData(); setData(d, src); },
    showTab, toast, esc, pill, tk, inr, pct, dfmt, stageState, finState, pendState,
  };

  async function init() {
    applyTheme(store.get("ppcr-theme"));
    const t = todayDay(); $("#asof").value = E.dayToISO(t);
    const h = (location.hash || "").replace("#", ""); if ($(`.tab-btn[data-tab="${h}"]`)) S.tab = h;
    if (window.Chart && window.ChartDataLabels) window.Chart.register(window.ChartDataLabels);
    if (window.Chart) { window.Chart.defaults.font.family = getComputedStyle(document.body).fontFamily; window.Chart.defaults.color = css("--ink-2"); }
    bind();
    try { const [d, src] = await loadData(); setData(d, src); }
    catch (e) {
      $("#loading").innerHTML = `Could not load the purchase data (${esc(e.message)}). Use <b>Load Excel</b> to open the purchase workbook, or check <code>data/purchase_data.json</code>.`;
      $("#src-chip").textContent = "No data";
    }
    if (CFG.dataSource === "apps_script" && CFG.appsScriptUrl && CFG.refreshMinutes > 0) {
      setInterval(async () => { try { const [d, src] = await loadData(); setData(d, src); } catch (e) { /* keep current view */ } }, CFG.refreshMinutes * 60000);
    }
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
