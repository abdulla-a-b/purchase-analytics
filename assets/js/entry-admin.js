/* Panda Procurement Control Room — Data Entry & Admin tabs
 * Saves to the Google Sheet through the Apps Script web app (apps-script/Code.gs).
 * Without that connection the tabs run in demo mode: changes stay in this browser tab.
 */
(function () {
  "use strict";
  const CFG = window.DASHBOARD_CONFIG || {};
  const E = window.PurchaseEngine;
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const P = () => window.PPCR;
  const esc = (s) => P().esc(s);
  const LIVE = CFG.dataSource === "apps_script" && !!CFG.appsScriptUrl;
  const FIELDS = window.ENTRY_FIELDS, SECTIONS = window.ENTRY_SECTIONS;
  const LABEL = {}; FIELDS.forEach((f) => { LABEL[f.k] = f.label; });

  /* ---------------------------------------------------------------- session (name + key) */
  const ss = { get(k) { try { return sessionStorage.getItem(k); } catch (e) { return null; } }, set(k, v) { try { if (v === null) sessionStorage.removeItem(k); else sessionStorage.setItem(k, v); } catch (e) { /* unavailable */ } } };
  const session = { name: ss.get("ppcr-name") || "", key: ss.get("ppcr-key") || "", role: ss.get("ppcr-role") || "" };

  /* ---------------------------------------------------------------- backend: live (Apps Script) or demo */
  async function liveCall(action, payload) {
    const r = await fetch(CFG.appsScriptUrl, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(Object.assign({ action, key: session.key, name: session.name }, payload || {})) });
    if (!r.ok) throw new Error("The Google Sheet did not answer (HTTP " + r.status + ")");
    return r.json();
  }
  const demoLog = [];
  function demoCall(action, p) {
    p = p || {};
    const d = P().data, lists = d.lists || {}, fx = (d.settings || {}).fx || {};
    const refs = () => d.lines.map((l) => ({ _row: l._row, pr_no: l.pr_no, item: l.item }));
    const who = { by: session.name || "Demo user", role: session.role || "admin" };
    const log = (action2, line, row, changes, reason, record) => {
      const id = demoLog.length + 2;
      const t = new Date(); const z = (n) => String(n).padStart(2, "0");
      demoLog.push({ id, time: `${t.getFullYear()}-${z(t.getMonth() + 1)}-${z(t.getDate())} ${z(t.getHours())}:${z(t.getMinutes())}`, by: who.by, role: who.role,
        action: action2, pr_no: line ? line.pr_no : "", item: line ? line.item : "", row, changes: changes || null, reason: reason || "", record: record || null });
      return id;
    };
    const add = (raw, act) => {
      const n = window.normalizeEntry(raw || {}, lists), c = window.checkEntry(n.line, { existing: refs(), fx });
      if (String(act).startsWith("Restored")) c.errors = c.errors.filter((x) => /already has the item/.test(x));
      const errors = n.errors.concat(c.errors), warnings = n.warnings.concat(c.warnings);
      if (errors.length) return { ok: false, pr_no: n.line.pr_no, item: n.line.item, errors, warnings };
      const row = Math.max(5, ...d.lines.map((l) => l._row || 0)) + 1;
      d.lines.push(Object.assign({ _row: row }, n.line));
      log(act, n.line, row, null, "", n.line);
      return { ok: true, row, pr_no: n.line.pr_no, item: n.line.item, warnings };
    };
    const KINDS = Object.keys(window.LOG_KINDS);
    const lineLogsD = (pr, item) => { const key = window.lineKey(pr, item), o = {}; KINDS.forEach((k) => { o[k] = (d[k] || []).filter((e) => window.lineKey(e.pr_no, e.item) === key); }); return o; };
    const nextIdD = (kind) => { let mx = 0; (d[kind] || []).forEach((e) => { const m = String(e.id || "").match(/-(\d+)$/); if (m) mx = Math.max(mx, +m[1]); }); return window.LOG_KINDS[kind].prefix + "-" + String(mx + 1).padStart(4, "0"); };
    const nowISO = () => new Date().toISOString().slice(0, 16);
    const applySum = (line, kind) => Object.assign(line, window.rebuildPatch(line, lineLogsD(line.pr_no, line.item), kind));
    const addLogsD = (kind, raws, act) => raws.map((raw) => {
      const n = window.normalizeLog(kind, raw || {}, lists), e = n.entry;
      const line = d.lines.find((l) => window.lineKey(l.pr_no, l.item) === window.lineKey(e.pr_no, e.item));
      const logs = line ? lineLogsD(line.pr_no, line.item) : {};
      const c = n.errors.length ? { errors: [], warnings: [] } : window.checkLog(kind, e, { line, logs, today: localToday(), fx, invTol: (d.settings || {}).inv_tol, allInvoices: invoiceRefs(d) });
      const errors = n.errors.concat(c.errors);
      if (errors.length) return { ok: false, pr_no: e.pr_no, item: e.item, errors };
      d[kind] = d[kind] || [];
      const opening = (kind === "deliveries" || kind === "payments") && !(logs[kind] || []).length ? window.openingLogs(kind, line, localToday()) : [];
      opening.forEach((o) => d[kind].push(Object.assign({ id: nextIdD(kind), pr_no: line.pr_no, item: line.item, po_no: line.po_no || null, by: "System (opening balance)", at: nowISO() }, o)));
      const entry = Object.assign({}, e, { id: nextIdD(kind), pr_no: line.pr_no, item: line.item, po_no: line.po_no || null, by: who.by, at: nowISO() });
      d[kind].push(entry);
      if (kind === "deliveries" || kind === "payments") applySum(line, kind);
      const rec = Object.assign({ __kind: kind }, entry); delete rec.by; delete rec.at;
      log(act || ("Added " + window.LOG_KINDS[kind].title.split(" ")[0].toLowerCase() + " " + entry.id), line, line._row, null, "", rec);
      return { ok: true, id: entry.id, row: line._row, pr_no: line.pr_no, item: line.item, warnings: n.warnings.concat(c.warnings), opening: opening.length };
    });
    const find = () => { const l = d.lines.find((x) => x._row === Number(p.row)); if (!l || l.pr_no !== p.pr_no || String(l.item || "") !== String(p.item || "")) throw new Error("This line was changed or moved since you opened it. Refresh and try again."); return l; };
    const clean = (l) => { const o = {}; FIELDS.forEach((f) => { o[f.k] = l[f.k] === undefined ? null : l[f.k]; }); return o; };
    switch (action) {
      case "auth": return { ok: true, role: "admin" };
      case "add": return add(p.line, "Created");
      case "bulkAdd": { const res = p.lines.map((x) => add(x, "Created (bulk)")); return { ok: true, results: res, saved: res.filter((r) => r.ok).length }; }
      case "update": {
        if (!String(p.reason || "").trim()) return { ok: false, error: "Please give a reason for the change." };
        const l = find(), old = clean(l), n = window.normalizeEntry(p.line || {}, lists);
        const c = window.checkEntry(n.line, { existing: refs(), selfRow: l._row, fx });
        if (n.errors.length || c.errors.length) return { ok: false, errors: n.errors.concat(c.errors) };
        const ch = window.diffEntry(old, n.line);
        if (!Object.keys(ch).length) return { ok: true, unchanged: true, row: l._row };
        const locked = window.ownedFields(lineLogsD(old.pr_no, old.item)).filter((k) => ch[k]);
        if (locked.length) return { ok: false, error: "These fields come from the delivery / payment logs and cannot be edited directly: " + locked.map((k) => LABEL[k]).join(", ") + ". Add or delete a log entry instead." };
        if (ch.pr_no || ch.item || ch.po_no) KINDS.forEach((k) => (d[k] || []).forEach((e) => {
          if (window.lineKey(e.pr_no, e.item) === window.lineKey(old.pr_no, old.item)) { e.pr_no = n.line.pr_no; e.item = n.line.item; e.po_no = n.line.po_no; }
        }));
        Object.assign(l, n.line);
        return { ok: true, row: l._row, changes: ch, logId: log("Updated", n.line, l._row, ch, p.reason) };
      }
      case "delete": {
        if (!String(p.reason || "").trim()) return { ok: false, error: "Please give a reason for deleting." };
        const l = find(), old = clean(l);
        const lg = lineLogsD(old.pr_no, old.item), nl = KINDS.reduce((s2, k) => s2 + lg[k].length, 0);
        if (nl) return { ok: false, error: "This line has " + nl + " delivery / payment / quotation / invoice entries. Delete those first (Admin ▸ Edit ▸ Logs), then delete the line." };
        d.lines.splice(d.lines.indexOf(l), 1);
        return { ok: true, row: l._row, logId: log("Deleted", old, l._row, null, p.reason, old) };
      }
      case "revert": {
        if (!String(p.reason || "").trim()) return { ok: false, error: "Please give a reason for the undo." };
        const e = demoLog.find((x) => x.id === Number(p.logId));
        if (!e) return { ok: false, error: "That history entry does not exist." };
        if (e.action.startsWith("Deleted") && e.record && e.record.__kind) {
          const rec = Object.assign({}, e.record), kind = rec.__kind; delete rec.__kind;
          const r = addLogsD(kind, [rec], "Restored #" + e.id)[0];
          return r.ok ? { ok: true, restored: true, id: r.id, row: r.row } : { ok: false, errors: r.errors };
        }
        if (e.action.startsWith("Deleted")) { const r = add(e.record, "Restored #" + e.id); return r.ok ? { ok: true, row: r.row, restored: true } : { ok: false, errors: r.errors }; }
        if (e.action.startsWith("Updated") || e.action.startsWith("Reverted")) {
          const l = d.lines.find((x) => x._row === e.row);
          if (!l) return { ok: false, error: "The line on row " + e.row + " no longer exists." };
          const clash = Object.keys(e.changes).filter((k) => Object.keys(window.diffEntry({ [k]: l[k] }, { [k]: e.changes[k][1] })).length);
          if (clash.length) return { ok: false, error: "This line was changed again after that entry (" + clash.join(", ") + "). Undo the later change first." };
          const own = window.ownedFields(lineLogsD(l.pr_no, l.item)).filter((k) => e.changes[k]);
          if (own.length) return { ok: false, error: "That change touched fields now controlled by the delivery / payment logs (" + own.map((k) => LABEL[k]).join(", ") + "). Adjust the log entries instead." };
          const undo = {}; Object.keys(e.changes).forEach((k) => { l[k] = e.changes[k][0]; undo[k] = [e.changes[k][1], e.changes[k][0]]; });
          return { ok: true, row: l._row, logId: log("Reverted #" + e.id, l, l._row, undo, p.reason) };
        }
        return { ok: false, error: "Only updates and deletions can be undone." };
      }
      case "history": { let out = demoLog.slice().reverse(); if (p.pr_no) out = out.filter((x) => String(x.pr_no).toLowerCase() === String(p.pr_no).toLowerCase()); return { ok: true, entries: out }; }
      case "addListValue": {
        const v = String(p.value || "").trim(); if (!v || !p.list) return { ok: false, error: "Choose a list and type the new value." };
        const L = (d.lists[p.list] = d.lists[p.list] || []);
        if (L.some((x) => x.toLowerCase() === v.toLowerCase())) return { ok: false, error: "'" + v + "' is already in the list." };
        L.push(v); log("List value added", { pr_no: "", item: p.list + ": " + v }, 0, null, ""); return { ok: true };
      }
      case "addLog": {
        if (!window.LOG_KINDS[p.kind]) return { ok: false, error: "Unknown log type." };
        if (!p.entries || !p.entries.length) return { ok: false, error: "Nothing to save." };
        const res = addLogsD(p.kind, p.entries); return { ok: true, results: res, saved: res.filter((r) => r.ok).length };
      }
      case "deleteLog": {
        if (!String(p.reason || "").trim()) return { ok: false, error: "Please give a reason for deleting." };
        const K = window.LOG_KINDS[p.kind]; if (!K) return { ok: false, error: "Unknown log type." };
        const e = (d[p.kind] || []).find((x) => x.id === p.id);
        if (!e) return { ok: false, error: "Entry " + p.id + " was not found — refresh and try again." };
        if (p.kind === "payments" && e.entry === "Planned" && (d.payments || []).some((x) => x.settles === e.id)) return { ok: false, error: "Planned instalment " + e.id + " is settled by a payment. Delete that payment first." };
        d[p.kind] = d[p.kind].filter((x) => x.id !== p.id);
        const line = d.lines.find((l) => window.lineKey(l.pr_no, l.item) === window.lineKey(e.pr_no, e.item));
        if (line && (p.kind === "deliveries" || p.kind === "payments")) applySum(line, p.kind);
        const rec = Object.assign({ __kind: p.kind }, e); delete rec.by; delete rec.at;
        return { ok: true, logId: log("Deleted " + K.title.split(" ")[0].toLowerCase() + " " + e.id, e, line ? line._row : 0, null, p.reason, rec) };
      }
      default: return { ok: false, error: "Unknown action" };
    }
  }
  const WRITES = ["add", "bulkAdd", "update", "delete", "revert", "addListValue", "addLog", "deleteLog"];
  function localToday() { const t = new Date(), z = (n) => String(n).padStart(2, "0"); return `${t.getFullYear()}-${z(t.getMonth() + 1)}-${z(t.getDate())}`; }
  function invoiceRefs(d) {
    const sup = {}; d.lines.forEach((l) => { sup[window.lineKey(l.pr_no, l.item)] = String(l.supplier || "").toLowerCase(); });
    return (d.invoices || []).map((i) => ({ invoice_no: i.invoice_no, supplier_key: sup[window.lineKey(i.pr_no, i.item)] }));
  }
  async function call(action, payload) {
    let res;
    try { res = LIVE ? await liveCall(action, payload) : demoCall(action, payload); }
    catch (e) { return { ok: false, error: e.message }; }
    if (res && res.ok && WRITES.includes(action) && !res.unchanged) {
      if (LIVE) { try { await P().reload(); } catch (e) { P().toast("Saved, but the dashboard could not refresh: " + e.message); } }
      else P().setData(P().data);
    }
    return res;
  }

  /* ---------------------------------------------------------------- sign-in strip (shared by both tabs) */
  function signinHTML(needAdmin) {
    if (session.role && (!needAdmin || session.role === "admin"))
      return `<div class="signin on"><span>${P().pill("good", "Signed in")} <b>${esc(session.name)}</b> · ${session.role === "admin" ? "Admin" : "Data entry"}</span>
        <button class="linkbtn" type="button" data-signout>Sign out</button></div>`;
    return `<form class="signin" data-signin novalidate>
      <label>Your name<input id="si-name-${needAdmin ? "a" : "e"}" autocomplete="name" value="${esc(session.name)}" required></label>
      <label>${needAdmin ? "Admin key" : "Access key"}<input id="si-key-${needAdmin ? "a" : "e"}" type="password" autocomplete="current-password" ${LIVE ? "required" : `placeholder="not needed in demo"`}></label>
      <button class="btn primary" type="submit">Sign in</button>
      ${session.role === "entry" && needAdmin ? `<span class="cell-sub">You are signed in for data entry. Admin work needs the admin key.</span>` : ""}
      <span class="form-msg" role="status"></span></form>`;
  }
  function bindSignin(root, needAdmin, onDone) {
    const f = $("form[data-signin]", root);
    if (f) f.addEventListener("submit", async (e) => {
      e.preventDefault();
      const name = $("input[autocomplete=name]", f).value.trim(), key = $("input[type=password]", f).value.trim();
      const msg = $(".form-msg", f);
      if (!name) { msg.textContent = "Please enter your name — every change is logged with it."; return; }
      if (LIVE && !key) { msg.textContent = "Please enter the key."; return; }
      session.name = name; session.key = key || "demo";
      msg.textContent = "Checking…";
      const r = await call("auth", {});
      if (!r.ok) { msg.textContent = r.error || "Sign-in failed."; session.key = ""; return; }
      if (needAdmin && r.role !== "admin") { msg.textContent = "This key is for data entry. Admin work needs the admin key."; }
      session.role = r.role; ss.set("ppcr-name", name); ss.set("ppcr-key", session.key); ss.set("ppcr-role", r.role);
      renderEntry(); renderAdmin(); onDone && onDone();
    });
    const so = $("[data-signout]", root);
    if (so) so.addEventListener("click", () => { session.key = ""; session.role = ""; ss.set("ppcr-key", null); ss.set("ppcr-role", null); renderEntry(); renderAdmin(); });
  }
  const modeBanner = () => LIVE
    ? `<div class="mode live">${P().pill("good", "Live")} Saving to the Google Sheet <b>${esc((P().data.meta || {}).source || "")}</b>. Every change is logged on its Change_Log sheet.</div>`
    : `<div class="mode demo">${P().pill("warning", "Demo mode")} Changes stay in this browser tab only and are lost on reload. To save for everyone, connect the Google Sheet (README ▸ Live Google Sheet) and set <code>dataSource: "apps_script"</code> in config.js.</div>`;

  /* ---------------------------------------------------------------- the entry form */
  function fieldHTML(f, v, lists, pre, locked) {
    const id = `${pre}-${f.k}`, req = f.req ? ` <i class="req" title="Required">*</i>` : "";
    if (locked) {
      const shown = v === null || v === undefined || v === "" ? "" : f.type === "percent" ? +(v * 100).toFixed(4) : v;
      return `<label class="fld${f.wide ? " wide" : ""} locked" for="${id}" title="Comes from the ${f.k === "qty_recv" || f.k === "inhouse_start" ? "Deliveries" : "Payments"} log"><span>${esc(f.label)} <i class="lock">from log</i></span>
        <input id="${id}" data-k="${f.k}" type="${f.type === "date" ? "date" : f.type === "number" ? "number" : "text"}" value="${esc(shown)}" readonly></label>`;
    }
    let input;
    if (f.type === "list") {
      const opts = (lists[f.k] || []).slice();
      if (v && !opts.includes(v)) opts.unshift(v);
      input = `<select id="${id}" data-k="${f.k}"><option value=""></option>${opts.map((o) => `<option${o === v ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
    } else if (f.type === "date") input = `<input id="${id}" data-k="${f.k}" type="date" value="${esc(v || "")}">`;
    else if (f.type === "percent") input = `<span class="suffix"><input id="${id}" data-k="${f.k}" type="number" step="any" min="0" max="100" value="${v === null || v === undefined || v === "" ? "" : +(v * 100).toFixed(4)}"><b>%</b></span>`;
    else if (f.type === "number") input = `<input id="${id}" data-k="${f.k}" type="number" step="any" min="0" value="${v === null || v === undefined ? "" : esc(v)}">`;
    else input = `<input id="${id}" data-k="${f.k}" type="text" value="${esc(v || "")}"${f.ph ? ` placeholder="${esc(f.ph)}"` : ""}>`;
    return `<label class="fld${f.wide ? " wide" : ""}" for="${id}"><span>${esc(f.label)}${req}</span>${input}</label>`;
  }
  function formHTML(values, pre, locked) {
    const lists = (P().data && P().data.lists) || {}, L = new Set(locked || []);
    return Object.keys(SECTIONS).map((s) => `<fieldset class="fsec"><legend><b>${s}</b> ${esc(SECTIONS[s])}</legend><div class="fgrid">
      ${FIELDS.filter((f) => String(f.sec) === s).map((f) => fieldHTML(f, values[f.k], lists, pre, L.has(f.k))).join("")}</div></fieldset>`).join("");
  }
  function readForm(root) {
    const raw = {};
    $$("[data-k]", root).forEach((el) => {
      const f = FIELDS.find((x) => x.k === el.dataset.k);
      raw[f.k] = f.type === "percent" ? (el.value === "" ? "" : el.value + "%") : el.value;
    });
    return raw;
  }
  function evaluate(raw, selfRow) {
    const d = P().data, n = window.normalizeEntry(raw, d.lists || {});
    const refs = d.lines.map((l) => ({ _row: l._row, pr_no: l.pr_no, item: l.item }));
    const c = window.checkEntry(n.line, { existing: refs, selfRow, fx: (d.settings || {}).fx });
    return { line: n.line, errors: n.errors.concat(c.errors), warnings: n.warnings.concat(c.warnings) };
  }
  function previewHTML(ev) {
    const d = P().data, p = P();
    const l = E.enrich({ settings: d.settings, lines: [Object.assign({ _row: 0 }, ev.line, { pr_no: ev.line.pr_no || "(new)" })] }, p.R)[0] || {};
    const row = (k, v) => `<div><span>${k}</span><b>${v}</b></div>`;
    return `<div class="kv one">${[
      row("PO value (BDT)", l.po_value === null || l.po_value === undefined ? "—" : p.tk(l.po_value)),
      row("Total incl. tax", l.total_cost === null || l.total_cost === undefined ? "—" : p.tk(l.total_cost)),
      row("Paid", l.total_cost ? p.pct(l.pct_paid) : "—"), row("Outstanding", l.outstanding === null || l.outstanding === undefined ? "—" : p.tk(l.outstanding)),
      row("Stage", l.stage ? p.pill(p.stageState(l.stage), l.stage) : "—"), row("Payment", l.fin_status ? p.pill(p.finState(l.fin_status), l.fin_status) : "—"),
      row("In-house", l.pend_status ? p.pill(p.pendState(l.pend_status), l.pend_status) : "—"),
      row("Price vs budget", l.pv_pct === null || l.pv_pct === undefined ? "—" : (l.pv_pct >= 0 ? "+" : "") + p.pct(l.pv_pct, 1)),
    ].join("")}</div>
    ${ev.errors.length ? `<ul class="msgs err">${ev.errors.map((e) => `<li>${esc(e)}</li>`).join("")}</ul>` : `<p class="ok-line">${p.pill("good", "Ready to save")}</p>`}
    ${ev.warnings.length ? `<ul class="msgs warn">${ev.warnings.map((e) => `<li>${esc(e)}</li>`).join("")}</ul>` : ""}`;
  }

  /* ---------------------------------------------------------------- DATA ENTRY tab */
  const entryState = { mode: "single", draft: {}, bulk: null, bulkKind: "lines", log: {} };
  const MODES = [["single", "Purchase line"], ["deliveries", "Delivery"], ["payments", "Payment"], ["quotations", "Quotation"], ["invoices", "Invoice"], ["bulk", "Bulk entry"]];
  function renderEntry() {
    const root = $("#entry-root"); if (!root || !P() || !P().data) return;
    const signed = !!session.role;
    root.innerHTML = `${modeBanner()}
      <div class="card"><div class="card-h"><h2>Who is entering data</h2><span class="sub">${LIVE ? "Use the access key given by the admin" : "Demo mode: type your name, no key needed"}</span></div>${signinHTML(false)}</div>
      ${signed ? `<div class="seg wrap-seg" id="entry-seg">${MODES.map(([m, t]) => `<button type="button" data-m="${m}" aria-pressed="${entryState.mode === m}">${t}</button>`).join("")}</div>
      <div id="entry-body"></div>` : ""}`;
    bindSignin(root, false);
    if (!signed) return;
    $$("#entry-seg button").forEach((b) => b.addEventListener("click", () => { entryState.mode = b.dataset.m; renderEntry(); }));
    if (entryState.mode === "single") renderSingle(); else if (entryState.mode === "bulk") renderBulk(); else renderLog(entryState.mode);
  }

  function renderSingle() {
    const body = $("#entry-body"), d = P().data;
    const prs = Array.from(new Set(d.lines.map((l) => l.pr_no))).filter(Boolean).sort().reverse();
    body.innerHTML = `<div class="entry-grid">
      <form class="card" id="single-form" novalidate>
        <div class="card-h"><h2>New purchase line</h2>
          <label class="inline">Copy PR details from <select id="prefill"><option value="">— new PR —</option>${prs.map((p) => `<option>${esc(p)}</option>`).join("")}</select></label></div>
        <p class="note" style="margin-top:-6px">Fields marked <i class="req">*</i> are required. Fill later sections (PO, payments, in-house) when they happen — an admin can update the line.</p>
        ${formHTML(entryState.draft, "ne")}
        <div class="form-actions"><button class="btn primary" type="submit" id="ne-save">Save line</button>
          <button class="btn" type="button" id="ne-save-next">Save &amp; add another item to this PR</button>
          <button class="btn ghost" type="button" id="ne-clear">Clear form</button><span class="form-msg" role="status"></span></div>
      </form>
      <aside class="card sticky"><div class="card-h"><h2>Live check</h2><span class="sub">What the dashboard will show</span></div><div id="ne-preview"></div></aside></div>`;
    const form = $("#single-form");
    const refresh = () => { entryState.draft = window.normalizeEntry(readForm(form), {}).line; const ev = evaluate(readForm(form)); $("#ne-preview").innerHTML = previewHTML(ev); return ev; };
    form.addEventListener("input", refresh); form.addEventListener("change", refresh);
    refresh();
    $("#prefill").addEventListener("change", (e) => {
      const l = d.lines.find((x) => x.pr_no === e.target.value); if (!l) return;
      const keep = {}; FIELDS.filter((f) => f.sec === 1 && !["item", "spec", "category"].includes(f.k)).forEach((f) => { keep[f.k] = l[f.k]; });
      entryState.draft = keep; renderSingle();
    });
    $("#ne-clear").addEventListener("click", () => { entryState.draft = {}; renderSingle(); });
    const save = async (next) => {
      const ev = refresh(), msg = $(".form-msg", form);
      if (ev.errors.length) { msg.textContent = "Please fix the items listed under Live check."; $("#ne-preview").scrollIntoView({ block: "nearest" }); return; }
      msg.textContent = "Saving…"; $$("button", form).forEach((b) => { b.disabled = true; });
      const r = await call("add", { line: readForm(form) });
      $$("button", form).forEach((b) => { b.disabled = false; });
      if (!r.ok) { msg.textContent = r.error || (r.errors || []).join(" · "); return; }
      P().toast(`Saved ${r.pr_no} · ${r.item} to row ${r.row}.`);
      if (next) { const keep = {}; FIELDS.filter((f) => f.sec === 1 && !["item", "spec", "category"].includes(f.k)).forEach((f) => { keep[f.k] = ev.line[f.k]; }); entryState.draft = keep; }
      else entryState.draft = {};
      renderSingle();
      $("#ne-item") && next && $("#ne-item").focus();
    };
    form.addEventListener("submit", (e) => { e.preventDefault(); save(false); });
    $("#ne-save-next").addEventListener("click", () => save(true));
  }

  /* ---------------------------------------------------------------- log entries: delivery · payment · quotation · invoice */
  const lineLabel = (l) => `${l.pr_no} · ${l.item}${l.po_no ? " · " + l.po_no : ""}${l.supplier ? " · " + l.supplier : ""}`;
  function lineLogs(d, l) { const key = window.lineKey(l.pr_no, l.item), o = {}; Object.keys(window.LOG_KINDS).forEach((k) => { o[k] = (d[k] || []).filter((e) => window.lineKey(e.pr_no, e.item) === key); }); return o; }
  function logFieldHTML(f, v, lists, pre, openPlans) {
    const id = `${pre}-${f.k}`, req = f.req ? ` <i class="req" title="Required">*</i>` : "";
    let input;
    if (f.k === "settles") input = `<select id="${id}" data-k="${f.k}"><option value="">— none —</option>${(openPlans || []).map((x) => `<option value="${esc(x.id)}"${x.id === v ? " selected" : ""}>${esc(x.id)} · ${esc(x.date)} · ৳${P().inr(x.amount)}</option>`).join("")}</select>`;
    else if (f.type === "choice" || f.type === "list") {
      const opts = (f.type === "choice" ? f.opts : (lists[f.list] || [])).slice();
      if (v && !opts.includes(v)) opts.unshift(v);
      input = `<select id="${id}" data-k="${f.k}">${f.req && f.type === "choice" ? "" : `<option value=""></option>`}${opts.map((o) => `<option${o === v ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
    } else if (f.type === "date") input = `<input id="${id}" data-k="${f.k}" type="date" value="${esc(v || "")}">`;
    else if (f.type === "number") input = `<input id="${id}" data-k="${f.k}" type="number" step="any" min="0" value="${v === null || v === undefined ? "" : esc(v)}">`;
    else input = `<input id="${id}" data-k="${f.k}" type="text" value="${esc(v || "")}">`;
    return `<label class="fld${f.wide ? " wide" : ""}" for="${id}"><span>${esc(f.label)}${req}</span>${input}</label>`;
  }
  /* Simulate the effect of new log entries on the dashboard (same engine, copied data). */
  function simulate(kind, entries) {
    const d = P().data, sim = Object.assign({}, d, { lines: d.lines.map((l) => Object.assign({}, l)) });
    Object.keys(window.LOG_KINDS).forEach((k) => { if (d[k] || k === kind) sim[k] = (d[k] || []).slice(); });
    entries.forEach((e, i) => {
      const line = sim.lines.find((l) => window.lineKey(l.pr_no, l.item) === window.lineKey(e.pr_no, e.item)); if (!line) return;
      const logs = lineLogs(sim, line);
      if ((kind === "deliveries" || kind === "payments") && !logs[kind].length)
        window.openingLogs(kind, line, localToday()).forEach((o, j) => sim[kind].push(Object.assign({ id: "OPEN-" + i + "-" + j, pr_no: line.pr_no, item: line.item }, o)));
      sim[kind].push(Object.assign({ id: "NEW-" + i }, e, { pr_no: line.pr_no, item: line.item }));
      if (kind === "deliveries" || kind === "payments") Object.assign(line, window.rebuildPatch(line, lineLogs(sim, line), kind));
    });
    const L = E.enrich(sim, P().R); E.analyze(L, sim, P().R);
    return L;
  }
  function lineStatusHTML(before, after) {
    const p = P();
    const rows = [
      ["In-house", (l) => l.pend_status ? p.pill(p.pendState(l.pend_status), l.pend_status) : "—"],
      ["Received", (l) => l.qty === null ? "—" : `${p.inr(l.qty_recv || 0)} of ${p.inr(l.qty)} ${esc(l.uom || "")}`],
      ["Payment", (l) => l.fin_status ? p.pill(p.finState(l.fin_status), l.fin_status) : "—"],
      ["Paid / outstanding", (l) => l.total_cost === null ? "—" : `${p.tk(l.total_paid)} / ${p.tk(l.outstanding)}`],
      ["Next payment due", (l) => p.dfmt(l.due_date)],
      ["Invoice match", (l) => l.match_status ? p.pill(l.match_status === "Matched" ? "good" : l.match_status === "Mismatch" ? "critical" : "info", l.match_status) : "—"],
      ["Sourcing", (l) => l.q_status ? p.pill(l.q_issue ? "serious" : /needed/.test(l.q_status) ? "warning" : "good", l.q_status + (l.q_count ? ` · ${l.q_count} quote(s)` : "")) : "—"],
    ];
    return `<div class="nowafter">${rows.map(([n, f]) => {
      const a = f(before), b = after ? f(after) : a;
      return `<div class="na-row${a !== b ? " chg" : ""}"><div class="na-l">${n}${a !== b ? ' <span class="na-tag">changes</span>' : ""}</div>
        <div class="na-v">${a !== b ? `<span class="na-old">${a}</span><span class="na-arrow" aria-label="becomes">→</span><span class="na-new">${b}</span>` : a}</div></div>`; }).join("")}</div>`;
  }
  function renderLog(kind) {
    const body = $("#entry-body"), d = P().data, K = window.LOG_KINDS[kind], lists = d.lists || {};
    const st = (entryState.log[kind] = entryState.log[kind] || { key: "", draft: {} });
    const elig = d.lines.filter((l) => l.pr_no && l.item && (K.needsPO ? !!l.po_no : true));
    const byLabel = {}; elig.forEach((l) => { byLabel[lineLabel(l)] = l; });
    const line = elig.find((l) => window.lineKey(l.pr_no, l.item) === st.key) || null;
    const logs = line ? lineLogs(d, line) : {};
    const settled = new Set((logs.payments || []).filter((x) => x.settles).map((x) => x.settles));
    const openPlans = (logs.payments || []).filter((x) => x.entry === "Planned" && !settled.has(x.id));
    const dflt = Object.assign({}, kind === "quotations" && line ? { ccy: line.ccy } : {}, kind === "invoices" && line ? { unit_price: line.unit_price } : {},
      kind === "payments" ? { entry: "Paid", ptype: "Advance" } : {}, { date: localToday() },
      Object.fromEntries(Object.entries(st.draft).filter(([, v]) => v !== "" && v !== null && v !== undefined)));
    const hint = { deliveries: "Record each consignment as it arrives. Received quantity, first in-house date and (when complete) the closing date update on the register automatically.",
      payments: "Record a payment made (Paid) or a scheduled instalment (Planned). Paid amounts and the next due date update on the register automatically.",
      quotations: "Record every supplier quote for the line — before or after the PO. The dashboard checks that enough quotes were taken and that the lowest was chosen.",
      invoices: "Record the supplier's invoice. It is matched against the PO price and the quantity accepted at QC (three-way match) before payment." }[kind];
    body.innerHTML = `<div class="entry-grid"><form class="card" id="log-form" novalidate>
      <div class="card-h"><h2>New ${esc(K.title.toLowerCase())}</h2></div>
      <p class="note" style="margin-top:-6px">${esc(hint)}</p>
      <label class="fld wide" for="lg-line"><span>Purchase line <i class="req">*</i></span>
        <input id="lg-line" list="lg-lines" placeholder="Type a PR No, PO No, item or supplier…" value="${line ? esc(lineLabel(line)) : ""}" autocomplete="off"></label>
      <datalist id="lg-lines">${elig.map((l) => `<option value="${esc(lineLabel(l))}"></option>`).join("")}</datalist>
      ${K.needsPO ? `<p class="cell-sub">Only lines with a PO are listed.</p>` : ""}
      <fieldset class="fsec" style="margin-top:10px"><legend>Details</legend><div class="fgrid">${K.fields.filter((f) => !(f.k === "settles" && !openPlans.length)).map((f) => logFieldHTML(f, dflt[f.k], lists, "lg", openPlans)).join("")}</div></fieldset>
      <div class="form-actions"><button class="btn primary" type="submit" id="lg-save">Save</button><button class="btn" type="button" id="lg-next">Save &amp; add another for this line</button>
        <button class="btn ghost" type="button" id="lg-clear">Clear</button><span class="form-msg" role="status"></span></div>
      </form>
      <aside class="card sticky"><div class="card-h"><h2>Live check</h2><span class="sub">Now → after saving</span></div><div id="lg-preview"></div></aside></div>
      <div class="card" id="lg-existing"></div>`;
    const form = $("#log-form");
    const read = () => { const raw = {}; $$("[data-k]", form).forEach((el) => { raw[el.dataset.k] = el.value; }); if (line) { raw.pr_no = line.pr_no; raw.item = line.item; } return raw; };
    const check = () => {
      const raw = read(); st.draft = Object.assign({}, raw); delete st.draft.pr_no; delete st.draft.item;
      if (!line) { $("#lg-preview").innerHTML = `<p class="note">Choose the purchase line first.</p>`; return { errors: ["Choose the purchase line"] }; }
      const n = window.normalizeLog(kind, raw, lists);
      const c = n.errors.length ? { errors: [], warnings: [] } : window.checkLog(kind, n.entry, { line, logs, today: localToday(), fx: (d.settings || {}).fx, invTol: (d.settings || {}).inv_tol, allInvoices: invoiceRefs(d) });
      const errors = n.errors.concat(c.errors), warnings = n.warnings.concat(c.warnings);
      const before = P().all.find((l) => window.lineKey(l.pr_no, l.item) === st.key);
      let after = null; if (!errors.length) after = simulate(kind, [n.entry]).find((l) => window.lineKey(l.pr_no, l.item) === st.key);
      $("#lg-preview").innerHTML = (before ? lineStatusHTML(before, after) : "") +
        (errors.length ? `<ul class="msgs err">${errors.map((e) => `<li>${esc(e)}</li>`).join("")}</ul>` : `<p class="ok-line">${P().pill("good", "Ready to save")}</p>`) +
        (warnings.length ? `<ul class="msgs warn">${warnings.map((e) => `<li>${esc(e)}</li>`).join("")}</ul>` : "") +
        ((kind === "deliveries" || kind === "payments") && !(logs[kind] || []).length && (kind === "deliveries" ? line.qty_recv : (line.adv_paid || line.interim_paid || line.closing_paid))
          ? `<p class="note">This is the first ${kind === "deliveries" ? "delivery" : "payment"} logged for this line: the amount already on the register is kept as an opening-balance entry.</p>` : "");
      return { errors, raw };
    };
    $("#lg-line").addEventListener("change", (e) => { const l = byLabel[e.target.value]; st.key = l ? window.lineKey(l.pr_no, l.item) : ""; renderLog(kind); });
    form.addEventListener("input", (e) => { if (e.target.id !== "lg-line") check(); }); form.addEventListener("change", (e) => { if (e.target.id !== "lg-line") check(); });
    check();
    const ex = line ? logs[kind] : [];
    $("#lg-existing").innerHTML = `<div class="card-h"><h2>${esc(K.sheet)} for this line</h2><span class="sub">${line ? ex.length + " entr" + (ex.length === 1 ? "y" : "ies") : "Choose a line"}</span></div>` +
      (line && ex.length ? logTableHTML(kind, ex, false) : `<div class="empty">${line ? "Nothing recorded yet." : "Choose a purchase line to see its history."}</div>`);
    $("#lg-clear").addEventListener("click", () => { st.draft = {}; renderLog(kind); });
    const save = async (again) => {
      const r0 = check(), msg = $(".form-msg", form);
      if (r0.errors.length) { msg.textContent = "Please fix the items listed under Live check."; return; }
      msg.textContent = "Saving…"; $$("button", form).forEach((b) => { b.disabled = true; });
      const r = await call("addLog", { kind, entries: [r0.raw] });
      $$("button", form).forEach((b) => { b.disabled = false; });
      const res = r.ok ? r.results[0] : null;
      if (!r.ok || !res.ok) { msg.textContent = r.error || (res && res.errors || []).join(" · "); return; }
      P().toast(`Saved ${res.id} for ${res.pr_no}${res.opening ? " (plus an opening-balance entry)" : ""}.`);
      st.draft = {}; if (!again) st.key = "";
      renderLog(kind);
    };
    form.addEventListener("submit", (e) => { e.preventDefault(); save(false); });
    $("#lg-next").addEventListener("click", () => save(true));
  }
  function logTableHTML(kind, entries, withDelete) {
    const K = window.LOG_KINDS[kind], p = P();
    const fmt = (f, v) => v === null || v === undefined || v === "" ? "—" : f.type === "number" ? p.inr(v, Number.isInteger(v) ? 0 : 2) : f.type === "date" ? p.dfmt(E.toDay(v)) : esc(v);
    const cols = K.fields.filter((f) => f.k !== "remarks");
    return `<div class="tbl-wrap"><table><thead><tr><th>ID</th>${cols.map((f) => `<th${f.type === "number" ? ' class="n"' : ""}>${esc(f.label)}</th>`).join("")}<th>By</th>${withDelete ? "<th></th>" : ""}</tr></thead><tbody>
      ${entries.map((e) => `<tr><td class="mono">${esc(e.id)}</td>${cols.map((f) => `<td${f.type === "number" ? ' class="n"' : ""}>${fmt(f, e[f.k])}</td>`).join("")}
        <td class="cell-sub">${esc(e.by || "")}</td>${withDelete ? `<td><button class="btn sm danger" type="button" data-dellog="${esc(kind)}|${esc(e.id)}">Delete</button></td>` : ""}</tr>`).join("")}</tbody></table></div>`;
  }

  /* bulk: parse pasted text / CSV / Excel into rows keyed by field */
  function parseDelimited(text) {
    const delim = text.includes("\t") ? "\t" : ",";
    const rows = []; let row = [], cell = "", q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; continue; }
      if (ch === '"' && cell === "") { q = true; continue; }
      if (ch === delim) { row.push(cell); cell = ""; continue; }
      if (ch === "\n" || ch === "\r") { if (ch === "\r" && text[i + 1] === "\n") i++; row.push(cell); rows.push(row); row = []; cell = ""; continue; }
      cell += ch;
    }
    if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }
  const bulkFields = (kind) => kind === "lines" ? FIELDS : [{ k: "pr_no", label: "PR No" }, { k: "item", label: "Item / Material Purchased" }].concat(window.LOG_KINDS[kind].fields);
  function rowsToRaw(rows, kind) {
    const norm = (s) => String(s === null || s === undefined ? "" : s).trim().toLowerCase();
    const FIELDS = bulkFields(kind || "lines");
    const byLabel = { item: "item" }; FIELDS.forEach((f) => { byLabel[norm(f.label)] = f.k; byLabel[f.k] = f.k; });
    let hi = rows.findIndex((r, i) => i < 15 && r.some((c) => byLabel[norm(c)] === "pr_no"));
    let map;
    if (hi >= 0) map = rows[hi].map((c) => byLabel[norm(c)] || null);
    else { hi = -1; map = FIELDS.map((f) => f.k); }
    const out = [];
    rows.slice(hi + 1).forEach((r, i) => {
      if (!r.some((c) => String(c === null || c === undefined ? "" : c).trim() !== "")) return;
      const raw = { _src: hi + 2 + i };
      map.forEach((k, j) => { if (k) raw[k] = r[j] === undefined ? "" : r[j]; });
      out.push(raw);
    });
    return { rows: out, mapped: map.filter(Boolean).length, header: hi >= 0 };
  }
  function checkLogBatch(kind, raws) {
    const d = P().data, sim = Object.assign({}, d); sim[kind] = (d[kind] || []).slice();
    return raws.map((raw, i) => {
      const n = window.normalizeLog(kind, raw, d.lists || {}), e = n.entry;
      const line = d.lines.find((l) => window.lineKey(l.pr_no, l.item) === window.lineKey(e.pr_no, e.item)) || null;
      const logs = line ? lineLogs(sim, line) : {};
      const c = n.errors.length ? { errors: [], warnings: [] } : window.checkLog(kind, e, { line, logs, today: localToday(), fx: (d.settings || {}).fx, invTol: (d.settings || {}).inv_tol, allInvoices: invoiceRefs(sim) });
      const errors = n.errors.concat(c.errors);
      if (!errors.length) sim[kind].push(Object.assign({ id: "BATCH-" + i }, e, { pr_no: line.pr_no, item: line.item }));
      return { raw, line: Object.assign({ supplier: line && line.supplier, po_no: line && line.po_no }, e), errors, warnings: n.warnings.concat(c.warnings) };
    });
  }
  function checkBatch(raws) {
    const d = P().data, refs = d.lines.map((l) => ({ _row: l._row, pr_no: l.pr_no, item: l.item }));
    return raws.map((raw, i) => {
      const n = window.normalizeEntry(raw, d.lists || {});
      const c = window.checkEntry(n.line, { existing: refs, fx: (d.settings || {}).fx });
      const errors = n.errors.concat(c.errors);
      if (!errors.length) refs.push({ _row: -1 - i, pr_no: n.line.pr_no, item: n.line.item });
      return { raw, line: n.line, errors, warnings: n.warnings.concat(c.warnings) };
    });
  }
  function renderBulk() {
    const body = $("#entry-body");
    const b = entryState.bulk, bk = entryState.bulkKind;
    const BF = bulkFields(bk);
    const tsvHeader = BF.map((f) => f.label).join("\t");
    const KN = [["lines", "Purchase lines"], ["deliveries", "Deliveries"], ["payments", "Payments"], ["quotations", "Quotations"], ["invoices", "Invoices"]];
    body.innerHTML = `<div class="card">
      <div class="card-h"><h2>Bulk entry</h2><span class="sub">Paste from Excel, or choose a CSV / Excel file. Rows with problems are skipped, never half-saved.</span></div>
      <div class="seg wrap-seg" id="bk-kind" style="margin-bottom:14px">${KN.map(([k, t]) => `<button type="button" data-k="${k}" aria-pressed="${bk === k}">${t}</button>`).join("")}</div>
      <ol class="bulk-steps">
        <li><b>Get the column layout.</b> <button class="btn" type="button" id="bk-copy">Copy column headers</button> <button class="btn" type="button" id="bk-tpl">Download CSV template</button>
          <div class="cell-sub">${bk === "lines" ? "Paste the headers into row 1 of a blank Excel sheet and fill one purchase line per row. Columns may be in any order; extra columns are ignored. You can also paste straight from the Purchase_Register sheet (header row included)."
            : "Each row needs the PR No and the item exactly as on the register. Columns may be in any order; extra columns are ignored. You can also paste from the " + esc(window.LOG_KINDS[bk].sheet) + " sheet."}</div></li>
        <li><b>Paste the rows</b> (with the header row) or <label class="btn" for="bk-file">choose a file</label><input type="file" id="bk-file" accept=".csv,.txt,.xlsx,.xlsm" class="sr-only">
          <textarea id="bk-text" rows="7" placeholder="${esc(BF.slice(0, 6).map((f) => f.label).join("\t"))}\t…"></textarea>
          <div class="form-actions"><button class="btn primary" type="button" id="bk-check">Check rows</button><button class="btn ghost" type="button" id="bk-reset">Start again</button><span class="form-msg" role="status"></span></div></li>
      </ol>
      <div id="bk-review">${b ? bulkReviewHTML(b) : ""}</div></div>`;
    $("#bk-copy").addEventListener("click", () => {
      const done = () => P().toast("Column headers copied — paste them into row 1 of an Excel sheet.");
      try { navigator.clipboard.writeText(tsvHeader).then(done, () => { $("#bk-text").value = tsvHeader + "\n"; $("#bk-text").select(); P().toast("Headers placed in the box — copy them from there."); }); }
      catch (e) { $("#bk-text").value = tsvHeader + "\n"; $("#bk-text").select(); }
    });
    $("#bk-tpl").addEventListener("click", () => {
      const csv = "﻿" + BF.map((f) => `"${f.label.replace(/"/g, '""')}"`).join(",") + "\r\n";
      const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" })); a.download = (bk === "lines" ? "purchase_lines" : bk) + "_template.csv";
      document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
    });
    const runCheck = (rows, label) => {
      const parsed = rowsToRaw(rows, bk);
      if (!parsed.rows.length) { $(".form-msg", body).textContent = "No rows found."; return; }
      if (parsed.header && parsed.mapped < Math.min(5, BF.filter((f) => f.req || f.k === "pr_no" || f.k === "item").length)) { $(".form-msg", body).textContent = "Only " + parsed.mapped + " columns were recognised — check the header row."; }
      entryState.bulk = { kind: bk, source: label, items: bk === "lines" ? checkBatch(parsed.rows) : checkLogBatch(bk, parsed.rows), header: parsed.header };
      $("#bk-review").innerHTML = bulkReviewHTML(entryState.bulk); bindBulkReview();
    };
    $("#bk-check").addEventListener("click", () => {
      const t = $("#bk-text").value; if (!t.trim()) { $(".form-msg", body).textContent = "Paste some rows first."; return; }
      runCheck(parseDelimited(t.replace(/\s+$/, "")), "pasted rows");
    });
    $("#bk-reset").addEventListener("click", () => { entryState.bulk = null; renderBulk(); });
    $("#bk-file").addEventListener("change", async (e) => {
      const f = e.target.files && e.target.files[0]; if (!f) return;
      try {
        if (/\.(xlsx|xlsm)$/i.test(f.name)) {
          const XLSX = await window.PurchaseXlsx.loadLib();
          const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: false });
          const sh = (bk === "lines" ? wb.Sheets.Purchase_Register : wb.Sheets[window.LOG_KINDS[bk].sheet]) || wb.Sheets[wb.SheetNames[0]];
          runCheck(XLSX.utils.sheet_to_json(sh, { header: 1, raw: true, defval: "" }), f.name);
        } else runCheck(parseDelimited((await f.text()).replace(/^﻿/, "").replace(/\s+$/, "")), f.name);
      } catch (err) { P().toast("Could not read " + f.name + ": " + err.message); }
      e.target.value = "";
    });
    $$("#bk-kind button").forEach((x) => x.addEventListener("click", () => { entryState.bulkKind = x.dataset.k; entryState.bulk = null; renderBulk(); }));
    if (b) bindBulkReview();
  }
  function bulkReviewHTML(b) {
    const good = b.items.filter((x) => !x.errors.length && !x.saved), bad = b.items.filter((x) => x.errors.length && !x.saved), saved = b.items.filter((x) => x.saved);
    const p = P();
    return `<div class="bulk-sum">${p.pill("good", good.length + " ready")} ${bad.length ? p.pill("critical", bad.length + " with problems") : ""} ${saved.length ? p.pill("info", saved.length + " saved") : ""}
      <span class="cell-sub">from ${esc(b.source)}</span>
      ${good.length ? `<button class="btn primary" type="button" id="bk-save">Save ${good.length} ${b.kind && b.kind !== "lines" ? "entr" + (good.length > 1 ? "ies" : "y") : "line" + (good.length > 1 ? "s" : "")}</button>` : ""}</div>
      <div class="tbl-wrap"><table><thead><tr><th>Row</th><th>Status</th><th>PR No</th><th>Item</th><th>Supplier</th><th>PO No</th><th class="n">${b.kind === "payments" ? "Amount" : b.kind === "quotations" ? "Unit price" : "Qty"}</th><th>Problems / notes</th></tr></thead><tbody>
      ${b.items.map((x) => `<tr><td class="n">${x.raw._src}</td><td>${x.saved ? p.pill("info", "Saved · " + x.saved) : x.errors.length ? p.pill("critical", "Fix") : p.pill("good", "Ready")}</td>
        <td class="mono">${esc(x.line.pr_no || "—")}</td><td>${esc(x.line.item || "—")}</td><td>${esc(x.line.supplier || "—")}</td><td class="mono">${esc(x.line.po_no || "—")}</td>
        <td class="n">${(() => { const v = b.kind === "payments" ? x.line.amount : b.kind === "quotations" ? x.line.unit_price : x.line.qty; return v === null || v === undefined ? "—" : p.inr(v, Number.isInteger(v) ? 0 : 2); })()}</td>
        <td>${x.errors.map((e) => `<div class="t-err">${esc(e)}</div>`).join("")}${x.warnings.map((w) => `<div class="t-warn">${esc(w)}</div>`).join("")}</td></tr>`).join("")}
      </tbody></table></div>`;
  }
  function bindBulkReview() {
    const btn = $("#bk-save"); if (!btn) return;
    btn.addEventListener("click", async () => {
      const b = entryState.bulk, todo = b.items.filter((x) => !x.errors.length && !x.saved);
      btn.disabled = true; btn.textContent = "Saving…";
      const payload = todo.map((x) => { const o = Object.assign({}, x.raw); delete o._src; return o; });
      const r = b.kind && b.kind !== "lines" ? await call("addLog", { kind: b.kind, entries: payload }) : await call("bulkAdd", { lines: payload });
      if (!r.ok) { btn.disabled = false; btn.textContent = "Save"; P().toast(r.error || "Could not save."); return; }
      r.results.forEach((res, i) => { if (res.ok) todo[i].saved = res.id || ("row " + res.row); else todo[i].errors = res.errors || ["Not saved"]; });
      P().toast(`${r.saved} line(s) saved${r.saved < todo.length ? ", " + (todo.length - r.saved) + " rejected — see the table" : ""}.`);
      $("#bk-review").innerHTML = bulkReviewHTML(b); bindBulkReview();
    });
  }

  /* ---------------------------------------------------------------- ADMIN tab */
  const adminState = { q: "", edit: null, del: null, hist: [], histPR: "", undo: null };
  function renderAdmin() {
    const root = $("#admin-root"); if (!root || !P() || !P().data) return;
    const isAdmin = session.role === "admin";
    root.innerHTML = `${modeBanner()}
      <div class="card"><div class="card-h"><h2>Admin sign-in</h2><span class="sub">Edit, delete and undo need the admin key</span></div>${signinHTML(true)}</div>
      ${isAdmin ? `<div id="admin-body"></div>` : ""}`;
    bindSignin(root, true, () => loadHistory());
    if (!isAdmin) return;
    const body = $("#admin-body");
    if (adminState.edit) return renderEdit(body);
    const p = P(), q = adminState.q.toLowerCase();
    const L = p.all.filter((l) => !q || [l.pr_no, l.po_no, l.item, l.supplier, l.dept].join(" ").toLowerCase().includes(q)).sort((a, b) => (b._row || 0) - (a._row || 0));
    body.innerHTML = `<div class="grid">
      <div class="card c7"><div class="card-h"><h2>Purchase lines</h2><span class="sub">${L.length} line(s) · newest first</span></div>
        <div class="tbl-tools"><input type="search" id="ad-q" placeholder="Search PR, PO, item, supplier…" value="${esc(adminState.q)}" aria-label="Search lines"></div>
        ${adminState.del ? deletePanel() : ""}
        <div class="tbl-wrap" style="max-height:620px;overflow:auto"><table><thead><tr><th class="n">Row</th><th>PR No</th><th>Item · supplier</th><th>Stage</th><th></th></tr></thead><tbody>
        ${L.slice(0, 300).map((l) => `<tr><td class="n">${l._row || "—"}</td><td class="mono">${esc(l.pr_no)}</td><td>${esc(l.item || "—")}<div class="cell-sub">${esc(l.supplier || "—")}</div></td>
          <td>${p.pill(p.stageState(l.stage), l.stage)}</td>
          <td style="white-space:nowrap"><button class="btn sm" type="button" data-edit="${l._row}">Edit</button> <button class="btn sm danger" type="button" data-del="${l._row}">Delete</button></td></tr>`).join("")}
        </tbody></table></div></div>
      <div class="c5" style="display:grid;gap:16px;align-content:start">
        <div class="card"><div class="card-h"><h2>Change history</h2><span class="sub">Undo any update or deletion</span></div>
          <div class="tbl-tools"><input type="search" id="ad-hpr" placeholder="Filter by PR No" value="${esc(adminState.histPR)}" aria-label="Filter history by PR"><button class="btn" type="button" id="ad-hload">Show</button></div>
          <div id="ad-hist">${historyHTML()}</div></div>
        <div class="card"><div class="card-h"><h2>Dropdown lists</h2><span class="sub">Add a new supplier, department, category…</span></div>
          <form class="list-add" id="ad-list" novalidate><select id="ad-list-k" aria-label="List">${["supplier", "dept", "category", "uom", "pay_terms", "ccy", "origin", "qc_status"].map((k) => `<option value="${k}">${esc(LABEL[k])}</option>`).join("")}</select>
            <input id="ad-list-v" placeholder="New value" aria-label="New value"><button class="btn primary" type="submit">Add</button><span class="form-msg" role="status"></span></form></div>
      </div></div>`;
    $("#ad-q").addEventListener("input", (e) => { adminState.q = e.target.value; const pos = e.target.selectionStart; renderAdmin(); const n = $("#ad-q"); n.focus(); n.setSelectionRange(pos, pos); });
    $$("[data-edit]", body).forEach((b) => b.addEventListener("click", () => { const l = p.all.find((x) => x._row === Number(b.dataset.edit)); adminState.edit = { row: l._row, pr_no: l.pr_no, item: l.item, orig: rawOf(l) }; adminState.del = null; renderAdmin(); window.scrollTo({ top: 0 }); }));
    $$("[data-del]", body).forEach((b) => b.addEventListener("click", () => { const l = p.all.find((x) => x._row === Number(b.dataset.del)); adminState.del = { row: l._row, pr_no: l.pr_no, item: l.item, supplier: l.supplier, stage: l.stage }; renderAdmin(); }));
    const dp = $("#ad-delform");
    if (dp) {
      $("#ad-delcancel").addEventListener("click", () => { adminState.del = null; renderAdmin(); });
      dp.addEventListener("submit", async (e) => {
        e.preventDefault();
        const reason = $("#ad-delreason").value.trim(), msg = $(".form-msg", dp);
        if (!reason) { msg.textContent = "Please give a reason."; return; }
        msg.textContent = "Deleting…";
        const r = await call("delete", { row: adminState.del.row, pr_no: adminState.del.pr_no, item: adminState.del.item, reason });
        if (!r.ok) { msg.textContent = r.error || (r.errors || []).join(" · "); return; }
        P().toast(`Deleted ${adminState.del.pr_no} · ${adminState.del.item}. You can restore it from Change history.`);
        adminState.del = null; await loadHistory(); renderAdmin();
      });
    }
    $("#ad-hload").addEventListener("click", () => { adminState.histPR = $("#ad-hpr").value.trim(); loadHistory(); });
    bindHistory();
    $("#ad-list").addEventListener("submit", async (e) => {
      e.preventDefault();
      const msg = $(".form-msg", e.target), list = $("#ad-list-k").value, value = $("#ad-list-v").value.trim();
      msg.textContent = "Adding…";
      const r = await call("addListValue", { list, value });
      if (!r.ok) { msg.textContent = r.error; return; }
      P().toast(`'${value}' added to the ${LABEL[list]} list.`); await loadHistory(); renderAdmin();
    });
  }
  const rawOf = (l) => { const o = {}; FIELDS.forEach((f) => { const v = l[f.k]; o[f.k] = f.type === "date" ? (v === null || v === undefined ? null : (typeof v === "number" ? E.dayToISO(v) : v)) : v; }); return o; };
  function deletePanel() {
    const x = adminState.del, p = P();
    return `<form class="danger-box" id="ad-delform" novalidate><b>Delete row ${x.row}: ${esc(x.pr_no)} · ${esc(x.item || "")}</b>
      <div class="cell-sub">${esc(x.supplier || "—")} · ${esc(x.stage)}. The line leaves the register and the dashboard. Its full record stays in Change history, where it can be restored.</div>
      <label class="fld wide">Reason (required)<input id="ad-delreason" placeholder="e.g. Raised in error — duplicate of PR-0045"></label>
      <div class="form-actions"><button class="btn danger" type="submit">Delete line</button><button class="btn ghost" type="button" id="ad-delcancel">Cancel</button><span class="form-msg" role="status"></span></div></form>`;
  }
  function renderEdit(body) {
    const x = adminState.edit, p = P();
    const cur = p.all.find((l) => l._row === x.row && l.pr_no === x.pr_no && String(l.item || "") === String(x.item || ""));
    if (!cur) { adminState.edit = null; P().toast("That line changed or was removed — showing the list."); renderAdmin(); return; }
    x.orig = rawOf(cur);
    const raw = p.data.lines.find((l) => l._row === x.row) || {};
    const lg = lineLogs(p.data, raw), locked = window.ownedFields(lg);
    const vals = Object.assign({}, x.draft || x.orig); locked.forEach((k) => { vals[k] = x.orig[k]; });
    const nLogs = Object.keys(lg).reduce((a, k) => a + lg[k].length, 0);
    body.innerHTML = `<div class="entry-grid"><form class="card" id="ed-form" novalidate>
      <div class="card-h"><h2>Edit row ${x.row}: ${esc(x.pr_no)}</h2><button class="linkbtn" type="button" id="ed-back">← Back to the list</button></div>
      ${locked.length ? `<p class="note" style="margin-top:-6px">Fields marked <i class="lock">from log</i> are kept up to date by this line's delivery and payment logs — change them by adding or deleting log entries below.</p>` : ""}
      ${formHTML(vals, "ed", locked)}
      <div id="ed-diff"></div>
      <label class="fld wide">Reason for the change (required)<input id="ed-reason" value="${esc(x.reason || "")}" placeholder="e.g. Supplier confirmed new delivery date"></label>
      <div class="form-actions"><button class="btn primary" type="submit">Save changes</button><button class="btn ghost" type="button" id="ed-cancel">Cancel</button><span class="form-msg" role="status"></span></div>
      </form><aside class="card sticky"><div class="card-h"><h2>Live check</h2><span class="sub">After your changes</span></div><div id="ed-preview"></div></aside></div>
      <div class="card" id="ed-logs"><div class="card-h"><h2>Logs for this line</h2><span class="sub">${nLogs} entr${nLogs === 1 ? "y" : "ies"} · deleting an entry updates the register and can be undone from Change history</span></div>
        ${adminState.delLog ? `<form class="danger-box" id="ed-dellog" novalidate><b>Delete ${esc(adminState.delLog.id)}</b>
          <label class="fld wide">Reason (required)<input id="ed-dellog-r" placeholder="e.g. Entered twice"></label>
          <div class="form-actions"><button class="btn danger" type="submit">Delete entry</button><button class="btn ghost" type="button" id="ed-dellog-c">Cancel</button><span class="form-msg" role="status"></span></div></form>` : ""}
        ${Object.keys(window.LOG_KINDS).map((k) => lg[k].length ? `<h3 class="subh">${esc(window.LOG_KINDS[k].sheet)}</h3>` + logTableHTML(k, lg[k], true) : "").join("") || `<div class="empty">No deliveries, payments, quotations or invoices recorded for this line.</div>`}
      </div>`;
    const form = $("#ed-form");
    $$("[data-dellog]", body).forEach((b) => b.addEventListener("click", () => { const [kind, id] = b.dataset.dellog.split("|"); adminState.delLog = { kind, id }; renderEdit(body); $("#ed-dellog-r").focus(); }));
    const dl = $("#ed-dellog");
    if (dl) {
      $("#ed-dellog-c").addEventListener("click", () => { adminState.delLog = null; renderEdit(body); });
      dl.addEventListener("submit", async (e) => {
        e.preventDefault();
        const reason = $("#ed-dellog-r").value.trim(), msg = $(".form-msg", dl);
        if (!reason) { msg.textContent = "Please give a reason."; return; }
        msg.textContent = "Deleting…";
        const r = await call("deleteLog", { kind: adminState.delLog.kind, id: adminState.delLog.id, reason });
        if (!r.ok) { msg.textContent = r.error || (r.errors || []).join(" · "); return; }
        P().toast(`Deleted ${adminState.delLog.id}. You can restore it from Change history.`);
        adminState.delLog = null; x.draft = null; await loadHistory(); renderEdit(body);
      });
    }
    const refresh = () => {
      const ev = evaluate(readForm(form), x.row);
      x.draft = readFormClean(form);
      locked.forEach((k) => { x.draft[k] = x.orig[k]; });
      $("#ed-preview").innerHTML = previewHTML(ev);
      const before = window.normalizeEntry(x.orig, {}).line;
      const ch = window.diffEntry(before, ev.line);
      const fmt = (k, v) => v === null || v === undefined ? "(blank)" : (FIELDS.find((f) => f.k === k).type === "percent" ? p.pct(v, 2) : String(v));
      $("#ed-diff").innerHTML = Object.keys(ch).length ? `<div class="diff"><b>${Object.keys(ch).length} change(s):</b><table><tbody>${Object.keys(ch).map((k) =>
        `<tr><td>${esc(LABEL[k])}</td><td class="old">${esc(fmt(k, ch[k][0]))}</td><td>→</td><td class="new">${esc(fmt(k, ch[k][1]))}</td></tr>`).join("")}</tbody></table></div>`
        : `<p class="note">No changes yet.</p>`;
      return { ev, ch };
    };
    form.addEventListener("input", refresh); form.addEventListener("change", refresh); refresh();
    const back = () => { adminState.edit = null; renderAdmin(); };
    $("#ed-back").addEventListener("click", back); $("#ed-cancel").addEventListener("click", back);
    $("#ed-reason").addEventListener("input", (e) => { x.reason = e.target.value; });
    form.addEventListener("submit", async (e) => {
      e.preventDefault();
      const { ev, ch } = refresh(), msg = $(".form-msg", form), reason = $("#ed-reason").value.trim();
      if (!Object.keys(ch).length) { msg.textContent = "Nothing has changed."; return; }
      if (ev.errors.length) { msg.textContent = "Please fix the items listed under Live check."; return; }
      if (!reason) { msg.textContent = "Please give a reason for the change."; $("#ed-reason").focus(); return; }
      msg.textContent = "Saving…";
      const r = await call("update", { row: x.row, pr_no: x.pr_no, item: x.item, line: readForm(form), reason });
      if (!r.ok) { msg.textContent = r.error || (r.errors || []).join(" · "); return; }
      P().toast(r.unchanged ? "Nothing changed." : `Row ${x.row} updated (${Object.keys(r.changes).length} field(s)).`);
      adminState.edit = null; await loadHistory(); renderAdmin();
    });
  }
  function readFormClean(form) { return window.normalizeEntry(readForm(form), {}).line; }

  async function loadHistory() {
    if (session.role !== "admin") return;
    const r = await call("history", { pr_no: adminState.histPR });
    adminState.hist = r.ok ? r.entries : [];
    const h = $("#ad-hist"); if (h) { h.innerHTML = historyHTML(); bindHistory(); }
  }
  function historyHTML() {
    const H = adminState.hist, p = P();
    if (!H.length) return `<div class="empty">No changes recorded${adminState.histPR ? " for " + esc(adminState.histPR) : " yet"}.</div>`;
    const can = (e) => /^(Updated|Reverted|Deleted)/.test(e.action);
    const ast = (a) => a.startsWith("Deleted") ? "critical" : a.startsWith("Updated") ? "warning" : a.startsWith("Revert") || a.startsWith("Restor") ? "info" : "good";
    return `<div class="hist">${H.slice(0, 120).map((e) => `<div class="h-item">
      <div class="h-top">${p.pill(ast(e.action), e.action)} <span class="mono">${esc(e.pr_no || "")}</span> <span class="cell-sub">#${e.id} · ${esc(e.time)} · ${esc(e.by)}</span></div>
      <div class="h-what">${esc(e.item || "")}${e.row ? ` <span class="cell-sub">row ${e.row}</span>` : ""}</div>
      ${e.changes ? `<div class="h-ch">${Object.keys(e.changes).map((k) => `<span><b>${esc(LABEL[k] || k)}</b>: ${esc(e.changes[k][0] === null ? "(blank)" : e.changes[k][0])} → ${esc(e.changes[k][1] === null ? "(blank)" : e.changes[k][1])}</span>`).join("")}</div>` : ""}
      ${e.reason ? `<div class="cell-sub">Reason: ${esc(e.reason)}</div>` : ""}
      ${can(e) ? (adminState.undo === e.id ? `<form class="undo-form" data-undo-form="${e.id}" novalidate><input placeholder="Reason for undo (required)" aria-label="Reason for undo"><button class="btn sm primary" type="submit">${e.action.startsWith("Deleted") ? "Restore line" : "Undo change"}</button><button class="btn sm ghost" type="button" data-undo-cancel>Cancel</button><span class="form-msg" role="status"></span></form>`
        : `<button class="btn sm" type="button" data-undo="${e.id}">${e.action.startsWith("Deleted") ? "Restore" : "Undo"}</button>`) : ""}
      </div>`).join("")}</div>`;
  }
  function bindHistory() {
    $$("[data-undo]").forEach((b) => b.addEventListener("click", () => { adminState.undo = Number(b.dataset.undo); $("#ad-hist").innerHTML = historyHTML(); bindHistory(); }));
    $$("[data-undo-cancel]").forEach((b) => b.addEventListener("click", () => { adminState.undo = null; $("#ad-hist").innerHTML = historyHTML(); bindHistory(); }));
    $$("[data-undo-form]").forEach((f) => f.addEventListener("submit", async (e) => {
      e.preventDefault();
      const reason = $("input", f).value.trim(), msg = $(".form-msg", f);
      if (!reason) { msg.textContent = "Please give a reason."; return; }
      msg.textContent = "Working…";
      const r = await call("revert", { logId: Number(f.dataset.undoForm), reason });
      if (!r.ok) { msg.textContent = r.error || (r.errors || []).join(" · "); return; }
      P().toast(r.restored ? `Line restored to row ${r.row}.` : `Change undone on row ${r.row}.`);
      adminState.undo = null; await loadHistory(); renderAdmin();
    }));
  }

  /* ---------------------------------------------------------------- wire-up: re-render after every data load */
  function hook() {
    if (!window.PPCR || !window.PPCR.data) { setTimeout(hook, 150); return; }
    renderEntry(); renderAdmin(); if (session.role === "admin") loadHistory();
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", hook); else hook();
  window.PPCR_ENTRY = { renderEntry, renderAdmin, call };
})();
