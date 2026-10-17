/* Engine.gs — exact copy of assets/js/engine.js (written by scripts/sync_rules.py; do not edit here).
 * Lets the email alerts calculate exactly like the web dashboard. */
/*!
 * Panda Purchase Analytics — calculation engine
 * --------------------------------------------------------------
 * Reproduces every formula of the Excel "Purchase_Register" (row level)
 * and the analysis sheets (Supplier_Scorecard, Spend_Analysis,
 * Budget_Control, Payables_Aging, Cash_Forecast) so the web dashboard
 * and the workbook always give the same numbers.
 *
 * Works in the browser (window.PurchaseEngine) and in Node (require) so
 * the test-suite can compare it line-by-line with the Excel results.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.PurchaseEngine = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  /* ---------------------------------------------------------------- helpers */
  const DAY = 86400000;
  function toDay(v) {                       // -> integer day number (UTC) or null
    if (v === null || v === undefined || v === "") return null;
    if (v instanceof Date) return Math.round(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()) / DAY);
    if (typeof v === "number") return Math.round(v);
    const m = String(v).trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return null;
    return Math.round(Date.UTC(+m[1], +m[2] - 1, +m[3]) / DAY);
  }
  function dayToISO(d) { return d === null ? null : new Date(d * DAY).toISOString().slice(0, 10); }
  function dayToDate(d) { return new Date(d * DAY); }
  const isNum = (v) => typeof v === "number" && isFinite(v);
  const N = (v) => (isNum(v) ? v : 0);                       // Excel N()
  const blank = (v) => v === null || v === undefined || v === "";
  const txt = (v) => (blank(v) ? "" : String(v).trim());
  const num = (v) => {
    if (blank(v)) return null;
    if (isNum(v)) return v;
    const x = Number(String(v).replace(/,/g, ""));
    return isFinite(x) ? x : null;
  };
  // Excel ROUND (half away from zero)
  const xround = (x, d = 0) => { const f = Math.pow(10, d); return Math.sign(x) * Math.round(Math.abs(x) * f + 1e-9) / f; };
  // Excel TEXT(x,"#,##0")
  // Excel TEXT(x,"#,##0") — plain code, no Intl, so it runs the same in browsers, Node and Apps Script
  const fmt0 = (x) => { const v = xround(x, 0); return (v < 0 ? "-" : "") + String(Math.abs(v)).replace(/\B(?=(\d{3})+(?!\d))/g, ","); };
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  // Excel TEXT(x,"0%")
  const pct0 = (x) => xround(x * 100, 0) + "%";
  const mean = (a) => { const v = a.filter(isNum); return v.length ? v.reduce((s, x) => s + x, 0) / v.length : null; };
  const sum = (a, f) => a.reduce((s, x) => s + N(f ? f(x) : x), 0);
  const lc = (s) => txt(s).toLowerCase();

  const INPUT_KEYS = ["pr_no", "pr_date", "dept", "req_by", "item", "category", "spec", "spend", "appr_ref", "appr_date",
    "appr_status", "appr_by", "supplier", "origin", "po_no", "po_date", "qty", "uom", "unit_price", "ccy", "budget_price",
    "tax_pct", "pay_terms", "adv_paid", "interim_paid", "closing_paid", "last_pay_date", "due_date", "commit_date",
    "inhouse_start", "qty_recv", "close_date", "qc_status", "photo_link", "remarks"];
  const DATE_KEYS = ["pr_date", "appr_date", "po_date", "last_pay_date", "due_date", "commit_date", "inhouse_start", "close_date"];
  const NUM_KEYS = ["qty", "unit_price", "budget_price", "tax_pct", "adv_paid", "interim_paid", "closing_paid", "qty_recv"];

  const DEFAULT_SETTINGS = { otif_tol: 0.98, pv_alert: 0.05, watch: 0.85, fy_start: null, sla_appr: 3, sla_po: 3,
    abc_a: 0.8, abc_b: 0.95, min_quotes: 3, quote_threshold: 100000, inv_tol: 0.01, fx: { BDT: 1, USD: 122, CNY: 17, EUR: 133, INR: 1.45 } };

  /* ------------------------------------------------- row-level (Purchase_Register) */
  function normaliseLine(raw) {
    const o = {};
    INPUT_KEYS.forEach((k) => {
      const v = raw[k];
      if (DATE_KEYS.includes(k)) o[k] = toDay(v);
      else if (NUM_KEYS.includes(k)) o[k] = num(v);
      else o[k] = blank(v) ? null : txt(v);
    });
    o._row = raw._row || null;
    return o;
  }

  function fxLookup(fx, ccy) {               // Excel MATCH is case-insensitive; unknown currency -> 1
    const k = Object.keys(fx || {}).find((x) => x.toLowerCase() === String(ccy).toLowerCase());
    return k !== undefined && isNum(fx[k]) ? fx[k] : 1;
  }

  function enrich(data, reportDay) {
    const S = Object.assign({}, DEFAULT_SETTINGS, data.settings || {});
    const R = reportDay;
    const lines = (data.lines || []).map(normaliseLine).filter((l) => !blank(l.pr_no));

    // pass 1 — everything that only needs its own row
    lines.forEach((l) => {
      const fxv = blank(l.ccy) && blank(l.unit_price) ? null : blank(l.ccy) ? 1 : fxLookup(S.fx, l.ccy);
      l.fx = fxv;
      l.po_value = l.qty === null || l.unit_price === null ? null : l.qty * l.unit_price * fxv;
      l.pv_bdt = l.po_value === null || l.budget_price === null ? null : l.po_value - l.qty * l.budget_price;
      l.pv_pct = l.pv_bdt === null ? null : (N(l.budget_price) * N(l.qty) === 0 ? null : l.pv_bdt / (l.qty * l.budget_price));
      l.total_cost = l.po_value === null ? null : l.po_value * (1 + N(l.tax_pct));
      l.total_paid = N(l.adv_paid) + N(l.interim_paid) + N(l.closing_paid);
      l.pct_paid = l.total_cost === null ? null : (N(l.total_cost) === 0 ? null : l.total_paid / l.total_cost);
      l.outstanding = l.total_cost === null ? null : (l.total_cost - l.total_paid < 1 ? 0 : xround(l.total_cost - l.total_paid, 2));
      l.fin_status = blank(l.po_no) ? "No PO Yet" : l.total_paid === 0 ? "Not Paid" : N(l.pct_paid) >= 0.999 ? "Fully Paid"
        : N(l.interim_paid) + N(l.closing_paid) > 0 ? "Partially Paid" : "Initial / Advance Paid";
      l.days_overdue = l.outstanding === null || N(l.outstanding) === 0 || l.due_date === null ? null : Math.max(0, R - l.due_date);
      l.aging = l.outstanding === null || N(l.outstanding) <= 0 ? null : l.due_date === null ? "No Due Date"
        : R <= l.due_date ? "Not Yet Due" : R - l.due_date <= 30 ? "1-30 Days" : R - l.due_date <= 60 ? "31-60 Days"
        : R - l.due_date <= 90 ? "61-90 Days" : "90+ Days";
      l.recv_pct = l.qty === null ? null : (l.qty === 0 ? null : N(l.qty_recv) / l.qty);
      l.pend_qty = blank(l.po_no) || l.qty === null ? null : Math.max(0, l.qty - N(l.qty_recv));
      l.pend_val = l.pend_qty === null || l.po_value === null ? null : (N(l.qty) === 0 ? 0 : l.pend_qty / l.qty * l.total_cost);
      const u = l.uom || "";
      l.pend_status = blank(l.po_no) ? "No PO Yet" : l.qty === null ? "Order Qty Missing"
        : N(l.qty_recv) >= l.qty ? (N(l.qty_recv) > l.qty ? `All In-house (+${fmt0(N(l.qty_recv) - l.qty)} ${u} excess)` : "All In-house")
        : l.close_date !== null ? `Short Closed - ${fmt0(l.pend_qty)} ${u} not received`
        : N(l.qty_recv) === 0 ? `Not Started - ${fmt0(l.pend_qty)} ${u} pending`
        : `Pending ${fmt0(l.pend_qty)} ${u} (${pct0(l.recv_pct)} received)`;
      l.d_pr_appr = l.pr_date === null || l.appr_date === null ? null : l.appr_date - l.pr_date;
      l.d_appr_po = l.appr_date === null || l.po_date === null ? null : l.po_date - l.appr_date;
      l.d_po_close = l.po_date === null || l.close_date === null ? null : l.close_date - l.po_date;
      l.d_cycle = l.pr_date === null || l.close_date === null ? null : l.close_date - l.pr_date;
      l.delay = l.commit_date === null || blank(l.po_no) ? null : Math.max(0, (l.close_date === null ? R : l.close_date) - l.commit_date);
      l.otif = l.close_date === null || l.commit_date === null ? null
        : (l.close_date <= l.commit_date && N(l.recv_pct) >= S.otif_tol ? "OTIF" : "Not OTIF");
      l.stage = l.appr_status === "Rejected" ? "Rejected" : l.appr_status !== "Approved" ? "Awaiting Approval"
        : blank(l.po_no) ? "Approved - PO Pending"
        : l.close_date !== null ? (N(l.pct_paid) >= 0.999 ? "Closed" : "In-house Closed - Payment Due")
        : l.inhouse_start !== null ? "Receiving In-house" : "PO Issued - Awaiting Delivery";
      l.unit_bdt = l.po_value === null || N(l.qty) === 0 ? null : l.po_value / l.qty;
      l.paid_ahead = blank(l.po_no) || l.total_cost === null ? null : Math.max(0, xround(l.total_paid - (l.total_cost - N(l.pend_val)), 0));
    });

    // pass 2 — previous PO price for the same item (needs all rows)
    lines.forEach((l) => {
      l.prev_price = null; l.price_chg = null;
      if (l.unit_bdt === null || l.po_date === null || blank(l.item)) return;
      const key = lc(l.item);
      const earlier = lines.filter((o) => lc(o.item) === key && o.po_date !== null && o.po_date < l.po_date && N(o.unit_bdt) > 0);
      if (!earlier.length) return;
      const md = Math.max(...earlier.map((o) => o.po_date));
      l.prev_price = mean(earlier.filter((o) => o.po_date === md).map((o) => o.unit_bdt));
      l.price_chg = l.prev_price ? l.unit_bdt / l.prev_price - 1 : null;
    });

    // pass 3 — data check, actions, priority
    lines.forEach((l) => {
      const c = [];
      if (l.pr_date === null || blank(l.dept) || blank(l.item) || blank(l.category) || blank(l.spend)) c.push("PR info missing; ");
      if (l.appr_status === "Approved" && (l.appr_date === null || blank(l.appr_ref))) c.push("Approval ref/date missing; ");
      if (!blank(l.po_no) && l.appr_status !== "Approved") c.push("PO without approval; ");
      if (!blank(l.po_no) && (l.po_date === null || blank(l.supplier) || l.qty === null || l.unit_price === null || blank(l.ccy))) c.push("PO details missing; ");
      if (!blank(l.po_no) && l.commit_date === null) c.push("Committed date missing; ");
      if (l.appr_date !== null && l.pr_date !== null && l.appr_date < l.pr_date) c.push("Approval dated before PR; ");
      if (l.po_date !== null && l.appr_date !== null && l.po_date < l.appr_date) c.push("PO dated before approval; ");
      if (l.close_date !== null && l.inhouse_start !== null && l.close_date < l.inhouse_start) c.push("Close before in-house start; ");
      if (l.close_date !== null && l.inhouse_start === null) c.push("In-house start date missing; ");
      if (l.close_date !== null && l.qty_recv === null) c.push("Qty received missing; ");
      if (l.close_date !== null && blank(l.qc_status)) c.push("QC status missing; ");
      if (N(l.qty_recv) > N(l.qty) * 1.05) c.push("Received > ordered +5%; ");
      if (l.total_paid > 0 && blank(l.pay_terms)) c.push("Payment terms missing; ");
      if (l.total_paid > N(l.total_cost) + 1) c.push("Overpaid vs PO; ");
      if (N(l.outstanding) > 0 && l.close_date !== null && l.due_date === null) c.push("Payment due date missing; ");
      l.data_check = c.length ? c.join("") : "OK";

      const a = [];
      const od = N(l.days_overdue), dl = N(l.delay), pa = N(l.paid_ahead), open = !blank(l.po_no) && l.close_date === null;
      const waitAppr = l.stage === "Awaiting Approval" && l.pr_date !== null ? R - l.pr_date : null;
      const waitPO = l.stage === "Approved - PO Pending" && l.appr_date !== null ? R - l.appr_date : null;
      const short = (l.pend_status || "").startsWith("Short Closed");
      if (od > 0) a.push(`Pay supplier: ${od} days overdue; `);
      if (open && dl > 0) a.push(`Expedite delivery: ${dl} days late; `);
      if (waitAppr !== null && waitAppr > S.sla_appr) a.push(`Get approval: PR waiting ${waitAppr} days; `);
      if (waitPO !== null && waitPO > S.sla_po) a.push(`Issue PO: approved ${waitPO} days ago; `);
      if ((l.stage === "Awaiting Approval" || l.stage === "Approved - PO Pending") && N(l.pv_pct) > S.pv_alert) a.push(`Negotiate price: ${pct0(l.pv_pct)} above budget; `);
      if (short) a.push("Short closed: claim balance / credit note" + (pa > 0 ? ` BDT ${fmt0(pa)}` : "") + "; ");
      if (l.qc_status === "Rejected") a.push("QC rejected: return / replace; ");
      if (open && pa > 0 && dl > 0) a.push(`Advance at risk: BDT ${fmt0(pa)} paid, goods late; `);
      if (l.data_check !== "OK") a.push("Fix data entry (see Data Check); ");
      l.action = a.length ? a.join("") : null;
      l.priority = null;
      if (l.action) {
        const high = od > 30 || (l.close_date === null && dl > 15) || l.qc_status === "Rejected" || (short && pa > 0) || (open && pa > 0 && dl > 0);
        const med = od > 0 || (l.close_date === null && dl > 7) || short || (waitAppr !== null && waitAppr > 2 * S.sla_appr) || (waitPO !== null && waitPO > 2 * S.sla_po);
        l.priority = high ? "High" : med ? "Medium" : "Low";
      }
      l.wait_days = waitAppr !== null ? waitAppr : waitPO !== null ? waitPO : null;
      l.is_open = open;
      l.days_to_commit = open && l.commit_date !== null ? l.commit_date - R : null;
    });
    return lines;
  }

  /* ------------------------------------------------- controls from the four logs */
  const PRANK = { High: 1, Medium: 2, Low: 3 };
  const worst = (a, b) => (!a ? b : !b ? a : PRANK[a] <= PRANK[b] ? a : b);
  const lkey = (p, i) => txt(p).toLowerCase() + "|" + txt(i).toLowerCase();

  /* Adds to every line: logs, deliveries punctuality, quotation status, invoice match status,
     ctrl_action / ctrl_priority, and all_action / all_priority (register action + control action). */
  function attachControls(lines, data, R, S) {
    const has = { deliveries: Array.isArray(data.deliveries), payments: Array.isArray(data.payments), quotations: Array.isArray(data.quotations), invoices: Array.isArray(data.invoices) };
    const groups = {};
    ["deliveries", "payments", "quotations", "invoices"].forEach((kind) => (data[kind] || []).forEach((e) => {
      const g = (groups[lkey(e.pr_no, e.item)] = groups[lkey(e.pr_no, e.item)] || { deliveries: [], payments: [], quotations: [], invoices: [] });
      g[kind].push(Object.assign({}, e, { day: toDay(e.date) }));
    }));
    lines.forEach((l) => {
      const g = groups[lkey(l.pr_no, l.item)] || { deliveries: [], payments: [], quotations: [], invoices: [] };
      l.logs = g;
      const ca = [];
      let cp = null;
      const push = (text, pr) => { ca.push(text + "; "); cp = worst(cp, pr); };
      // deliveries
      l.n_deliveries = g.deliveries.length;
      l.receipts_on_time = l.commit_date === null ? null : g.deliveries.filter((d) => d.day !== null && d.day <= l.commit_date).length;
      // quotations
      l.q_status = null; l.q_issue = false; l.q_count = 0; l.q_l1 = null; l.q_l1_supplier = null; l.q_sel = null; l.q_saving = null;
      if (has.quotations) {
        const best = {};
        g.quotations.forEach((q) => {
          const b = N(num(q.unit_price)) * fxLookup(S.fx, q.ccy || l.ccy || "BDT");
          if (b > 0 && (!best[q.supplier] || b < best[q.supplier])) best[q.supplier] = b;
        });
        const sup = Object.keys(best);
        l.q_count = sup.length;
        if (sup.length) { const s1 = sup.reduce((a, b) => (best[b] < best[a] ? b : a)); l.q_l1 = best[s1]; l.q_l1_supplier = s1; }
        const selKey = sup.find((x) => x.toLowerCase() === txt(l.supplier).toLowerCase());
        l.q_sel = selKey ? best[selKey] : null;
        const value = N(l.total_cost);
        const need = value >= S.quote_threshold ? S.min_quotes : 1;
        const prePO = blank(l.po_no);
        if (l.stage === "Rejected") l.q_status = null;
        else if (prePO && !l.q_sel) l.q_status = l.q_count >= need ? "Quotes complete" : value >= S.quote_threshold ? `Quotes needed (${l.q_count} of ${need})` : "Below threshold";
        else if (!l.q_count) l.q_status = value >= S.quote_threshold ? "No quotation" : "Below threshold";
        else if (l.q_sel === null) l.q_status = "Supplier not quoted";
        else if (l.q_count < need) l.q_status = `Insufficient quotes (${l.q_count} of ${need})`;
        else if (l.q_sel > l.q_l1 * (1 + 1e-6)) l.q_status = "Not lowest quote";
        else l.q_status = value >= S.quote_threshold ? "Compliant" : "Below threshold";
        l.q_issue = /^(No quotation|Supplier not|Insufficient|Not lowest)/.test(l.q_status || "");
        if (l.q_count >= 2 && l.q_sel !== null && l.qty) {
          const avg = sup.reduce((s, x) => s + best[x], 0) / sup.length;
          l.q_saving = (avg - l.q_sel) * l.qty;
        }
        if (!prePO && l.q_status === "Not lowest quote") push(`Justify supplier: ${pct0(l.q_sel / l.q_l1 - 1)} above the lowest quote (${l.q_l1_supplier})`, "Medium");
        else if (!prePO && /^(No quotation|Insufficient|Supplier not)/.test(l.q_status)) push(`Sourcing: ${l.q_status.toLowerCase()} for a PO of BDT ${fmt0(value)}`, "Medium");
        else if (prePO && /^Quotes needed/.test(l.q_status || "") && l.stage !== "Rejected") push(`Collect quotations: ${l.q_count} of ${need} received`, "Low");
      }
      // three-way match: PO · goods received (accepted) · supplier invoice
      l.match_status = null; l.match_issues = []; l.inv_qty = 0; l.inv_amt = 0; l.accepted = null;
      if (has.invoices && !blank(l.po_no)) {
        l.inv_qty = g.invoices.reduce((s, i) => s + N(num(i.qty)), 0);
        l.inv_amt = g.invoices.reduce((s, i) => s + N(num(i.amount)), 0);
        l.accepted = g.deliveries.length ? g.deliveries.reduce((s, d) => s + (num(d.qty_ok) === null ? N(num(d.qty)) : N(num(d.qty_ok))), 0) : N(l.qty_recv);
        const iss = l.match_issues;
        // goods still to come: an invoice may arrive first (e.g. LC imports) — check it against the order;
        // once the line is closed, the invoice must not exceed what QC accepted
        if (g.invoices.length && l.close_date === null && l.inv_qty > N(l.qty) + 1e-9) iss.push(`Invoiced ${fmt0(l.inv_qty - N(l.qty))} ${l.uom || ""} more than ordered`.replace("  ", " "));
        if (g.invoices.length && l.close_date !== null && l.inv_qty > l.accepted + 1e-9) iss.push(`Invoiced ${fmt0(l.inv_qty - l.accepted)} ${l.uom || ""} more than accepted`.replace("  ", " "));
        const bad = g.invoices.filter((i) => l.unit_price && Math.abs(N(num(i.unit_price)) - l.unit_price) / l.unit_price > S.inv_tol);
        if (bad.length) iss.push(`Invoice price differs from PO (${bad.map((i) => i.invoice_no).join(", ")})`);
        if (g.invoices.length && l.inv_amt > N(l.total_cost) + 1) iss.push("Invoiced more than the PO value");
        const paidAfter = N(l.interim_paid) + N(l.closing_paid);
        if (paidAfter > l.inv_amt + 1) iss.push(`Paid BDT ${fmt0(paidAfter - l.inv_amt)} more than invoiced`);
        l.match_status = iss.length ? "Mismatch" : !g.invoices.length ? (N(l.qty_recv) > 0 ? "Awaiting invoice" : "Not due")
          : l.inv_qty > l.accepted + 1e-9 ? "Invoiced, goods pending" : l.inv_qty + 1e-9 < l.accepted ? "Partly invoiced" : "Matched";
        iss.forEach((t) => push(t.startsWith("Invoice price") ? `Check ${t.toLowerCase()}` : t.startsWith("Paid") ? `Recover or invoice: ${t.toLowerCase()}` : `Hold invoice: ${t.toLowerCase()}`,
          t.startsWith("Invoice price") ? "Medium" : "High"));
        if (l.match_status === "Awaiting invoice" && l.close_date !== null && R - l.close_date > 15) push(`Collect supplier invoice: goods in-house ${R - l.close_date} days`, "Low");
      }
      l.ctrl_action = ca.length ? ca.join("") : null;
      l.ctrl_priority = cp;
      l.all_action = [l.action, l.ctrl_action].filter(Boolean).join("") || null;
      l.all_priority = worst(l.priority, cp);
    });
    return has;
  }

  /* Cash plan: open planned instalments from the payments log; lines without a plan use the
     register's next due date for the whole outstanding amount. */
  function cashPlan(lines, has) {
    const out = [];
    lines.forEach((l) => {
      if (!(N(l.outstanding) > 0)) return;
      const plans = has.payments ? l.logs.payments.filter((p) => p.entry === "Planned") : [];
      if (plans.length) {
        const settled = new Set(l.logs.payments.filter((p) => p.entry === "Paid" && p.settles).map((p) => p.settles));
        plans.filter((p) => !settled.has(p.id) && p.day !== null).forEach((p) => out.push({ day: p.day, amount: N(num(p.amount)), line: l, source: "Planned " + (p.ptype || "") }));
      } else if (l.due_date !== null) out.push({ day: l.due_date, amount: N(l.outstanding), line: l, source: "Next due date" });
    });
    return out.sort((a, b) => a.day - b.day);
  }

  /* ------------------------------------------------------------- aggregations */
  function supplierStats(lines, names) {
    const totalSpend = sum(lines, (l) => l.total_cost);
    const all = Array.from(new Set([...(names || []), ...lines.map((l) => l.supplier).filter(Boolean)]));
    return all.map((s) => {
      const L = lines.filter((l) => l.supplier === s);
      const ot = L.filter((l) => l.otif !== null);
      const ap = L.filter((l) => (l.qc_status || "").startsWith("Approved")).length, rj = L.filter((l) => l.qc_status === "Rejected").length;
      const spend = sum(L, (l) => l.total_cost);
      const otif = ot.length ? ot.filter((l) => l.otif === "OTIF").length / ot.length : null;
      const infull = ot.length ? ot.filter((l) => (l.pend_status || "").startsWith("All In-house")).length / ot.length : null;
      const qc = ap + rj ? ap / (ap + rj) : null;
      const apv = mean(L.map((l) => l.pv_pct));
      const score = otif === null || qc === null ? null
        : xround(40 * otif + 20 * N(infull) + 25 * qc + 15 * (1 - Math.min(1, Math.max(0, N(apv)) / 0.1)), 0);
      return {
        supplier: s, pos: L.filter((l) => !blank(l.po_no)).length, spend, share: totalSpend ? spend / totalSpend : null,
        paid: sum(L, (l) => l.total_paid), outstanding: sum(L, (l) => l.outstanding), closed: ot.length, otif, infull,
        lead: mean(L.map((l) => l.d_po_close)), delay: mean(L.filter((l) => l.close_date !== null).map((l) => l.delay)),
        lateOpen: L.filter((l) => l.is_open && N(l.delay) > 0).length,
        shortClosed: L.filter((l) => (l.pend_status || "").startsWith("Short Closed")).length,
        pending: sum(L.filter((l) => l.close_date === null), (l) => l.pend_val), qc, apv, score,
        rating: score === null ? "Not Rated" : score >= 80 ? "A — Preferred" : score >= 60 ? "B — Approved" : "C — Review",
        paidAhead: sum(L, (l) => l.paid_ahead), overdue: sum(L.filter((l) => N(l.days_overdue) > 0), (l) => l.outstanding),
      };
    }).filter((x) => x.pos > 0 || x.spend > 0);
  }

  function categoryABC(lines, S) {
    const m = new Map();
    lines.forEach((l) => { if (l.category) m.set(l.category, (m.get(l.category) || 0) + N(l.total_cost)); });
    const total = Array.from(m.values()).reduce((a, b) => a + b, 0);
    let cum = 0;
    return Array.from(m.entries()).sort((a, b) => b[1] - a[1]).map(([cat, v]) => {
      const share = total ? v / total : 0; const prev = cum; cum += share;
      const L = lines.filter((l) => l.category === cat);
      return { category: cat, spend: v, share, cum, cls: v === 0 ? "-" : prev < S.abc_a ? "A" : prev < S.abc_b ? "B" : "C",
        lines: L.filter((l) => !blank(l.po_no)).length, apv: mean(L.map((l) => l.pv_pct)),
        savings: -sum(L.filter((l) => N(l.pv_bdt) < 0), (l) => l.pv_bdt), overspend: sum(L.filter((l) => N(l.pv_bdt) > 0), (l) => l.pv_bdt),
        paid: sum(L, (l) => l.total_paid), outstanding: sum(L, (l) => l.outstanding),
        pending: sum(L.filter((l) => l.close_date === null), (l) => l.pend_val) };
    });
  }

  function budgetStats(lines, budgets, S) {
    const names = Array.from(new Set([...(budgets || []).map((b) => b.dept), ...lines.map((l) => l.dept).filter(Boolean)]));
    return names.map((d) => {
      const b = (budgets || []).find((x) => x.dept === d) || {};
      const L = lines.filter((l) => l.dept === d);
      const appr = L.filter((l) => l.appr_status === "Approved");
      const opexC = sum(appr.filter((l) => l.spend === "OPEX"), (l) => l.total_cost);
      const capexC = sum(appr.filter((l) => l.spend === "CAPEX"), (l) => l.total_cost);
      const budget = N(b.opex) + N(b.capex), committed = opexC + capexC;
      const util = budget ? committed / budget : null;
      const status = budget === 0 ? (committed > 0 ? "No Budget - Spending" : "No Budget")
        : committed > budget ? "Over Budget" : util >= S.watch ? "Watch" : "Within Budget";
      return { dept: d, opexBudget: N(b.opex), capexBudget: N(b.capex), opexC, capexC, budget, committed,
        paid: sum(L, (l) => l.total_paid), remaining: budget - committed, util, status };
    }).filter((x) => x.budget > 0 || x.committed > 0);
  }

  const BUCKETS = ["Not Yet Due", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days", "No Due Date"];
  const STAGES = ["Awaiting Approval", "Approved - PO Pending", "PO Issued - Awaiting Delivery", "Receiving In-house",
    "In-house Closed - Payment Due", "Closed", "Rejected"];
  const FIN = ["No PO Yet", "Not Paid", "Initial / Advance Paid", "Partially Paid", "Fully Paid"];

  function monthly(lines, fyStartDay) {
    if (fyStartDay === null) return [];
    const d0 = dayToDate(fyStartDay);
    const out = [];
    for (let i = 0; i < 12; i++) {
      const a = Math.round(Date.UTC(d0.getUTCFullYear(), d0.getUTCMonth() + i, 1) / DAY);
      const b = Math.round(Date.UTC(d0.getUTCFullYear(), d0.getUTCMonth() + i + 1, 0) / DAY);
      const inR = (x) => x !== null && x >= a && x <= b;
      out.push({ start: a, end: b, label: MONTHS[dayToDate(a).getUTCMonth()] + " " + String(dayToDate(a).getUTCFullYear()).slice(2),
        prs: lines.filter((l) => inR(l.pr_date)).length, pos: lines.filter((l) => inR(l.po_date)).length,
        poValue: sum(lines.filter((l) => inR(l.po_date)), (l) => l.total_cost),
        paid: sum(lines.filter((l) => inR(l.last_pay_date)), (l) => l.total_paid),
        due: sum(lines.filter((l) => inR(l.due_date)), (l) => l.outstanding) });
    }
    return out;
  }

  function analyze(lines, data, R, targets) {
    const S = Object.assign({}, DEFAULT_SETTINGS, data.settings || {});
    const T = Object.assign({ otif: 0.9, infull: 0.95, qc: 0.95, priceVar: 0.02, quoteCompliance: 0.9 }, targets || {});
    const has = attachControls(lines, data, R, S);
    const po = lines.filter((l) => !blank(l.po_no));
    const ot = lines.filter((l) => l.otif !== null);
    const committed = sum(lines, (l) => l.total_cost), paid = sum(lines, (l) => l.total_paid);
    const outstanding = sum(lines, (l) => l.outstanding);
    const overdueLines = lines.filter((l) => N(l.days_overdue) > 0);
    const overdue = sum(overdueLines, (l) => l.outstanding);
    const netPV = sum(lines, (l) => l.pv_bdt);
    const budgetBase = sum(lines.filter((l) => l.qty !== null && l.budget_price !== null), (l) => l.qty * l.budget_price);
    const qcA = lines.filter((l) => (l.qc_status || "").startsWith("Approved")).length, qcR = lines.filter((l) => l.qc_status === "Rejected").length;
    const open = lines.filter((l) => l.is_open);
    const budgets = budgetStats(lines, data.budgets, S);
    const sup = supplierStats(lines, (data.lists || {}).supplier);
    const hhi = sup.reduce((s, x) => s + Math.pow(N(x.share), 2), 0) * 10000;
    const apprDone = lines.filter((l) => l.d_pr_appr !== null), poDone = lines.filter((l) => l.d_appr_po !== null);
    const k = {
      prs: lines.length, pos: po.length, committed, paid, outstanding, overdue, pctPaid: committed ? paid / committed : 0,
      opex: sum(lines.filter((l) => l.spend === "OPEX"), (l) => l.total_cost),
      capex: sum(lines.filter((l) => l.spend === "CAPEX"), (l) => l.total_cost),
      budget: sum(budgets, (b) => b.budget), budgetCommitted: sum(budgets, (b) => b.committed),
      netPV, netPVpct: budgetBase ? netPV / budgetBase : 0,
      savings: -sum(lines.filter((l) => N(l.pv_bdt) < 0), (l) => l.pv_bdt), overspend: sum(lines.filter((l) => N(l.pv_bdt) > 0), (l) => l.pv_bdt),
      otif: ot.length ? ot.filter((l) => l.otif === "OTIF").length / ot.length : null,
      infull: ot.length ? ot.filter((l) => (l.pend_status || "").startsWith("All In-house")).length / ot.length : null,
      qc: qcA + qcR ? qcA / (qcA + qcR) : null,
      dPrAppr: mean(lines.map((l) => l.d_pr_appr)), dApprPo: mean(lines.map((l) => l.d_appr_po)),
      dPoClose: mean(lines.map((l) => l.d_po_close)), dCycle: mean(lines.map((l) => l.d_cycle)),
      apprInSla: apprDone.length ? apprDone.filter((l) => l.d_pr_appr <= S.sla_appr).length / apprDone.length : null,
      poInSla: poDone.length ? poDone.filter((l) => l.d_appr_po <= S.sla_po).length / poDone.length : null,
      awaiting: lines.filter((l) => l.stage === "Awaiting Approval").length,
      openPOs: open.length, lateOpen: open.filter((l) => N(l.delay) > 0).length,
      dueSoon: open.filter((l) => l.days_to_commit !== null && l.days_to_commit >= 0 && l.days_to_commit <= 3 && N(l.delay) === 0).length,
      maxLate: Math.max(0, ...open.map((l) => N(l.delay))),
      lateAll: lines.filter((l) => N(l.delay) > 0).length,
      breach: lines.filter((l) => !blank(l.po_no) && l.appr_status !== "Approved").length,
      pendingVal: sum(lines.filter((l) => l.close_date === null), (l) => l.pend_val),
      pendingLines: lines.filter((l) => /^(Pending|Not Started)/.test(l.pend_status || "")).length,
      shortClosed: lines.filter((l) => (l.pend_status || "").startsWith("Short Closed")).length,
      paidAhead: sum(lines, (l) => l.paid_ahead),
      atRisk: sum(open.filter((l) => N(l.delay) > 0), (l) => l.paid_ahead),
      high: lines.filter((l) => l.all_priority === "High").length, medium: lines.filter((l) => l.all_priority === "Medium").length,
      low: lines.filter((l) => l.all_priority === "Low").length,
      dataIssues: lines.filter((l) => l.data_check !== "OK").length,
      maxOverdue: Math.max(0, ...lines.map((l) => N(l.days_overdue))),
      dearer: lines.filter((l) => N(l.price_chg) > 0).length,
      hhi, topSupplier: sup.slice().sort((a, b) => b.spend - a.spend)[0] || null,
      overBudget: budgets.filter((b) => b.status === "Over Budget").length, watchBudget: budgets.filter((b) => b.status === "Watch").length,
      noBudget: budgets.filter((b) => b.status === "No Budget - Spending").length,
      slaBreaches: lines.filter((l) => l.wait_days !== null && ((l.stage === "Awaiting Approval" && l.wait_days > S.sla_appr) || (l.stage === "Approved - PO Pending" && l.wait_days > S.sla_po))).length,
    };
    k.budgetUtil = k.budget ? k.budgetCommitted / k.budget : null;

    const pipeline = STAGES.map((s) => {
      const L = lines.filter((l) => l.stage === s);
      const age = L.map((l) => s === "Awaiting Approval" ? (l.pr_date !== null ? R - l.pr_date : null)
        : s === "Approved - PO Pending" ? (l.appr_date !== null ? R - l.appr_date : null)
        : s === "PO Issued - Awaiting Delivery" ? (l.po_date !== null ? R - l.po_date : null)
        : s === "Receiving In-house" ? (l.inhouse_start !== null ? R - l.inhouse_start : null)
        : s === "In-house Closed - Payment Due" ? (l.close_date !== null ? R - l.close_date : null) : null);
      return { stage: s, count: L.length, value: sum(L, (l) => l.total_cost), avgAge: mean(age) };
    });

    /* ---- risk signals (status: good / warning / serious / critical) */
    const band = (v, good, warn, ser, higherBetter = true) => {
      if (v === null || v === undefined) return "none";
      if (higherBetter) return v >= good ? "good" : v >= warn ? "warning" : v >= ser ? "serious" : "critical";
      return v <= good ? "good" : v <= warn ? "warning" : v <= ser ? "serious" : "critical";
    };
    const P = (x) => (x === null ? "—" : Math.round(x * 100) + "%");
    const L_ = (x) => "BDT " + (x / 1e5).toFixed(2) + " L";
    const signals = [
      { area: "On-time delivery", status: band(k.otif, T.otif, 0.75, 0.6), metric: P(k.otif) + " OTIF",
        note: `Target ${P(T.otif)} · ${k.lateAll} late deliveries recorded` },
      { area: "Open PO lead time", status: k.maxLate > 15 ? "critical" : k.lateOpen > 0 ? "serious" : k.dueSoon > 0 ? "warning" : "good",
        metric: `${k.lateOpen} late / ${k.openPOs} open`, note: k.lateOpen ? `Worst ${k.maxLate} days past committed date` : `${k.dueSoon} due within 3 days` },
      { area: "In-full delivery", status: band(k.infull, T.infull, 0.85, 0.7), metric: P(k.infull) + " in-full",
        note: `${k.shortClosed} short-closed line(s) to claim` },
      { area: "Supplier payments", status: k.maxOverdue > 60 ? "critical" : k.maxOverdue > 30 ? "serious" : k.overdue > 0 ? "warning" : "good",
        metric: L_(k.overdue) + " overdue", note: k.overdue ? `Oldest ${k.maxOverdue} days past due` : "Nothing past due" },
      { area: "Budget control", status: k.overBudget ? "critical" : k.noBudget ? "serious" : k.watchBudget ? "warning" : "good",
        metric: k.overBudget ? `${k.overBudget} dept over budget` : P(k.budgetUtil) + " used",
        note: `${P(k.budgetUtil)} of total budget used · ${k.watchBudget} on watch · ${k.noBudget} spending without budget` },
      { area: "Material quality", status: band(k.qc, T.qc, 0.9, 0.8), metric: P(k.qc) + " accepted",
        note: `${qcR} rejection(s) of ${qcA + qcR} inspected` },
      { area: "Approval & PO speed", status: k.slaBreaches >= 3 ? "serious" : k.slaBreaches > 0 ? "warning" : "good",
        metric: `${k.slaBreaches} beyond SLA`, note: `Avg ${k.dPrAppr === null ? "—" : k.dPrAppr.toFixed(1)} d to approve · ${k.dApprPo === null ? "—" : k.dApprPo.toFixed(1)} d to PO` },
      { area: "Price discipline", status: band(k.netPVpct, 0, T.priceVar, 0.05, false), metric: (k.netPVpct >= 0 ? "+" : "") + (k.netPVpct * 100).toFixed(1) + "% vs budget",
        note: `Saved ${L_(k.savings)} · overspent ${L_(k.overspend)}` },
      { area: "Cash at risk", status: k.atRisk > 0 ? "critical" : k.paidAhead > 0 ? "warning" : "good", metric: L_(k.paidAhead) + " paid ahead",
        note: k.atRisk > 0 ? `${L_(k.atRisk)} paid on late POs` : "Advances are on POs still on schedule" },
      { area: "Supplier concentration", status: hhi > 2500 ? "serious" : hhi > 1500 ? "warning" : "good", metric: "HHI " + fmt0(hhi),
        note: k.topSupplier ? `${k.topSupplier.supplier}: ${P(k.topSupplier.share)} of spend` : "—" },
      { area: "Control & data", status: k.breach ? "critical" : k.dataIssues ? "warning" : "good",
        metric: `${k.breach} breach · ${k.dataIssues} data issue(s)`, note: k.breach ? "PO issued without approval" : "Every PO has an approved OPEX/CAPEX" },
    ];

    /* ---- sourcing & invoice-matching controls */
    const poL = lines.filter((l) => !blank(l.po_no) && l.stage !== "Rejected");
    const qScope = poL.filter((l) => N(l.total_cost) >= S.quote_threshold);
    const qOK = qScope.filter((l) => !l.q_issue);
    const rec = lines.reduce((a, l) => a.concat(l.logs.deliveries.filter((d) => l.commit_date !== null && d.day !== null).map((d) => d.day <= l.commit_date)), []);
    const mStat = ["Matched", "Partly invoiced", "Invoiced, goods pending", "Awaiting invoice", "Mismatch", "Not due"];
    const controls = {
      has,
      quoteScope: qScope.length, quoteCompliant: qOK.length,
      quoteCompliance: qScope.length ? sum(qOK, (l) => l.total_cost) / sum(qScope, (l) => l.total_cost) : null,
      quoteIssues: poL.filter((l) => l.q_issue), notL1: poL.filter((l) => l.q_status === "Not lowest quote"),
      notL1Extra: sum(poL.filter((l) => l.q_status === "Not lowest quote"), (l) => (l.q_sel - l.q_l1) * N(l.qty)),
      quoteSaving: sum(lines, (l) => l.q_saving), quotesNeeded: lines.filter((l) => /^Quotes needed/.test(l.q_status || "")),
      quotesPerPO: poL.length ? sum(poL, (l) => l.q_count) / poL.length : null,
      match: mStat.map((st) => ({ status: st, lines: poL.filter((l) => l.match_status === st).length, value: sum(poL.filter((l) => l.match_status === st), (l) => l.total_cost) })),
      mismatches: poL.filter((l) => l.match_status === "Mismatch"),
      paidOverInvoiced: sum(poL, (l) => Math.max(0, N(l.interim_paid) + N(l.closing_paid) - l.inv_amt)),
      receipts: rec.length, receiptsOnTime: rec.length ? rec.filter(Boolean).length / rec.length : null,
      splitPOs: poL.filter((l) => l.n_deliveries >= 2).length,
      avgReceipts: (() => { const c = poL.filter((l) => l.close_date !== null && l.n_deliveries); return c.length ? sum(c, (l) => l.n_deliveries) / c.length : null; })(),
      invoicedValue: sum(poL, (l) => l.inv_amt),
    };
    const plan = cashPlan(lines, has);
    controls.plan = plan;
    const highMis = controls.mismatches.some((l) => l.match_issues.some((t) => !t.startsWith("Invoice price")));
    signals.splice(8, 0,
      { area: "Sourcing compliance", status: !has.quotations || controls.quoteCompliance === null ? "none" : band(controls.quoteCompliance, T.quoteCompliance, 0.75, 0.5),
        metric: controls.quoteCompliance === null ? "No POs above threshold" : P(controls.quoteCompliance) + " of value compliant",
        note: has.quotations ? `${controls.quoteIssues.length} PO line(s) short of quotes or not at lowest quote` : "Quotations log not in use yet" },
      { area: "Invoice matching", status: !has.invoices ? "none" : highMis ? "critical" : controls.mismatches.length ? "warning" : "good",
        metric: has.invoices ? `${controls.mismatches.length} mismatch(es)` : "—",
        note: has.invoices ? `${controls.match.find((m) => m.status === "Awaiting invoice").lines} awaiting invoice · ${L_(controls.paidOverInvoiced)} paid beyond invoices` : "Invoices log not in use yet" });

    /* ---- procurement health score (0–100), transparent components */
    const comp = [
      { name: "On-time delivery (OTIF)", w: 20, s: k.otif === null ? 1 : Math.min(1, k.otif / T.otif) },
      { name: "In-full delivery", w: 10, s: k.infull === null ? 1 : Math.min(1, k.infull / T.infull) },
      { name: "Material quality", w: 10, s: k.qc === null ? 1 : Math.min(1, k.qc / T.qc) },
      { name: "Payment discipline", w: 15, s: k.outstanding ? 1 - k.overdue / k.outstanding : 1 },
      { name: "Budget compliance", w: 10, s: budgets.length ? budgets.filter((b) => b.status !== "Over Budget" && b.status !== "No Budget - Spending").length / budgets.length : 1 },
      { name: "Approval & PO speed", w: 10, s: ((k.apprInSla === null ? 1 : k.apprInSla) + (k.poInSla === null ? 1 : k.poInSla)) / 2 },
      { name: "Price discipline", w: 10, s: 1 - Math.min(1, Math.max(0, k.netPVpct) / 0.1) },
      { name: "Control & data quality", w: 15, s: k.breach ? 0 : lines.length ? 1 - k.dataIssues / lines.length : 1 },
    ];
    const health = Math.round(comp.reduce((s, c) => s + c.w * c.s, 0));

    return {
      kpis: k, pipeline, signals, health, healthParts: comp,
      suppliers: sup, categories: categoryABC(lines, S), budgets,
      aging: BUCKETS.map((b) => ({ bucket: b, value: sum(lines.filter((l) => l.aging === b), (l) => l.outstanding), lines: lines.filter((l) => l.aging === b).length })),
      finance: FIN.map((f) => ({ status: f, lines: lines.filter((l) => l.fin_status === f).length, value: sum(lines.filter((l) => l.fin_status === f), (l) => l.total_cost) })),
      monthly: monthly(lines, toDay(S.fy_start)),
      cashDue: [7, 30, 60, 90].map((d) => ({ days: d, value: plan.filter((x) => x.day > R && x.day <= R + d).reduce((s, x) => s + x.amount, 0) })),
      cashOverdue: plan.filter((x) => x.day <= R).reduce((s, x) => s + x.amount, 0),
      controls,
      payTerms: Array.from(new Set(po.map((l) => l.pay_terms || "Not set"))).map((t) => {
        const L = po.filter((l) => (l.pay_terms || "Not set") === t);
        return { terms: t, lines: L.length, value: sum(L, (l) => l.total_cost), paidAhead: sum(L, (l) => l.paid_ahead) };
      }).sort((a, b) => b.value - a.value),
      origin: ["Local", "Import"].map((o) => {
        const L = po.filter((l) => l.origin === o), O = L.filter((l) => l.otif !== null);
        return { origin: o, lines: L.length, value: sum(L, (l) => l.total_cost), lead: mean(L.map((l) => l.d_po_close)),
          otif: O.length ? O.filter((l) => l.otif === "OTIF").length / O.length : null };
      }),
      open: open.slice().sort((a, b) => N(b.delay) - N(a.delay) || N(a.days_to_commit) - N(b.days_to_commit)),
      waiting: lines.filter((l) => l.stage === "Awaiting Approval" || l.stage === "Approved - PO Pending").sort((a, b) => N(b.wait_days) - N(a.wait_days)),
      actions: lines.filter((l) => l.all_priority).sort((a, b) => PRANK[a.all_priority] - PRANK[b.all_priority] || (a._row || 0) - (b._row || 0)),
      settings: S,
    };
  }

  return { enrich, analyze, attachControls, cashPlan, toDay, dayToISO, dayToDate, INPUT_KEYS, DATE_KEYS, NUM_KEYS, STAGES, BUCKETS, FIN, xround, fmt0, pct0 };
});
