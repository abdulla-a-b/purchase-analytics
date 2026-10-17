/*
 * Panda Purchase Analytics — data-entry rules
 * --------------------------------------------------------------
 * One set of rules used in three places, so a line is judged the same way
 * everywhere:
 *   • the Data Entry and Admin tabs (browser)
 *   • the Google Sheets backend (apps-script/Code.gs holds an exact copy of the
 *     block between the <entry-rules> markers — tests/entry.test.js checks it)
 *   • the test-suite (Node)
 * Plain ES2015 functions, no imports, so it runs in all three.
 */
// <entry-rules>
var ENTRY_FIELDS = [
  { k: "pr_no", label: "PR No", type: "text", req: true, sec: 1, ph: "PSIL-PR-2026-0046" },
  { k: "pr_date", label: "PR Received Date", type: "date", req: true, sec: 1 },
  { k: "dept", label: "Department", type: "list", req: true, sec: 1 },
  { k: "req_by", label: "Requested By", type: "text", sec: 1 },
  { k: "item", label: "Item / Material Purchased", type: "text", req: true, sec: 1, wide: true },
  { k: "category", label: "Category", type: "list", req: true, sec: 1 },
  { k: "spec", label: "Spec / Size / Colour", type: "text", sec: 1 },
  { k: "spend", label: "Spend Type (OPEX / CAPEX)", type: "list", req: true, sec: 1 },
  { k: "appr_ref", label: "OPEX / CAPEX Approval Ref", type: "text", sec: 1 },
  { k: "appr_date", label: "Approval Date", type: "date", sec: 1 },
  { k: "appr_status", label: "Approval Status", type: "list", sec: 1 },
  { k: "appr_by", label: "Approved By", type: "text", sec: 1 },
  { k: "supplier", label: "Supplier", type: "list", sec: 2, wide: true },
  { k: "origin", label: "Local / Import", type: "list", sec: 2 },
  { k: "po_no", label: "PO No", type: "text", sec: 2, ph: "PSIL-PO-2026-0046" },
  { k: "po_date", label: "PO Date", type: "date", sec: 2 },
  { k: "qty", label: "Order Quantity", type: "number", sec: 2 },
  { k: "uom", label: "UOM", type: "list", sec: 2 },
  { k: "unit_price", label: "Unit Price (PO Currency)", type: "number", sec: 2 },
  { k: "ccy", label: "Currency", type: "list", sec: 2 },
  { k: "budget_price", label: "Budget / Standard Unit Price (BDT)", type: "number", sec: 2 },
  { k: "tax_pct", label: "VAT / AIT / Duty %", type: "percent", sec: 2 },
  { k: "pay_terms", label: "Payment Terms", type: "list", sec: 3, wide: true },
  { k: "adv_paid", label: "Initial / Advance Paid (BDT)", type: "number", sec: 3 },
  { k: "interim_paid", label: "Interim Payment (BDT)", type: "number", sec: 3 },
  { k: "closing_paid", label: "Closing Payment (BDT)", type: "number", sec: 3 },
  { k: "last_pay_date", label: "Last Payment Date", type: "date", sec: 3 },
  { k: "due_date", label: "Next Payment Due Date", type: "date", sec: 3 },
  { k: "commit_date", label: "Supplier Committed Delivery Date", type: "date", sec: 4 },
  { k: "inhouse_start", label: "Starting to In-house Date", type: "date", sec: 4 },
  { k: "qty_recv", label: "Qty Received In-house", type: "number", sec: 4 },
  { k: "close_date", label: "Final In-house Closing Date", type: "date", sec: 4 },
  { k: "qc_status", label: "QC / Material Approval", type: "list", sec: 5 },
  { k: "photo_link", label: "Approved Material Photo — File / Link", type: "text", sec: 5, wide: true },
  { k: "remarks", label: "Remarks", type: "text", sec: 5, wide: true },
];
var ENTRY_SECTIONS = { 1: "Requisition & OPEX/CAPEX approval", 2: "Purchase order", 3: "Finance & payments", 4: "In-house receiving", 5: "Material approval & notes" };
var ENTRY_MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12 };

function entryBlank(v) { return v === null || v === undefined || (typeof v === "string" && v.trim() === ""); }

/* Date from: 2026-10-15 · 15/10/2026 · 15-10-26 · 15 Oct 2026 · 15-Oct-26 · Excel serial.
   Day-first (Bangladesh convention). Returns ISO text, null when blank, undefined when invalid. */
function entryParseDate(v) {
  if (entryBlank(v)) return null;
  var y, m, d, mm;
  if (typeof v === "number") {
    if (v < 20000 || v > 80000) return undefined;
    return new Date(Math.round((Math.floor(v) - 25569) * 86400000)).toISOString().slice(0, 10);
  }
  if (Object.prototype.toString.call(v) === "[object Date]") {
    if (isNaN(v.getTime())) return undefined;
    y = v.getFullYear(); m = v.getMonth() + 1; d = v.getDate();
  } else {
    var s = String(v).trim();
    if ((mm = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ].*)?$/))) { y = +mm[1]; m = +mm[2]; d = +mm[3]; }
    else if ((mm = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})$/))) { d = +mm[1]; m = +mm[2]; y = +mm[3]; }
    else if ((mm = s.match(/^(\d{1,2})[\s\-]+([A-Za-z]{3,9})[\s\-,]+(\d{2}|\d{4})$/))) {
      d = +mm[1]; m = ENTRY_MONTHS[mm[2].toLowerCase()] || ENTRY_MONTHS[mm[2].slice(0, 3).toLowerCase()]; y = +mm[3];
    } else return undefined;
    if (y < 100) y += 2000;
  }
  if (!m || m < 1 || m > 12 || d < 1 || d > 31) return undefined;
  var t = new Date(Date.UTC(y, m - 1, d));
  if (t.getUTCFullYear() !== y || t.getUTCMonth() !== m - 1 || t.getUTCDate() !== d) return undefined;
  if (y < 2000 || y > 2100) return undefined;
  return t.toISOString().slice(0, 10);
}

/* Number from 1,250 · 1250.5 · "7.5%" (percent fields). null blank, undefined invalid. */
function entryParseNum(v, isPct) {
  if (entryBlank(v)) return null;
  if (typeof v === "number") return isFinite(v) ? v : undefined;
  var s = String(v).trim().replace(/,/g, "").replace(/^(BDT|Tk\.?|৳)\s*/i, "");
  var pct = /%$/.test(s);
  if (pct) s = s.slice(0, -1).trim();
  if (!/^-?\d+(\.\d+)?$/.test(s)) return undefined;
  var n = Number(s);
  return pct ? n / 100 : n;
}

/* Turn raw values (form, pasted row, file row) into a clean line.
   lists: { dept: [...], category: [...], ... } from the Lists sheet. */
function normalizeEntry(raw, lists) {
  var line = {}, errors = [], warnings = [];
  lists = lists || {};
  ENTRY_FIELDS.forEach(function (f) {
    var v = raw[f.k];
    if (f.type === "date") {
      var d = entryParseDate(v);
      if (d === undefined) { errors.push(f.label + ": '" + v + "' is not a valid date (use 15/10/2026 or 2026-10-15)"); d = null; }
      line[f.k] = d;
    } else if (f.type === "number" || f.type === "percent") {
      var n = entryParseNum(v, f.type === "percent");
      if (n === undefined) { errors.push(f.label + ": '" + v + "' is not a number"); n = null; }
      else if (n !== null && n < 0) { errors.push(f.label + " cannot be negative"); }
      if (f.type === "percent" && n !== null && n > 1) errors.push(f.label + ": enter as a percent such as 7.5% (not " + v + ")");
      line[f.k] = n;
    } else if (f.type === "list") {
      if (entryBlank(v)) { line[f.k] = null; return; }
      var s = String(v).trim(), opts = lists[f.k] || [], hit = null;
      for (var i = 0; i < opts.length; i++) { if (String(opts[i]).toLowerCase() === s.toLowerCase()) { hit = opts[i]; break; } }
      if (hit === null && opts.length) errors.push(f.label + ": '" + s + "' is not in the " + f.label + " list (an admin can add it)");
      line[f.k] = hit === null ? s : hit;
    } else {
      line[f.k] = entryBlank(v) ? null : String(v).trim();
    }
  });
  if (line.unit_price !== null && line.ccy === null) { line.ccy = "BDT"; warnings.push("Currency was blank — set to BDT"); }
  return { line: line, errors: errors, warnings: warnings };
}

/* Business rules. ctx: { existing: [{_row, pr_no, item}], selfRow, fx: {USD:122,...} } */
function checkEntry(line, ctx) {
  ctx = ctx || {};
  var errors = [], warnings = [];
  var lab = {}; ENTRY_FIELDS.forEach(function (f) { lab[f.k] = f.label; });
  ENTRY_FIELDS.forEach(function (f) { if (f.req && entryBlank(line[f.k])) errors.push(f.label + " is required"); });
  if (line.appr_status === "Approved" && entryBlank(line.appr_date)) errors.push("Approval Date is required when the status is Approved");
  if (!entryBlank(line.po_no)) {
    if (line.appr_status !== "Approved") errors.push("A PO can only be entered after the OPEX/CAPEX approval is Approved");
    ["po_date", "supplier", "qty", "unit_price"].forEach(function (k) { if (entryBlank(line[k])) errors.push(lab[k] + " is required once a PO No is entered"); });
    if (entryBlank(line.commit_date)) warnings.push("Supplier Committed Delivery Date is missing — lead-time follow-up needs it");
  }
  var before = function (a, b) { return !entryBlank(line[a]) && !entryBlank(line[b]) && line[a] < line[b]; };
  if (before("appr_date", "pr_date")) errors.push("Approval Date is before the PR Received Date");
  if (before("po_date", "appr_date")) errors.push("PO Date is before the Approval Date");
  if (before("close_date", "inhouse_start")) errors.push("Final In-house Closing Date is before Starting to In-house Date");
  if (before("inhouse_start", "po_date")) warnings.push("Goods started coming in-house before the PO Date");
  if (!entryBlank(line.close_date)) {
    if (entryBlank(line.inhouse_start)) warnings.push("Starting to In-house Date is missing for a closed line");
    if (entryBlank(line.qty_recv)) warnings.push("Qty Received is missing for a closed line");
    if (entryBlank(line.qc_status)) warnings.push("QC / Material Approval is missing for a closed line");
  }
  if (line.qty_recv !== null && line.qty !== null && line.qty_recv > line.qty * 1.05) warnings.push("Received quantity is more than 5% above the order");
  var paid = (line.adv_paid || 0) + (line.interim_paid || 0) + (line.closing_paid || 0);
  if (paid > 0 && entryBlank(line.pay_terms)) warnings.push("Payment Terms are missing although a payment is recorded");
  if (paid > 0 && line.qty !== null && line.unit_price !== null) {
    var fx = 1, fxs = ctx.fx || {};
    Object.keys(fxs).forEach(function (c) { if (line.ccy && c.toLowerCase() === String(line.ccy).toLowerCase()) fx = fxs[c]; });
    var total = line.qty * line.unit_price * fx * (1 + (line.tax_pct || 0));
    if (paid > total + 1) warnings.push("Total paid is more than the PO value incl. tax");
  }
  if (!entryBlank(line.pr_no) && !entryBlank(line.item)) {
    var key = String(line.pr_no).trim().toLowerCase() + "|" + String(line.item).trim().toLowerCase();
    (ctx.existing || []).forEach(function (e) {
      if (e._row !== ctx.selfRow && String(e.pr_no || "").trim().toLowerCase() + "|" + String(e.item || "").trim().toLowerCase() === key)
        errors.push("PR " + line.pr_no + " already has the item '" + line.item + "' (row " + e._row + ") — change it in Admin instead");
    });
  }
  return { errors: errors, warnings: warnings };
}

/* Fields that differ between two clean lines -> { key: [old, new] } */
function diffEntry(oldLine, newLine) {
  var out = {};
  ENTRY_FIELDS.forEach(function (f) {
    var a = entryBlank(oldLine[f.k]) ? null : oldLine[f.k], b = entryBlank(newLine[f.k]) ? null : newLine[f.k];
    if (typeof a === "number" && typeof b === "number" ? Math.abs(a - b) > 1e-9 : a !== b) out[f.k] = [a, b];
  });
  return out;
}
/* ===================== Logs: deliveries · payments · quotations · invoices ===================== */
var LOG_KINDS = {
  deliveries: { sheet: "Deliveries", prefix: "GRN", title: "Delivery (goods received)", needsPO: true, fields: [
    { k: "date", label: "Receipt Date", type: "date", req: true },
    { k: "qty", label: "Qty Received", type: "number", req: true },
    { k: "qty_ok", label: "Qty Accepted at QC", type: "number" },
    { k: "challan", label: "Delivery Challan No", type: "text" },
    { k: "remarks", label: "Remarks", type: "text", wide: true }] },
  payments: { sheet: "Payments", prefix: "PAY", title: "Payment or planned instalment", needsPO: true, fields: [
    { k: "entry", label: "Entry", type: "choice", opts: ["Paid", "Planned"], req: true },
    { k: "ptype", label: "Payment Type", type: "choice", opts: ["Advance", "Interim", "Closing"], req: true },
    { k: "date", label: "Date (paid or due)", type: "date", req: true },
    { k: "amount", label: "Amount (BDT)", type: "number", req: true },
    { k: "invoice_no", label: "Invoice No", type: "text" },
    { k: "settles", label: "Settles Planned ID", type: "text" },
    { k: "ref", label: "Bank / Cheque Ref", type: "text" },
    { k: "remarks", label: "Remarks", type: "text", wide: true }] },
  quotations: { sheet: "Quotations", prefix: "QT", title: "Supplier quotation", needsPO: false, fields: [
    { k: "supplier", label: "Supplier", type: "list", list: "supplier", req: true, wide: true },
    { k: "date", label: "Quote Date", type: "date", req: true },
    { k: "unit_price", label: "Unit Price", type: "number", req: true },
    { k: "ccy", label: "Currency", type: "list", list: "ccy" },
    { k: "lead_days", label: "Quoted Lead Time (days)", type: "number" },
    { k: "pay_terms", label: "Payment Terms", type: "list", list: "pay_terms" },
    { k: "remarks", label: "Remarks", type: "text", wide: true }] },
  invoices: { sheet: "Invoices", prefix: "INV", title: "Supplier invoice", needsPO: true, fields: [
    { k: "invoice_no", label: "Supplier Invoice No", type: "text", req: true },
    { k: "date", label: "Invoice Date", type: "date", req: true },
    { k: "qty", label: "Qty Invoiced", type: "number", req: true },
    { k: "unit_price", label: "Unit Price (PO Currency)", type: "number", req: true },
    { k: "amount", label: "Invoice Amount incl. Tax (BDT)", type: "number", req: true },
    { k: "remarks", label: "Remarks", type: "text", wide: true }] },
};
/* columns every log sheet has, before and after the kind's own fields */
var LOG_HEAD = [{ k: "id", label: "ID" }, { k: "pr_no", label: "PR No" }, { k: "item", label: "Item / Material Purchased" }, { k: "po_no", label: "PO No" }];
var LOG_TAIL = [{ k: "by", label: "Entered By" }, { k: "at", label: "Entered At" }];
function logColumns(kind) { return LOG_HEAD.concat(LOG_KINDS[kind].fields, LOG_TAIL); }
function lineKey(prNo, item) { return String(prNo || "").trim().toLowerCase() + "|" + String(item || "").trim().toLowerCase(); }

/* Clean one log entry. Returns { entry, errors, warnings }. */
function normalizeLog(kind, raw, lists) {
  var K = LOG_KINDS[kind], e = {}, errors = [], warnings = [];
  if (!K) return { entry: {}, errors: ["Unknown log type: " + kind], warnings: [] };
  lists = lists || {};
  e.pr_no = entryBlank(raw.pr_no) ? null : String(raw.pr_no).trim();
  e.item = entryBlank(raw.item) ? null : String(raw.item).trim();
  if (!e.pr_no) errors.push("PR No is required");
  if (!e.item) errors.push("Item is required");
  K.fields.forEach(function (f) {
    var v = raw[f.k];
    if (f.type === "date") { var d = entryParseDate(v); if (d === undefined) { errors.push(f.label + ": '" + v + "' is not a valid date"); d = null; } e[f.k] = d; }
    else if (f.type === "number") {
      var n = entryParseNum(v); if (n === undefined) { errors.push(f.label + ": '" + v + "' is not a number"); n = null; }
      else if (n !== null && n < 0) errors.push(f.label + " cannot be negative");
      e[f.k] = n;
    } else if (f.type === "choice" || f.type === "list") {
      if (entryBlank(v)) { e[f.k] = null; }
      else {
        var opts = f.type === "choice" ? f.opts : (lists[f.list] || []), s = String(v).trim(), hit = null;
        for (var i = 0; i < opts.length; i++) if (String(opts[i]).toLowerCase() === s.toLowerCase()) hit = opts[i];
        if (hit === null && opts.length) errors.push(f.label + ": '" + s + "' is not one of: " + opts.slice(0, 8).join(", ") + (opts.length > 8 ? "…" : ""));
        e[f.k] = hit === null ? s : hit;
      }
    } else e[f.k] = entryBlank(v) ? null : String(v).trim();
    if (f.req && (e[f.k] === null || e[f.k] === undefined)) errors.push(f.label + " is required");
  });
  return { entry: e, errors: errors, warnings: warnings };
}

/* Business rules for one log entry. ctx: { line (clean register line or null), logs: {kind: [entries of this line]}, today: ISO, fx } */
function checkLog(kind, e, ctx) {
  var errors = [], warnings = [], L = ctx.line, logs = ctx.logs || {};
  if (!L) { errors.push("No purchase line " + e.pr_no + " · " + e.item + " — check the PR No and item spelling"); return { errors: errors, warnings: warnings }; }
  if (LOG_KINDS[kind].needsPO && entryBlank(L.po_no)) errors.push("This line has no PO yet — a " + LOG_KINDS[kind].title.toLowerCase() + " needs a PO");
  var sum = function (arr, k, f) { return (arr || []).filter(f || function () { return true; }).reduce(function (s, x) { return s + (x[k] || 0); }, 0); };
  if (kind === "deliveries") {
    if (e.qty !== null && e.qty <= 0) errors.push("Qty Received must be more than 0");
    if (e.qty_ok !== null && e.qty !== null && e.qty_ok > e.qty) errors.push("Qty Accepted cannot be more than Qty Received");
    if (e.date && ctx.today && e.date > ctx.today) errors.push("Receipt Date is in the future");
    if (e.date && L.po_date && e.date < L.po_date) warnings.push("Receipt Date is before the PO Date");
    var after = sum(logs.deliveries, "qty") + (e.qty || 0);
    if (L.qty !== null && L.qty !== undefined && after > L.qty * 1.05) warnings.push("Total received (" + after + ") will be more than 5% above the order (" + L.qty + ")");
  }
  if (kind === "payments") {
    if (e.amount !== null && e.amount <= 0) errors.push("Amount must be more than 0");
    if (e.entry === "Paid" && e.date && ctx.today && e.date > ctx.today) errors.push("A payment cannot be dated in the future — record it as Planned");
    if (e.settles) {
      var plan = (logs.payments || []).filter(function (p) { return p.id === e.settles && p.entry === "Planned"; })[0];
      if (!plan) errors.push("Planned instalment " + e.settles + " was not found on this line");
      else if ((logs.payments || []).some(function (p) { return p.settles === e.settles; })) errors.push("Planned instalment " + e.settles + " is already settled");
      if (e.entry !== "Paid") errors.push("Only a Paid entry can settle a planned instalment");
    }
    if (e.entry === "Paid") {
      var fx = 1, fxs = ctx.fx || {};
      Object.keys(fxs).forEach(function (c) { if (L.ccy && c.toLowerCase() === String(L.ccy).toLowerCase()) fx = fxs[c]; });
      var total = (L.qty || 0) * (L.unit_price || 0) * fx * (1 + (L.tax_pct || 0));
      var paid = sum(logs.payments, "amount", function (p) { return p.entry === "Paid"; }) + (e.amount || 0);
      if (total > 0 && paid > total + 1) warnings.push("Total paid will be more than the PO value incl. tax");
    }
  }
  if (kind === "quotations") {
    if (e.unit_price !== null && e.unit_price <= 0) errors.push("Unit Price must be more than 0");
    if ((logs.quotations || []).some(function (q) { return String(q.supplier).toLowerCase() === String(e.supplier || "").toLowerCase(); }))
      warnings.push(e.supplier + " already quoted for this line — both quotes are kept, the lowest is used");
  }
  if (kind === "invoices") {
    if (e.qty !== null && e.qty <= 0) errors.push("Qty Invoiced must be more than 0");
    if (e.amount !== null && e.amount <= 0) errors.push("Invoice Amount must be more than 0");
    if ((ctx.allInvoices || []).some(function (i) { return String(i.invoice_no).toLowerCase() === String(e.invoice_no || "").toLowerCase() && i.supplier_key === String(L.supplier || "").toLowerCase(); }))
      errors.push("Invoice " + e.invoice_no + " from " + L.supplier + " is already recorded");
    if (e.date && L.po_date && e.date < L.po_date) warnings.push("Invoice Date is before the PO Date");
    if (e.unit_price !== null && L.unit_price && Math.abs(e.unit_price - L.unit_price) / L.unit_price > (ctx.invTol === undefined ? 0.01 : ctx.invTol))
      warnings.push("Invoice price differs from the PO price (" + L.unit_price + ")");
  }
  return { errors: errors, warnings: warnings };
}

/* Register summary columns rebuilt from a line's logs. Only kinds with entries are touched.
   Returns a patch: { qty_recv, inhouse_start, close_date, adv_paid, interim_paid, closing_paid, last_pay_date, due_date } */
function summarizeLogs(line, logs) {
  var p = {}, D = logs.deliveries || [], P = logs.payments || [];
  if (D.length) {
    var ds = D.slice().sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    var run = 0, crossed = null;
    ds.forEach(function (d) { run += d.qty || 0; if (crossed === null && line.qty !== null && line.qty !== undefined && run >= line.qty) crossed = d.date; });
    p.qty_recv = run; p.inhouse_start = ds[0].date;
    if (entryBlank(line.close_date) && crossed) p.close_date = crossed;
  }
  var paid = P.filter(function (x) { return x.entry === "Paid"; }), planned = P.filter(function (x) { return x.entry === "Planned"; });
  if (paid.length) {
    var by = { Advance: 0, Interim: 0, Closing: 0 };
    paid.forEach(function (x) { by[x.ptype] = (by[x.ptype] || 0) + (x.amount || 0); });
    p.adv_paid = by.Advance || null; p.interim_paid = by.Interim || null; p.closing_paid = by.Closing || null;
    p.last_pay_date = paid.map(function (x) { return x.date; }).sort().pop();
  }
  if (planned.length) {
    var settled = {}; paid.forEach(function (x) { if (x.settles) settled[x.settles] = true; });
    var open = planned.filter(function (x) { return !settled[x.id]; }).map(function (x) { return x.date; }).sort();
    p.due_date = open.length ? open[0] : null;
  }
  return p;
}

/* Fields owned by a log once it has entries — the forms lock them and the server refuses direct edits. */
var LOG_OWNED = { deliveries: ["qty_recv", "inhouse_start"], payments: ["adv_paid", "interim_paid", "closing_paid", "last_pay_date"] };

/* Opening entries that carry existing register values into a log the first time it is used, so nothing is lost. */
function openingLogs(kind, line, today) {
  var out = [], note = "Opening balance — value entered on the register before the log was used";
  if (kind === "deliveries" && line.qty_recv) out.push({ date: line.inhouse_start || line.close_date || today, qty: line.qty_recv, qty_ok: null, challan: null, remarks: note });
  if (kind === "payments") {
    [["adv_paid", "Advance"], ["interim_paid", "Interim"], ["closing_paid", "Closing"]].forEach(function (x) {
      if (line[x[0]]) out.push({ entry: "Paid", ptype: x[1], date: line.last_pay_date || line.po_date || today, amount: line[x[0]], invoice_no: null, settles: null, ref: null, remarks: note });
    });
  }
  return out;
}
/* Register patch after a log change. When the last entry of a kind is removed, the fields it fed are cleared. */
function rebuildPatch(line, logs, kind) {
  var p = summarizeLogs(line, logs);
  var paid = (logs.payments || []).filter(function (x) { return x.entry === "Paid"; });
  if (kind === "deliveries" && !(logs.deliveries || []).length) { p.qty_recv = null; p.inhouse_start = null; }
  if (kind === "payments" && !paid.length) { p.adv_paid = null; p.interim_paid = null; p.closing_paid = null; p.last_pay_date = null; }
  return p;
}
/* Register fields a line's logs currently control (locked in forms, refused in direct edits). */
function ownedFields(logs) {
  var o = [];
  if ((logs.deliveries || []).length) o = o.concat(LOG_OWNED.deliveries);
  if ((logs.payments || []).some(function (x) { return x.entry === "Paid"; })) o = o.concat(LOG_OWNED.payments);
  if ((logs.payments || []).some(function (x) { return x.entry === "Planned"; })) o.push("due_date");
  return o;
}
// </entry-rules>

if (typeof module === "object" && module.exports) {
  module.exports = { ENTRY_FIELDS: ENTRY_FIELDS, ENTRY_SECTIONS: ENTRY_SECTIONS, entryParseDate: entryParseDate, entryParseNum: entryParseNum,
    normalizeEntry: normalizeEntry, checkEntry: checkEntry, diffEntry: diffEntry,
    LOG_KINDS: LOG_KINDS, LOG_OWNED: LOG_OWNED, logColumns: logColumns, lineKey: lineKey, normalizeLog: normalizeLog, checkLog: checkLog, summarizeLogs: summarizeLogs, openingLogs: openingLogs, rebuildPatch: rebuildPatch, ownedFields: ownedFields };
}
