/**
 * Panda Purchase Analytics — live data feed from Google Sheets
 * ============================================================
 * Turns the Google Sheet version of the purchase workbook into a JSON feed
 * the web dashboard reads directly, so the dashboard is live as soon as
 * anyone edits the sheet.
 *
 * SET-UP (once)
 *  1. Upload Panda_Purchase_Tracking_Dashboard.xlsx to Google Drive and open it
 *     with Google Sheets (File ▸ Save as Google Sheets). Keep the sheet names.
 *  2. Extensions ▸ Apps Script. Paste this file as Code.gs and the provided
 *     appsscript.json (Project Settings ▸ show "appsscript.json" manifest).
 *  3. Deploy ▸ New deployment ▸ type "Web app"
 *       Execute as: Me      Who has access: Anyone
 *     Copy the URL that ends in /exec.
 *  4. In the dashboard repo, edit assets/js/config.js:
 *       dataSource: "apps_script",
 *       appsScriptUrl: "<the /exec URL>"
 *
 *  5. In the sheet: Purchase Dashboard ▸ Set access keys — choose an ENTRY key
 *     (for people who add purchase lines) and an ADMIN key (edit / delete / undo).
 *
 * The feed only exports the INPUT columns. All metrics are recalculated in the
 * browser by the same engine that is tested against the Excel formulas.
 * Writes (doPost) touch only the input cells of Purchase_Register, never the
 * formula columns, and every change is recorded on the Change_Log sheet.
 */

var HEADERS = {
  "PR No": "pr_no", "PR Received Date": "pr_date", "Department": "dept", "Requested By": "req_by",
  "Item / Material Purchased": "item", "Category": "category", "Spec / Size / Colour": "spec",
  "Spend Type (OPEX / CAPEX)": "spend", "OPEX / CAPEX Approval Ref": "appr_ref", "Approval Date": "appr_date",
  "Approval Status": "appr_status", "Approved By": "appr_by", "Supplier": "supplier", "Local / Import": "origin",
  "PO No": "po_no", "PO Date": "po_date", "Order Quantity": "qty", "UOM": "uom",
  "Unit Price (PO Currency)": "unit_price", "Currency": "ccy", "Budget / Standard Unit Price (BDT)": "budget_price",
  "VAT / AIT / Duty %": "tax_pct", "Payment Terms": "pay_terms", "Initial / Advance Paid (BDT)": "adv_paid",
  "Interim Payment (BDT)": "interim_paid", "Closing Payment (BDT)": "closing_paid", "Last Payment Date": "last_pay_date",
  "Next Payment Due Date": "due_date", "Supplier Committed Delivery Date": "commit_date",
  "Starting to In-house Date": "inhouse_start", "Qty Received In-house": "qty_recv",
  "Final In-house Closing Date": "close_date", "QC / Material Approval": "qc_status",
  "Approved Material Photo — File / Link": "photo_link", "Remarks": "remarks"
};
var DATE_KEYS = ["pr_date", "appr_date", "po_date", "last_pay_date", "due_date", "commit_date", "inhouse_start", "close_date"];
var NUM_KEYS = ["qty", "unit_price", "budget_price", "tax_pct", "adv_paid", "interim_paid", "closing_paid", "qty_recv"];
var LIST_KEYS = { "Department": "dept", "Category": "category", "Spend Type": "spend", "Approval Status": "appr_status",
  "Origin": "origin", "UOM": "uom", "Currency": "ccy", "Payment Terms": "pay_terms", "QC / Material Approval": "qc_status", "Supplier": "supplier" };
var SETTING_LABELS = { "Report Date": "report_date", "Company": "company", "OTIF quantity tolerance": "otif_tol",
  "Price variance alert threshold": "pv_alert", "Budget 'Watch' threshold": "watch", "Financial year start": "fy_start",
  "Approval SLA": "sla_appr", "PO issue SLA": "sla_po", "ABC class A cut-off": "abc_a", "ABC class B cut-off": "abc_b",
  "Quotations required": "min_quotes", "Quotation threshold": "quote_threshold", "Invoice price tolerance": "inv_tol" };

/** Web-app entry point: returns the dashboard data as JSON. */
function doGet() {
  var data = buildData_();
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function isBlank_(v) { return v === null || v === undefined || (typeof v === "string" && v.trim() === ""); }

function isoDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, "yyyy-MM-dd");
  if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
  return null;
}

function buildData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var warnings = [];

  // --- Purchase_Register (header row found by its "PR No" cell)
  var info = registerInfo_(ss), vals = info.values, hr = info.hr, idx = info.idx;

  var lines = [];
  for (var r2 = hr + 1; r2 < vals.length; r2++) {
    var row = vals[r2];
    if (isBlank_(row[idx.pr_no])) continue;
    var o = { _row: r2 + 1 };
    Object.keys(idx).forEach(function (k) {
      var v = row[idx[k]];
      if (isBlank_(v)) { o[k] = null; return; }
      if (DATE_KEYS.indexOf(k) >= 0) {
        var d = isoDate_(v, tz);
        if (!d) warnings.push("Row " + (r2 + 1) + " · " + k + ": '" + v + "' is not a date");
        o[k] = d;
      } else if (NUM_KEYS.indexOf(k) >= 0) {
        var n = typeof v === "number" ? v : Number(String(v).replace(/,/g, ""));
        if (!isFinite(n)) { warnings.push("Row " + (r2 + 1) + " · " + k + ": '" + v + "' is not a number"); n = null; }
        o[k] = n;
      } else {
        o[k] = String(v).trim();
      }
    });
    lines.push(o);
  }

  // --- Settings
  var settings = { fx: {} }, company = "Panda Shoes Industry Ltd", reportDate = null, inFx = false;
  ss.getSheetByName("Settings").getDataRange().getValues().forEach(function (row) {
    var a = row[0], b = row[1];
    if (isBlank_(a)) { inFx = false; return; }
    a = String(a);
    if (a === "Currency" && String(b).indexOf("Rate") === 0) { inFx = true; return; }
    if (inFx && typeof b === "number") { settings.fx[a.trim()] = b; return; }
    Object.keys(SETTING_LABELS).forEach(function (lab) {
      if (a.indexOf(lab) === 0) {
        var key = SETTING_LABELS[lab];
        if (key === "company") company = b;
        else if (key === "report_date") reportDate = isoDate_(b, tz);
        else if (key === "fy_start") settings.fy_start = isoDate_(b, tz);
        else settings[key] = b;
      }
    });
  });

  // --- Budget_Control
  var bv = ss.getSheetByName("Budget_Control").getDataRange().getValues();
  var bh = -1;
  for (var i = 0; i < bv.length; i++) { if (bv[i][0] === "Department") { bh = i; break; } }
  var budgets = [];
  if (bh >= 0) {
    var cO = bv[bh].indexOf("OPEX Budget"), cC = bv[bh].indexOf("CAPEX Budget");
    for (var j = bh + 1; j < bv.length; j++) {
      var dName = bv[j][0];
      if (isBlank_(dName) || dName === "TOTAL") continue;
      budgets.push({ dept: String(dName).trim(), opex: Number(bv[j][cO]) || 0, capex: Number(bv[j][cC]) || 0 });
    }
  }

  // --- Lists
  var lv = ss.getSheetByName("Lists").getRange(1, 1, 41, ss.getSheetByName("Lists").getLastColumn()).getValues();
  var lists = {};
  lv[0].forEach(function (h, c) {
    if (!LIST_KEYS[h]) return;
    lists[LIST_KEYS[h]] = lv.slice(1).map(function (row) { return row[c]; })
      .filter(function (v) { return !isBlank_(v); }).map(function (v) { return String(v).trim(); });
  });

  var out = {
    meta: { schema: 1, company: company, source: ss.getName() + " (Google Sheets)", generated_at: new Date().toISOString(),
            excel_report_date: reportDate, rows: lines.length, warnings: warnings },
    settings: settings, budgets: budgets, lists: lists, lines: lines
  };
  Object.keys(LOG_KINDS).forEach(function (kind) {
    var sh = ss.getSheetByName(LOG_KINDS[kind].sheet);
    if (sh) out[kind] = readLog_(sh, kind, tz, warnings);
  });
  return out;
}

/** One log sheet -> array of entries (keys from LOG_KINDS). */
function readLog_(sh, kind, tz, warnings) {
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  var cols = logColumns(kind), at = {};
  vals[0].forEach(function (h, i) { cols.forEach(function (c) { if (c.label === h) at[c.k] = { i: i, type: c.type || "text" }; }); });
  var out = [];
  for (var r = 1; r < vals.length; r++) {
    var row = vals[r];
    if (Object.keys(at).every(function (k) { return isBlank_(row[at[k].i]); })) continue;
    var e = {};
    cols.forEach(function (c) {
      var a = at[c.k];
      var v = a ? row[a.i] : null;
      if (isBlank_(v)) { e[c.k] = null; return; }
      if (c.k === "at") e[c.k] = v instanceof Date ? Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm") : String(v);
      else if (a.type === "date") { e[c.k] = isoDate_(v, tz); if (!e[c.k]) warnings.push(sh.getName() + " row " + (r + 1) + " · " + c.label + ": '" + v + "' is not a date"); }
      else if (a.type === "number") { var n = typeof v === "number" ? v : Number(String(v).replace(/,/g, "")); e[c.k] = isFinite(n) ? n : null; }
      else e[c.k] = String(v).trim();
    });
    out.push(e);
  }
  return out;
}

/** Register layout: header row, input column positions (0-based), all values. */
function registerInfo_(ss) {
  var reg = ss.getSheetByName("Purchase_Register");
  if (!reg) throw new Error("Sheet 'Purchase_Register' not found");
  var vals = reg.getDataRange().getValues();
  var hr = -1;
  for (var r = 0; r < Math.min(20, vals.length); r++) { if (vals[r].indexOf("PR No") >= 0) { hr = r; break; } }
  if (hr < 0) throw new Error("Header 'PR No' not found in Purchase_Register");
  var idx = {};
  vals[hr].forEach(function (h, i) { if (HEADERS[h]) idx[HEADERS[h]] = i; });
  var missing = Object.keys(HEADERS).filter(function (h) { return idx[HEADERS[h]] === undefined; });
  if (missing.length) throw new Error("Purchase_Register is missing columns: " + missing.join(", "));
  // contiguous runs of input columns, so writes never touch a formula column
  var cols = Object.keys(idx).map(function (k) { return idx[k]; }).sort(function (a, b) { return a - b; });
  var runs = [];
  cols.forEach(function (c) { var last = runs[runs.length - 1]; if (last && last.end === c - 1) last.end = c; else runs.push({ start: c, end: c }); });
  var keyAt = {}; Object.keys(idx).forEach(function (k) { keyAt[idx[k]] = k; });
  return { sheet: reg, values: vals, hr: hr, idx: idx, runs: runs, keyAt: keyAt };
}

/* ======================================================================
 *  WRITE API  (doPost)  — add · bulkAdd · update · delete · revert ·
 *  history · addListValue · auth
 * ==================================================================== */
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

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents || "{}");
    out = handle_(req);
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function role_(key) {
  var p = PropertiesService.getScriptProperties();
  var admin = p.getProperty("ADMIN_KEY"), entry = p.getProperty("ENTRY_KEY");
  if (!admin && !entry) throw new Error("Access keys are not set yet. In the Google Sheet use Purchase Dashboard ▸ Set access keys.");
  if (key && admin && key === admin) return "admin";
  if (key && entry && key === entry) return "entry";
  return null;
}

function handle_(req) {
  var role = role_(req.key);
  if (!role) return { ok: false, error: "The access key is not correct." };
  var name = String(req.name || "").trim();
  if (req.action === "auth") return { ok: true, role: role };
  if (!name) return { ok: false, error: "Please enter your name — every change is logged with it." };
  var adminOnly = ["update", "delete", "revert", "history", "addListValue", "deleteLog"];
  if (adminOnly.indexOf(req.action) >= 0 && role !== "admin") return { ok: false, error: "Only an admin can do this." };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = buildData_();
    var info = registerInfo_(ss);
    var who = { name: name, role: role };
    switch (req.action) {
      case "add": return addLines_(ss, info, data, [req.line], who, "Created")[0];
      case "bulkAdd": {
        if (!req.lines || !req.lines.length) return { ok: false, error: "No lines to save." };
        if (req.lines.length > 500) return { ok: false, error: "Please save at most 500 lines at a time." };
        var res = addLines_(ss, info, data, req.lines, who, "Created (bulk)");
        return { ok: true, results: res, saved: res.filter(function (r) { return r.ok; }).length };
      }
      case "update": return updateLine_(ss, info, data, req, who);
      case "delete": return deleteLine_(ss, info, data, req, who);
      case "revert": return revert_(ss, info, data, req, who);
      case "history": return history_(ss, req);
      case "addListValue": return addListValue_(ss, req, who);
      case "addLog": {
        if (!LOG_KINDS[req.kind]) return { ok: false, error: "Unknown log type." };
        if (!req.entries || !req.entries.length) return { ok: false, error: "Nothing to save." };
        if (req.entries.length > 500) return { ok: false, error: "Please save at most 500 entries at a time." };
        var lr = addLogs_(ss, info, data, req.kind, req.entries, who, null);
        return { ok: true, results: lr, saved: lr.filter(function (r) { return r.ok; }).length };
      }
      case "deleteLog": return deleteLog_(ss, info, data, req, who);
      default: return { ok: false, error: "Unknown action: " + req.action };
    }
  } finally {
    lock.releaseLock();
  }
}

function existingRefs_(data) { return data.lines.map(function (l) { return { _row: l._row, pr_no: l.pr_no, item: l.item }; }); }

/* Midnight of an ISO date in the spreadsheet's own time zone, so the cell shows exactly that date. */
function sheetDate_(iso) {
  return Utilities.parseDate(String(iso).slice(0, 10), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
}

function toCell_(k, v) {
  if (v === null || v === undefined) return "";
  if (DATE_KEYS.indexOf(k) >= 0) return sheetDate_(v);
  return v;
}

function writeRow_(info, row, line) {
  info.runs.forEach(function (run) {
    var vals = [];
    for (var c = run.start; c <= run.end; c++) { var k = info.keyAt[c]; vals.push(line ? toCell_(k, line[k]) : ""); }
    info.sheet.getRange(row, run.start + 1, 1, vals.length).setValues([vals]);
  });
}

/* First register row with an empty PR No; if none, a new row is added that copies
   the formulas and formats of the last row. */
function blankRow_(info, taken) {
  var sh = info.sheet, first = info.hr + 2, last = sh.getLastRow();
  var col = info.idx.pr_no + 1;
  if (last >= first) {
    var v = sh.getRange(first, col, last - first + 1, 1).getValues();
    for (var i = 0; i < v.length; i++) {
      var r = first + i;
      if ((v[i][0] === "" || v[i][0] === null) && !taken[r]) return r;
    }
  }
  sh.insertRowAfter(last);
  sh.getRange(last, 1, 1, sh.getLastColumn()).copyTo(sh.getRange(last + 1, 1, 1, sh.getLastColumn()));
  writeRow_(info, last + 1, null);
  return last + 1;
}

function addLines_(ss, info, data, rawLines, who, action) {
  var existing = existingRefs_(data), taken = {}, results = [];
  rawLines.forEach(function (raw) {
    var n = normalizeEntry(raw || {}, data.lists);
    var c = checkEntry(n.line, { existing: existing, fx: data.settings.fx });
    // a restore brings back exactly what was deleted: only a duplicate can stop it
    if (String(action).indexOf("Restored") === 0) c.errors = c.errors.filter(function (x) { return /already has the item/.test(x); });
    var errors = n.errors.concat(c.errors), warnings = n.warnings.concat(c.warnings);
    if (errors.length) { results.push({ ok: false, pr_no: n.line.pr_no, item: n.line.item, errors: errors, warnings: warnings }); return; }
    var row = blankRow_(info, taken);
    taken[row] = true;
    writeRow_(info, row, n.line);
    existing.push({ _row: row, pr_no: n.line.pr_no, item: n.line.item });
    log_(ss, who, action, n.line.pr_no, n.line.item, row, null, "", n.line);
    results.push({ ok: true, row: row, pr_no: n.line.pr_no, item: n.line.item, warnings: warnings });
  });
  return results;
}

function findLine_(data, req) {
  var l = null;
  data.lines.forEach(function (x) { if (x._row === Number(req.row)) l = x; });
  if (!l || l.pr_no !== req.pr_no || String(l.item || "") !== String(req.item || ""))
    throw new Error("This line was changed or moved since you opened it. Refresh and try again.");
  return l;
}

function cleanOld_(l) { var o = {}; ENTRY_FIELDS.forEach(function (f) { o[f.k] = l[f.k] === undefined ? null : l[f.k]; }); return o; }

function updateLine_(ss, info, data, req, who) {
  if (!String(req.reason || "").trim()) return { ok: false, error: "Please give a reason for the change." };
  var old = cleanOld_(findLine_(data, req));
  var n = normalizeEntry(req.line || {}, data.lists);
  var c = checkEntry(n.line, { existing: existingRefs_(data), selfRow: Number(req.row), fx: data.settings.fx });
  var errors = n.errors.concat(c.errors);
  if (errors.length) return { ok: false, errors: errors, warnings: n.warnings.concat(c.warnings) };
  var changes = diffEntry(old, n.line);
  if (!Object.keys(changes).length) return { ok: true, unchanged: true, row: Number(req.row) };
  var logs = lineLogs_(data, old.pr_no, old.item), owned = ownedFields(logs);
  var locked = owned.filter(function (k) { return changes[k]; });
  if (locked.length) return { ok: false, error: "These fields come from the delivery / payment logs and cannot be edited directly: " +
    locked.map(function (k) { return labelOf_(k); }).join(", ") + ". Add or delete a log entry instead." };
  if (changes.pr_no || changes.item || changes.po_no) renameInLogs_(ss, old, n.line);
  writeRow_(info, Number(req.row), n.line);
  var id = log_(ss, who, "Updated", n.line.pr_no, n.line.item, Number(req.row), changes, req.reason, null);
  return { ok: true, row: Number(req.row), changes: changes, logId: id, warnings: n.warnings.concat(c.warnings) };
}

function deleteLine_(ss, info, data, req, who) {
  if (!String(req.reason || "").trim()) return { ok: false, error: "Please give a reason for deleting." };
  var old = cleanOld_(findLine_(data, req));
  var lg = lineLogs_(data, old.pr_no, old.item);
  var nLogs = Object.keys(lg).reduce(function (s, k) { return s + lg[k].length; }, 0);
  if (nLogs) return { ok: false, error: "This line has " + nLogs + " delivery / payment / quotation / invoice entries. Delete those first (Admin ▸ Edit ▸ Logs), then delete the line." };
  writeRow_(info, Number(req.row), null);
  var id = log_(ss, who, "Deleted", old.pr_no, old.item, Number(req.row), null, req.reason, old);
  return { ok: true, row: Number(req.row), logId: id };
}

/* Undo a logged Update (puts the old values back) or a Delete (restores the line). */
function revert_(ss, info, data, req, who) {
  if (!String(req.reason || "").trim()) return { ok: false, error: "Please give a reason for the undo." };
  var sh = logSheet_(ss), id = Number(req.logId);
  if (!(id >= 2 && id <= sh.getLastRow())) return { ok: false, error: "That history entry does not exist." };
  var e = sh.getRange(id, 1, 1, 11).getValues()[0];
  var action = String(e[4]), row = Number(e[7]);
  if (action.indexOf("Deleted") === 0 && String(e[10] || "").indexOf('"__kind"') >= 0) {
    var lrec = JSON.parse(e[10]);
    var kind = lrec.__kind; delete lrec.__kind;
    var rr = addLogs_(ss, info, data, kind, [lrec], who, "Restored #" + id)[0];
    return rr.ok ? { ok: true, restored: true, id: rr.id, row: rr.row } : { ok: false, errors: rr.errors };
  }
  if (action.indexOf("Deleted") === 0) {
    var rec = JSON.parse(e[10] || "{}");
    var r = addLines_(ss, info, data, [rec], who, "Restored #" + id)[0];
    return r.ok ? { ok: true, row: r.row, restored: true } : { ok: false, errors: r.errors };
  }
  if (action.indexOf("Updated") === 0 || action.indexOf("Reverted") === 0) {
    var changes = JSON.parse(e[8] || "{}");
    var cur = null;
    data.lines.forEach(function (x) { if (x._row === row) cur = x; });
    if (!cur) return { ok: false, error: "The line on row " + row + " no longer exists." };
    var now = cleanOld_(cur), back = cleanOld_(cur), clash = [];
    Object.keys(changes).forEach(function (k) {
      var chk = {}; chk[k] = changes[k][1];
      var cmp = {}; cmp[k] = now[k];
      if (Object.keys(diffEntry(cmp, chk)).length) clash.push(k);
      back[k] = changes[k][0];
    });
    if (clash.length) return { ok: false, error: "This line was changed again after that entry (" + clash.join(", ") + "). Undo the later change first, or edit it directly." };
    var own = ownedFields(lineLogs_(data, cur.pr_no, cur.item)).filter(function (k) { return changes[k]; });
    if (own.length) return { ok: false, error: "That change touched fields now controlled by the delivery / payment logs (" + own.map(labelOf_).join(", ") + "). Adjust the log entries instead." };
    var c = checkEntry(back, { existing: existingRefs_(data), selfRow: row, fx: data.settings.fx });
    if (c.errors.length) return { ok: false, errors: c.errors };
    writeRow_(info, row, back);
    var undo = {}; Object.keys(changes).forEach(function (k) { undo[k] = [changes[k][1], changes[k][0]]; });
    var nid = log_(ss, who, "Reverted #" + id, back.pr_no, back.item, row, undo, req.reason, null);
    return { ok: true, row: row, logId: nid };
  }
  return { ok: false, error: "Only updates and deletions can be undone." };
}

/* ---------------------------------------------------------------- the four logs */
function labelOf_(k) { var f = ENTRY_FIELDS.filter(function (x) { return x.k === k; })[0]; return f ? f.label : k; }

function lineLogs_(data, prNo, item) {
  var key = lineKey(prNo, item), out = {};
  Object.keys(LOG_KINDS).forEach(function (kind) { out[kind] = (data[kind] || []).filter(function (e) { return lineKey(e.pr_no, e.item) === key; }); });
  return out;
}

function logSheetFor_(ss, kind) {
  var K = LOG_KINDS[kind], sh = ss.getSheetByName(K.sheet);
  if (!sh) {
    sh = ss.insertSheet(K.sheet);
    sh.appendRow(logColumns(kind).map(function (c) { return c.label; }));
    sh.setFrozenRows(1);
  }
  return sh;
}

function nextLogId_(data, kind) {
  var pre = LOG_KINDS[kind].prefix + "-", max = 0;
  (data[kind] || []).forEach(function (e) { var m = String(e.id || "").match(/-(\d+)$/); if (m) max = Math.max(max, Number(m[1])); });
  return pre + ("0000" + (max + 1)).slice(-Math.max(4, String(max + 1).length));
}

/* Write a log entry into the first row of its sheet whose ID cell is empty (or a new row). */
function appendLog_(ss, kind, entry) {
  var sh = logSheetFor_(ss, kind), cols = logColumns(kind);
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var row = cols.map(function () { return ""; });
  var pos = {}; head.forEach(function (h, i) { cols.forEach(function (c) { if (c.label === h) pos[c.k] = i; }); });
  var width = head.length, vals = head.map(function () { return ""; });
  cols.forEach(function (c) {
    if (pos[c.k] === undefined) return;
    var v = entry[c.k];
    if (v === null || v === undefined) v = "";
    else if (c.type === "date") v = sheetDate_(v);
    vals[pos[c.k]] = v;
  });
  var last = sh.getLastRow(), target = last + 1;
  if (last >= 2) {
    var ids = sh.getRange(2, pos.id + 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) if (isBlank_(ids[i][0])) { target = i + 2; break; }
  }
  sh.getRange(target, 1, 1, width).setValues([vals]);
  return target;
}

function today_(ss) { return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd"); }

/* Write the register summary (received qty, payments, dates) rebuilt from the logs. */
function applySummary_(info, line, logs, kind) {
  var patch = rebuildPatch(line, logs, kind), changed = false, next = cleanOld_(line);
  Object.keys(patch).forEach(function (k) {
    var a = next[k] === undefined ? null : next[k], b = patch[k] === undefined ? null : patch[k];
    if (typeof a === "number" && typeof b === "number" ? Math.abs(a - b) > 1e-9 : a !== b) { next[k] = b; changed = true; }
  });
  if (changed) { writeRow_(info, line._row, next); Object.keys(patch).forEach(function (k) { line[k] = patch[k]; }); }
  return changed;
}

function addLogs_(ss, info, data, kind, raws, who, actionOverride) {
  var results = [], today = today_(ss), fx = data.settings.fx, tol = data.settings.inv_tol;
  data[kind] = data[kind] || [];
  var supplierOf = {}; data.lines.forEach(function (l) { supplierOf[lineKey(l.pr_no, l.item)] = String(l.supplier || "").toLowerCase(); });
  raws.forEach(function (raw) {
    var n = normalizeLog(kind, raw || {}, data.lists), e = n.entry;
    var line = null; data.lines.forEach(function (l) { if (lineKey(l.pr_no, l.item) === lineKey(e.pr_no, e.item)) line = l; });
    var logs = line ? lineLogs_(data, line.pr_no, line.item) : {};
    var allInv = (data.invoices || []).map(function (i) { return { invoice_no: i.invoice_no, supplier_key: supplierOf[lineKey(i.pr_no, i.item)] }; });
    var c = n.errors.length ? { errors: [], warnings: [] } : checkLog(kind, e, { line: line, logs: logs, today: today, fx: fx, invTol: tol, allInvoices: allInv });
    var errors = n.errors.concat(c.errors);
    if (errors.length) { results.push({ ok: false, pr_no: e.pr_no, item: e.item, errors: errors }); return; }
    var opening = (kind === "deliveries" || kind === "payments") && !(logs[kind] || []).length ? openingLogs(kind, line, today) : [];
    opening.forEach(function (o) {
      var oe = Object.assign({ id: nextLogId_(data, kind), pr_no: line.pr_no, item: line.item, po_no: line.po_no, by: "System (opening balance)", at: new Date() }, o);
      appendLog_(ss, kind, oe); data[kind].push(oe);
    });
    var entry = Object.assign({}, e, { id: nextLogId_(data, kind), pr_no: line.pr_no, item: line.item, po_no: line.po_no || null, by: who.name, at: new Date() });
    appendLog_(ss, kind, entry); data[kind].push(entry);
    var after = lineLogs_(data, line.pr_no, line.item);
    if (kind === "deliveries" || kind === "payments") applySummary_(info, line, after, kind);
    var rec = Object.assign({ __kind: kind }, entry); delete rec.by; delete rec.at;
    log_(ss, who, actionOverride || ("Added " + LOG_KINDS[kind].title.split(" ")[0].toLowerCase() + " " + entry.id), line.pr_no, line.item, line._row, null, "", rec);
    results.push({ ok: true, id: entry.id, row: line._row, pr_no: line.pr_no, item: line.item, warnings: n.warnings.concat(c.warnings), opening: opening.length });
  });
  return results;
}

function deleteLog_(ss, info, data, req, who) {
  if (!String(req.reason || "").trim()) return { ok: false, error: "Please give a reason for deleting." };
  var kind = req.kind, K = LOG_KINDS[kind];
  if (!K) return { ok: false, error: "Unknown log type." };
  var entry = (data[kind] || []).filter(function (e) { return e.id === req.id; })[0];
  if (!entry) return { ok: false, error: "Entry " + req.id + " was not found — refresh and try again." };
  if (kind === "payments" && entry.entry === "Planned" && (data.payments || []).some(function (p) { return p.settles === entry.id; }))
    return { ok: false, error: "Planned instalment " + entry.id + " is settled by a payment. Delete that payment first." };
  var sh = logSheetFor_(ss, kind), vals = sh.getDataRange().getValues(), idCol = vals[0].indexOf("ID"), rowNo = -1;
  for (var r = 1; r < vals.length; r++) if (vals[r][idCol] === req.id) { rowNo = r + 1; break; }
  if (rowNo < 0) return { ok: false, error: "Entry " + req.id + " was not found on the " + K.sheet + " sheet." };
  sh.deleteRow(rowNo);
  data[kind] = data[kind].filter(function (e) { return e.id !== req.id; });
  var line = null; data.lines.forEach(function (l) { if (lineKey(l.pr_no, l.item) === lineKey(entry.pr_no, entry.item)) line = l; });
  if (line && (kind === "deliveries" || kind === "payments")) applySummary_(info, line, lineLogs_(data, line.pr_no, line.item), kind);
  var rec = Object.assign({ __kind: kind }, entry); delete rec.by; delete rec.at;
  var id = log_(ss, who, "Deleted " + K.title.split(" ")[0].toLowerCase() + " " + entry.id, entry.pr_no, entry.item, line ? line._row : 0, null, req.reason, rec);
  return { ok: true, logId: id };
}

/* A renamed line (PR No / item / PO No) keeps its log entries. */
function renameInLogs_(ss, oldL, newL) {
  var key = lineKey(oldL.pr_no, oldL.item);
  Object.keys(LOG_KINDS).forEach(function (kind) {
    var sh = ss.getSheetByName(LOG_KINDS[kind].sheet); if (!sh || sh.getLastRow() < 2) return;
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var cp = head.indexOf("PR No"), ci = head.indexOf("Item / Material Purchased"), co = head.indexOf("PO No");
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues();
    v.forEach(function (row, i) {
      if (lineKey(row[cp], row[ci]) !== key) return;
      sh.getRange(i + 2, cp + 1).setValue(newL.pr_no); sh.getRange(i + 2, ci + 1).setValue(newL.item);
      if (co >= 0) sh.getRange(i + 2, co + 1).setValue(newL.po_no || "");
    });
  });
}

function logSheet_(ss) {
  var sh = ss.getSheetByName("Change_Log");
  if (!sh) {
    sh = ss.insertSheet("Change_Log");
    sh.appendRow(["Log ID", "Time", "By", "Role", "Action", "PR No", "Item", "Row", "Changes (JSON)", "Reason", "Record (JSON)"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function log_(ss, who, action, prNo, item, row, changes, reason, record) {
  var sh = logSheet_(ss);
  var id = sh.getLastRow() + 1;
  sh.appendRow([id, new Date(), who.name, who.role, action, prNo || "", item || "", row,
    changes ? JSON.stringify(changes) : "", reason || "", record ? JSON.stringify(record) : ""]);
  return id;
}

function history_(ss, req) {
  var sh = logSheet_(ss), last = sh.getLastRow();
  if (last < 2) return { ok: true, entries: [] };
  var tz = ss.getSpreadsheetTimeZone();
  var v = sh.getRange(2, 1, last - 1, 11).getValues();
  var out = v.map(function (r) {
    return { id: Number(r[0]), time: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, "yyyy-MM-dd HH:mm") : String(r[1]),
      by: r[2], role: r[3], action: r[4], pr_no: r[5], item: r[6], row: Number(r[7]),
      changes: r[8] ? JSON.parse(r[8]) : null, reason: r[9], record: r[10] ? JSON.parse(r[10]) : null };
  }).reverse();
  if (req.pr_no) out = out.filter(function (x) { return String(x.pr_no).toLowerCase() === String(req.pr_no).toLowerCase(); });
  return { ok: true, entries: out.slice(0, 400) };
}

function addListValue_(ss, req, who) {
  var names = {}; Object.keys(LIST_KEYS).forEach(function (h) { names[LIST_KEYS[h]] = h; });
  var header = names[req.list], value = String(req.value || "").trim();
  if (!header || !value) return { ok: false, error: "Choose a list and type the new value." };
  var sh = ss.getSheetByName("Lists");
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var c = head.indexOf(header);
  if (c < 0) return { ok: false, error: "List '" + header + "' not found on the Lists sheet." };
  var v = sh.getRange(2, c + 1, 40, 1).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase() === value.toLowerCase()) return { ok: false, error: "'" + value + "' is already in the list." };
  }
  for (var j = 0; j < v.length; j++) {
    if (v[j][0] === "" || v[j][0] === null) {
      sh.getRange(j + 2, c + 1).setValue(value);
      log_(ss, who, "List value added", "", header + ": " + value, 0, null, "", null);
      return { ok: true, list: req.list, value: value };
    }
  }
  return { ok: false, error: "The " + header + " list is full (40 values)." };
}

/** Menu with a quick self-test, so the sheet owner can check the feed. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Purchase Dashboard")
    .addItem("Test the data feed", "testFeed")
    .addItem("Set access keys", "setAccessKeys")
    .addSeparator()
    .addItem("Set up email alerts", "setupAlerts")
    .addItem("Send a test alert now", "sendTestAlert")
    .addItem("Stop email alerts", "stopAlerts")
    .addToUi();
}

function testFeed() {
  var d = buildData_();
  var msg = d.lines.length + " purchase lines exported.\n" +
            (d.meta.warnings.length ? d.meta.warnings.length + " cell(s) could not be read:\n• " + d.meta.warnings.slice(0, 10).join("\n• ")
                                    : "No data problems found.");
  SpreadsheetApp.getUi().alert("Purchase Dashboard feed", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Owner sets the two access keys (stored in Script Properties, never in the website). */
function setAccessKeys() {
  var ui = SpreadsheetApp.getUi(), p = PropertiesService.getScriptProperties();
  var a = ui.prompt("Purchase Dashboard", "ENTRY key — for people who add purchase lines (min. 6 characters):", ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  var b = ui.prompt("Purchase Dashboard", "ADMIN key — for edit, delete and undo (min. 8 characters, different from the entry key):", ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  var ek = a.getResponseText().trim(), ak = b.getResponseText().trim();
  if (ek.length < 6 || ak.length < 8 || ek === ak) { ui.alert("Keys not saved: the entry key needs 6+ characters, the admin key 8+, and they must differ."); return; }
  p.setProperty("ENTRY_KEY", ek); p.setProperty("ADMIN_KEY", ak);
  ui.alert("Access keys saved. Share the entry key with data-entry staff and keep the admin key to admins.");
}
