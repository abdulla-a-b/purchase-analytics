/**
 * Load the purchase Excel file directly in the browser ("Load Excel" button).
 * Mirrors scripts/excel_to_json.py: reads only the input columns, by header text.
 * Nothing is uploaded anywhere — the file is read locally in the viewer's browser.
 */
(function () {
  "use strict";
  const SHEETJS = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
  const HEADERS = {
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
    "Approved Material Photo — File / Link": "photo_link", "Remarks": "remarks",
  };
  const DATES = new Set(["pr_date", "appr_date", "po_date", "last_pay_date", "due_date", "commit_date", "inhouse_start", "close_date"]);
  const LISTS = { "Department": "dept", "Category": "category", "Spend Type": "spend", "Approval Status": "appr_status", "Origin": "origin",
    "UOM": "uom", "Currency": "ccy", "Payment Terms": "pay_terms", "QC / Material Approval": "qc_status", "Supplier": "supplier" };
  const SETTINGS = { "OTIF quantity tolerance": "otif_tol", "Price variance alert threshold": "pv_alert", "Budget 'Watch' threshold": "watch",
    "Financial year start": "fy_start", "Approval SLA": "sla_appr", "PO issue SLA": "sla_po", "ABC class A cut-off": "abc_a", "ABC class B cut-off": "abc_b",
    "Report Date": "report_date", "Company": "company", "Quotations required": "min_quotes", "Quotation threshold": "quote_threshold",
    "Invoice price tolerance": "inv_tol" };

  function loadLib() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    return new Promise((res, rej) => {
      const s = document.createElement("script"); s.src = SHEETJS;
      s.onload = () => res(window.XLSX); s.onerror = () => rej(new Error("Could not load the Excel reader. Check the internet connection."));
      document.head.appendChild(s);
    });
  }
  // Excel serial date -> ISO (1900 date system)
  const serialToISO = (n) => new Date(Math.round((Math.floor(n + 1e-9) - 25569) * 86400000)).toISOString().slice(0, 10);
  const blank = (v) => v === null || v === undefined || (typeof v === "string" && !v.trim());
  function rows(ws, XLSX) { return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true }); }

  async function parse(file) {
    const XLSX = await loadLib();
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: false });
    const need = ["Purchase_Register", "Settings", "Budget_Control", "Lists"];
    const miss = need.filter((n) => !wb.Sheets[n]);
    if (miss.length) throw new Error("This file is missing the sheet(s): " + miss.join(", "));

    const R = rows(wb.Sheets.Purchase_Register, XLSX);
    const hr = R.findIndex((r) => r && r.includes("PR No"));
    if (hr < 0) throw new Error("Could not find the 'PR No' header in Purchase_Register.");
    const head = R[hr];
    const idx = {}; head.forEach((h, i) => { if (h && HEADERS[h]) idx[HEADERS[h]] = i; });
    const missing = Object.keys(HEADERS).filter((h) => idx[HEADERS[h]] === undefined);
    if (missing.length) throw new Error("Purchase_Register is missing columns: " + missing.join(", "));
    const warnings = [];
    const lines = [];
    for (let r = hr + 1; r < R.length; r++) {
      const row = R[r] || [];
      if (blank(row[idx.pr_no])) continue;
      const o = { _row: r + 1 };
      Object.entries(idx).forEach(([k, i]) => {
        let v = row[i];
        if (blank(v)) { o[k] = null; return; }
        if (DATES.has(k)) {
          if (typeof v === "number") v = serialToISO(v);
          else if (!/^\d{4}-\d{2}-\d{2}/.test(String(v))) { warnings.push(`Row ${r + 1} · ${k}: '${v}' is not a date`); v = null; }
        }
        o[k] = typeof v === "string" ? v.trim() : v;
      });
      lines.push(o);
    }

    const S = rows(wb.Sheets.Settings, XLSX);
    const settings = { fx: {} }; let inFx = false, company = null, excelDate = null;
    S.forEach((r) => {
      const a = r && r[0], b = r && r[1];
      if (blank(a)) { inFx = false; return; }
      if (a === "Currency" && String(b).startsWith("Rate")) { inFx = true; return; }
      if (inFx && typeof b === "number") { settings.fx[String(a).trim()] = b; return; }
      Object.entries(SETTINGS).forEach(([lab, key]) => {
        if (String(a).startsWith(lab)) {
          if (key === "company") company = b;
          else if (key === "report_date") excelDate = typeof b === "number" ? serialToISO(b) : null;
          else if (key === "fy_start") settings.fy_start = typeof b === "number" ? serialToISO(b) : b;
          else settings[key] = b;
        }
      });
    });

    const B = rows(wb.Sheets.Budget_Control, XLSX);
    const bh = B.findIndex((r) => r && r[0] === "Department");
    const bcol = (name) => B[bh].indexOf(name);
    const budgets = [];
    for (let r = bh + 1; r < B.length; r++) {
      const d = B[r] && B[r][0];
      if (blank(d) || d === "TOTAL") continue;
      budgets.push({ dept: String(d).trim(), opex: B[r][bcol("OPEX Budget")] || 0, capex: B[r][bcol("CAPEX Budget")] || 0 });
    }

    const Lr = rows(wb.Sheets.Lists, XLSX);
    const lists = {};
    (Lr[0] || []).forEach((h, c) => {
      if (LISTS[h]) lists[LISTS[h]] = Lr.slice(1, 41).map((r) => r && r[c]).filter((v) => !blank(v) && typeof v !== "object").map((v) => String(v).trim());
    });

    // the four logs (optional sheets), columns by header text from LOG_KINDS (entry-rules.js)
    const logs = {};
    if (window.LOG_KINDS) Object.keys(window.LOG_KINDS).forEach((kind) => {
      const ws = wb.Sheets[window.LOG_KINDS[kind].sheet]; if (!ws) return;
      const rr = rows(ws, XLSX), cols = window.logColumns(kind), at = {};
      (rr[0] || []).forEach((h, i) => cols.forEach((c) => { if (c.label === h) at[c.k] = { i, type: c.type || "text" }; }));
      logs[kind] = [];
      rr.slice(1).forEach((r, ri) => {
        if (!r || Object.keys(at).every((k) => blank(r[at[k].i]))) return;
        const e = {};
        cols.forEach((c) => {
          const a = at[c.k]; let v = a ? r[a.i] : null;
          if (blank(v)) { e[c.k] = null; return; }
          if (c.k === "at") e[c.k] = typeof v === "number" ? new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 16) : String(v);
          else if (a.type === "date") { e[c.k] = typeof v === "number" ? serialToISO(v) : (/^\d{4}-\d{2}-\d{2}/.test(String(v)) ? String(v).slice(0, 10) : null); if (!e[c.k]) warnings.push(`${window.LOG_KINDS[kind].sheet} row ${ri + 2}: '${v}' is not a date`); }
          else if (a.type === "number") { const n = typeof v === "number" ? v : Number(String(v).replace(/,/g, "")); e[c.k] = isFinite(n) ? n : null; }
          else e[c.k] = String(v).trim();
        });
        logs[kind].push(e);
      });
    });

    return Object.assign({ meta: { schema: 1, company: company || "Panda Shoes Industry Ltd", source: file.name, generated_at: new Date().toISOString(),
      excel_report_date: excelDate, rows: lines.length, warnings }, settings, budgets, lists, lines }, logs);
  }

  window.PurchaseXlsx = { parse, loadLib };
})();
