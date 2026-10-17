/**
 * Logs test:  node tests/logs.test.js
 *  1. every log entry in data/purchase_data.json has exactly the columns defined in LOG_KINDS
 *  2. rebuilding the register summary from the logs (summarizeLogs) gives the values
 *     already on the register — so logs and register never disagree
 *  3. the log rules (normalizeLog / checkLog / openingLogs) behave as documented
 */
const fs = require("fs"), path = require("path");
const R = require("../assets/js/entry-rules.js");
const data = JSON.parse(fs.readFileSync(path.join(__dirname, "../data/purchase_data.json"), "utf8"));
let n = 0; const fails = [];
const yes = (c, m) => { n++; if (!c) fails.push(m); };
const same = (a, b) => (a === null || a === undefined || a === 0 ? null : a) === (b === null || b === undefined || b === 0 ? null : b) ||
  (typeof a === "number" && typeof b === "number" && Math.abs(a - b) < 0.01);

// 1. columns
Object.keys(R.LOG_KINDS).forEach((kind) => {
  if (!data[kind]) return;
  const keys = R.logColumns(kind).map((c) => c.k).sort().join(",");
  data[kind].forEach((e, i) => yes(Object.keys(e).sort().join(",") === keys, `${kind}[${i}] columns differ from LOG_KINDS`));
});

// 2. logs -> register summary
const byLine = {};
Object.keys(R.LOG_KINDS).forEach((kind) => (data[kind] || []).forEach((e) => {
  const k = R.lineKey(e.pr_no, e.item); (byLine[k] = byLine[k] || {})[kind] = ((byLine[k] || {})[kind] || []).concat([e]);
}));
let linesChecked = 0;
data.lines.forEach((l) => {
  const logs = byLine[R.lineKey(l.pr_no, l.item)]; if (!logs) return;
  linesChecked++;
  const p = R.summarizeLogs(l, logs);
  Object.keys(p).forEach((f) => yes(same(p[f], l[f]), `${l.pr_no} ${l.item}: ${f} from logs = ${p[f]}, register = ${l[f]}`));
});
yes(linesChecked > 30, "too few lines with logs: " + linesChecked);
// every log points at a real line
Object.keys(R.LOG_KINDS).forEach((kind) => (data[kind] || []).forEach((e) =>
  yes(data.lines.some((l) => R.lineKey(l.pr_no, l.item) === R.lineKey(e.pr_no, e.item)), `${kind} ${e.id} points to no line`)));

// 3. rules
const lists = { supplier: ["Savar Leather Works Ltd"], ccy: ["BDT", "USD"], pay_terms: ["Credit 30 Days"] };
const line = { pr_no: "PR-1", item: "Leather", po_no: "PO-1", po_date: "2026-10-01", qty: 100, unit_price: 10, ccy: "BDT", tax_pct: 0, supplier: "Savar Leather Works Ltd" };
const chk = (kind, raw, logs, re, msg, wantWarn) => {
  const a = R.normalizeLog(kind, Object.assign({ pr_no: "PR-1", item: "Leather" }, raw), lists);
  const c = R.checkLog(kind, a.entry, { line, logs: logs || {}, today: "2026-10-15", fx: { BDT: 1 }, allInvoices: [{ invoice_no: "A-1", supplier_key: "savar leather works ltd" }] });
  const pool = wantWarn ? c.warnings : a.errors.concat(c.errors);
  yes(re ? pool.some((x) => re.test(x)) : !a.errors.concat(c.errors).length, msg + " :: " + JSON.stringify(a.errors.concat(c.errors, c.warnings)));
};
chk("deliveries", { date: "14/10/2026", qty: "40" }, null, null, "valid delivery");
chk("deliveries", { date: "20/10/2026", qty: "40" }, null, /future/, "future receipt rejected");
chk("deliveries", { date: "14/10/2026", qty: "40", qty_ok: "50" }, null, /Accepted cannot/, "accepted > received rejected");
chk("deliveries", { date: "14/10/2026", qty: "40" }, { deliveries: [{ qty: 70 }] }, /more than 5%/, "over-receipt warns", true);
chk("payments", { entry: "paid", ptype: "advance", date: "2026-10-02", amount: "300" }, null, null, "valid payment (case-insensitive choices)");
chk("payments", { entry: "Paid", ptype: "Advance", date: "2026-10-20", amount: "300" }, null, /record it as Planned/, "future paid rejected");
chk("payments", { entry: "Paid", ptype: "Bonus", date: "2026-10-02", amount: "300" }, null, /not one of/, "unknown payment type");
chk("payments", { entry: "Paid", ptype: "Closing", date: "2026-10-02", amount: "300", settles: "PAY-9" }, { payments: [] }, /not found/, "settles unknown plan");
chk("payments", { entry: "Paid", ptype: "Closing", date: "2026-10-02", amount: "1200" }, null, /more than the PO value/, "overpayment warns", true);
chk("quotations", { supplier: "savar leather works ltd", date: "2026-09-25", unit_price: "9.5" }, null, null, "valid quotation");
chk("quotations", { supplier: "Nobody", date: "2026-09-25", unit_price: "9.5" }, null, /not one of/, "quote supplier must be listed");
chk("invoices", { invoice_no: "A-1", date: "2026-10-10", qty: "100", unit_price: "10", amount: "1000" }, null, /already recorded/, "duplicate invoice rejected");
chk("invoices", { invoice_no: "A-2", date: "2026-10-10", qty: "100", unit_price: "10.5", amount: "1050" }, null, /differs from the PO price/, "invoice price warns", true);
const noPO = R.checkLog("deliveries", { pr_no: "PR-1", item: "Leather", date: "2026-10-10", qty: 5 }, { line: Object.assign({}, line, { po_no: null }), today: "2026-10-15" });
yes(noPO.errors.some((x) => /no PO yet/.test(x)), "delivery needs PO");
yes(R.checkLog("quotations", { pr_no: "PR-1", item: "Leather", supplier: "Savar Leather Works Ltd", date: "2026-10-01", unit_price: 5 }, { line: Object.assign({}, line, { po_no: null }) }).errors.length === 0, "quotation allowed before PO");
const op = R.openingLogs("payments", { adv_paid: 300, closing_paid: 700, last_pay_date: "2026-10-05" }, "2026-10-15");
yes(op.length === 2 && op[0].ptype === "Advance" && op[1].amount === 700 && op[0].date === "2026-10-05", "opening payments carry register values");
const sm = R.summarizeLogs({ qty: 100, close_date: null }, { deliveries: [{ date: "2026-10-03", qty: 60 }, { date: "2026-10-01", qty: 50 }],
  payments: [{ id: "P1", entry: "Planned", date: "2026-10-20" }, { id: "P2", entry: "Planned", date: "2026-10-10" }, { entry: "Paid", ptype: "Advance", date: "2026-10-02", amount: 30, settles: "P2" }] });
yes(sm.qty_recv === 110 && sm.inhouse_start === "2026-10-01" && sm.close_date === "2026-10-03" && sm.adv_paid === 30 && sm.due_date === "2026-10-20",
  "summary: totals, first receipt, auto-close on crossing, next open instalment " + JSON.stringify(sm));

console.log(`Logs: ${n} checks, ${fails.length} failure(s) — ${linesChecked} lines rebuilt from their logs`);
fails.slice(0, 25).forEach((f) => console.log("  ✗ " + f));
process.exit(fails.length ? 1 : 0);
