/**
 * Proves the web engine gives exactly the same results as the Excel workbook.
 *
 *   python scripts/excel_to_json.py --snapshot tests/excel_snapshot.json
 *   node tests/engine.test.js
 *
 * Every calculated column of every register line is compared with Excel's own
 * value, plus the Supplier_Scorecard, Spend_Analysis (ABC) and Budget_Control
 * results.  Exit code 1 on any mismatch (used by the GitHub Action).
 */
const fs = require("fs");
const path = require("path");
const E = require("../assets/js/engine.js");

const root = path.join(__dirname, "..");
const data = JSON.parse(fs.readFileSync(path.join(root, "data/purchase_data.json"), "utf8"));
const snapPath = path.join(__dirname, "excel_snapshot.json");
if (!fs.existsSync(snapPath)) { console.log("No excel_snapshot.json — run the extractor with --snapshot first."); process.exit(0); }
const snap = JSON.parse(fs.readFileSync(snapPath, "utf8"));

const R = E.toDay(snap.report_date);
const lines = E.enrich(data, R);
const A = E.analyze(lines, data, R);
let checks = 0, fails = [];

function same(a, b) {
  const ea = a === null || a === undefined || a === "", eb = b === null || b === undefined || b === "";
  if (ea || eb) return ea === eb;
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) <= 1e-6 * Math.max(1, Math.abs(b)) + 1e-6;
  return String(a) === String(b);
}
function cmp(label, got, exp) { checks++; if (!same(got, exp)) fails.push(`${label}: web=${JSON.stringify(got)} excel=${JSON.stringify(exp)}`); }

// 1. row level
snap.lines.forEach((s) => {
  const l = lines.find((x) => x.pr_no === s.pr_no);
  if (!l) { fails.push(`missing line ${s.pr_no}`); return; }
  Object.keys(s).forEach((k) => { if (k !== "_row" && k !== "pr_no") cmp(`${s.pr_no}.${k}`, l[k], s[k]); });
});
// 2. supplier scorecard
(snap.suppliers || []).forEach((s) => {
  const x = A.suppliers.find((y) => y.supplier === s.supplier);
  if (!x) { if (s.spend) fails.push(`missing supplier ${s.supplier}`); return; }
  ["spend", "otif", "infull", "qc", "score", "rating"].forEach((k) => cmp(`supplier ${s.supplier}.${k}`, x[k], s[k]));
});
// 3. ABC categories (order and class)
(snap.categories || []).filter((c) => c.spend).forEach((c, i) => {
  const x = A.categories[i] || {};
  cmp(`category #${i + 1} name`, x.category, c.category); cmp(`category ${c.category}.spend`, x.spend, c.spend); cmp(`category ${c.category}.class`, x.cls, c.cls);
});
// 4. budget status
(snap.budgets || []).forEach((b) => {
  const x = A.budgets.find((y) => y.dept === b.dept);
  if (!x) { if (b.committed || (b.status && b.status !== "No Budget")) fails.push(`missing budget ${b.dept}`); return; }
  cmp(`budget ${b.dept}.committed`, x.committed, b.committed); cmp(`budget ${b.dept}.status`, x.status, b.status);
});

console.log(`Engine vs Excel: ${checks} checks, ${fails.length} mismatch(es) — report date ${snap.report_date}`);
fails.slice(0, 40).forEach((f) => console.log("  ✗ " + f));
process.exit(fails.length ? 1 : 0);
