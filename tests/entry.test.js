/**
 * Data-entry rules test:  node tests/entry.test.js
 *  1. the rules block in apps-script/Code.gs is identical to assets/js/entry-rules.js
 *     (run `python scripts/sync_rules.py` after editing the rules)
 *  2. date / number parsing and the business rules behave as documented
 */
const fs = require("fs"), path = require("path");
const R = require("../assets/js/entry-rules.js");
const root = path.join(__dirname, "..");
let n = 0, fails = [];
const eq = (got, exp, msg) => { n++; if (JSON.stringify(got) !== JSON.stringify(exp)) fails.push(`${msg}: got ${JSON.stringify(got)}, expected ${JSON.stringify(exp)}`); };
const yes = (c, msg) => { n++; if (!c) fails.push(msg); };

// 1. rules identical in website and Google Sheet backend
const block = (s) => s.slice(s.indexOf("// <entry-rules>"), s.indexOf("// </entry-rules>"));
yes(block(fs.readFileSync(path.join(root, "assets/js/entry-rules.js"), "utf8")) === block(fs.readFileSync(path.join(root, "apps-script/Code.gs"), "utf8")),
  "Code.gs rules differ from entry-rules.js — run: python scripts/sync_rules.py");

// 2. parsing
[["2026-10-15", "2026-10-15"], ["15/10/2026", "2026-10-15"], ["15-10-26", "2026-10-15"], ["5.1.2027", "2027-01-05"], ["15 Oct 2026", "2026-10-15"],
 ["15-Oct-26", "2026-10-15"], ["1 September 2026", "2026-09-01"], [46310, "2026-10-15"], ["", null], [null, null],
 ["31/02/2026", undefined], ["13/13/2026", undefined], ["next week", undefined], ["10/15/2026", undefined]]
  .forEach(([v, e]) => eq(R.entryParseDate(v), e, `date ${JSON.stringify(v)}`));
[["1,250", 1250], ["1250.5", 1250.5], ["৳ 2,000", 2000], ["7.5%", 0.075], ["abc", undefined], ["", null], [12, 12]]
  .forEach(([v, e]) => eq(R.entryParseNum(v), e, `number ${JSON.stringify(v)}`));

// 3. normalising and rules
const lists = { dept: ["Cutting"], category: ["Leather & Upper"], spend: ["OPEX", "CAPEX"], appr_status: ["Approved", "Pending", "Rejected"],
  supplier: ["Savar Leather Works Ltd"], uom: ["sqft"], ccy: ["BDT", "USD"], origin: ["Local", "Import"] };
const good = { pr_no: "PR-1", pr_date: "01/10/2026", dept: "cutting", item: "Leather", category: "leather & upper", spend: "opex",
  appr_status: "approved", appr_date: "02/10/2026", po_no: "PO-1", po_date: "03/10/2026", supplier: "savar leather works ltd", qty: "100", unit_price: "250", commit_date: "2026-10-20", tax_pct: "5%" };
let a = R.normalizeEntry(good, lists);
eq(a.errors, [], "clean line has no errors");
eq([a.line.dept, a.line.spend, a.line.supplier, a.line.ccy, a.line.tax_pct], ["Cutting", "OPEX", "Savar Leather Works Ltd", "BDT", 0.05], "list values use list spelling; blank currency -> BDT");
eq(R.checkEntry(a.line, {}).errors, [], "clean line passes rules");
const rule = (patch, re, msg) => { const x = R.normalizeEntry(Object.assign({}, good, patch), lists); const c = R.checkEntry(x.line, { existing: [{ _row: 9, pr_no: "PR-1", item: "Old" }] });
  yes(x.errors.concat(c.errors).some((e) => re.test(e)), msg); };
rule({ dept: "" }, /Department is required/, "required field");
rule({ appr_status: "Pending" }, /only be entered after/, "PO needs approval");
rule({ supplier: "Nobody Ltd" }, /not in the Supplier list/, "supplier must be in list");
rule({ appr_date: "30/09/2026" }, /before the PR Received Date/, "approval before PR");
rule({ qty: "-1" }, /cannot be negative/, "negative number");
rule({ tax_pct: "7.5" }, /percent/, "tax given as 7.5 instead of 7.5%");
rule({ item: "old" }, /already has the item/, "duplicate PR + item");
const self = R.normalizeEntry(Object.assign({}, good, { item: "Old" }), lists).line;
eq(R.checkEntry(self, { existing: [{ _row: 9, pr_no: "PR-1", item: "Old" }], selfRow: 9 }).errors, [], "editing the same row is not a duplicate");
yes(R.checkEntry(R.normalizeEntry(Object.assign({}, good, { commit_date: "" }), lists).line, {}).warnings.some((w) => /Committed/.test(w)), "missing committed date warns");
eq(R.diffEntry({ qty: 100, remarks: null }, { qty: 120, remarks: "" }), { qty: [100, 120] }, "diff ignores blank-vs-null");

console.log(`Entry rules: ${n} checks, ${fails.length} failure(s)`);
fails.forEach((f) => console.log("  ✗ " + f));
process.exit(fails.length ? 1 : 0);
