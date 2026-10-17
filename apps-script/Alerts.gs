/**
 * Panda Purchase Analytics — email alerts
 * =======================================
 * Sends management a procurement alert email on a schedule, built with the same
 * calculation engine as the web dashboard (Engine.gs is an exact copy of
 * assets/js/engine.js), so the email and the dashboard always agree.
 *
 * In the Google Sheet:  Purchase Dashboard ▸ Set up email alerts
 *   recipients · dashboard link · daily (every morning) or weekly (Sunday morning)
 *   Purchase Dashboard ▸ Send a test alert now   ·   Purchase Dashboard ▸ Stop email alerts
 */

function setupAlerts() {
  var ui = SpreadsheetApp.getUi(), p = PropertiesService.getScriptProperties();
  var a = ui.prompt("Email alerts — recipients", "Email addresses, separated by commas:\n(current: " + (p.getProperty("ALERT_TO") || "none") + ")", ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  var to = a.getResponseText().split(/[,;\s]+/).map(function (x) { return x.trim(); }).filter(function (x) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x); });
  if (!to.length) { ui.alert("No valid email address was entered — alerts not changed."); return; }
  var b = ui.prompt("Email alerts — dashboard link", "Link to the web dashboard (shown as a button in the email). Leave blank to skip:\n(current: " + (p.getProperty("DASHBOARD_URL") || "none") + ")", ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  var c = ui.prompt("Email alerts — how often", "Type  daily  (every morning at 8:00) or  weekly  (Sunday at 8:00):", ui.ButtonSet.OK_CANCEL);
  if (c.getSelectedButton() !== ui.Button.OK) return;
  var freq = c.getResponseText().trim().toLowerCase().indexOf("week") === 0 ? "weekly" : "daily";
  p.setProperty("ALERT_TO", to.join(","));
  if (b.getResponseText().trim()) p.setProperty("DASHBOARD_URL", b.getResponseText().trim());
  p.setProperty("ALERT_FREQ", freq);
  removeAlertTriggers_();
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var t = ScriptApp.newTrigger("sendDigest").timeBased().atHour(8).inTimezone(tz);
  if (freq === "weekly") t.onWeekDay(ScriptApp.WeekDay.SUNDAY).create(); else t.everyDays(1).create();
  ui.alert("Email alerts are on: " + freq + " at 8:00 (" + tz + ") to " + to.join(", ") + ".\nUse 'Send a test alert now' to see one straight away.");
}

function stopAlerts() {
  removeAlertTriggers_();
  SpreadsheetApp.getUi().alert("Email alerts are stopped. Recipients are kept; 'Set up email alerts' turns them on again.");
}

function sendTestAlert() {
  var r = sendDigest({ test: true });
  SpreadsheetApp.getUi().alert(r.sent ? "Test alert sent to " + r.to + "." : r.reason);
}

function removeAlertTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === "sendDigest") ScriptApp.deleteTrigger(t); });
}

/** Builds and sends the alert email. Returns { sent, to, subject } (used by the test). */
function sendDigest(opts) {
  opts = opts && opts.test !== undefined ? opts : {};
  var p = PropertiesService.getScriptProperties();
  var to = p.getProperty("ALERT_TO");
  if (!to) return { sent: false, reason: "No recipients yet — use Purchase Dashboard ▸ Set up email alerts." };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = buildData_();
  var E = PurchaseEngine, R = E.toDay(today_(ss));
  var lines = E.enrich(data, R), A = E.analyze(lines, data, R);
  var d = digest_(A, lines, data, R, p.getProperty("DASHBOARD_URL"), !!opts.test);
  MailApp.sendEmail({ to: to, subject: d.subject, htmlBody: d.html, name: "Procurement Control Room" });
  return { sent: true, to: to, subject: d.subject, html: d.html };
}

/* ---------------------------------------------------------------- email content */
function digest_(A, lines, data, R, url, isTest) {
  var E = PurchaseEngine, k = A.kpis, S = A.settings;
  var esc = function (s) { return String(s === null || s === undefined ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); };
  var tk = function (x) { var v = x || 0, a = Math.abs(v); return a >= 1e7 ? "৳" + (v / 1e7).toFixed(2) + " Cr" : a >= 1e5 ? "৳" + (v / 1e5).toFixed(2) + " L" : "৳" + E.fmt0(v); };
  var dt = function (day) { var x = E.dayToDate(day); return day === null ? "—" : x.getUTCDate() + " " + ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][x.getUTCMonth()]; };
  var COL = { critical: ["#d03b3b", "#fbe3e3", "#8f1d1d", "Critical"], serious: ["#ec835a", "#fde9e0", "#8a3412", "At risk"], warning: ["#fab219", "#fff4d6", "#6b4a00", "Watch"], good: ["#0ca30c", "#e3f4e3", "#075c07", "On track"], info: ["#2a78d6", "#e2edf9", "#1c4f8a", "Info"] };
  var tag = function (st, text) { var c = COL[st] || COL.info; return '<span style="display:inline-block;padding:1px 8px;border-radius:10px;background:' + c[1] + ';color:' + c[2] + ';font-weight:700;font-size:12px;white-space:nowrap">' + esc(text || c[3]) + "</span>"; };
  var th = function (cols) { return "<tr>" + cols.map(function (c) { return '<th style="text-align:' + (c[1] || "left") + ';font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#44504d;padding:6px 8px;border-bottom:1px solid #d9e1df;background:#f5f8f7">' + esc(c[0]) + "</th>"; }).join("") + "</tr>"; };
  var td = function (v, al) { return '<td style="padding:7px 8px;border-bottom:1px solid #e6ecea;font-size:13px;text-align:' + (al || "left") + ';vertical-align:top">' + v + "</td>"; };
  var section = function (title, sub, rows, head) {
    if (!rows.length) return "";
    return '<h3 style="font:700 15px Arial,sans-serif;color:#0f1b19;margin:22px 0 4px">' + esc(title) + '</h3><div style="font-size:12px;color:#6f7b78;margin-bottom:6px">' + esc(sub) + "</div>" +
      '<table role="presentation" cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e1df">' + th(head) + rows.join("") + "</table>";
  };
  var hs = A.health >= 80 ? "good" : A.health >= 65 ? "warning" : A.health >= 50 ? "serious" : "critical";
  var bad = A.signals.filter(function (s) { return s.status === "critical" || s.status === "serious" || s.status === "warning"; })
    .sort(function (a, b) { var r = { critical: 3, serious: 2, warning: 1 }; return r[b.status] - r[a.status]; });
  var crit = A.signals.filter(function (s) { return s.status === "critical"; }).length;

  var late = A.open.filter(function (l) { return (l.delay || 0) > 0; });
  var soon = A.open.filter(function (l) { return !(l.delay > 0) && l.days_to_commit !== null && l.days_to_commit >= 0 && l.days_to_commit <= 7; });
  var overdue = lines.filter(function (l) { return (l.days_overdue || 0) > 0; }).sort(function (a, b) { return b.days_overdue - a.days_overdue; });
  var payNext = A.controls.plan.filter(function (x) { return x.day > R && x.day <= R + 7; });
  var waiting = A.waiting.filter(function (l) { return l.wait_days !== null && l.wait_days > (l.stage === "Awaiting Approval" ? S.sla_appr : S.sla_po); });
  var ctrl = lines.filter(function (l) { return l.ctrl_action; }).sort(function (a, b) { return ({ High: 1, Medium: 2, Low: 3 })[a.ctrl_priority] - ({ High: 1, Medium: 2, Low: 3 })[b.ctrl_priority]; });
  var high = A.actions.filter(function (l) { return l.all_priority === "High"; });
  var MAX = 10, more = function (arr) { return arr.length > MAX ? '<div style="font-size:12px;color:#6f7b78;margin-top:4px">+' + (arr.length - MAX) + " more on the dashboard</div>" : ""; };

  var html = '<div style="background:#edf1f0;padding:20px 0;font-family:Arial,Helvetica,sans-serif;color:#0f1b19">' +
    '<div style="max-width:700px;margin:0 auto;background:#ffffff;border:1px solid #d9e1df;border-radius:10px;overflow:hidden">' +
    '<div style="background:#0b2a2e;color:#e8f2f1;padding:18px 22px">' +
      '<div style="font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:#e0954d;font-weight:700">' + esc(data.meta.company || "") + (isTest ? " · TEST" : "") + "</div>" +
      '<div style="font:800 21px Arial,sans-serif;margin-top:2px">Procurement alerts · ' + dt(R) + " " + String(E.dayToDate(R).getUTCFullYear()) + "</div>" +
      '<div style="font-size:13px;color:#9fb9b7;margin-top:4px">' + crit + " critical signal(s) · " + k.high + " high-priority action(s) · " + tk(k.overdue) + " overdue to suppliers</div></div>" +
    '<div style="padding:4px 22px 22px">' +
      '<table role="presentation" style="width:100%;margin-top:16px;border-collapse:collapse"><tr>' +
        '<td style="width:120px;text-align:center;vertical-align:middle;padding:10px;border:1px solid #d9e1df;border-radius:8px">' +
          '<div style="font:800 40px Arial,sans-serif;color:' + COL[hs][0] + '">' + A.health + '</div><div style="font-size:11px;color:#6f7b78">health / 100</div><div style="margin-top:4px">' + tag(hs, { good: "Healthy", warning: "Watch", serious: "At risk", critical: "Critical" }[hs]) + "</div></td>" +
        '<td style="padding:0 0 0 14px;font-size:13px;line-height:1.7;vertical-align:middle">' +
          "Committed <b>" + tk(k.committed) + "</b> on <b>" + k.pos + "</b> POs · paid <b>" + Math.round(k.pctPaid * 100) + "%</b><br>" +
          "Outstanding <b>" + tk(k.outstanding) + "</b> · overdue <b>" + tk(k.overdue) + "</b><br>" +
          "OTIF <b>" + (k.otif === null ? "—" : Math.round(k.otif * 100) + "%") + "</b> · open POs <b>" + k.openPOs + "</b> (" + k.lateOpen + " late) · still to come in-house <b>" + tk(k.pendingVal) + "</b><br>" +
          "Sourcing compliance <b>" + (A.controls.quoteCompliance === null ? "—" : Math.round(A.controls.quoteCompliance * 100) + "%") + "</b> · invoice mismatches <b>" + A.controls.mismatches.length + "</b></td></tr></table>" +
      section("Risk signals needing attention", "From the 13 signals on the dashboard", bad.map(function (s) { return "<tr>" + td(tag(s.status)) + td("<b>" + esc(s.area) + "</b><br><span style='color:#6f7b78;font-size:12px'>" + esc(s.note) + "</span>") + td(esc(s.metric), "right") + "</tr>"; }), [["Status"], ["Area"], ["Now", "right"]]) +
      section("Late purchase orders", "Past the supplier's committed delivery date", late.slice(0, MAX).map(function (l) { return "<tr>" + td(tag("critical", "Late " + l.delay + " d")) + td("<b>" + esc(l.po_no) + "</b> · " + esc(l.supplier) + "<br><span style='color:#6f7b78;font-size:12px'>" + esc(l.item) + "</span>") + td(Math.round((l.recv_pct || 0) * 100) + "% received", "right") + "</tr>"; }), [["Status"], ["PO · supplier"], ["Received", "right"]]) + more(late) +
      section("Deliveries due in the next 7 days", "Confirm with the supplier", soon.slice(0, MAX).map(function (l) { return "<tr>" + td(tag("warning", l.days_to_commit === 0 ? "Today" : "In " + l.days_to_commit + " d")) + td("<b>" + esc(l.po_no) + "</b> · " + esc(l.supplier) + "<br><span style='color:#6f7b78;font-size:12px'>" + esc(l.item) + "</span>") + td(dt(l.commit_date), "right") + "</tr>"; }), [["When"], ["PO · supplier"], ["Committed", "right"]]) + more(soon) +
      section("Overdue supplier payments", "Outstanding past the due date", overdue.slice(0, MAX).map(function (l) { return "<tr>" + td(tag(l.days_overdue > 30 ? "critical" : "warning", l.days_overdue + " d overdue")) + td("<b>" + esc(l.supplier) + "</b><br><span style='color:#6f7b78;font-size:12px'>" + esc(l.po_no) + " · " + esc(l.item) + "</span>") + td("<b>" + tk(l.outstanding) + "</b>", "right") + "</tr>"; }), [["Status"], ["Supplier"], ["Amount", "right"]]) + more(overdue) +
      section("Payments due in the next 7 days", "From the payment schedule", payNext.slice(0, MAX).map(function (x) { return "<tr>" + td(tag("info", dt(x.day))) + td("<b>" + esc(x.line.supplier) + "</b><br><span style='color:#6f7b78;font-size:12px'>" + esc(x.line.po_no) + " · " + esc(x.source) + "</span>") + td("<b>" + tk(x.amount) + "</b>", "right") + "</tr>"; }), [["Due"], ["Supplier"], ["Amount", "right"]]) + more(payNext) +
      section("Waiting beyond the service level", "Approval SLA " + S.sla_appr + " d · PO issue SLA " + S.sla_po + " d", waiting.slice(0, MAX).map(function (l) { return "<tr>" + td(tag(l.wait_days > 2 * (l.stage === "Awaiting Approval" ? S.sla_appr : S.sla_po) ? "serious" : "warning", l.wait_days + " d")) + td("<b>" + esc(l.pr_no) + "</b> · " + esc(l.item) + "<br><span style='color:#6f7b78;font-size:12px'>" + esc(l.dept || "") + "</span>") + td(l.stage === "Awaiting Approval" ? "Approval" : "Issue PO", "right") + "</tr>"; }), [["Waiting"], ["PR · item"], ["Step", "right"]]) + more(waiting) +
      section("Invoice matching and sourcing", "Three-way match and quotation checks", ctrl.slice(0, MAX).map(function (l) { return "<tr>" + td(tag({ High: "critical", Medium: "serious", Low: "warning" }[l.ctrl_priority], l.ctrl_priority)) + td("<b>" + esc(l.pr_no) + "</b> · " + esc(l.supplier || "") + "<br><span style='color:#6f7b78;font-size:12px'>" + esc(l.ctrl_action.replace(/; $/, "")) + "</span>") + "</tr>"; }), [["Priority"], ["Line · what to do"]]) + more(ctrl) +
      section("High-priority actions", "Everything flagged High on the Action Center", high.slice(0, MAX).map(function (l) { return "<tr>" + td("<b>" + esc(l.pr_no) + "</b><br><span style='color:#6f7b78;font-size:12px'>" + esc(l.supplier || "—") + "</span>") + td(esc(l.all_action.replace(/; $/, "")).split("; ").join("<br>")) + "</tr>"; }), [["Line"], ["What to do"]]) + more(high) +
      (url ? '<div style="margin-top:22px"><a href="' + esc(url) + '" style="display:inline-block;background:#0f5c63;color:#ffffff;text-decoration:none;font-weight:700;padding:11px 18px;border-radius:8px">Open the dashboard</a></div>' : "") +
    '</div><div style="padding:12px 22px;background:#f5f8f7;border-top:1px solid #d9e1df;font-size:11.5px;color:#6f7b78">Sent by the Purchase Dashboard Google Sheet. Change recipients or frequency: Purchase Dashboard ▸ Set up email alerts.<div style="margin-top:6px"><a href="https://guulba.com" style="color:#0f5c63;text-decoration:none;font-weight:700">guulba</a> · Applied AI for the work that matters.</div></div></div></div>';
  var subject = "Procurement alerts — " + crit + " critical · " + k.high + " high-priority · " + dt(R) + (isTest ? " (test)" : "");
  return { subject: subject, html: html };
}
