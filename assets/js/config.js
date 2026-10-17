/**
 * Dashboard configuration — the only file you normally edit.
 */
window.DASHBOARD_CONFIG = {
  company: "Panda Shoes Industry Ltd",
  title: "Procurement Control Room",
  preparedBy: "Prepared by Abdulla, Manufacturing Resilience, Sustainability & Transformation Division",

  /* Where the data comes from:
   *   "json"        -> data/purchase_data.json (built from the Excel file by
   *                    scripts/excel_to_json.py or the GitHub Action)
   *   "apps_script" -> live Google Sheet through the Apps Script web app
   *                    (paste its /exec URL below). If it cannot be reached the
   *                    dashboard falls back to the JSON file.
   *                    This is also what turns on saving in the Data Entry and
   *                    Admin tabs; without it they run in demo mode.           */
  dataSource: "json",
  jsonUrl: "data/purchase_data.json",
  appsScriptUrl: "",          // e.g. "https://script.google.com/macros/s/XXXX/exec"

  /* Management targets used by the risk signals and the health score. */
  targets: {
    otif: 0.90,      // on-time in-full delivery
    infull: 0.95,    // closed deliveries received in full
    qc: 0.95,        // material approval rate at QC
    priceVar: 0.02,  // acceptable net price variance vs budget price
  },

  /* Auto-refresh (minutes) when the data source is the live Google Sheet. 0 = off */
  refreshMinutes: 10,
};
