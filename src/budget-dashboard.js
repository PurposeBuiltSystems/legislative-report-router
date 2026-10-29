/* Budget dashboard rendering shared by the router and native SharePoint page. */
/* global LrrFiscal */
(function (root) {
  "use strict";
  var Fiscal = typeof module !== "undefined" && module.exports ? require("./fiscal") : LrrFiscal;
  function esc(value) { return String(value == null ? "" : value).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function cost(row) { var n = Fiscal.parseCost(row.cost); return Number.isFinite(n) ? n : null; }
  function model(rows, filters) {
    filters = filters || {};
    var selected = (rows || []).filter(function (r) {
      return (!filters.fy || (r.fy || "(no FY)") === filters.fy) && (!filters.division || (r.division || "(none)") === filters.division);
    });
    var out = { rows: selected, total: 0, costs: 0, savings: 0, pending: 0, estimated: 0, divisions: Object.create(null), severity: Object.create(null), bills: new Set(), closed: 0, unchecked: 0, duplicates: 0 };
    var seen = new Set();
    selected.forEach(function (r) {
      var n = cost(r), division = r.division || "(none)", severity = r.severity || "Unknown";
      var key = [r.fy, r.division, r.bill].map(function (v) { return String(v || "").toUpperCase().trim(); }).join("|");
      if (seen.has(key)) { out.duplicates++; } seen.add(key);
      if (r.bill) { out.bills.add(r.bill); }
      if (r.closed == null) { out.unchecked++; }
      out.severity[severity] = (out.severity[severity] || 0) + 1;
      if (n === null) { out.pending++; return; }
      out.estimated++; out.total += n;
      if (n >= 0) { out.costs += n; } else { out.savings -= n; }
      out.divisions[division] = (out.divisions[division] || 0) + n;
      if (r.closed === true) { out.closed += n; }
    });
    out.billCount = out.bills.size; delete out.bills;
    return out;
  }
  function bars(values, money) {
    var keys = Object.keys(values).sort(function (a, b) { return Math.abs(values[b]) - Math.abs(values[a]); });
    var maximum = Math.max.apply(null, keys.map(function (k) { return Math.abs(values[k]); }).concat([1]));
    return '<table style="width:100%;border-collapse:collapse"><thead><tr><th scope="col">' + (money ? "Division" : "Impact severity") + '</th><th scope="col">' + (money ? "Net estimate" : "Tracker rows") + '</th><th scope="col">Magnitude</th></tr></thead><tbody>' + keys.map(function (key) {
      var value = values[key], width = Math.round(Math.abs(value) / maximum * 100);
      return '<tr><th scope="row" style="text-align:left;padding:6px">' + esc(key) + '</th><td style="padding:6px;white-space:nowrap">' + esc(money ? Fiscal.fmtMoney(value) : value) + '</td><td style="width:45%;padding:6px"><div style="height:14px;width:' + width + '%;background-color:' + (value < 0 ? "#168463" : "#2563a6") + '"></div></td></tr>';
    }).join("") + '</tbody></table>';
  }
  function html(rows, options) {
    options = options || {};
    var m = model(rows, options.filters);
    var cards = [["Net estimated cost", Fiscal.fmtMoney(m.total)], ["Estimated spending", Fiscal.fmtMoney(m.costs)], ["Estimated savings", Fiscal.fmtMoney(m.savings)], ["Awaiting estimates", m.pending]];
    var out = '<p><strong>BillTracker budget overview</strong> · ' + m.billCount + ' bill(s), ' + m.rows.length + ' division/FY row(s). ' + esc(options.label || "All fiscal years and divisions") + '</p>';
    if (options.asOf) { out += '<p>Data loaded: ' + esc(options.asOf) + '. Snapshot; not an automatic live feed.</p>'; }
    out += '<table style="width:100%;border-collapse:separate;border-spacing:8px"><tbody><tr>' + cards.map(function (card, index) {
      return '<td style="padding:14px;background-color:#edf3fa;vertical-align:top"><strong>' + esc(card[0]) + '</strong><p style="font-size:24px;color:#164471"><strong>' + esc(card[1]) + '</strong></p></td>';
    }).join("") + '</tr></tbody></table>';
    out += '<p>' + m.estimated + ' row(s) have estimates. Missing estimates are counted separately, never treated as $0. Totals include all tracked bills, including any that have closed.' +
      ' Estimates on known closed bills: ' + esc(Fiscal.fmtMoney(m.closed)) + '.' + (m.unchecked ? ' Bill status has not been checked for ' + m.unchecked + ' row(s).' : '') + '</p>';
    if (m.duplicates) { out += '<p><strong>Check source data:</strong> ' + m.duplicates + ' repeated bill/division/FY row(s) are included in these totals.</p>'; }
    if (!m.rows.length) { return out + '<p>No tracker rows match this view.</p>'; }
    out += '<h2>By division</h2><p>Bar length shows magnitude; green values are savings, blue values are costs.</p>' + bars(m.divisions, true);
    out += '<h2>Impact severity</h2>' + bars(m.severity, false);
    var fys = Array.from(new Set(m.rows.map(function (r) { return r.fy || "(no FY)"; }))).sort().reverse();
    out += '<h2>Fiscal years</h2><table style="width:100%"><thead><tr><th scope="col">Fiscal year</th><th scope="col">Net estimate</th><th scope="col">Estimated rows</th><th scope="col">Awaiting estimate</th></tr></thead><tbody>' + fys.map(function (fy) {
      var f = model(m.rows, { fy: fy });
      return '<tr><th scope="row">' + esc(fy) + '</th><td>' + esc(Fiscal.fmtMoney(f.total)) + '</td><td>' + f.estimated + '</td><td>' + f.pending + '</td></tr>';
    }).join("") + '</tbody></table>';
    var detail = m.rows.slice().sort(function (a, b) { return String(b.fy).localeCompare(String(a.fy)) || String(a.division).localeCompare(String(b.division)) || String(a.bill).localeCompare(String(b.bill)); });
    out += '<h2>Bill details</h2><p>Showing ' + Math.min(150, detail.length) + ' of ' + detail.length + ' row(s). Open BillTracker for the complete editable list.</p><table style="width:100%;border-collapse:collapse"><thead><tr><th scope="col">Bill</th><th scope="col">Division</th><th scope="col">FY</th><th scope="col">Estimate</th><th scope="col">Severity</th></tr></thead><tbody>' + detail.slice(0, 150).map(function (r) {
      return '<tr><td>' + esc(r.bill) + '</td><td>' + esc(r.division) + '</td><td>' + esc(r.fy) + '</td><td>' + esc(cost(r) === null ? "Awaiting estimate" : Fiscal.fmtMoney(cost(r))) + '</td><td>' + esc(r.severity || "Unknown") + '</td></tr>';
    }).join("") + '</tbody></table>';
    return out;
  }
  function page(rows, options) {
    var content = options.contentHtml || html(rows, { asOf: options.asOf });
    if (options.trackerUrl && /^https:\/\//i.test(options.trackerUrl)) { content += '<p><a href="' + esc(options.trackerUrl) + '">Open live BillTracker to edit estimates and review statuses</a></p>'; }
    return { "@odata.type": "#microsoft.graph.sitePage", name: "lrr-budget-" + String(options.trackerId).replace(/[^a-zA-Z0-9]/g, "") + ".aspx",
      title: "Legislative budget dashboard", description: "LRR_BUDGET_V1:" + options.trackerId,
      pageLayout: "article", showComments: false, showRecommendedPages: false,
      canvasLayout: { horizontalSections: [{ id: "1", layout: "oneColumn", emphasis: "none", columns: [{ id: "1", width: 12,
        webparts: [{ "@odata.type": "#microsoft.graph.textWebPart", id: "5025db0e-bc9d-4254-9fc8-b501338b4e7a", innerHtml: content }] }] }] } };
  }
  var api = { model: model, html: html, page: page };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrBudget = api; }
})(typeof self !== "undefined" ? self : this);
