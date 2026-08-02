/*
 * Legislative Report Router — fiscal-impact rollup (pure).
 *
 * Divisions report estimated costs and impact severity on tracker rows
 * (they edit the BillTracker list directly — the list IS the form). This
 * module turns those rows into running totals per fiscal year for budget
 * requests: totals by FY, by division, and by severity, plus a CSV.
 *
 * Iowa SFY runs July–June and is named by its ENDING year (SFY2027 =
 * Jul 2026–Jun 2027). A bill's default FY assumes the standard July 1
 * effective date after the session: published Jan–Jun of year Y →
 * effective Jul 1 Y → SFY(Y+1); published Jul–Dec → SFY(Y+2). Divisions
 * can overwrite the FY on any row — the default is a starting point.
 */
(function (root) {
  "use strict";

  function defaultFy(date) {
    var d = date instanceof Date ? date : new Date(date);
    var y = d.getFullYear();
    return "SFY" + (d.getMonth() <= 5 ? y + 1 : y + 2);
  }

  /**
   * "$1.2M" → 1200000 · "250,000" → 250000 · "300k" → 300000 ·
   * "1.5 million" → 1500000 · "(50,000)" / "-50000" → -50000 (savings) ·
   * "", "unknown", "TBD" → null (counted separately, never as zero).
   */
  function parseCost(text) {
    if (typeof text === "number") { return isNaN(text) ? null : text; }
    var s = String(text == null ? "" : text).trim().toLowerCase();
    if (!s || /^(unknown|tbd|n\/?a|none|\?+)$/.test(s)) { return null; }
    var negative = /^\(.*\)$/.test(s) || /^-/.test(s) || /saving/.test(s);
    // The magnitude suffix MUST end on a word boundary: without it the bare
    // letters k/m/b swallow the first letter of the next word, so
    // "50000 max" parsed as 50000 × 1e6. Costs are typed by humans into a
    // free-text column and feed the Power BI totals — silent 1000× errors
    // are the worst possible failure here.
    var m = /([0-9][0-9,]*\.?[0-9]*)\s*(?:(k|m|b|thousand|million|billion)\b)?/.exec(s.replace(/[$()]/g, ""));
    if (!m) { return null; }
    var n = Number(m[1].replace(/,/g, ""));
    if (isNaN(n)) { return null; }
    var mult = { k: 1e3, thousand: 1e3, m: 1e6, million: 1e6, b: 1e9, billion: 1e9 }[m[2]] || 1;
    return (negative ? -1 : 1) * n * mult;
  }

  function fmtMoney(n) {
    var sign = n < 0 ? "−" : "";
    var v = Math.abs(Math.round(n));
    return sign + "$" + String(v).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  }

  /**
   * rows: [{fy, division, bill, severity, cost(raw), notes}] straight off
   * the tracker. Returns FYs newest first, each with totals, per-division
   * totals, severity counts, and how many rows had no usable estimate.
   */
  function aggregate(rows) {
    var byFy = {};
    (rows || []).forEach(function (r) {
      var fy = r.fy || "(no FY)";
      var f = byFy[fy] || (byFy[fy] = {
        fy: fy, total: 0, estimated: 0, unestimated: 0,
        byDivision: {}, bySeverity: {}, bills: {},
      });
      var cost = parseCost(r.cost);
      if (cost === null) { f.unestimated++; }
      else {
        f.estimated++;
        f.total += cost;
        var div = r.division || "(none)";
        f.byDivision[div] = (f.byDivision[div] || 0) + cost;
      }
      var sev = (r.severity || "Unknown");
      f.bySeverity[sev] = (f.bySeverity[sev] || 0) + 1;
      if (r.bill) { f.bills[r.bill] = true; }
    });
    return Object.keys(byFy).sort().reverse().map(function (k) {
      var f = byFy[k];
      f.billCount = Object.keys(f.bills).length;
      delete f.bills;
      return f;
    });
  }

  function csvEscape(v) {
    var s = String(v == null ? "" : v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  /** Detail rows plus a TOTAL line per FY — pivots cleanly in Excel. */
  function rollupCsv(rows) {
    var header = ["fiscal_year", "division", "bill", "severity", "estimated_cost", "notes"];
    var lines = [header.join(",")];
    var sorted = (rows || []).slice().sort(function (a, b) {
      return String(b.fy).localeCompare(String(a.fy)) ||
        String(a.division || "").localeCompare(String(b.division || "")) ||
        String(a.bill || "").localeCompare(String(b.bill || ""));
    });
    var fyTotals = {};
    sorted.forEach(function (r) {
      var cost = parseCost(r.cost);
      if (cost !== null) { fyTotals[r.fy || "(no FY)"] = (fyTotals[r.fy || "(no FY)"] || 0) + cost; }
      lines.push([r.fy, r.division, r.bill, r.severity || "Unknown",
        cost === null ? "" : cost, r.notes || ""].map(csvEscape).join(","));
    });
    Object.keys(fyTotals).sort().reverse().forEach(function (fy) {
      lines.push([fy, "TOTAL", "", "", fyTotals[fy], ""].map(csvEscape).join(","));
    });
    return lines.join("\r\n") + "\r\n";
  }

  var api = {
    defaultFy: defaultFy,
    parseCost: parseCost,
    fmtMoney: fmtMoney,
    aggregate: aggregate,
    rollupCsv: rollupCsv,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrFiscal = api; }
})(typeof self !== "undefined" ? self : this);
