/*
 * Legislative Report Router — comment-deadline math (pure).
 *
 * DOT convention: divisions get 48 business hours to comment. "48 business
 * hours" is interpreted the way coordinators mean it — two business days
 * later at the same clock time — so a Friday 3 PM post is due Tuesday 3 PM,
 * never Sunday. Weekends always skip; state holidays skip when the
 * coordinator pastes them into Settings (comma-separated YYYY-MM-DD; blank
 * = weekends only, no code change needed when the list arrives).
 */
(function (root) {
  "use strict";

  /** "2026-09-07, 2026-11-26  2026-11-27" → {"2026-09-07":true, …} */
  function parseHolidays(text) {
    var map = {};
    String(text || "").split(/[\s,;]+/).forEach(function (tok) {
      if (/^\d{4}-\d{2}-\d{2}$/.test(tok)) { map[tok] = true; }
    });
    return map;
  }

  function ymd(d) {
    var p = function (n) { return (n < 10 ? "0" : "") + n; };
    return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate());
  }

  function isBusinessDay(d, holidays) {
    var dow = d.getDay();
    if (dow === 0 || dow === 6) { return false; }
    return !(holidays && holidays[ymd(d)]);
  }

  /**
   * start + N business hours. Whole 24-hour blocks step whole business
   * days (same clock time); a remainder adds clock hours and then rolls
   * forward to the next business day if it landed on one that isn't.
   */
  function addBusinessHours(start, hours, holidays) {
    var d = new Date(start);
    var days = Math.floor((Number(hours) || 0) / 24);
    var rem = (Number(hours) || 0) % 24;
    while (days > 0) {
      d.setDate(d.getDate() + 1);
      if (isBusinessDay(d, holidays)) { days--; }
    }
    if (rem > 0) { d.setHours(d.getHours() + rem); }
    while (!isBusinessDay(d, holidays)) { d.setDate(d.getDate() + 1); }
    return d;
  }

  var DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  var MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  /** "Tue, Aug 4, 3:00 PM" — fixed format, no locale surprises. */
  function formatDeadline(d) {
    var h = d.getHours();
    var ampm = h >= 12 ? "PM" : "AM";
    var h12 = h % 12 || 12;
    var m = d.getMinutes();
    return DAYS[d.getDay()] + ", " + MONTHS[d.getMonth()] + " " + d.getDate() +
      ", " + h12 + ":" + (m < 10 ? "0" : "") + m + " " + ampm;
  }

  var api = {
    parseHolidays: parseHolidays,
    isBusinessDay: isBusinessDay,
    addBusinessHours: addBusinessHours,
    formatDeadline: formatDeadline,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrDeadline = api; }
})(typeof self !== "undefined" ? self : this);
