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

  /* ------------------------------------------------------------------
   * US federal holidays, computed rather than typed. Weekends and the
   * six holidays essentially every US public employer closes for are
   * automatic. The rest genuinely vary by agency — Iowa DOT and a city
   * clerk's office do not observe the same list — so those are asked
   * rather than assumed, and an agency can still add its own closures.
   * ------------------------------------------------------------------ */

  /** nth (1-based) weekday of a month; n = -1 means the last one. */
  function nthWeekday(year, month, weekday, n) {
    if (n > 0) {
      var d = new Date(year, month, 1);
      var shift = (weekday - d.getDay() + 7) % 7;
      return new Date(year, month, 1 + shift + (n - 1) * 7);
    }
    var last = new Date(year, month + 1, 0);
    var back = (last.getDay() - weekday + 7) % 7;
    return new Date(year, month, last.getDate() - back);
  }

  /** Federal rule: a fixed-date holiday on Sat is observed Fri, on Sun Mon. */
  function observed(date) {
    var d = new Date(date);
    if (d.getDay() === 6) { d.setDate(d.getDate() - 1); }
    else if (d.getDay() === 0) { d.setDate(d.getDate() + 1); }
    return d;
  }

  // Always applied — closures you can rely on across US public employers.
  var CORE = ["newYears", "memorial", "independence", "labor", "thanksgiving", "christmas"];
  // Asked, because observance genuinely differs between agencies.
  var OPTIONAL = [
    { key: "mlk", label: "Martin Luther King Jr. Day" },
    { key: "presidents", label: "Presidents Day / Washington's Birthday" },
    { key: "juneteenth", label: "Juneteenth" },
    { key: "columbus", label: "Columbus Day / Indigenous Peoples' Day" },
    { key: "veterans", label: "Veterans Day" },
    { key: "dayAfterThanksgiving", label: "Day after Thanksgiving" },
  ];

  function holidayDates(year, which) {
    var t = nthWeekday(year, 10, 4, 4); // Thanksgiving: 4th Thursday of November
    var dayAfter = new Date(year, 10, t.getDate() + 1);
    var map = {
      newYears: observed(new Date(year, 0, 1)),
      mlk: nthWeekday(year, 0, 1, 3),
      presidents: nthWeekday(year, 1, 1, 3),
      memorial: nthWeekday(year, 4, 1, -1),
      juneteenth: observed(new Date(year, 5, 19)),
      independence: observed(new Date(year, 6, 4)),
      labor: nthWeekday(year, 8, 1, 1),
      columbus: nthWeekday(year, 9, 1, 2),
      veterans: observed(new Date(year, 10, 11)),
      thanksgiving: t,
      dayAfterThanksgiving: dayAfter,
      christmas: observed(new Date(year, 11, 25)),
    };
    var out = {};
    Object.keys(map).forEach(function (k) {
      if (which && which[k] === false) { return; }
      if (which && which[k] === true) { out[ymd(map[k])] = k; return; }
      if (CORE.indexOf(k) !== -1) { out[ymd(map[k])] = k; }
    });
    return out;
  }

  /**
   * The holiday set actually used: computed federal holidays for the years
   * in play, plus any extra dates an agency pastes in.
   * opts: {observe:{mlk:true,…}, extra:"2026-12-24, …", years:[…]}
   */
  function holidaySet(opts) {
    opts = opts || {};
    var years = opts.years && opts.years.length ? opts.years
      : [new Date().getFullYear(), new Date().getFullYear() + 1];
    var out = {};
    years.forEach(function (y) {
      var h = holidayDates(y, opts.observe);
      Object.keys(h).forEach(function (d) { out[d] = true; });
    });
    Object.keys(parseHolidays(opts.extra)).forEach(function (d) { out[d] = true; });
    return out;
  }

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

  /**
   * DOT convention (confirmed 2026-07-31): comments are due END OF BUSINESS
   * — the business-hours math picks WHICH day, then the clock snaps to the
   * due time (default 5:00 PM). Friday 3 PM + 48 = Tuesday 5:00 PM.
   * Pass a falsy dueTime to keep same-clock-time semantics.
   */
  function deadlineFor(start, hours, holidays, dueTime) {
    var d = addBusinessHours(start, hours, holidays);
    var m = /^(\d{1,2}):(\d{2})$/.exec(String(dueTime || "").trim());
    if (!m) { return d; }
    var snapped = new Date(d);
    snapped.setHours(Number(m[1]), Number(m[2]), 0, 0);
    // Snapping to end-of-business must never SHORTEN the window: a 6 PM
    // publish with a short response window would otherwise come out "due"
    // at 5 PM the same day — a deadline already in the past.
    var guard = 0;
    while ((snapped.getTime() < d.getTime() || !isBusinessDay(snapped, holidays)) && guard++ < 400) {
      snapped.setDate(snapped.getDate() + 1);
      snapped.setHours(Number(m[1]), Number(m[2]), 0, 0);
    }
    return snapped;
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
    deadlineFor: deadlineFor,
    nthWeekday: nthWeekday,
    observed: observed,
    holidayDates: holidayDates,
    holidaySet: holidaySet,
    OPTIONAL_HOLIDAYS: OPTIONAL,
    ymd: ymd, // LOCAL calendar date — toISOString() would shift the day west of UTC
    parseHolidays: parseHolidays,
    isBusinessDay: isBusinessDay,
    addBusinessHours: addBusinessHours,
    formatDeadline: formatDeadline,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrDeadline = api; }
})(typeof self !== "undefined" ? self : this);
