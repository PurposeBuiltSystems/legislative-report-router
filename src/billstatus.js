/*
 * Bill status, derived from the legislature's own feed.
 *
 * WHAT THIS CAN AND CANNOT TELL YOU
 *
 * The Iowa feed never says a bill died. Across 4,224 real entries there are
 * zero mentions of "funnel", "died", "failed" or "withdrawn". What it does
 * carry is the opposite: 364 entries record a disposition of
 * "Effective date: 07/01/2025", meaning enacted, and 1,323 carry
 * "(Formerly HSB 52.)", recording that a study bill became a numbered bill.
 *
 * Note the 364. A loose match on "effective date" finds 948 - because 627
 * descriptions contain ordinary drafting language, "...and including effective
 * date provisions", which says only that the bill CONTAINS such a clause. That
 * mistake would have declared 584 bills law that never passed. The colon and
 * the date are what distinguish a disposition from a description.
 *
 * So death is inferred from absence plus time, never from a statement:
 *
 *   - A study bill (HSB/SSB) that never appears as the "Formerly" of a
 *     numbered bill did not advance. No study bill in the feed carries an
 *     enactment disposition, so advancing is the only survival they have.
 *     263 of 1,069 never did.
 *   - A numbered bill (HF/SF) with no effective date, once the session has
 *     ended, did not become law.
 *
 * WHICH funnel a bill died in is NOT derivable here: the feed carries no
 * funnel dates and no committee actions. A status of "Did not advance" is
 * therefore exactly that, and deliberately not labelled "died in funnel" —
 * it would be a more useful claim than the evidence supports.
 *
 * Nothing is marked dead without a session-end date to justify it. While a
 * session is running, a bill with no outcome yet is simply Active.
 */
(function (root) {
  "use strict";

  var STATUS = {
    ENACTED: "Enacted",
    VETOED: "Vetoed",
    ADVANCED: "Advanced",          // study bill → numbered bill
    NOT_ADVANCED: "Did not advance",
    NOT_ENACTED: "Did not pass",
    ACTIVE: "Active",
    UNKNOWN: "Unknown",
  };

  function norm(bill) {
    return String(bill || "").toUpperCase().replace(/[\s.]/g, "");
  }

  function isStudyBill(bill) {
    return /^(HSB|SSB)\d+$/.test(norm(bill));
  }

  /**
   * Index the feed once, so a tracker of any size costs one pass.
   * entries: [{bill, description}] — the shape LrrFeed already produces.
   */
  function buildIndex(entries) {
    var idx = { successorOf: {}, enacted: {}, vetoed: {}, seen: {}, title: {} };
    (entries || []).forEach(function (e) {
      var bill = norm(e.bill);
      if (!bill) { return; }
      var desc = String(e.description || "");
      idx.seen[bill] = true;
      if (e.title || desc) { idx.title[bill] = e.title || desc.slice(0, 160); }
      // MUST be "Effective date: 07/01/2025" - the disposition the clerk
      // appends once a bill is enacted. A loose /effective date/i match also
      // hits ordinary drafting language: "...and including effective date
      // provisions" appears in 627 of these descriptions and means only that
      // the bill CONTAINS such a clause. Measured against the real feed, the
      // loose form matched 948 entries where only 364 were enacted - 584
      // bills wrongly declared law.
      if (/Effective date:\s*\d/.test(desc)) { idx.enacted[bill] = true; }
      if (/\bvetoed\b/i.test(desc)) { idx.vetoed[bill] = true; }
      // "(Formerly HSB 52.)" — the study bill this one grew out of.
      var m = /\(Formerly\s+([A-Z]{2,4}\s?\d+)/i.exec(desc);
      if (m) { idx.successorOf[norm(m[1])] = bill; }
    });
    return idx;
  }

  /**
   * opts.sessionEnd — ISO date the session adjourned. Without it nothing is
   * called dead, because "we haven't seen it pass" and "it failed" are only
   * the same statement after the session is over.
   * opts.now — for testing.
   */
  function statusFor(bill, idx, opts) {
    opts = opts || {};
    var b = norm(bill);
    if (!b) { return { status: STATUS.UNKNOWN, detail: "No bill number." }; }

    if (idx.vetoed[b]) {
      return { status: STATUS.VETOED, detail: "Vetoed by the Governor." };
    }
    if (idx.enacted[b]) {
      return { status: STATUS.ENACTED, detail: "Enacted — the feed carries an effective date." };
    }

    var successor = idx.successorOf[b] || "";
    if (successor) {
      var s = statusFor(successor, idx, opts);
      // A study bill's fate is its successor's fate; say both.
      return {
        status: s.status === STATUS.ACTIVE ? STATUS.ADVANCED : s.status,
        successor: successor,
        detail: "Became " + successor + (s.status === STATUS.ACTIVE ? "." : " — " + s.detail.toLowerCase()),
      };
    }

    var ended = sessionHasEnded(opts);
    if (!ended) {
      return {
        status: STATUS.ACTIVE,
        detail: idx.seen[b] ? "No outcome recorded yet; the session is still running."
                            : "Not found in the feed yet.",
      };
    }
    if (!idx.seen[b]) {
      return { status: STATUS.UNKNOWN, detail: "Not in the feed — check the bill number." };
    }
    return isStudyBill(b)
      ? { status: STATUS.NOT_ADVANCED,
          detail: "Never became a numbered bill before the session ended." }
      : { status: STATUS.NOT_ENACTED,
          detail: "No effective date by the end of the session." };
  }

  function sessionHasEnded(opts) {
    if (!opts.sessionEnd) { return false; }
    var end = new Date(String(opts.sessionEnd) + (String(opts.sessionEnd).length === 10 ? "T23:59:59" : ""));
    if (isNaN(end)) { return false; }
    var now = opts.now instanceof Date ? opts.now : new Date();
    return now.getTime() > end.getTime();
  }

  /** True when a status means the bill can no longer cost anyone anything. */
  function isClosed(status) {
    return status === STATUS.NOT_ADVANCED || status === STATUS.NOT_ENACTED ||
           status === STATUS.VETOED;
  }

  /**
   * Summarise a tracker's worth of bills. Separates money still at stake from
   * money attached to bills that are over — mixing them is how a budget
   * request ends up defended with numbers that include dead bills.
   */
  function summarize(bills, idx, opts) {
    var out = { active: [], closed: [], enacted: [], unknown: [], byStatus: {} };
    (bills || []).forEach(function (b) {
      var r = statusFor(b, idx, opts);
      r.bill = norm(b);
      out.byStatus[r.status] = (out.byStatus[r.status] || 0) + 1;
      if (r.status === STATUS.ENACTED) { out.enacted.push(r); }
      else if (isClosed(r.status)) { out.closed.push(r); }
      else if (r.status === STATUS.UNKNOWN) { out.unknown.push(r); }
      else { out.active.push(r); }
    });
    return out;
  }

  var api = {
    STATUS: STATUS,
    buildIndex: buildIndex,
    statusFor: statusFor,
    summarize: summarize,
    isClosed: isClosed,
    isStudyBill: isStudyBill,
    _internals: { norm: norm, sessionHasEnded: sessionHasEnded },
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrBillStatus = api; }
})(typeof self !== "undefined" ? self : this);
