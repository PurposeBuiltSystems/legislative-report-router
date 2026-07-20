/*
 * Legislative Report Router — Iowa Code chapter/section tracking (pure).
 *
 * Bill briefs constantly reference the Code ("Amends Code Chapter 9.",
 * "for purposes of Code chapters 6A and 6B", "Creates new language in
 * Code 307."). This module extracts those references and matches them
 * against the agency's tracked-chapters list (seeded from the 2015 DOT
 * list — EDITABLE in Settings, because the list needs periodic review).
 *
 * Tracked entries may be: a chapter ("321J"), an exact section
 * ("232.52"), or a section range ("331.301-331.440").
 */
(function (root) {
  "use strict";

  // 2015 DOT tracking list (chapters/sections DOT followed in the Iowa
  // Acts). Starting point only — Settings overrides.
  var DEFAULT_TRACKED = [
    "6A", "6B", "8C", "9E", "17A", "25", "28L", "28M", "28N", "29C", "72",
    "73", "73A", "161E", "124D", "232.52", "263B", "306", "306A", "306B",
    "306C", "306D", "307", "307A", "307C", "308", "308A", "309", "310",
    "311", "312", "312A", "313", "313A", "314", "315", "316", "317", "318",
    "320", "321", "321A", "321C", "321D", "321E", "321F", "321G", "321H",
    "321I", "321J", "321K", "321L", "321M", "322", "322A", "322C", "322D",
    "322G", "324A", "325A", "326", "327B", "327C", "327D", "327E", "327F",
    "327G", "327H", "327J", "328", "329", "330", "330A", "331.301-331.440",
    "331.551-331.600", "354", "355", "362", "364", "384.37-384.79",
    "384.95-384.109", "423.26", "423.26A", "423.40", "435.26B", "452A",
    "460", "461A", "465B", "468.335-468.354", "468.600-468.634", "480",
    "573", "573A", "589", "613", "614", "657", "668", "669", "670",
    "707.6A", "714", "801", "804", "805", "808", "809", "809A",
  ];

  var TOKEN_RE = /^\d{1,3}[A-Z]{0,2}(\.\d+[A-Z]{0,2})?$/;

  function normToken(t) {
    return String(t || "").toUpperCase().replace(/\s+/g, "").replace(/[.,;]$/, "");
  }

  /**
   * Extract Code chapter/section references from brief text.
   * Catches "Code chapter 22", "chapters 6A and 6B", "Code 307",
   * "section 321.276", "Amends 6A." — but never bill numbers (those
   * carry letter prefixes like HF/SSB, which TOKEN_RE rejects).
   */
  function extractChapters(text) {
    var out = [];
    function add(tok) {
      var t = normToken(tok);
      if (TOKEN_RE.test(t) && out.indexOf(t) === -1) { out.push(t); }
    }
    var s = String(text || "");
    // keyword followed by an enumeration: chapter(s)/section(s)/Code X
    var re = /\b(?:code\s+)?(?:chapters?|sections?)\s+([0-9][0-9A-Z.]*(?:\s*(?:,|and|&)\s*[0-9][0-9A-Z.]*)*)/gi;
    var m;
    while ((m = re.exec(s)) !== null) {
      m[1].split(/\s*(?:,|and|&)\s*/i).forEach(add);
    }
    // bare "Code 307" / "in Code 321J"
    var re2 = /\bcode\s+([0-9][0-9A-Z.]*)/gi;
    while ((m = re2.exec(s)) !== null) { add(m[1]); }
    // "Amends 6A." / "Amends 6A and 321J" (never "Amends 2026 Iowa Acts")
    var re3 = /\bamends\s+([0-9][0-9A-Z.]*(?:\s*(?:,|and|&)\s*[0-9][0-9A-Z.]*)*)\b(?!\s*iowa\s+acts)/gi;
    while ((m = re3.exec(s)) !== null) {
      m[1].split(/\s*(?:,|and|&)\s*/i).forEach(function (tok) {
        if (!/^\d{4}$/.test(normToken(tok))) { add(tok); } // skip years
      });
    }
    return out;
  }

  function chapterOf(token) { return String(token).split(".")[0]; }

  function parseTrackedEntry(e) {
    var t = normToken(e);
    var range = t.match(/^(\d{1,3}[A-Z]{0,2})\.(\d+)-(?:\1\.)?(\d+)$/);
    if (range) { return { kind: "range", chapter: range[1], lo: Number(range[2]), hi: Number(range[3]) }; }
    if (t.indexOf(".") !== -1) { return { kind: "section", token: t }; }
    return { kind: "chapter", chapter: t };
  }

  /**
   * Which extracted tokens hit the tracked list. A tracked CHAPTER
   * matches the chapter itself and any of its sections; a tracked
   * SECTION matches exactly; a RANGE matches sections inside it.
   */
  function matchTracked(tokens, trackedList) {
    var tracked = (trackedList && trackedList.length ? trackedList : DEFAULT_TRACKED).map(parseTrackedEntry);
    var hits = [];
    (tokens || []).forEach(function (tok) {
      var t = normToken(tok);
      var ch = chapterOf(t);
      var sec = t.indexOf(".") !== -1 ? Number(t.split(".")[1]) : null;
      var hit = tracked.some(function (tr) {
        if (tr.kind === "chapter") { return tr.chapter === ch; }
        if (tr.kind === "section") { return tr.token === t; }
        return tr.chapter === ch && sec != null && sec >= tr.lo && sec <= tr.hi;
      });
      if (hit && hits.indexOf(t) === -1) { hits.push(t); }
    });
    return hits;
  }

  /** Parse the Settings textarea (commas/whitespace/newlines). */
  function parseTrackedText(text) {
    return String(text || "").split(/[\s,;]+/).map(normToken).filter(function (t) {
      return /^\d/.test(t);
    });
  }

  /**
   * Suggest routing rules by chapter: rules whose CodeChapters list
   * intersects the item's extracted chapters (chapter-level match).
   */
  function suggestRules(itemChapters, rules) {
    var chapters = (itemChapters || []).map(chapterOf);
    var out = [];
    (rules || []).forEach(function (r) {
      var claims = (r.codeChapters || []).map(function (c) { return chapterOf(normToken(c)); });
      var overlap = claims.filter(function (c) { return chapters.indexOf(c) !== -1; });
      if (overlap.length) { out.push({ rule: r, chapters: overlap }); }
    });
    return out;
  }

  var api = {
    DEFAULT_TRACKED: DEFAULT_TRACKED,
    extractChapters: extractChapters,
    matchTracked: matchTracked,
    parseTrackedText: parseTrackedText,
    suggestRules: suggestRules,
    _internals: { parseTrackedEntry: parseTrackedEntry, chapterOf: chapterOf },
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrChapters = api; }
})(typeof self !== "undefined" ? self : this);
