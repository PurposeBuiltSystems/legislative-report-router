/*
 * Legislative Report Router — report parser (pure logic, no Office/Graph).
 *
 * Staged pipeline:
 *   HTML → normalized plain text (links preserved)
 *   → header removal → entry boundary detection (line-aware)
 *   → field extraction → division normalization → confidence scoring
 *
 * Boundaries are STANDALONE identifier lines only ("HF935" alone on a line).
 * References inside a brief ("Successor to HSB171") can never split an
 * entry, because they never stand alone on a line. Deterministic — no AI.
 *
 * Works in browser (global `LrrParser`) and Node (module.exports).
 */
(function (root) {
  "use strict";

  // ---------- configuration (overridable via opts) ----------

  var DEFAULT_IDENTIFIERS = [
    "HF", "SF", "HSB", "SSB", "HJR", "SJR", "HCR", "SCR", "HR", "SR",
  ];

  // Division-ish line: short, and made of code-like tokens (allows
  // "MVD/TDD", "ELT and MVD", "AG Office", "MVD, TDD, and SOD").
  var DIVISION_LINE_MAX = 60;

  var ENTITIES = {
    "&nbsp;": " ", "&amp;": "&", "&lt;": "<", "&gt;": ">",
    "&quot;": '"', "&#39;": "'", "&apos;": "'", "&ndash;": "-",
    "&mdash;": "-", "&rsquo;": "’", "&lsquo;": "‘",
    "&rdquo;": "”", "&ldquo;": "“",
  };

  function decodeEntities(s) {
    return String(s || "")
      .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(Number(n)); })
      .replace(/&[a-z]+;|&#\d+;/gi, function (e) { return ENTITIES[e.toLowerCase()] || e; });
  }

  /**
   * HTML → plain text with line semantics, capturing hyperlinks.
   * Returns {text, links: [{text, href}]}.
   */
  function htmlToText(html) {
    var s = String(html || "");
    var links = [];
    // capture anchors before stripping
    s = s.replace(/<a\b[^>]*href\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi,
      function (_, href, inner) {
        var text = inner.replace(/<[^>]+>/g, "").trim();
        if (href && !/^javascript:/i.test(href)) { links.push({ text: text, href: href }); }
        return text;
      });
    s = s
      .replace(/<(style|script|head)\b[\s\S]*?<\/\1>/gi, "")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(p|div|li|tr|h[1-6]|table)>/gi, "\n")
      .replace(/<(p|div|li|tr|h[1-6])\b[^>]*>/gi, "\n")
      .replace(/<td\b[^>]*>/gi, "\t")
      .replace(/<[^>]+>/g, "");
    s = decodeEntities(s);
    return { text: s, links: links };
  }

  /** CRLF/LF/entity-normalized, trimmed lines. */
  function toLines(text) {
    return String(text || "")
      .replace(/\r\n?/g, "\n")
      .replace(/ /g, " ")
      .split("\n")
      .map(function (l) { return l.replace(/[ \t]+$/g, "").replace(/^[ \t]+/g, ""); });
  }

  function identifierLineRe(identifiers) {
    // The WHOLE line must be one identifier + number (e.g. "HF935", "HF 935",
    // "HSB-171", "SF2310A"). That's what makes internal references safe.
    var alts = (identifiers || DEFAULT_IDENTIFIERS)
      .slice()
      .sort(function (a, b) { return b.length - a.length; })
      .map(function (i) { return i.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); })
      .join("|");
    return new RegExp("^(" + alts + ")\\s*-?\\s*(\\d+[A-Z]?)$", "i");
  }

  // Amendment entries: "S-5235 filed on SF2284" / "H-8437 filed on SF2478".
  // The amendment number + "filed on" + parent bill IS the boundary line.
  var AMENDMENT_LINE_RE = /^([HS])\s*-?\s*(\d{1,6})\s+filed\s+on\s+([A-Z]{1,4})\s*-?\s*(\d+[A-Z]?)\b.*$/i;

  // Column-label rows that appear mid-report ("Amendment Number",
  // "Distributed To", "Title", ...) — layout furniture, never content.
  var LABEL_LINE_RE = /^(bill\s+number|bill\s+distributed\s+to|bill\s+comment\s+requested\s+from|bill\s+brief[^:]*|amendment\s+number|distributed\s+to|comment\s+requested\s+from|title)\s*:?\s*$/i;

  function referenceRe(identifiers) {
    var alts = (identifiers || DEFAULT_IDENTIFIERS)
      .slice()
      .sort(function (a, b) { return b.length - a.length; })
      .map(function (i) { return i.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); })
      .join("|");
    return new RegExp("\\b(" + alts + ")\\s*-?\\s*(\\d+[A-Z]?)\\b", "gi");
  }

  /**
   * Split "MVD/TDD", "MVD; SOD", "ELT and MVD", "MVD, TDD, and SOD" into
   * normalized division tokens. Multi-word divisions ("AG Office") survive.
   */
  function normalizeDivisions(s) {
    return String(s || "")
      .split(/[\/;,]|\band\b|&/i)
      .map(function (t) { return t.replace(/\s+/g, " ").trim(); })
      .filter(function (t) { return t && !/^and$/i.test(t); });
  }

  /** Heuristic: does this line look like a division designation, per known codes? */
  function looksLikeDivisionLine(line, knownCodes) {
    var l = String(line || "").trim();
    if (!l || l.length > DIVISION_LINE_MAX) { return false; }
    if (/[.?!]$/.test(l) && !/\b(inc|dept|div)\.$/i.test(l)) { return false; }
    var toks = normalizeDivisions(l);
    if (!toks.length) { return false; }
    var known = (knownCodes || []).map(function (c) { return c.toLowerCase(); });
    return toks.every(function (t) {
      var tl = t.toLowerCase();
      if (known.length && (known.indexOf(tl) !== -1)) { return true; }
      // qualifier form: "ELT - awareness" — code plus a lowercase note
      var codePart = t.split(/\s*[-–]\s*/)[0];
      if (known.length && known.indexOf(codePart.toLowerCase()) !== -1) { return true; }
      // fall back: short ALL-CAPS-ish code (with optional qualifier) or
      // "<Word> Office/Bureau/Division"
      return /^[A-Z]{2,6}(\s*[-–]\s*\S.*)?$/.test(t) ||
        /^[A-Z][A-Za-z]{1,10}\s(office|bureau|division)$/i.test(t);
    });
  }

  /** Strip report header: everything before the first boundary line. */
  function headerLines(lines, idRe) {
    for (var i = 0; i < lines.length; i++) {
      if (idRe.test(lines[i])) { return i; }
    }
    return lines.length;
  }

  var SIGNATURE_RE = /^(--\s*$|best regards|regards,|thank you,|thanks,|sincerely)/i;
  var FORWARD_RE = /^(from|sent|to|subject|cc):\s/i;

  /**
   * Parse a report. opts: {identifiers, knownDivisions, reportId}
   * Returns {items, header, warnings, links}.
   */
  function parseReport(input, opts) {
    opts = opts || {};
    var idRe = identifierLineRe(opts.identifiers);
    var refRe = referenceRe(opts.identifiers);
    var known = opts.knownDivisions || [];

    var html = /<\s*(html|body|div|p|br|table)\b/i.test(String(input || ""));
    var conv = html ? htmlToText(input) : { text: String(input || ""), links: [] };
    var lines = toLines(conv.text);

    var start = headerLines(lines, idRe);
    var header = lines.slice(0, start).filter(Boolean).join("\n");
    var warnings = [];
    if (start === lines.length) {
      warnings.push("No legislative entries found. Check that bill numbers appear alone on their own line (e.g. \"HF935\").");
      return { items: [], header: header, warnings: warnings, links: conv.links };
    }

    // slice into blocks at standalone identifier lines or amendment lines
    var blocks = [];
    var cur = null;
    for (var i = start; i < lines.length; i++) {
      var line = lines[i];
      if (SIGNATURE_RE.test(line)) { break; } // signature onward: stop
      if (LABEL_LINE_RE.test(line)) { continue; } // column-label furniture
      var m = idRe.exec(line);
      var am = AMENDMENT_LINE_RE.exec(line);
      if (m) {
        if (cur) { blocks.push(cur); }
        cur = { billNumber: (m[1].toUpperCase() + m[2].toUpperCase()), documentType: m[1].toUpperCase(), lines: [], source: [line], relatedBill: "" };
      } else if (am) {
        if (cur) { blocks.push(cur); }
        cur = {
          billNumber: am[1].toUpperCase() + "-" + am[2],
          documentType: "Amendment",
          relatedBill: am[3].toUpperCase() + am[4].toUpperCase(),
          lines: [], source: [line],
        };
      } else if (cur) {
        if (FORWARD_RE.test(line) && !cur.lines.length) { continue; }
        cur.lines.push(line);
        cur.source.push(line);
      }
    }
    if (cur) { blocks.push(cur); }

    var items = blocks.map(function (b, idx) {
      var body = b.lines.slice();
      // drop leading/trailing blanks
      while (body.length && !body[0]) { body.shift(); }
      while (body.length && !body[body.length - 1]) { body.pop(); }

      var itemWarnings = [];
      var confidence = 1.0;

      var nonblank = [];
      for (var j = 0; j < body.length; j++) {
        if (body[j]) { nonblank.push({ text: body[j], idx: j }); }
      }

      var distributedTo = [];
      var commentFrom = [];
      var briefStart = 0;

      if (nonblank.length && looksLikeDivisionLine(nonblank[0].text, known)) {
        distributedTo = normalizeDivisions(nonblank[0].text);
        briefStart = nonblank[0].idx + 1;
        if (nonblank.length > 1 && looksLikeDivisionLine(nonblank[1].text, known)) {
          commentFrom = normalizeDivisions(nonblank[1].text);
          briefStart = nonblank[1].idx + 1;
        } else {
          commentFrom = distributedTo.slice();
          itemWarnings.push("Single division line found; using it for both Distributed To and Comment Requested From.");
          confidence -= 0.1;
        }
      } else {
        itemWarnings.push("No division line detected after the bill number.");
        confidence -= 0.35;
      }

      var briefLines = body.slice(briefStart).filter(function (l, k, arr) {
        return l || (k > 0 && arr[k - 1]); // collapse runs of blanks
      });
      var brief = briefLines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
      if (!brief) {
        itemWarnings.push("No brief text found.");
        confidence -= 0.25;
      }
      var title = "";
      if (briefLines.length && briefLines[0].length <= 120) { title = briefLines[0]; }

      // referenced bills inside the brief (never boundaries); an
      // amendment's parent bill leads the list
      var refs = [];
      if (b.relatedBill) { refs.push(b.relatedBill); }
      var rm;
      refRe.lastIndex = 0;
      while ((rm = refRe.exec(brief)) !== null) {
        var r = rm[1].toUpperCase() + rm[2].toUpperCase();
        if (r !== b.billNumber && refs.indexOf(r) === -1) { refs.push(r); }
      }

      // attach links whose anchor text appears in this block
      var blockText = b.source.join("\n");
      var links = (conv.links || []).filter(function (l) {
        return l.text && blockText.indexOf(l.text) !== -1;
      });

      return {
        id: "item-" + (idx + 1),
        reportId: opts.reportId || "",
        billNumber: b.billNumber,
        documentType: b.documentType,
        relatedBill: b.relatedBill || "",
        distributedTo: distributedTo,
        commentRequestedFrom: commentFrom,
        title: title,
        brief: brief,
        referencedBills: refs,
        sourceLinks: links,
        sourceBlock: [b.billNumber].concat(body).join("\n"),
        parserConfidence: Math.max(0, Math.round(confidence * 100) / 100),
        parserWarnings: itemWarnings,
        routingStatus: "unmatched",
        matchedRoutingRules: [],
      };
    });

    return { items: items, header: header, warnings: warnings, links: conv.links };
  }

  var api = {
    parseReport: parseReport,
    normalizeDivisions: normalizeDivisions,
    _internals: {
      htmlToText: htmlToText,
      toLines: toLines,
      identifierLineRe: identifierLineRe,
      looksLikeDivisionLine: looksLikeDivisionLine,
      decodeEntities: decodeEntities,
      DEFAULT_IDENTIFIERS: DEFAULT_IDENTIFIERS,
    },
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrParser = api; }
})(typeof self !== "undefined" ? self : this);
