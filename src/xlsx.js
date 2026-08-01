/*
 * Legislative Report Router — minimal .xlsx text extraction (browser).
 *
 * An .xlsx is a ZIP: xl/sharedStrings.xml holds the string table,
 * xl/worksheets/sheetN.xml holds cells (<c t="s"><v>idx</v></c> for
 * shared strings, plain <v> for numbers). This walks the first few
 * sheets and emits tab-joined rows of text — enough for the fiscal
 * harvest to find "Estimated cost   $1,200,000" wherever it lives.
 * Same JSZip + DOMParser stack as the DOCX reader; no new libraries.
 */
/* global JSZip, DOMParser */
(function (root) {
  "use strict";

  var MAX_BYTES = 15 * 1024 * 1024;
  var MAX_SHEETS = 3;
  var MAX_ROWS = 500;

  function b64ToBytes(b64) {
    var bin = atob(b64);
    var arr = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) { arr[i] = bin.charCodeAt(i); }
    return arr;
  }

  function isSupported(name, size) {
    if (size && size > MAX_BYTES) { return { ok: false, reason: "File exceeds the 15 MB limit." }; }
    if (/\.xlsm$/i.test(name || "")) { return { ok: false, reason: "Macro-enabled workbooks (.xlsm) are not accepted." }; }
    if (!/\.xlsx$/i.test(name || "")) { return { ok: false, reason: "Only .xlsx workbooks are supported." }; }
    return { ok: true };
  }

  function parseSharedStrings(xmlText) {
    var out = [];
    if (!xmlText) { return out; }
    var xml = new DOMParser().parseFromString(xmlText, "application/xml");
    var sis = xml.getElementsByTagName("*");
    for (var i = 0; i < sis.length; i++) {
      if (sis[i].localName === "si") {
        var parts = [];
        var kids = sis[i].getElementsByTagName("*");
        for (var j = 0; j < kids.length; j++) {
          if (kids[j].localName === "t") { parts.push(kids[j].textContent); }
        }
        out.push(parts.join(""));
      }
    }
    return out;
  }

  function sheetToLines(xmlText, shared) {
    var xml = new DOMParser().parseFromString(xmlText, "application/xml");
    var lines = [];
    var rows = xml.getElementsByTagName("*");
    var count = 0;
    for (var i = 0; i < rows.length && count < MAX_ROWS; i++) {
      if (rows[i].localName !== "row") { continue; }
      count++;
      var cells = [];
      var kids = rows[i].getElementsByTagName("*");
      for (var j = 0; j < kids.length; j++) {
        if (kids[j].localName !== "c") { continue; }
        var c = kids[j];
        var type = c.getAttribute("t") || "";
        var v = "";
        var inner = c.getElementsByTagName("*");
        for (var k = 0; k < inner.length; k++) {
          if (inner[k].localName === "v" || inner[k].localName === "t") { v = inner[k].textContent; break; }
        }
        if (type === "s") { v = shared[Number(v)] || ""; }
        cells.push(v);
      }
      var line = cells.join("\t").replace(/\t+$/, "");
      if (line.trim()) { lines.push(line); }
    }
    return lines;
  }

  /** base64 XLSX → plain text (tab-joined rows, first sheets only). */
  async function extractText(base64, name, size) {
    var gate = isSupported(name, size);
    if (!gate.ok) { throw new Error(gate.reason); }
    var zip = await JSZip.loadAsync(b64ToBytes(base64));
    var sharedFile = zip.file("xl/sharedStrings.xml");
    var shared = parseSharedStrings(sharedFile ? await sharedFile.async("string") : "");
    var lines = [];
    for (var n = 1; n <= MAX_SHEETS; n++) {
      var f = zip.file("xl/worksheets/sheet" + n + ".xml");
      if (!f) { break; }
      lines = lines.concat(sheetToLines(await f.async("string"), shared));
    }
    if (!lines.length) { throw new Error("Workbook has no readable cells."); }
    return lines.join("\n");
  }

  var api = { extractText: extractText, isSupported: isSupported };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrXlsx = api; }
})(typeof self !== "undefined" ? self : this);
