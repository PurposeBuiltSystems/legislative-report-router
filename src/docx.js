/*
 * Legislative Report Router — DOCX text extraction (client-side, no Word).
 *
 * A .docx is a zip; the document body is word/document.xml. We unzip with
 * JSZip (CDN) and walk the XML with DOMParser: paragraphs become lines,
 * table rows become tab-joined lines, hyperlink relationships resolve from
 * word/_rels/document.xml.rels. Content is NEVER executed — text only.
 * Macro formats (.docm) and non-docx types are rejected before parsing.
 */
/* global JSZip, DOMParser */
(function (root) {
  "use strict";

  var MAX_BYTES = 15 * 1024 * 1024; // 15 MB cap

  function b64ToBytes(b64) {
    var bin = atob(b64);
    var arr = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) { arr[i] = bin.charCodeAt(i); }
    return arr;
  }

  function isSupported(name, contentType, size) {
    if (size && size > MAX_BYTES) { return { ok: false, reason: "File exceeds the 15 MB limit." }; }
    if (/\.docm$/i.test(name || "")) { return { ok: false, reason: "Macro-enabled documents (.docm) are not accepted." }; }
    if (!/\.docx$/i.test(name || "") &&
        String(contentType || "").indexOf("officedocument.wordprocessingml.document") === -1) {
      return { ok: false, reason: "Only .docx documents are supported." };
    }
    return { ok: true };
  }

  function textOfParagraph(p, rels) {
    var parts = [];
    var walker = p.getElementsByTagName("*");
    for (var i = 0; i < walker.length; i++) {
      var el = walker[i];
      var tag = el.localName;
      if (tag === "t") { parts.push(el.textContent); }
      else if (tag === "tab") { parts.push("\t"); }
      else if (tag === "br") { parts.push("\n"); }
    }
    var text = parts.join("");
    // append resolved hyperlink target once per w:hyperlink
    var links = p.getElementsByTagName("*");
    for (var j = 0; j < links.length; j++) {
      if (links[j].localName === "hyperlink") {
        var rid = links[j].getAttribute("r:id") || links[j].getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
        if (rid && rels[rid]) { text += " (" + rels[rid] + ")"; }
      }
    }
    return text;
  }

  /**
   * base64 DOCX → {text, links[]}. Paragraph per line; table rows become
   * tab-joined lines so bill blocks living in table cells still parse.
   */
  async function extractText(base64, name, contentType, size) {
    var gate = isSupported(name, contentType, size);
    if (!gate.ok) { throw new Error(gate.reason); }
    var zip = await JSZip.loadAsync(b64ToBytes(base64));
    var docFile = zip.file("word/document.xml");
    if (!docFile) { throw new Error("Not a valid .docx (missing word/document.xml)."); }

    var rels = {};
    var links = [];
    var relFile = zip.file("word/_rels/document.xml.rels");
    if (relFile) {
      var relXml = new DOMParser().parseFromString(await relFile.async("string"), "application/xml");
      var relNodes = relXml.getElementsByTagName("*");
      for (var i = 0; i < relNodes.length; i++) {
        if (relNodes[i].localName === "Relationship" &&
            /hyperlink/i.test(relNodes[i].getAttribute("Type") || "")) {
          rels[relNodes[i].getAttribute("Id")] = relNodes[i].getAttribute("Target");
          links.push({ text: "", href: relNodes[i].getAttribute("Target") });
        }
      }
    }

    var xml = new DOMParser().parseFromString(await docFile.async("string"), "application/xml");
    var body = xml.getElementsByTagName("*");
    var lines = [];
    for (var k = 0; k < body.length; k++) {
      var el = body[k];
      if (el.localName === "p" && !hasAncestor(el, "tbl")) {
        lines.push(textOfParagraph(el, rels));
      } else if (el.localName === "tr") {
        var cells = [];
        var kids = el.getElementsByTagName("*");
        var cellTexts = [];
        for (var c = 0; c < kids.length; c++) {
          if (kids[c].localName === "tc") {
            var ps = kids[c].getElementsByTagName("*");
            var cellParts = [];
            for (var q = 0; q < ps.length; q++) {
              if (ps[q].localName === "p") { cellParts.push(textOfParagraph(ps[q], rels)); }
            }
            cellTexts.push(cellParts.join(" ").trim());
          }
        }
        cells = cellTexts;
        // one line per cell keeps "HF935 | MVD | MVD | brief" rows parseable
        cells.forEach(function (ct) { if (ct) { lines.push(ct); } });
        lines.push("");
      }
    }
    return { text: lines.join("\n"), links: links };
  }

  function hasAncestor(el, localName) {
    var p = el.parentNode;
    while (p) {
      if (p.localName === localName) { return true; }
      p = p.parentNode;
    }
    return false;
  }

  root.LrrDocx = { extractText: extractText, isSupported: isSupported, _internals: { b64ToBytes: b64ToBytes } };
})(typeof self !== "undefined" ? self : this);
