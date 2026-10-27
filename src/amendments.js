/* Parent-bill context from the same assembly's feed/mirrored bill text.
 * Context is attributed separately; it is not a summary of amendment changes.
 */
(function (root) {
  "use strict";
  function assembly(entry) { var m = /[?&]ga=(\d+)\b/i.exec(entry.link || ""); return m ? m[1] : ""; }
  function parentEntry(entry) {
    if (!entry.relatedBill || !/^[A-Z]{2,4}\d+[A-Z]?$/.test(entry.relatedBill)) { return null; }
    var ga = assembly(entry);
    return { bill: entry.relatedBill, link: ga ? "https://www.legis.iowa.gov/legislation/BillBook?ga=" + ga + "&ba=" + entry.relatedBill : "" };
  }
  function attach(recent, all, documents) {
    documents = documents || {};
    return recent.map(function (entry) {
      var parent = parentEntry(entry);
      if (!parent) { return entry; }
      var ga = assembly(entry), key = (ga ? ga + ":" : "") + parent.bill;
      var source = (all || []).find(function (e) { return !e.relatedBill && e.bill === parent.bill && assembly(e) === ga; });
      var doc = documents[key] || {};
      var description = source && source.description || "";
      var excerpt = "";
      if (!description && doc.text) {
        // Quote a bounded excerpt instead of inventing a legal summary.
        var text = String(doc.text).replace(/\s+/g, " ");
        var at = text.search(/\bAn Act\b/i);
        excerpt = text.slice(at >= 0 ? at : 0, (at >= 0 ? at : 0) + 420);
      }
      return Object.assign({}, entry, { parentBill: { bill: parent.bill, link: parent.link,
        description: description, excerpt: excerpt, version: doc.version || "",
        textUrl: /^https:\/\/www\.legis\.iowa\.gov\//i.test(doc.url || "") ? doc.url : "",
        available: !!(description || excerpt) } });
    });
  }
  function text(parent) {
    if (!parent) { return ""; }
    return "Parent bill " + parent.bill + ": " + (parent.description || (parent.excerpt ? "Bill text excerpt: " + parent.excerpt : "Summary unavailable in the mirror; open the parent bill.")) +
      (parent.version ? " [" + parent.version + "]" : "") + " This is bill context, not an assessment of the amendment's changes.";
  }
  var api = { parentEntry: parentEntry, attach: attach, text: text };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrAmendments = api; }
})(typeof self !== "undefined" ? self : this);
