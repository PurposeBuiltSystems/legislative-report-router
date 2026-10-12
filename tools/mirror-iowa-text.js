/* Public Iowa document text, cached for agency screening; no tenant data.
 * Bounded batches gradually backfill the entire feed. Missing/failing text
 * remains summary-only in the app, never an implicit negative result.
 */
"use strict";
const fs = require("node:fs");
const path = require("node:path");
const Feed = require("../src/feed.js");
const Screening = require("../src/screening.js");
const Parser = require("../src/parser.js");

function documentUrl(entry) {
  const match = String(entry.link || "").match(/^https:\/\/www\.legis\.iowa\.gov\/legislation\/BillBook\?[^#]*\bga=(\d+)\b/i);
  if (!match || !/^(HF|SF|HSB|SSB|HJR|SJR|HCR|SCR|HR|SR|H-|S-|CCS|CCH)\d+[A-Z]?$/.test(entry.bill)) return null;
  const type = /^(H-|S-|CCS|CCH)/.test(entry.bill) ? "AMDI" : "LGI";
  return `https://www.legis.iowa.gov/docs/publications/${type}/${match[1]}/attachments/${entry.bill.replace("-", "")}.html`;
}
function documentText(html) {
  // Iowa positions each word in a span; inserting spaces preserves phrases.
  return Parser._internals.htmlToText(String(html).replace(/<\/span>/gi, " </span>")).text.replace(/\s+/g, " ").trim();
}
async function update(entries, previous, opts = {}) {
  const now = opts.now || new Date();
  const mirror = { documents: Object.assign({}, previous.documents), failures: Object.assign({}, previous.failures) };
  const unique = new Map();
  entries.slice().sort((a, b) => b.pubDate - a.pubDate).forEach(entry => {
    const key = Screening.entryKey(entry);
    if (!unique.has(key)) unique.set(key, entry);
  });
  let attempted = 0, succeeded = 0;
  for (const [key, entry] of unique) {
    const url = documentUrl(entry);
    if (!url) continue;
    const doc = mirror.documents[key];
    if (doc && new Date(doc.fetchedAt) >= entry.pubDate) continue;
    const failed = mirror.failures[key];
    if (failed && now - new Date(failed.at) < 6 * 3600000) continue;
    if (attempted >= (opts.limit || 80)) break;
    attempted++;
    try {
      const res = await (opts.fetch || fetch)(url, { signal: AbortSignal.timeout(10000) });
      if (!res.ok) throw new Error("HTTP " + res.status);
      const html = await res.text();
      if (html.length > 5e6) throw new Error("Document too large");
      const text = documentText(html);
      if (text.length < 100 || !/\b(bill|amend|resolution|conference)\b/i.test(text)) throw new Error("Unrecognized legislative document");
      if (text.length > 500000) throw new Error("Document exceeds screening limit");
      mirror.documents[key] = { url, version: /AMDI/.test(url) ? "filed amendment/conference" : "introduced bill", fetchedAt: now.toISOString(), text };
      delete mirror.failures[key];
      succeeded++;
    } catch (err) {
      mirror.failures[key] = { at: now.toISOString(), reason: err.message };
    }
  }
  // A failed batch preserves the last successful documents.
  return { mirror, attempted, succeeded, pending: [...unique.keys()].filter(key => !mirror.documents[key]).length };
}
async function main() {
  const base = path.join(__dirname, "../feeds");
  const file = path.join(base, "iowa-screening.json");
  const previous = fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, "utf8")) : {};
  // Keep the browser index small; document bodies live in separate files.
  for (const [key, metadata] of Object.entries(previous.documents || {})) {
    if (metadata.text) continue;
    const documentFile = path.join(base, "iowa-text", key.replace(":", "/") + ".json");
    if (fs.existsSync(documentFile)) previous.documents[key] = JSON.parse(fs.readFileSync(documentFile, "utf8"));
    else delete previous.documents[key];
  }
  const entries = ["IowaBills.xml", "Amendments.xml"].flatMap(name => Feed.parseFeed(fs.readFileSync(path.join(base, name), "utf8")));
  const result = await update(entries, previous);
  const index = { documents: {}, failures: result.mirror.failures };
  for (const [key, doc] of Object.entries(result.mirror.documents)) {
    const documentFile = path.join(base, "iowa-text", key.replace(":", "/") + ".json");
    fs.mkdirSync(path.dirname(documentFile), { recursive: true });
    const body = JSON.stringify(doc) + "\n";
    if (!fs.existsSync(documentFile) || fs.readFileSync(documentFile, "utf8") !== body) fs.writeFileSync(documentFile, body);
    index.documents[key] = { url: doc.url, version: doc.version, fetchedAt: doc.fetchedAt };
  }
  fs.writeFileSync(file + ".tmp", JSON.stringify(index) + "\n");
  fs.renameSync(file + ".tmp", file);
  console.log(`Iowa text mirror: ${result.succeeded}/${result.attempted} fetched; ${result.pending} awaiting text (summary-only review available).`);
}
module.exports = { documentUrl, documentText, update };
if (require.main === module) main().catch(err => { console.error(err.message); process.exitCode = 1; });
