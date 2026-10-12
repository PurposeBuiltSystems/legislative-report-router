"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const F = require("../src/feed.js");
const S = require("../src/screening.js");
const C = require("../src/chapters.js");
const G = require("../src/reportgen.js");
const P = require("../src/parser.js");
const Mirror = require("../tools/mirror-iowa-text.js");
const now = new Date("2026-02-12T12:00:00Z");
const contract = { bill: "HF2500", description: "A bill for an act relating to contracts entered into by state agencies and including applicability provisions.", link: "https://www.legis.iowa.gov/legislation/BillBook?ga=91&ba=HF2500", pubDate: now };
const unknown = { bill: "HF100", description: "A bill relating to miscellaneous provisions.", pubDate: now };
const terms = ["transportation", "motor vehicle", "highway"];
assert.equal(F.watchFilter([contract], terms, 3, now).length, 0, "reproduces the old contracting gap");
const hits = S.screen([contract, unknown, contract, { ...unknown, bill: "HF1", pubDate: new Date("2025-01-01") }], { terms, days: 3, now });
assert.equal(hits.length, 2, "unknown filings retained; duplicates and old filings excluded");
assert(hits[0].screening.reasons.includes("Contracts and procurement"));
assert.equal(hits[1].screening.status, "needs-screening");
assert.equal(hits[1].screening.fullTextAvailable, false);
assert(S.assess({description:"Amends Code chapter 8A."}).reasons.some(r=>r.includes("Agency-wide Code: 8A")));
for (const [description, reason] of [
  ["Ownership of intellectual property from federally funded research", "Research and intellectual property"],
  ["State employee retirement benefits", "Employment and benefits"],
  ["Appropriations and grants", "Funding and grants"],
  ["Public records and cybersecurity", "Agency administration and legal duties"],
]) assert(S.assess({ description }).reasons.includes(reason));
const doc = "<html><head><script>evil()</script></head><body><span>State</span><span>agency</span><span>contract</span><span>intellectual</span><span>property</span><p>Research funding and procurement. Amends Code chapter 8A. This bill addresses contract terms and required provisions.</p></body></html>";
const text = Mirror.documentText(doc);
assert(text.includes("intellectual property"));
assert(!text.includes("evil"));
const enriched = S.enrich([contract], { documents: { "91:HF2500": { text, version: "introduced bill" } } });
assert(S.assess(enriched[0]).reasons.includes("Research and intellectual property"));
assert.equal(S.enrich([{ ...contract, link: contract.link.replace("91", "92") }], { documents: { "91:HF2500": { text } } })[0].fullText, undefined, "text never crosses assemblies");
const xml = '<rss><channel><item><title><![CDATA[S 5210 (SF 2472)]]></title><description><![CDATA[Amendment to SF 2472]]></description><link><![CDATA[https://www.legis.iowa.gov/legislation/BillBook?ga=91&ba=S5210]]></link><pubDate>Thu, 12 Feb 2026 09:00:00</pubDate></item><item><title><![CDATA[CCS 2218 (SF 2218)]]></title><description>Conference report</description><pubDate>Thu, 12 Feb 2026 09:00:00</pubDate></item></channel></rss>';
const amendments = F.parseFeed(xml);
assert.equal(amendments.length, 2);
assert.equal(amendments[0].bill, "S-5210");
assert.equal(amendments[0].relatedBill, "SF2472");
assert.equal(amendments[1].documentType, "Conference report");
assert.equal(G.filterNew(amendments, ["SF2472"]).length, 2, "reporting a parent never hides its new amendment");
const amendmentItem = F.toLegislativeItem(amendments[0], "test", 0);
assert.equal(amendmentItem.documentType, "Amendment");
const round = P.parseReport(G.buildDailyReport([amendmentItem]).text);
assert.equal(round.items[0].billNumber, "S-5210");
assert.equal(round.items[0].relatedBill, "SF2472");
assert(Mirror.documentUrl(amendments[0]).includes("/AMDI/91/attachments/S5210.html"));
assert.equal(Mirror.documentUrl({ ...contract, link: "https://evil.example/?ga=91" }), null);
// Browser globals must work, not just the CommonJS implementation.
const browser = { self: {}, Date };
vm.createContext(browser);
for (const name of ["chapters", "feed", "screening"]) vm.runInContext(fs.readFileSync(require("node:path").join(__dirname, "../src/" + name + ".js"), "utf8"), browser);
assert(browser.self.LrrScreening.assess(contract).reasons.includes("Contracts and procurement"));
(async () => {
  const first = await Mirror.update([contract], {}, { now, fetch: async () => ({ ok: true, text: async () => doc }) });
  assert.equal(first.succeeded, 1);
  assert(first.mirror.documents["91:HF2500"].text.includes("intellectual property"));
  const next = await Mirror.update([contract], first.mirror, { now, fetch: async () => { throw new Error("must be cached"); } });
  assert.equal(next.attempted, 0);
  const failed = await Mirror.update([{ ...contract, pubDate: new Date(now.getTime() + 1000) }], first.mirror, { now, fetch: async () => ({ ok: false, status: 503 }) });
  assert.equal(failed.succeeded, 0);
  assert(failed.mirror.documents["91:HF2500"], "failure preserves previous text");
  console.log("All agency screening and document mirror tests passed.");
})().catch(e => { console.error(e); process.exitCode = 1; });
