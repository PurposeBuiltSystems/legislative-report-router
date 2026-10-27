"use strict";
const assert=require("node:assert/strict"),W=require("../src/watch-window");
(async()=>{
 const now=new Date("2026-10-27T21:00:00Z"),config={stateName:"Iowa",days:3,terms:["Contracts"],tracked:["8A"],matchChapters:true};
 const entry={bill:"HF1",link:"https://www.legis.iowa.gov/legislation/BillBook?ga=91&ba=HF1",description:"Bill",pubDate:new Date("2026-10-26T16:00:00Z"),fullText:"Introduced text",textVersion:"introduced"};
 const first=await W.prepare([entry],config,now,[]);assert.equal(first.entries.length,1);
 const key=first.entries[0].key;
 const rows=[{fields:{ReportKey:"watch:"+first.scope,IdempotencyKey:key,Status:"watch-reviewed"}}];
 const repeat=await W.prepare([entry],config,now,rows);assert.ok(repeat.reviewed.has(repeat.entries[0].key));
 for(const change of [{description:"Updated summary"},{fullText:"Amended text"},{link:entry.link.replace("ga=91","ga=92")},{parentBill:{bill:"SF1",description:"Changed parent"}}]) {
  assert.notEqual(await W.fingerprint({...entry,...change}),await W.fingerprint(entry),"every meaningful version/context change is new");
 }
 const late={...entry,bill:"HF2",pubDate:new Date("2026-10-25T12:00:00Z")};
 const later=await W.prepare([entry,late],config,new Date("2026-10-27T22:00:00Z"),rows);
 assert.equal(later.entries.filter(e=>!later.reviewed.has(e.key)).length,1,"late arrivals inside the rolling window are not skipped by last-check time");
 assert.equal((await W.prepare([{...entry,pubDate:new Date("2026-10-28T12:00:00Z")}],config,now,[])).entries.length,0);
 assert.equal((await W.prepare([{...entry,pubDate:new Date("2026-10-20T12:00:00Z")}],config,now,[])).entries.length,0);
 assert.equal(await W.scope(config),await W.scope({...config,days:30,terms:["contracts","contracts"]}),"window expansion reuses exact-version reviews");
 assert.notEqual(await W.scope(config),await W.scope({...config,terms:["aviation"]}),"different screening criteria require fresh review");
 assert.equal(W.history([{fields:{...rows[0].fields,Status:"published"}}],first.scope).reviewed.size,0,"publication is not an explicit review checkpoint");
 console.log("Watch version fingerprints, rolling overlap, scopes and history tests passed.");
})().catch(e=>{console.error(e);process.exitCode=1;});
