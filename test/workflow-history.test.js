'use strict';
const assert=require('node:assert/strict'),W=require('../src/workflow');
const roles=[{code:'MVD',coordinator:'me@agency.test',director:'director@agency.test',reviewers:[]}];
function request(id,bill,assembly,summary,notes,actor='me@agency.test'){
 const event={id:id+'-v1',kind:'artifact',division:'MVD',actor,at:'2026-01-01T10:00:00Z',data:{version:'v1',note:notes,narrative:{},links:['https://example.test/'+id]}};
 return {id,config:{bill,assembly,session:assembly==='91'?'2026':'2024',label:'Comment request',summary,related:'',divisions:roles},conflicts:[],appliedEvents:[event],events:[event]};
}
const current=request('current','HF123','91','Overweight vehicle permits under chapter 321.','');
const old=request('old','HF999','90','','Overweight vehicle permits will affect county roads.');
const exact=request('exact','HF123','91','','Earlier response.');
const reused=request('reused','HF123','90','','Unrelated subject.');
const other=request('other','SF20','90','','Overweight vehicle permits','other@agency.test');other.appliedEvents[0].division='OTHER';
const unrelated=request('unrelated','HF200','91','','Daily legislative fiscal comment response review.');
const conflict=request('conflict','HF123','91','','Overweight vehicle permits');conflict.conflicts=['bad'];
let results=W.relevantResponses(current,[current,old,exact,reused,other,unrelated,conflict],'me@agency.test','all');
assert.deepEqual(results.map(r=>r.request.id),['exact','old','other']);
assert.match(results[1].reasons.join(' '),/Shared keywords:.*overweight/);
assert.deepEqual(W.relevantResponses(current,[old,other],'me@agency.test','mine').map(r=>r.request.id),['old']);
assert.deepEqual(W.relevantResponses(current,[old,other],'me@agency.test','division').map(r=>r.request.id),['old']);
const related=request('related','SF555','91','','');related.config.related='HF 123';
assert.match(W.relevantResponses(current,[related],'me@agency.test','all')[0].reasons.join(' '),/Related bill reference/);
const chapter=request('chapter','SF90','90','','Section 321. requires a change.');
assert.match(W.relevantResponses(current,[chapter],'me@agency.test','all')[0].reasons.join(' '),/Shared code chapter/);
const newer={...old.appliedEvents[0],id:'old-v2',data:{...old.appliedEvents[0].data,version:'v2'}};
old.appliedEvents.push(newer);assert.equal(W.relevantResponses(current,[old],'me@agency.test','all').length,1);assert.equal(W.relevantResponses(current,[old],'me@agency.test','all')[0].event.id,'old-v2');
const wrongBill=request('wrong','HF1234','91','','');assert.equal(W.relevantResponses(current,[wrongBill],'me@agency.test','all').length,0);
console.log('Past response matching: automatic keywords, bill identity, related references, chapters, personal/division scope, conflict exclusion and latest-version dedupe passed.');
