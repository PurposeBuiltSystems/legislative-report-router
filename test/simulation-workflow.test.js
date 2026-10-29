"use strict";
const assert=require('node:assert/strict'),fs=require('node:fs'),{JSDOM}=require('jsdom');
const dom=new JSDOM('<div id="host"></div>',{runScripts:'outside-only',url:'https://router.test/'}),w=dom.window;
Object.defineProperty(w,'crypto',{value:require('node:crypto').webcrypto});w.TextEncoder=TextEncoder;
for(const file of ['staff-config','workflow','workflow-store','workflow-ui'])w.eval(fs.readFileSync('src/'+file+'.js','utf8'));
let rows=[],seq=0;const actor='test@agency.test',site={siteId:'site',auditListId:'audit'};
const graph={getToken:async()=>actor,authenticatedAccount:t=>t,listItems:async()=>structuredClone(rows),addListItem:async(t,s,l,fields)=>rows.push({id:String(++seq),fields})};
const view=w.LrrWorkflowUI.mount(w.document.getElementById('host'),{graph,site:()=>site});
const el=id=>w.document.getElementById(id),set=(id,v)=>{el(id).value=v;},click=id=>el(id).click();
const choose=(id,v)=>{set(id,v);el(id).dispatchEvent(new w.Event('change'));};
async function done(){for(let deadline=Date.now()+5000;Date.now()<deadline;){await new Promise(r=>setTimeout(r,5));if(!el('wfLoad').disabled)return;}throw Error('save timeout');}
async function action(kind){choose('wfAction',kind);el('wfConfirm').checked=true;click('wfSaveAction');await done();}
async function response(division,version){choose('wfActionDivision',division);choose('wfAction','artifact');set('wfVersion',version);set('wfDocs','https://example.test/'+division+'/'+version);el('wfConfirm').checked=true;click('wfSaveAction');await done();assert.equal(el('wfActionDivision').value,division,'saving must keep the selected division');}
(async()=>{
 await view.load();click('wfNew');set('wfBill','HF999998');set('wfSession','SIMULATION');set('wfLabel','Test');set('wfKind','Fiscal data');set('wfCfo',actor);set('wfSubmitter',actor);
 set('wfAssignments','MVD | '+actor+' | '+actor+' | '+actor+'\nTDD | '+actor+' | '+actor+' | '+actor);click('wfSaveConfig');await done();
 await response('TDD','v1');await action('review');assert.equal(el('wfActionDivision').value,'TDD');await action('coordinator');await action('director');
 assert.ok(!el('wfAction').querySelector('[value=cfo]'),'one division cannot unlock CFO approval');
 await response('MVD','v1');await action('review');await action('coordinator');await action('director');assert.match(el('wfQueue').textContent,/CFO review/);
 await action('cfo');assert.match(el('wfQueue').textContent,/CFO approved/);assert.ok(!el('wfAction').querySelector('[value=cfo]'),'current CFO approval is not offered again');
 assert.equal(el('wfAction').value,'submitted','next action goes to the submitter');assert.match(el('wfDetail').textContent,/record completed LSA submission/);
 await response('TDD','v2');assert.match(el('wfQueue').textContent,/Division work pending/);assert.match(el('wfDetail').textContent,/Earlier CFO approval is no longer current/);
 assert.ok(!el('wfAction').querySelector('[value=submitted]'),'revised response cannot use old approval to record submission');
 console.log('Live-simulation regressions: division selection survives saves, all-divisions gate, CFO-to-submitter handoff and visible stale approval passed.');dom.window.close();
})().catch(e=>{console.error(e);process.exitCode=1;dom.window.close();});
