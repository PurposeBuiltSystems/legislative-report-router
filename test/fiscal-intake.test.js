"use strict";
const assert=require('node:assert/strict'),fs=require('node:fs'),{JSDOM}=require('jsdom');
const W=require('../src/workflow'),S=require('../src/workflow-store'),I=require('../src/fiscal-intake');
const actor='staff@agency.test',site={siteId:'s',auditListId:'a'};
const thread=(id='root',team='team',channel='channel')=>`https://teams.microsoft.com/l/message/${channel}/${id}?groupId=${team}&parentMessageId=${id}`;
const divisions=['MVD','TDD','ASD','HIGHWAY','FINANCE'].map(code=>({code,coordinator:code.toLowerCase()+'@agency.test',director:code.toLowerCase()+'-director@agency.test',reviewers:[actor]}));
const config={bill:'HF123',assembly:'91',session:'Iowa / 2026',label:'Issued request',type:'Fiscal data',thread:thread(),submitter:'submit@agency.test',cfo:'cfo@agency.test',divisions};
let sequence=0;
function event(config,id='issued',at='2026-10-01T12:00:00Z'){return {...W.plan([],'create',{config},'owner@agency.test',{id:id+'-create',request:id,at}),order:++sequence};}
async function records(events){const rows=[];for(const e of events){const encoded=await S.encode(e);for(const fields of [...encoded.parts,encoded.commit])rows.push({id:String(++sequence),fields});}return rows;}
const publication=(status='published',id='root',team='team',channel='channel')=>({id:String(++sequence),createdDateTime:'2026-10-01T11:00:00Z',fields:{Status:status,Title:'HF123',ReportKey:status==='bill-thread'?'Iowa / 2026':'feed|fiscal-note',TeamId:team,ChannelId:channel,TeamsMessageId:id,Divisions:'MVD; TDD; ASD; HIGHWAY; FINANCE'}});
async function fixture(){
 const events=[event(config),event({...config,type:'Comment'},'comment'),event({...config,bill:'HF999',thread:thread('unpublished')},'unpublished'),event({...config,thread:thread('commentOnly')},'commentOnly'),event({...config,session:'Iowa / 2025'},'oldSession'),event({...config,thread:thread('root','other-team')},'otherTeam')];
 const rows=await records(events);rows.push(publication(),publication('bill-thread'),publication('bill-thread','commentOnly'),{...publication('published','commentOnly'),fields:{...publication('published','commentOnly').fields,ReportKey:'comment-feed'}},publication('bill-thread','root','other-team'));
 const replies=divisions.map((d,i)=>({id:'r'+i,createdDateTime:`2026-10-02T12:00:0${i}Z`,from:{user:{displayName:d.code+' analyst'}},body:{content:`<p>${d.code}: FY2027 estimate $${(i+1)*100000}. Needs validation.</p>`},attachments:i===0?[{name:'Estimate.xlsx',contentUrl:'https://agency.sharepoint.com/estimate.xlsx'}]:[]}));
 replies.push({id:'before',createdDateTime:'2026-09-01T12:00:00Z',body:{content:'Old estimate $990000'}},{id:'deleted',createdDateTime:'2026-10-02T12:00:00Z',deletedDateTime:'2026-10-03T12:00:00Z',body:{content:'$999999'}},{id:'our-request',createdDateTime:'2026-10-02T12:00:00Z',body:{content:'Please supply an estimate.'}});
 rows.push({id:String(++sequence),fields:{Status:'delivery-link',TeamId:'team',ChannelId:'channel',TeamsMessageId:'our-request'}});
 let writes=0;const reads=[],drafts=[];
 const graph={getToken:async()=>actor,authenticatedAccount:()=>actor,listItems:async()=>structuredClone(rows),listReplies:async(t,team,channel,id)=>{reads.push([team,channel,id]);assert.deepEqual([team,channel,id],['team','channel','root']);return structuredClone(replies);},driveItemFromUrl:async()=>({id:'file',size:500,eTag:'v1',lastModifiedDateTime:'2026-10-02T12:00:00Z',parentReference:{driveId:'drive'}}),driveItemContentB64:async()=>'',addListItem:async(t,s,l,fields)=>{writes++;rows.push({id:String(++sequence),fields});},createDraftMessage:async(t,to,subject,html)=>{drafts.push({to,subject,html});}};
 const options={graph,token:'t',rows,actor,commentLink:(t,id)=>thread(id),extract:async()=> 'FY2027 expenditure estimate $100,000',check:()=>{}};
 return {events,rows,replies,graph,options,reads,drafts,get writes(){return writes;}};
}
async function pure(){
 const f=await fixture(),result=await I.collect(f.options);
 assert.equal(I.scope([event({...config,session:'2026'},'year-label')],f.rows,actor).eligible.length,1,'year label matches recorded Iowa session and assembly');assert.equal(I.scope([event({...config,assembly:'90'},'wrong-assembly')],f.rows,actor).eligible.length,0,'reused bill from another assembly is excluded');
 assert.equal(result.requests,1);assert.equal(result.threads,1);assert.equal(result.candidates.length,5);assert.equal(result.excluded.length,4);assert.equal(f.writes,0,'collection is read only');assert.deepEqual(f.reads,[['team','channel','root']]);
 assert.equal(I.thread('https://evil.test/l/message/c/r?groupId=t'),null);
 assert.equal(I.thread(thread('root')+'&x=1').messageId,'root');
 const candidate=result.candidates.find(c=>c.author==='MVD analyst'),state=S.requests(result.events).find(s=>s.id==='issued'),draft=I.prepare(candidate,state,'MVD',actor);
 assert.equal(draft.financialText,'','unclassified scalar is never turned into a dashboard total');assert.match(draft.narrative.sources,/FY2027/);assert.match(draft.links,/estimate.xlsx/);
 assert.throws(()=>I.prepare(candidate,state,'MVD','unassigned@agency.test'),/assigned/);assert.throws(()=>I.prepare(candidate,{...state,revision:'changed'},'MVD',actor),/changed/);
 const accepted=W.plan(state.events,'artifact',{...draft,division:'MVD',revision:state.revision,requestReview:true},actor,{id:'accepted',at:'2026-10-03T12:00:00Z'});accepted.order=++sequence;
 const acceptedState=W.project([...state.events,accepted]);assert.equal(acceptedState.divisions.MVD.stage,'Review requested');assert.equal(acceptedState.cfo,null);assert.equal(acceptedState.divisions.MVD.director,null);
 f.rows.push(...await records([accepted]));const rerun=await I.collect(f.options);assert.equal(rerun.skipped,1);assert.equal(rerun.candidates.length,4,'unchanged collected response is not registered twice');
 assert.throws(()=>W.plan([...state.events,accepted],'artifact',{...draft,division:'MVD',revision:acceptedState.revision},actor,{id:'duplicate'}),/already registered/);
 f.replies[0].lastModifiedDateTime='2026-10-04T12:00:00Z';f.replies[0].body.content='<p>Revised estimate $200000</p>';assert.equal((await I.collect(f.options)).candidates.length,5,'edited reply is a revision candidate');
 f.graph.driveItemFromUrl=async()=>{throw new Error('access denied');};const warning=await I.collect(f.options);assert.equal(warning.warnings.length,1);assert.ok(warning.candidates.find(c=>c.author==='MVD analyst').files.length,'inaccessible attachment stays linked for manual review');
 let n=0;await assert.rejects(()=>I.collect({...f.options,check:()=>{if(++n>2)throw new Error('account changed');}}),/account changed/);assert.equal(f.writes,0);
 const foreign=await I.collect({...f.options,actor:'outsider@agency.test'});assert.equal(foreign.threads,0,'unassigned unrelated user does not scan requests');
 const completed={...state,submitted:{id:'sent'}};assert.throws(()=>I.prepare(candidate,completed,'MVD',actor),/eligible/);
}
async function ui(){
 const f=await fixture(),dom=new JSDOM('<div id="host"></div>',{runScripts:'outside-only',url:'https://router.test/'}),w=dom.window;Object.defineProperty(w,'crypto',{value:require('node:crypto').webcrypto});w.TextEncoder=TextEncoder;
 for(const file of ['workflow','workflow-store','harvest','fiscal-intake','workflow-ui'])w.eval(fs.readFileSync('src/'+file+'.js','utf8'));
 const host=w.document.getElementById('host'),el=id=>w.document.getElementById(id);let changed=0;
 const view=w.LrrWorkflowUI.mount(host,{graph:f.graph,site:()=>site,prepareFiscalResponse:I.prepare,onSaved:()=>changed++,collectFiscalResponses:async(t,a,check)=>I.collect({...f.options,token:t,actor:a,check})});
 async function settled(){for(let i=0;i<500;i++){await new Promise(r=>setTimeout(r,2));if(!el('wfLoad').disabled)return;}throw new Error('UI never settled');}
 await view.load();await view.collectFiscal();assert.equal(f.writes,0);assert.equal(host.querySelectorAll('[data-intake-prepare]').length,5);
 const select=host.querySelector('[data-intake-request="0"]');select.value='issued';select.dispatchEvent(new w.Event('change'));host.querySelector('[data-intake-division="0"]').value='FINANCE';host.querySelector('[data-intake-prepare="0"]').click();
 assert.equal(el('wfActionDivision').value,'FINANCE');assert.equal(el('wfFinance').value,'');assert.equal(el('wfConfirm').checked,false);el('wfSaveAction').click();assert.equal(f.writes,0);assert.match(el('wfInfo').textContent,/confirm/);
 await view.collectFiscal();assert.match(el('wfInfo').textContent,/Save or discard/,'collection never discards edits');
 el('wfDocs').value='https://agency.sharepoint.com/response?version=1';el('wfFinance').value='FY2027\tExpenditure\tIncrease\t500000\tRoad fund\tOngoing\tConfirmed estimate';el('wfConfirm').checked=true;el('wfSaveAction').click();await settled();
 let state=S.requests(await S.decode(f.rows)).find(s=>s.id==='issued');assert.equal(state.divisions.FINANCE.stage,'Review requested');assert.equal(state.divisions.FINANCE.artifact.data.finance[0].amount,500000);assert.equal(changed,1);assert.equal(f.drafts.length,1);assert.deepEqual(Array.from(f.drafts[0].to),['finance@agency.test']);assert.equal(el('wfCollection').textContent,'','successful save clears stale candidates');assert.match(el('wfInfo').textContent,/review and send/);
 await view.collectFiscal();assert.equal(host.querySelectorAll('[data-intake-prepare]').length,4);view.invalidate();assert.equal(el('wfCollection').textContent,'');assert.equal(el('wfCollectFiscal').disabled,true);
 // The Teams staff tab does not receive the reply-reading action or broader scopes.
 const staff=w.LrrWorkflowUI.mount(host,{graph:f.graph,site:()=>site,staff:true});assert.equal(el('wfCollectFiscal'),null);staff.invalidate();dom.window.close();
}
(async()=>{await pure();await ui();console.log('Fiscal intake: five bureaus, issued-only scope, no discovery/writes on collect, duplicates, edits, attachment failures, account changes, confirmation, atomic review handoff and coordinator draft passed.');})().catch(e=>{console.error(e);process.exitCode=1;});
