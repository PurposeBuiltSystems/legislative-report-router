"use strict";
const assert=require('node:assert/strict');
const W=require('../src/workflow'),S=require('../src/workflow-store');
const owner='owner@agency.test',config={bill:'HF123',assembly:'91',session:'2026',label:'Confirmed request',type:'Comment',due:'2026-11-02',dueSource:'Requester email',source:'https://agency.test/request',thread:'https://teams.test/thread',cfo:'cfo@agency.test',submitter:'submit@agency.test',divisions:[{code:'MVD',coordinator:'coord@agency.test',director:'director@agency.test',reviewers:[]}]};
let seq=0;const rows=[],graph={listItems:async()=>structuredClone(rows),addListItem:async(t,s,l,fields)=>rows.push({id:String(++seq),fields:structuredClone(fields)}),updateListItemFields:async(t,s,l,id,fields)=>Object.assign(rows.find(r=>r.id===id).fields,fields)};
const client=S.client(graph,{siteId:'site',auditListId:'audit'});
async function act(kind,input={},actor=owner){const events=await client.read('token'),state=W.project(events);const e=W.plan(events,kind,{...input,revision:state?state.revision:''},actor,{id:'e'+(seq+1),request:'request',at:'2026-10-29T12:00:00Z'});return client.append('token',e);}
(async()=>{
await act('create',{config});assert.equal(rows.filter(r=>r.fields.Status==='reminder-pending').length,0,'explicit opt in');
let result=await client.configureReminders('token',true);assert.equal(result.pending,1);
let row=rows.find(r=>r.fields.Status==='reminder-pending');assert.equal(row.fields.EmailRecipients,'coord@agency.test');assert.equal(row.fields.Error,'2026-11-02');assert.equal(row.fields.ChannelId+row.fields.TeamsMessageId,config.thread);
const commits=()=>rows.filter(r=>r.fields.Status==='workflow-event');assert.equal(row.fields.IdempotencyKey,commits().at(-1).id,'freshness uses committed SharePoint ID');
row.fields.PublishedBy='2026-10-30';await act('note',{intent:'FYI',note:'Context updated'});assert.equal(row.fields.PublishedBy,'2026-10-30','sync preserves send reservation');assert.equal(row.fields.IdempotencyKey,commits().at(-1).id);
await act('artifact',{division:'MVD',version:'v1',links:'https://agency.test/doc'},'coord@agency.test');await act('review',{division:'MVD'},'coord@agency.test');await act('coordinator',{division:'MVD'},'coord@agency.test');assert.equal(row.fields.EmailRecipients,'director@agency.test','current responsible person changes');
await act('director',{division:'MVD'},'director@agency.test');assert.equal(row.fields.Status,'reminder-closed','completed work closes reminder');
await act('artifact',{division:'MVD',version:'v2',links:'https://agency.test/doc2'},'coord@agency.test');assert.equal(row.fields.Status,'reminder-pending','rework reopens reminder');assert.equal(row.fields.PublishedBy,'2026-10-30');
await client.configureReminders('token',false);assert.equal(row.fields.Status,'reminder-closed','disable closes pending feed');
await client.configureReminders('token',true);assert.equal(row.fields.Status,'reminder-pending');
const oldRevision=row.fields.IdempotencyKey;graph.updateListItemFields=async()=>{throw new Error('projection unavailable');};const saved=await act('note',{intent:'FYI',note:'Newest context'});assert.match(saved.reminderWarning,/could not update/);assert.notEqual(oldRevision,commits().at(-1).id,'failed projection remains fenced off by latest commit');assert.equal(row.fields.IdempotencyKey,oldRevision);
console.log('Reminder feed opt-in, responsible roles, revision fence, completion, rework, daily reservation and sync failure tests passed.');
})().catch(e=>{console.error(e);process.exitCode=1;});
