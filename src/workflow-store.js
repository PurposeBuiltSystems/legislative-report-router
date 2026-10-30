/* Append-only workflow events in existing audit columns. Commit last; partial writes never become approvals. */
/* global LrrWorkflow */
(function (root) {
  "use strict";
  var PAYLOAD_FIELDS = ["Error", "SourceSubject", "Divisions", "EmailRecipients", "TeamId", "ChannelId", "TeamsMessageId"];
  var Workflow = typeof module!=="undefined" && module.exports ? require("./workflow") : LrrWorkflow;
  async function digest(value) { var h=await crypto.subtle.digest("SHA-256",new TextEncoder().encode(value)); return Array.from(new Uint8Array(h),function(n){return n.toString(16).padStart(2,"0");}).join(""); }
  async function encode(e) {
    var json=JSON.stringify(e); if (json.length>16000) { throw new Error("Request details are too long. Link supporting analysis instead of pasting its full contents."); }
    var parts=[];
    for(var i=0;i<json.length;i+=1680) { var fields={Title:"Workflow data", Status:"workflow-part", IdempotencyKey:e.id, ReportKey:String(parts.length)}; PAYLOAD_FIELDS.forEach(function(name,index){ fields[name]=json.slice(i+index*240,i+(index+1)*240); }); parts.push(fields); }
    return {parts:parts, commit:{Title:"Workflow "+e.kind,Status:"workflow-event",IdempotencyKey:e.id,ReportKey:e.request,PublishedBy:e.actor,SourceSubject:JSON.stringify({v:1,n:parts.length,sha:await digest(json)})}};
  }
  async function decode(rows) {
    var parts=Object.create(null), events=[], seen=Object.create(null);
    rows.forEach(function(row){var f=row.fields||{};if(f.Status==="workflow-part"){var key=f.IdempotencyKey+":"+f.ReportKey;var payload=PAYLOAD_FIELDS.map(function(name){return f[name]||"";}).join("");if(parts[key]!==undefined && parts[key]!==payload){throw new Error("Conflicting workflow data. Shared history needs repair before approvals can continue.");}parts[key]=payload;}});
    for(var i=0;i<rows.length;i++){
      var row=rows[i], f=row.fields||{};if(f.Status!=="workflow-event"){continue;}
      var m;try{m=JSON.parse(f.SourceSubject);}catch(e){throw new Error("Unreadable workflow commit; history needs repair.");}
      if(m.v!==1 || !Number.isInteger(m.n) || m.n<1 || m.n>80){throw new Error("Unsupported workflow history format.");}
      var json="";for(var n=0;n<m.n;n++){var part=parts[f.IdempotencyKey+":"+n];if(typeof part!=="string"){throw new Error("Incomplete committed workflow history. Refresh before continuing.");}json+=part;}
      if(await digest(json)!==m.sha){throw new Error("Workflow history checksum failed. No approval changes can be saved.");}
      if(seen[f.IdempotencyKey]){if(seen[f.IdempotencyKey].json!==json){throw new Error("Conflicting workflow commits.");}seen[f.IdempotencyKey].event.order=Math.min(seen[f.IdempotencyKey].event.order,Number(row.id));continue;}
      var event=JSON.parse(json);if(event.id!==f.IdempotencyKey || event.request!==f.ReportKey || event.actor!==f.PublishedBy || !event.data || !event.kind || !Number.isFinite(Number(row.id))){throw new Error("Invalid workflow event identity.");}
      event.order=Number(row.id); seen[f.IdempotencyKey]={json:json,event:event}; events.push(event);
    }
    return events;
  }
  function requests(events) { var groups=Object.create(null); events.forEach(function(e){(groups[e.request]||(groups[e.request]=[])).push(e);});var result=Object.keys(groups).map(function(k){return Workflow.project(groups[k]);});
    result.forEach(function(s){var latest=null;result.forEach(function(other){if(other.config.bill===s.config.bill && other.config.assembly===s.config.assembly && other.config.session===s.config.session && !other.conflicts.length){other.appliedEvents.forEach(function(e){if(e.kind==="legislation" && (!latest || e.order>latest.order)){latest=e;}});}});if(latest){s.legislative=latest.data.status;}});return result; }
  /** Flat, current reminder records for a standard SharePoint/Outlook cloud flow.
   * The newest committed workflow row ID is the freshness fence checked by the flow.
   * Existing PublishedBy (the daily send reservation) is never overwritten by sync.
   */
  function reminderRecords(events) {
    var out=[];
    requests(events).forEach(function(state){
      if(!state.config.due){return;}
      var revision=String(Math.max.apply(null,state.events.map(function(e){return e.order;})));
      Workflow.reminders([state],state.owner,state.config.due).forEach(function(group){group.items.forEach(function(item){
        var link=item.thread||item.source||"";if(link.length>480){link="";}
        out.push({key:state.id+"|"+item.division,fields:{Title:item.bill,Status:"reminder-pending",ReportKey:state.id,
          IdempotencyKey:revision,Divisions:item.division,EmailRecipients:group.to,Error:item.due,
          SourceSubject:item.label.slice(0,240),TeamId:item.task.slice(0,240),ChannelId:link.slice(0,240),TeamsMessageId:link.slice(240)}});
      });});
    });
    return out;
  }
  function reminderControl(rows){return rows.filter(function(r){return (r.fields||{}).Status==="reminder-control";}).sort(function(a,b){return Number(b.id)-Number(a.id);})[0];}
  async function syncReminderRecords(graph,site,token,events,rows){
    var control=reminderControl(rows);if(!control){return {enabled:false,pending:0};}
    var enabled=control.fields.Error==="enabled", desired=enabled?reminderRecords(events):[], current=rows.filter(function(r){return /^reminder-(pending|closed)$/.test((r.fields||{}).Status||"");});
    var keep=Object.create(null);
    for(var i=0;i<desired.length;i++){
      var record=desired[i],existing=current.find(function(r){return (r.fields||{}).ReportKey===record.fields.ReportKey && (r.fields||{}).Divisions===record.fields.Divisions;});
      if(existing){keep[existing.id]=true;var changed=Object.keys(record.fields).some(function(k){return (existing.fields[k]||"")!==(record.fields[k]||"");});if(changed){await graph.updateListItemFields(token,site.siteId,site.auditListId,existing.id,record.fields);}}
      else{await graph.addListItem(token,site.siteId,site.auditListId,record.fields);}
    }
    for(var j=0;j<current.length;j++){if(!keep[current[j].id]&&current[j].fields.Status==="reminder-pending"){await graph.updateListItemFields(token,site.siteId,site.auditListId,current[j].id,{Status:"reminder-closed"});}}
    return {enabled:enabled,pending:desired.length};
  }
  function client(graph, site) {
    var lastRows=[];
    async function read(token){lastRows=await graph.listItems(token,site.siteId,site.auditListId,2000);return decode(lastRows);}
    async function configureReminders(token,enabled){
      await read(token);var control=reminderControl(lastRows),fields={Title:"Scheduled reminder control",Status:"reminder-control",Error:enabled?"enabled":"disabled"};
      if(control){await graph.updateListItemFields(token,site.siteId,site.auditListId,control.id,fields);}else{await graph.addListItem(token,site.siteId,site.auditListId,fields);}
      var events=await read(token);return syncReminderRecords(graph,site,token,events,lastRows);
    }
    async function refreshReminders(token){var events=await read(token);return syncReminderRecords(graph,site,token,events,lastRows);}
    async function append(token,event){
      var before=await read(token), existing=before.find(function(e){return e.id===event.id;});if(existing){return before;}
      var related=before.filter(function(e){return e.request===event.request;});
      if(Workflow.revision(related)!==event.base){throw new Error("Shared work changed. Refresh and review before saving.");}
      var encoded=await encode(event);
      for(var i=0;i<encoded.parts.length;i++){await graph.addListItem(token,site.siteId,site.auditListId,encoded.parts[i]);}
      // Recheck after writing parts; another writer may have completed meanwhile.
      var latest=await read(token);if(Workflow.revision(latest.filter(function(e){return e.request===event.request;}))!==event.base){throw new Error("Someone else saved a change. Refresh before continuing; this incomplete write was not committed.");}
      await graph.addListItem(token,site.siteId,site.auditListId,encoded.commit);
      var saved=await read(token); // Concurrent commits are detected by the projection and block approvals.
      try{await syncReminderRecords(graph,site,token,saved,lastRows);}catch(e){saved.reminderWarning="Workflow saved, but the automatic reminder feed could not update: "+e.message+". The flow must skip records whose revision is no longer current.";}
      return saved;
    }
    return {read:read,append:append,configureReminders:configureReminders,refreshReminders:refreshReminders};
  }
  var api={encode:encode,decode:decode,requests:requests,client:client,reminderRecords:reminderRecords};
  if(typeof module!=="undefined" && module.exports){module.exports=api;}else{root.LrrWorkflowStore=api;}
})(typeof self!=="undefined" ? self : this);
