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
  function client(graph, site) {
    async function read(token){return decode(await graph.listItems(token,site.siteId,site.auditListId,2000));}
    async function append(token,event){
      var before=await read(token), existing=before.find(function(e){return e.id===event.id;});if(existing){return before;}
      var related=before.filter(function(e){return e.request===event.request;});
      if(Workflow.revision(related)!==event.base){throw new Error("Shared work changed. Refresh and review before saving.");}
      var encoded=await encode(event);
      for(var i=0;i<encoded.parts.length;i++){await graph.addListItem(token,site.siteId,site.auditListId,encoded.parts[i]);}
      // Recheck after writing parts; another writer may have completed meanwhile.
      var latest=await read(token);if(Workflow.revision(latest.filter(function(e){return e.request===event.request;}))!==event.base){throw new Error("Someone else saved a change. Refresh before continuing; this incomplete write was not committed.");}
      await graph.addListItem(token,site.siteId,site.auditListId,encoded.commit);
      return read(token); // Concurrent commits are detected by the projection and block approvals.
    }
    return {read:read,append:append};
  }
  var api={encode:encode,decode:decode,requests:requests,client:client};
  if(typeof module!=="undefined" && module.exports){module.exports=api;}else{root.LrrWorkflowStore=api;}
})(typeof self!=="undefined" ? self : this);
