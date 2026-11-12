/* Read-only collection for existing, issued fiscal requests. No channel discovery or writes. */
/* global LrrWorkflow, LrrWorkflowStore, LrrHarvest */
(function (root) {
  "use strict";
  var W = typeof module !== "undefined" && module.exports ? require("./workflow") : LrrWorkflow;
  var S = typeof module !== "undefined" && module.exports ? require("./workflow-store") : LrrWorkflowStore;
  var H = typeof module !== "undefined" && module.exports ? require("./harvest") : LrrHarvest;
  function thread(url) {
    try {
      var u = new URL(url), m = /^\/l\/message\/([^/]+)\/([^/]+)\/?$/.exec(u.pathname);
      if (u.protocol !== "https:" || !/^(teams\.microsoft\.com|gov\.teams\.microsoft\.us|dod\.teams\.microsoft\.us)$/.test(u.hostname) || !m) { return null; }
      var t = {teamId:u.searchParams.get("groupId"), channelId:decodeURIComponent(m[1]), messageId:u.searchParams.get("parentMessageId") || decodeURIComponent(m[2])};
      return t.teamId && t.channelId && t.messageId ? t : null;
    } catch (e) { return null; }
  }
  function key(t) { return JSON.stringify([t.teamId,t.channelId,t.messageId]); }
  function rowThread(f) { return {teamId:f.TeamId,channelId:f.ChannelId,messageId:f.TeamsMessageId}; }
  function sameSession(record, config) {
    var iowa=/^Iowa \/ (\d{4})$/.exec(record.ReportKey||"");
    if(iowa && String(91+Math.floor((Number(iowa[1])-2025)/2))!==config.assembly){return false;}
    return record.ReportKey===config.session || !!(iowa && config.session===iowa[1]);
  }
  function manages(s, actor) { return s.owner === actor || s.config.cfo === actor || s.config.submitter === actor || s.config.divisions.some(function(r){return r.coordinator===actor || r.director===actor || r.reviewers.indexOf(actor)>=0;}); }
  function writer(s, division, actor) { var r=s.config.divisions.find(function(d){return d.code===division;});return !!r && (r.coordinator===actor || r.reviewers.indexOf(actor)>=0); }
  function scope(events, rows, actor) {
    var eligible=[], excluded=[];
    S.requests(events).filter(function(s){return s.config.type==="Fiscal data" && manages(s,actor);}).forEach(function(s){
      var t=thread(s.config.thread), reason="", started=Date.parse(s.events[0].at);
      if(s.submitted || s.conflicts.length) { reason="Completed submission or unresolved concurrent changes"; }
      else if(!t) { reason="No exact supported Teams conversation link"; }
      var record=t && rows.find(function(row){var f=row.fields||{};return f.Status==="bill-thread" && key(rowThread(f))===key(t) && f.Title===s.config.bill && sameSession(f,s.config);});
      if(!reason && !record) { reason="No matching router bill/session publication record"; }
      // A comment post alone does not prove that a fiscal request was issued.
      var publications=t ? rows.filter(function(row){var f=row.fields||{};return f.Status==="published" && /\|fiscal-note$/.test(f.ReportKey||"") && key(rowThread(f))===key(t) && f.Title===s.config.bill && Number.isFinite(Date.parse(row.createdDateTime));}) : [];
      if(!reason && (!publications.length || !Number.isFinite(started))) { reason="No dated successful fiscal-request publication"; }
      if(reason) { excluded.push({request:s.id,label:s.config.bill+" — "+s.config.label,reason:reason});return; }
      var issued=Math.min.apply(null,publications.map(function(row){return Date.parse(row.createdDateTime);}));
      eligible.push({state:s,thread:t,since:Math.max(started,issued)});
    });
    return {eligible:eligible,excluded:excluded};
  }
  async function fingerprint(value) {
    var bytes=await crypto.subtle.digest("SHA-256",new TextEncoder().encode(JSON.stringify(value)));
    return Array.from(new Uint8Array(bytes),function(n){return n.toString(16).padStart(2,"0");}).join("");
  }
  async function collect(options) {
    var graph=options.graph, token=options.token, check=options.check || function(){}, rows=options.rows;
    check();var events=await S.decode(rows), scoped=scope(events,rows,options.actor), groups=new Map(), candidates=[], warnings=[], skipped=0;
    scoped.eligible.forEach(function(r){var k=key(r.thread);if(!groups.has(k)){groups.set(k,{thread:r.thread,requests:[]});}groups.get(k).requests.push(r);});
    if(groups.size>120) { throw new Error("More than 120 issued fiscal threads qualify. Narrow the connected queue before collecting; nothing was collected."); }
    var groupsList=Array.from(groups.values());
    for(var i=0;i<groupsList.length;i++) {
      check();var group=groupsList[i], t=group.thread, replies;
      if(options.progress) { options.progress("Reading issued fiscal thread "+(i+1)+" of "+groupsList.length+"…"); }
      try { replies=await graph.listReplies(token,t.teamId,t.channelId,t.messageId);check(); }
      catch(e) { check();warnings.push("Could not read "+group.requests[0].state.config.bill+": "+e.message);continue; }
      for(var j=0;j<replies.length;j++) {
        check();var rep=replies[j], date=Date.parse(rep.createdDateTime);
        if(!rep.id || rep.deletedDateTime || !Number.isFinite(date)) { continue; }
        // Exclude the router's own published request/update messages.
        if(rows.some(function(row){var f=row.fields||{};return f.Status==="delivery-link" && f.TeamId===t.teamId && f.ChannelId===t.channelId && f.TeamsMessageId===rep.id;})) { continue; }
        var choices=group.requests.filter(function(r){return date>=r.since;});if(!choices.length){continue;}
        var body=H.htmlToText((rep.body||{}).content||""), files=[], evidence=[], attachmentMarks=[];
        var author=((rep.from||{}).user||{}).displayName || "Unknown author";
        var replyKey=JSON.stringify([t.teamId,t.channelId,t.messageId,rep.id]);
        var commentLink=options.commentLink(t,rep.id), attachments=rep.attachments||[];
        var bodyEvidence=H.analyzeSource({text:body,source:"Teams reply",author:author,date:rep.createdDateTime});
        if(bodyEvidence){evidence.push(bodyEvidence);}
        for(var a=0;a<attachments.length;a++) {
          var att=attachments[a], url;
          try{url=new URL(att.contentUrl);}catch(e){url=null;}
          if(!url || url.protocol!=="https:"){continue;}
          files.push(url.href);
          if(a>=5 || !/\.(docx|xlsx)$/i.test(att.name||"")) { attachmentMarks.push([att.name,url.href,"not extracted"]);continue; }
          try {
            check();var item=await graph.driveItemFromUrl(token,url.href);check();
            attachmentMarks.push([url.href,item.id,item.eTag||"",item.lastModifiedDateTime||""]);
            if(!item.parentReference || !item.parentReference.driveId || !Number.isFinite(item.size) || item.size>20*1024*1024) { throw new Error("File metadata missing or file exceeds 20 MB"); }
            var b64=await graph.driveItemContentB64(token,item.parentReference.driveId,item.id);check();
            var extracted=await options.extract(b64,att.name,item.size);check();
            var found=H.analyzeSource({text:extracted,source:att.name,kind:"attachment",author:author,date:rep.createdDateTime});if(found){evidence.push(found);}
          } catch(e) { check();attachmentMarks.push([url.href,"extraction failed"]);warnings.push("Attachment "+(att.name||"file")+" on "+commentLink+": "+e.message); }
        }
        if(!body && !files.length){continue;}
        var hash=await fingerprint([rep.lastModifiedDateTime||rep.createdDateTime,body,attachmentMarks,files]);check();
        choices=choices.filter(function(r){return !r.state.appliedEvents.some(function(e){var intake=(e.data||{}).intake;return e.kind==="artifact" && intake && intake.key===replyKey && intake.fingerprint===hash;});});
        if(!choices.length){skipped++;continue;}
        candidates.push({key:replyKey,fingerprint:hash,commentLink:commentLink,body:body,author:author,date:rep.createdDateTime,files:files,evidence:evidence,
          requests:choices.map(function(r){return {id:r.state.id,revision:r.state.revision,bill:r.state.config.bill,label:r.state.config.label,assembly:r.state.config.assembly,session:r.state.config.session,divisions:r.state.config.divisions.map(function(d){return {code:d.code,canRegister:writer(r.state,d.code,options.actor)};})};})});
      }
    }
    candidates.sort(function(a,b){return b.date.localeCompare(a.date);});
    return {events:events,candidates:candidates,excluded:scoped.excluded,warnings:warnings,threads:groups.size,requests:scoped.eligible.length,skipped:skipped};
  }
  function prepare(candidate, state, division, actor) {
    var match=candidate.requests.find(function(r){return r.id===state.id;});
    if(!match || match.revision!==state.revision || state.config.type!=="Fiscal data" || state.submitted || state.conflicts.length) { throw new Error("The fiscal request changed or is no longer eligible. Collect again before preparing this reply."); }
    if(!writer(state,division,actor)) { throw new Error("Choose a division for which you are the assigned reviewer or coordinator. Other staff must open and confirm their own response."); }
    var hints=candidate.evidence.map(function(e){return [e.source,e.fy||"Fiscal year unspecified",e.costRaw||"Amount unconfirmed",e.evidence].filter(Boolean).join(" — ");}).join("\n");
    return {version:"Teams response "+candidate.date.slice(0,19).replace("T"," "),links:candidate.files.join("\n"),
      note:("Collected from "+candidate.author+" at "+candidate.date+"\n"+candidate.commentLink+"\n"+candidate.body).slice(0,4500),
      narrative:{impact:candidate.body.slice(0,3500),sources:(candidate.commentLink+"\nExtracted clues — verify against the response:\n"+hints).slice(0,2000)},financialText:"",
      intake:{key:candidate.key,fingerprint:candidate.fingerprint,commentLink:candidate.commentLink,date:candidate.date}};
  }
  var api={thread:thread,scope:scope,collect:collect,prepare:prepare};
  if(typeof module!=="undefined" && module.exports){module.exports=api;}else{root.LrrFiscalIntake=api;}
})(typeof self!=="undefined"?self:this);
