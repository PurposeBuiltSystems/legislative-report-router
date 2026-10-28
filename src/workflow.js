/* Shared legislative work requests. Pure transitions; no notification or mail APIs. */
(function (root) {
  "use strict";
  function fail(message) { throw new Error(message); }
  function text(v) { return String(v == null ? "" : v).trim(); }
  function email(v) { var s = text(v).toLowerCase(); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s)) { fail("Use a signed-in email address for each assigned role."); } return s; }
  function bill(v) { var s = text(v).toUpperCase().replace(/\s/g, ""); if (!/^(?:HF|SF|HSB|SSB|HJR|SJR|HR|SR|HCR|SCR)\d+$/.test(s)) { fail("Enter a bill identifier such as SSB3108 or HF2586. Link amendments to their parent bill."); } return s; }
  function url(v) { var s = text(v); if (s && !/^https:\/\/[^\s]+$/i.test(s)) { fail("Document and source links must start with https://."); } return s; }
  function date(v) { var s = text(v); if (s && (!/^\d{4}-\d{2}-\d{2}$/.test(s) || new Date(s + "T00:00:00Z").toISOString().slice(0,10) !== s)) { fail("Enter a valid calendar date."); } return s; }
  function config(data) {
    var d = JSON.parse(JSON.stringify(data || {}));
    d.bill = bill(d.bill); d.assembly = text(d.assembly); d.session = text(d.session); d.label = text(d.label);
    if (!/^\d{1,3}$/.test(d.assembly) || !d.session || !d.label) { fail("Assembly, session and request label are required."); }
    if (["Comment", "Fiscal data"].indexOf(d.type) < 0) { fail("Choose Comment or Fiscal data."); }
    d.due = date(d.due); d.dueSource = text(d.dueSource); d.source = url(d.source); d.thread = url(d.thread);
    if (d.due && !d.dueSource) { fail("Describe the source of the confirmed deadline."); }
    d.related = text(d.related); d.cfo = d.cfo ? email(d.cfo) : ""; d.submitter = email(d.submitter);
    var seen = {};
    if (!Array.isArray(d.divisions) || !d.divisions.length) { fail("Assign at least one division."); }
    d.divisions = d.divisions.map(function (r) {
      var code = text(r.code).toUpperCase();
      if (!/^[A-Z0-9_-]{1,30}$/.test(code) || seen[code]) { fail("Division codes must be unique."); } seen[code] = true;
      return { code: code, coordinator: email(r.coordinator), director: email(r.director), reviewers: (r.reviewers || []).map(email) };
    });
    return d;
  }
  function finance(tsv) {
    return text(tsv).split(/\r?\n/).filter(Boolean).map(function (line, index) {
      var a = line.split("\t").map(text);
      if (a.length !== 7 || !/^FY\d{4}$/.test(a[0]) || ["Revenue", "Expenditure"].indexOf(a[1]) < 0 || ["Increase", "Decrease"].indexOf(a[2]) < 0 || ["One-time", "Ongoing"].indexOf(a[5]) < 0 || !a[4] || !a[6]) { fail("Financial line " + (index+1) + ": use FY2027, Revenue/Expenditure, Increase/Decrease, amount or Unknown, fund, One-time/Ongoing, description (tab separated)."); }
      var amount = a[3] === "Unknown" ? null : Number(a[3]);
      if (amount !== null && (!/^\d+(?:\.\d{1,2})?$/.test(a[3]) || !Number.isFinite(amount))) { fail("Amounts must be nonnegative dollars without commas, or Unknown."); }
      return { fy:a[0], kind:a[1], direction:a[2], amount:amount, fund:a[4], frequency:a[5], description:a[6] };
    });
  }
  function artifact(data) {
    var d = JSON.parse(JSON.stringify(data || {})); d.version = text(d.version);
    d.links = text(d.links).split(/\r?\n/).filter(Boolean).map(url);
    if (!d.version || !d.links.length) { fail("A version label and at least one document link are required. Link the exact file version reviewed."); }
    d.note = text(d.note); d.narrative = d.narrative || {};
    ["summary","background","assumptions","impact","sources"].forEach(function (k) { d.narrative[k] = text(d.narrative[k]); });
    d.finance = finance(d.financialText); d.financialText = text(d.financialText);
    return d;
  }
  function financialSummary(s) {
    var groups=Object.create(null), missing=0;
    s.config.divisions.forEach(function(r){var a=s.divisions[r.code].artifact;if(!a || !a.data.finance.length){missing++;return;}a.data.finance.forEach(function(line){var key=[line.fy,line.kind,line.direction].join(" / ");if(!groups[key]){groups[key]={total:0,unknown:0};}if(line.amount===null){groups[key].unknown++;}else{groups[key].total+=line.amount;}});});
    return {groups:groups,missingDivisions:missing};
  }
  function ordered(events) { return events.slice().sort(function (a,b) { return Number(a.order)-Number(b.order) || a.id.localeCompare(b.id); }); }
  function revision(events) { return events.map(function (e) { return e.id; }).sort().join(","); }
  function project(events) {
    events = ordered(events);
    var s = null, visited = [], applied = [], ignored = [];
    events.forEach(function (e) {
      var d = e.data || {};
      if (s && e.base !== revision(visited)) { s.conflicts.push(e.id); ignored.push(e.id); visited.push(e); return; }
      visited.push(e); applied.push(e);
      if (e.kind === "create") {
        if (s) { fail("Duplicate request creation in shared history."); }
        s = { id:e.request, owner:e.actor, config:d, configId:e.id, divisions:{}, events:events, revision:revision(events), cfo:null, submitted:null, legislative:"Not recorded", conflicts:[], appliedEvents:applied, ignoredIds:ignored };
      }
      if (!s) { fail("Request history is incomplete."); }
      if (e.kind === "config") { s.config=d; s.configId=e.id; s.cfo=null; s.submitted=null; s.conflicts=[]; }
      if (e.kind === "create" || e.kind === "config") {
        s.config.divisions.forEach(function (r) { if (!s.divisions[r.code]) { s.divisions[r.code]={ stage:"Assigned", artifact:null, coordinator:null, director:null }; } });
        Object.keys(s.divisions).forEach(function (k) { s.divisions[k].coordinator=null; s.divisions[k].director=null; s.divisions[k].stage=s.divisions[k].artifact ? "Drafting" : "Assigned"; });
      }
      var v = s.divisions[e.division];
      if (e.kind === "artifact" && v) { v.artifact={ id:e.id, data:d, actor:e.actor, at:e.at }; v.stage="Drafting"; v.coordinator=null; v.director=null; s.cfo=null; s.submitted=null; }
      if (["start","review","changes"].indexOf(e.kind) >= 0 && v) {
        v.stage={start:"Drafting", review:"Review requested", changes:"Changes requested"}[e.kind];
        v.coordinator=null; v.director=null; s.cfo=null; s.submitted=null;
      }
      if (e.kind === "coordinator" && v) { v.coordinator=e; v.director=null; v.stage="Director review"; s.cfo=null; s.submitted=null; }
      if (e.kind === "director" && v) { v.director=e; v.stage="Division approved"; s.cfo=null; s.submitted=null; }
      if (e.kind === "cfo") { s.cfo=e; s.submitted=null; }
      if (e.kind === "submitted") { s.submitted=e; }
      if (e.kind === "legislation") { s.legislative=d.status; }
    });
    if (s) {
      s.ready = !s.conflicts.length && s.config.divisions.every(function (r) { var v=s.divisions[r.code]; return !!(v.director && v.artifact && v.director.data.artifact === v.artifact.id && v.director.data.config === s.configId); });
      if (s.conflicts.length) { s.cfo=null; s.submitted=null; }
      s.stage = s.conflicts.length ? "Concurrent changes — owner review needed" : s.submitted ? "Submitted to LSA" : s.cfo ? "CFO approved" : s.ready ? (s.config.type === "Fiscal data" ? "CFO review" : "Division review complete") : "Division work pending";
    }
    return s;
  }
  function plan(events, kind, input, actor, meta) {
    actor=email(actor); input=input || {}; var s=project(events), d={}, division=text(input.division).toUpperCase();
    if (!s && kind !== "create") { fail("Create the request first."); }
    if (s && kind === "create") { fail("This request already exists."); }
    if (s && input.revision !== s.revision) { fail("Shared work changed. Refresh and review the latest version before saving."); }
    if (s && s.conflicts.length && kind !== "config" && kind !== "note") { fail("Concurrent edits need owner review. The owner must review history and save request details again; approvals will reset."); }
    var r=s && s.config.divisions.find(function (x) { return x.code===division; }), v=s && s.divisions[division];
    var coordinator=!!(r && r.coordinator===actor), director=!!(r && r.director===actor), writer=!!(r && (coordinator || r.reviewers.indexOf(actor)>=0));
    if (kind === "create") { d=config(input.config); }
    else if (kind === "config") { if (actor!==s.owner) { fail("Only the request owner can change assignments and request details."); } d=config(input.config); if (d.bill!==s.config.bill || d.assembly!==s.config.assembly || d.session!==s.config.session || d.type!==s.config.type) { fail("Bill, assembly, session and request type cannot change; create a separate request."); } }
    else if (kind === "artifact") { if (!writer) { fail("Only assigned reviewers or the division coordinator can register a response version."); } d=artifact(input); }
    else if (kind === "start") { if (!writer) { fail("Only assigned reviewers or the coordinator can start division work."); } }
    else if (kind === "review") { if (!writer || !v.artifact) { fail("An assigned reviewer must register a response version before requesting review."); } if (v.stage === "Changes requested") { fail("Register a revised response version before requesting review again."); } }
    else if (kind === "changes") { if (!(coordinator || director || (s.config.type === "Fiscal data" && actor===s.config.cfo)) || !v || !v.artifact) { fail("Only an assigned coordinator, director or CFO can request changes to a response."); } d.note=text(input.note); if (!d.note) { fail("Describe the changes needed."); } }
    else if (kind === "coordinator") { if (!coordinator || v.stage!=="Review requested" || !v.artifact) { fail("The assigned coordinator must consolidate the current response after review is requested."); } d={ artifact:v.artifact.id, config:s.configId }; }
    else if (kind === "director") { if (!director || !v.coordinator || v.stage!=="Director review" || !v.artifact) { fail("The assigned director can approve only the current consolidated response."); } d={ artifact:v.artifact.id, config:s.configId, coordinator:v.coordinator.id }; }
    else if (kind === "cfo") { if (s.config.type!=="Fiscal data" || actor!==s.config.cfo || !s.ready) { fail("CFO approval requires the assigned CFO and current director approval from every required division."); } d={ approvals:s.config.divisions.map(function (x) { return s.divisions[x.code].director.id; }), config:s.configId }; }
    else if (kind === "submitted") { if (s.config.type!=="Fiscal data" || actor!==s.config.submitter || !s.ready || !s.cfo) { fail("Only the designated submitter can record LSA submission after all division and CFO approvals."); } d={ note:text(input.note), at:text(input.submittedAt), evidence:url(input.evidence), approval:s.cfo.id }; if (!d.note || !d.evidence || !d.at || !Number.isFinite(Date.parse(d.at)) || Date.parse(d.at)>Date.now()+60000) { fail("Record the actual submission time, document description and evidence link."); } }
    else if (kind === "note") { d={ note:text(input.note), intent:text(input.intent) }; if (!d.note || ["FYI","Consultation","Action requested","Review requested","Approval requested"].indexOf(d.intent)<0) { fail("Choose a note intent and enter its text. Notes do not assign roles or approve documents."); } }
    else if (kind === "legislation") { if (actor!==s.owner) { fail("Only the request owner can record legislative progress here."); } d={ status:text(input.status), evidence:url(input.evidence), date:date(input.date) }; if (!d.status || !d.evidence || !d.date) { fail("Legislative events need a status, date and source link."); } }
    else { fail("Unknown workflow action."); }
    return { id:meta.id, request:s ? s.id : meta.request, kind:kind, division:division, actor:actor, at:meta.at, base:s ? s.revision : "", data:d };
  }
  function due(s, today) { if (!s.config.due) { return "Deadline needs confirmation"; } if (s.submitted) { return "Submitted " + s.submitted.data.at; } var n=Math.round((Date.parse(s.config.due+"T00:00:00Z")-Date.parse(today+"T00:00:00Z"))/86400000); return n<0 ? Math.abs(n)+" day(s) overdue" : n===0 ? "Due today" : "Due in "+n+" day(s)"; }
  var api={config:config, finance:finance, artifact:artifact, project:project, plan:plan, revision:revision, due:due, email:email, financialSummary:financialSummary};
  if (typeof module!=="undefined" && module.exports) { module.exports=api; } else { root.LrrWorkflow=api; }
})(typeof self!=="undefined" ? self : this);
