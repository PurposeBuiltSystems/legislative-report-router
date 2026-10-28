/* Work queue UI. All writes are explicit; no Teams/email publication is performed here. */
/* global LrrWorkflow, LrrWorkflowStore, document */
(function(root){
  "use strict";
  function esc(v){return String(v==null?"":v).replace(/[&<>"']/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c];});}
  function uid(){return crypto.randomUUID();}
  function mount(host,context){
    var events=[], selected="", loadedSite=null, actor="", busy=false, writable=false, epoch=0;
    host.innerHTML='<h3>Division work queue</h3><p>Track comments and fiscal requests separately. Approvals apply to the registered response version. Changes here do not send messages or submit documents to LSA.</p><div class="actions-row"><button id="wfLoad" type="button" class="primary">Load shared work</button><button id="wfNew" type="button" disabled>New request</button></div><p id="wfInfo" role="status" aria-live="polite">Connect SharePoint in Setup, then load the queue.</p><div class="actions-row"><label>Division <select id="wfDivision"><option value="">All divisions</option></select></label><label>Request type <select id="wfType"><option value="">All types</option><option>Comment</option><option>Fiscal data</option></select></label><label>Stage <select id="wfStage"><option value="">All stages</option><option>Division work pending</option><option>Assigned</option><option>Drafting</option><option>Review requested</option><option>Changes requested</option><option>Director review</option><option>Division approved</option><option>Division review complete</option><option>CFO review</option><option>CFO approved</option><option>Submitted to LSA</option></select></label><label>Due <select id="wfDue"><option value="">Any deadline</option><option value="soon">Due within 2 days / overdue</option><option value="missing">Needs confirmation</option></select></label></div><div id="wfQueue"></div><div id="wfDetail"></div>';
    function el(id){return host.querySelector('#'+id);}
    function val(id){var n=el(id);return n?n.value:"";}
    function info(message){el('wfInfo').textContent=message;}
    function site(){var s=context.site();if(!s||!s.auditListId){throw new Error("Connect the SharePoint audit list in Setup first.");}return s;}
    function identity(token){var a=context.graph.authenticatedAccount(token);if(!a){throw new Error("The signed-in identity could not be verified. Sign in again before changing workflow records.");}return LrrWorkflow.email(a);}
    function today(){var d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');}
    function all(){return LrrWorkflowStore.requests(events);}
    function current(){return all().find(function(s){return s.id===selected;});}
    function link(href,label){return href?'<a href="'+esc(href)+'" target="_blank" rel="noopener noreferrer">'+esc(label)+'</a>':'';}
    function button(id,label){return '<button type="button" id="'+id+'">'+esc(label)+'</button>';}
    function field(id,label,value,type){return '<label>'+esc(label)+'<input id="'+id+'" type="'+(type||'text')+'" value="'+esc(value)+'" /></label>';}
    function area(id,label,value){return '<label>'+esc(label)+'<textarea id="'+id+'" rows="3">'+esc(value)+'</textarea></label>';}
    function bind(id,fn){var n=el(id);if(n){n.addEventListener('click',fn);}}
    function busyState(v){busy=v;host.querySelectorAll('button').forEach(function(b){b.disabled=v;});el('wfNew').disabled=v||!writable;}
    function validSite(){if(!writable||loadedSite!==context.site()){throw new Error("Connection changed or queue is stale. Load shared work again.");}return loadedSite;}
    async function load(){
      if(busy){return;}var generation=++epoch;busyState(true);writable=false;el('wfDetail').innerHTML='';
      try{var s=site(), token=await context.graph.getToken(), a=identity(token), result=await LrrWorkflowStore.client(context.graph,s).read(token);
        if(generation!==epoch||s!==context.site()||identity(token)!==a){throw new Error("Connection or account changed while loading. Load again.");}
        LrrWorkflowStore.requests(result);events=result;loadedSite=s;actor=a;writable=true;
        var previous=val('wfDivision'), divisions=Array.from(new Set(all().flatMap(function(r){return r.config.divisions.map(function(d){return d.code;});}))).sort();
        el('wfDivision').innerHTML='<option value="">All divisions</option>'+divisions.map(function(d){return '<option>'+esc(d)+'</option>';}).join('');el('wfDivision').value=previous;
        info('Loaded '+all().length+' request(s). Signed in as '+actor+'. Due-date reminders below use the confirmed calendar date.');render();
      }catch(e){writable=false;el('wfQueue').textContent='Shared queue unavailable. No workflow changes can be saved.';info(e.message);}finally{busyState(false);}
    }
    async function save(kind,input){
      if(busy){return;}var s;try{s=validSite();}catch(e){info(e.message);return;}
      var generation=epoch;
      var old=current(), expected=old?old.revision:'', request=old?old.id:uid();input.revision=expected;
      busyState(true);
      try{var token=await context.graph.getToken(), a=identity(token);if(a!==actor||s!==context.site()){throw new Error("Account or site changed. Load shared work again.");}
        var fresh=await LrrWorkflowStore.client(context.graph,s).read(token), related=fresh.filter(function(e){return e.request===request;});
        if(LrrWorkflow.revision(related)!==expected){throw new Error("Shared work changed. Load again and review the latest version.");}
        var event=LrrWorkflow.plan(related,kind,input,a,{id:uid(),request:request,at:new Date().toISOString()});
        var guarded={listItems:async function(t,siteId,listId,top){if(generation!==epoch||s!==context.site()||identity(t)!==a){throw new Error('Connection or account changed. Reload shared work.');}return context.graph.listItems(t,siteId,listId,top);},addListItem:async function(t,siteId,listId,fields){if(generation!==epoch||s!==context.site()||identity(t)!==a){throw new Error('Connection or account changed. Reload shared work.');}return context.graph.addListItem(t,siteId,listId,fields);}};
        var result=await LrrWorkflowStore.client(guarded,s).append(token,event);
        if(generation!==epoch||s!==context.site()){throw new Error("Connection changed. The save may have completed on the previous site; reload that site to check.");}
        events=result;selected=request;render();info('Saved '+kind+'. No Teams messages or emails were sent.');
      }catch(e){writable=false;info(e.message+' Load shared work before retrying; check history if the save result was uncertain.');}finally{busyState(false);}
    }
    function render(){
      var rows=all(), div=val('wfDivision'), type=val('wfType'), stage=val('wfStage'), due=val('wfDue'), now=today();
      rows=rows.filter(function(s){var day=s.config.due?Math.round((Date.parse(s.config.due+'T00:00:00Z')-Date.parse(now+'T00:00:00Z'))/86400000):null;return (!div||s.config.divisions.some(function(r){return r.code===div;}))&&(!type||s.config.type===type)&&(!stage||s.stage===stage||s.config.divisions.some(function(r){return (!div||r.code===div)&&s.divisions[r.code].stage===stage;}))&&(!due||(due==='missing'&&!s.config.due)||(due==='soon'&&day!==null&&day<=2&&!s.submitted));});
      rows.sort(function(a,b){return (a.config.due||'9999').localeCompare(b.config.due||'9999');});
      el('wfQueue').innerHTML=rows.length?rows.map(function(s,i){return '<article class="options"><strong>'+esc(s.config.bill)+' · '+esc(s.config.type)+' · '+esc(s.config.label)+'</strong><p>'+esc(s.config.divisions.map(function(d){return d.code;}).join(', '))+' · '+esc(s.stage)+' · '+esc(LrrWorkflow.due(s,now))+' · Legislative: '+esc(s.legislative)+'</p><button type="button" data-work="'+i+'">Open request</button></article>';}).join(''):'<p>No requests match this view. Create a separate request for each comment or fiscal-data cycle.</p>';
      host.querySelectorAll('[data-work]').forEach(function(b){b.addEventListener('click',function(){selected=rows[Number(b.getAttribute('data-work'))].id;detail();});});
      if(selected&&current()){detail();}else{el('wfDetail').innerHTML='';}
    }
    function configForm(s){
      var c=s?s.config:{bill:'',assembly:'91',session:'',label:'',type:'Comment',due:'',dueSource:'',source:'',thread:'',related:'',cfo:'',submitter:actor,divisions:[]};
      return '<details'+(!s?' open':'')+'><summary>'+(s?'Request details and assignments':'New request')+'</summary><p>Owner: '+esc(s?s.owner:actor)+'. Saving assignment changes resets existing approvals. Use signed-in usernames for roles; the broad distribution list is separate.</p>'+field('wfBill','Bill (parent bill for amendments)',c.bill)+field('wfAssembly','General Assembly',c.assembly)+field('wfSession','Session/channel label',c.session)+field('wfLabel','Request label (identify this request/amendment)',c.label)+'<label>Request type <select id="wfKind"><option'+(c.type==='Comment'?' selected':'')+'>Comment</option><option'+(c.type==='Fiscal data'?' selected':'')+'>Fiscal data</option></select></label>'+field('wfDeadline','Confirmed due date (leave blank if unknown)',c.due,'date')+field('wfDeadlineSource','Deadline source / confirmation',c.dueSource)+field('wfSource','Original request link',c.source)+field('wfThread','Existing Teams bill conversation link',c.thread)+field('wfRelated','Related / predecessor / successor bill references',c.related)+field('wfCfo','CFO signed-in email (fiscal data)',c.cfo)+field('wfSubmitter','LSA submitter signed-in email',c.submitter)+area('wfAssignments','One division per line: CODE | coordinator email | director email | reviewer emails separated by commas',c.divisions.map(function(d){return [d.code,d.coordinator,d.director,d.reviewers.join(', ')].join(' | ');}).join('\n'))+button('wfSaveConfig',s?'Save details and reset approvals':'Create request')+'<p class="hint">For the supplied LSA template, five days is relative to the original request; this queue does not invent its request date or assume a holiday convention.</p></details>';
    }
    function configInput(){return {bill:val('wfBill'),assembly:val('wfAssembly'),session:val('wfSession'),label:val('wfLabel'),type:val('wfKind'),due:val('wfDeadline'),dueSource:val('wfDeadlineSource'),source:val('wfSource'),thread:val('wfThread'),related:val('wfRelated'),cfo:val('wfCfo'),submitter:val('wfSubmitter'),divisions:val('wfAssignments').split(/\r?\n/).filter(function(x){return x.trim();}).map(function(line){var a=line.split('|').map(function(x){return x.trim();});if(a.length!==4){throw new Error('Each assignment line needs four fields separated by |.');}return {code:a[0],coordinator:a[1],director:a[2],reviewers:a[3]?a[3].split(',').map(function(x){return x.trim();}):[]};})};}
    function bindConfig(s){bind('wfSaveConfig',function(){try{save(s?'config':'create',{config:configInput()});}catch(e){info(e.message);}});}
    function newRequest(){if(busy||!writable){return;}selected='';el('wfDetail').innerHTML=configForm(null);bindConfig(null);}
    function detail(){
      var s=current();if(!s){return;}var c=s.config;
      var html='<hr><h3>'+esc(c.bill)+' — '+esc(c.label)+'</h3><p>'+esc(c.assembly)+' / '+esc(c.session)+' · '+esc(s.stage)+' · '+esc(LrrWorkflow.due(s,today()))+'</p><p>'+link(c.source,'Original request')+' '+link(c.thread,'Open bill conversation in Teams')+'</p>';
      html+='<p>Legislative progress: '+esc(s.legislative)+' (separate from this request).</p>';
      if(c.related){html+='<p>Related bills: '+esc(c.related)+'</p>';}
      if(s.conflicts.length){html+='<p role="alert">Simultaneous changes were detected. The owner must inspect the history and save the reviewed request details to resolve them; approvals will reset.</p>';}
      if(s.owner===actor){html+=configForm(s);}
      html+='<h4>Division responses</h4>';
      c.divisions.forEach(function(r){var v=s.divisions[r.code];html+='<article class="options"><strong>'+esc(r.code)+' · '+esc(v.stage)+'</strong><p>Coordinator: '+esc(r.coordinator)+'<br>Director: '+esc(r.director)+'<br>Reviewers: '+esc(r.reviewers.join(', ')||'Coordinator only')+'</p>';
        if(v.artifact){html+='<p>Response '+esc(v.artifact.data.version)+' · registered by '+esc(v.artifact.actor)+' at '+esc(v.artifact.at)+'</p>'+v.artifact.data.links.map(function(u,i){return link(u,'Document '+(i+1));}).join(' · ');}
        var next=v.stage==='Director review'?r.director:v.stage==='Review requested'?r.coordinator:v.stage==='Division approved'?(s.ready?(c.type==='Fiscal data'?c.cfo:'Complete'):'Waiting for other divisions'):(r.reviewers.join(', ')||r.coordinator);html+='<p>Next: '+esc(next||'Assignment needed')+'</p></article>';});
      if(c.type==='Fiscal data'){var f=LrrWorkflow.financialSummary(s);html+='<h4>Registered fiscal detail</h4><p>Staff-entered response figures; approval stage is shown above. Revenue and expenditure changes are separate. Missing detail is not zero.</p>';Object.keys(f.groups).sort().forEach(function(key){var g=f.groups[key];html+='<p>'+esc(key)+': $'+esc(g.total.toLocaleString())+' known'+(g.unknown?' + '+g.unknown+' unknown line(s)':'')+'</p>';});html+='<p>'+f.missingDivisions+' division(s) without structured fiscal lines. Full analysis may be in the linked written response.</p>'; }
      html+='<h4>Record a workflow action</h4><label>Division <select id="wfActionDivision">'+c.divisions.map(function(r){return '<option>'+esc(r.code)+'</option>';}).join('')+'</select></label><label>Action <select id="wfAction"><option value="artifact">Register draft / revised response version</option><option value="start">Start drafting</option><option value="review">Request review of registered response</option><option value="changes">Request changes</option><option value="coordinator">Coordinator: consolidate response</option><option value="director">Director: approve current response</option>'+(c.type==='Fiscal data'?'<option value="cfo">CFO: approve all division responses</option><option value="submitted">Record completed LSA submission</option>':'')+'<option value="note">Add FYI / consultation note</option><option value="legislation">Record legislative progress</option></select></label><div id="wfActionFields"></div><label><input type="checkbox" id="wfConfirm" /> I reviewed this action and the current document version(s).</label>'+button('wfSaveAction','Save action')+'<h4>History</h4><p>Intent notes and filename labels do not constitute assignment or approval.</p><div id="wfEvents"></div>';
      el('wfDetail').innerHTML=html;if(s.owner===actor){bindConfig(s);}
      el('wfEvents').innerHTML=s.events.slice().reverse().map(function(e){return '<details><summary>'+esc(e.at)+' · '+esc(e.kind)+(s.ignoredIds.indexOf(e.id)>=0?' [not applied: concurrent edit]':'')+' · '+esc(e.division)+' · '+esc(e.actor)+'</summary><pre style="white-space:pre-wrap;overflow-wrap:anywhere">'+esc(JSON.stringify(e.data,null,2))+'</pre></details>';}).join('');
      el('wfAction').addEventListener('change',actionFields);el('wfActionDivision').addEventListener('change',actionFields);actionFields();
      bind('wfSaveAction',function(){try{if(!el('wfConfirm').checked){info('Review and confirm this action first.');return;}save(val('wfAction'),{division:val('wfActionDivision'),version:val('wfVersion'),links:val('wfDocs'),note:val('wfNote'),financialText:val('wfFinance'),narrative:{summary:val('wfSummary'),background:val('wfBackground'),assumptions:val('wfAssumptions'),impact:val('wfImpact'),sources:val('wfSources')},intent:val('wfIntent'),evidence:val('wfEvidence'),submittedAt:val('wfSubmittedAt')?new Date(val('wfSubmittedAt')).toISOString():'',status:val('wfLegislative'),date:val('wfEventDate')});}catch(e){info(e.message);}});
    }
    function actionFields(){
      var kind=val('wfAction'), s=current(), v=s.divisions[val('wfActionDivision')], a=v.artifact?v.artifact.data:{}, n=a.narrative||{}, html='';el('wfConfirm').checked=false;
      if(kind==='artifact'){
        html=field('wfVersion','Document version label',a.version||'')+area('wfDocs','Exact document-version links (one per line)',(a.links||[]).join('\n'))+area('wfNote','Response / revision description',a.note||'');
        if(s.config.type==='Fiscal data'){html+='<details><summary>Fiscal response detail (or keep full analysis in linked form)</summary>'+area('wfSummary','Bill summary',n.summary)+area('wfBackground','Background / implementation',n.background)+area('wfAssumptions','Assumptions, including reason for no impact',n.assumptions)+area('wfImpact','Fiscal / operational impact, funds, FTE and local impacts',n.impact)+area('wfSources','Related data and sources',n.sources)+area('wfFinance','Optional financial lines — paste seven tab-separated columns: FY2027, Revenue/Expenditure, Increase/Decrease, amount or Unknown, fund, One-time/Ongoing, description',a.financialText||'')+'<p>Do not paste template zeros as confirmed estimates. Revenue and expenditure changes stay separate; these entries do not replace the required written response.</p></details>';}
      }else if(kind==='note'){html='<label>Intent <select id="wfIntent"><option>FYI</option><option>Consultation</option><option>Action requested</option><option>Review requested</option><option>Approval requested</option></select></label>'+area('wfNote','Note','');}
      else if(kind==='changes'){html=area('wfNote','Changes required','');}
      else if(kind==='submitted'){html=field('wfSubmittedAt','Actual completed submission time (local time)','','datetime-local')+field('wfEvidence','Submission receipt or source message link','')+area('wfNote','Documents actually submitted','');}
      else if(kind==='legislation'){html=field('wfLegislative','Legislative status / action','')+field('wfEventDate','Event date','','date')+field('wfEvidence','Official source or source message link','')+'<p>Legislative progress is recorded independently; it does not complete or approve this request.</p>';}
      else {html='<p>Current response: '+esc(v.artifact?v.artifact.data.version:'No registered response')+'. '+(kind==='cfo'?'Requires current director approvals for every required division.':kind==='director'?'Requires coordinator consolidation of this exact version.':'')+'</p>';}
      el('wfActionFields').innerHTML=html;
    }
    bind('wfLoad',load);bind('wfNew',newRequest);['wfDivision','wfType','wfStage','wfDue'].forEach(function(id){el(id).addEventListener('change',render);});
    return {load:load,invalidate:function(){epoch++;writable=false;loadedSite=null;events=[];selected='';actor='';el('wfQueue').innerHTML='';el('wfDetail').innerHTML='';el('wfNew').disabled=true;info('Connection changed. Load shared work again.');}};
  }
  root.LrrWorkflowUI={mount:mount};
})(typeof self!=="undefined"?self:this);
