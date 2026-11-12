/* Work queue UI. All writes are explicit; no Teams/email publication is performed here. */
/* global LrrWorkflow, LrrWorkflowStore, LrrStaffConfig, document */
(function(root){
  "use strict";
  function esc(v){return String(v==null?"":v).replace(/[&<>"']/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c];});}
  function uid(){return crypto.randomUUID();}
  function mount(host,context){
    var events=[], selected="", loadedSite=null, actor="", busy=false, writable=false, epoch=0, defaultsUsed=false, editing=false, intake=null, collection=null;
    host.innerHTML='<h3>Division work queue</h3><p>Track comments and fiscal requests separately. Approvals apply to the registered response version. Saving does not send messages or submit documents to LSA. Collected fiscal responses can prepare a coordinator notification draft for you to review and send. If connected, scheduled reminders use the current assignments.</p><div class="actions-row"><button id="wfLoad" type="button" class="primary">Load shared work</button><button id="wfNew" type="button" disabled>New request</button><button id="wfReminders" type="button" disabled>Draft due-date reminders</button><button id="wfAutoReminders" type="button" disabled>Connect automatic reminders</button><button id="wfPauseReminders" type="button" disabled>Pause automatic reminders</button><button id="wfDiscard" type="button" hidden>Discard edits</button></div><p id="wfInfo" role="status" aria-live="polite">Connect SharePoint in Setup, then load the queue.</p><div class="actions-row"><label>Session <select id="wfSessionFilter"><option value="">All sessions</option></select></label><label>Division <select id="wfDivision"><option value="">All divisions</option></select></label><label>Request type <select id="wfType"><option value="">All types</option><option>Comment</option><option>Fiscal data</option></select></label><label>Stage <select id="wfStage"><option value="">All stages</option><option>Division work pending</option><option>Assigned</option><option>Drafting</option><option>Review requested</option><option>Changes requested</option><option>Director review</option><option>Division approved</option><option>Division review complete</option><option>CFO review</option><option>CFO approved</option><option>Submitted to LSA</option></select></label><label>Due <select id="wfDue"><option value="">Any deadline</option><option value="soon">Due within 2 days / overdue</option><option value="missing">Needs confirmation</option></select></label></div><label><input type="checkbox" id="wfMine" /> My assigned work</label><div id="wfTeamsSetup"></div><div id="wfQueue"></div><div id="wfDetail"></div>';
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
    if(context.collectFiscalResponses){
      var collectButton=document.createElement('button');collectButton.id='wfCollectFiscal';collectButton.type='button';collectButton.textContent='Collect fiscal responses';collectButton.disabled=true;el('wfLoad').parentNode.appendChild(collectButton);
      var collectionPanel=document.createElement('section');collectionPanel.id='wfCollection';collectionPanel.setAttribute('aria-label','Collected fiscal responses');host.insertBefore(collectionPanel,el('wfQueue'));
    }
    function busyState(v){busy=v;host.querySelectorAll('button').forEach(function(b){b.disabled=v;});el('wfNew').disabled=v||!writable;el('wfReminders').disabled=v||!writable;el('wfAutoReminders').disabled=v||!writable;el('wfPauseReminders').disabled=v||!writable;if(el('wfCollectFiscal')){el('wfCollectFiscal').disabled=v||!writable;}}
    function validSite(){if(!writable||loadedSite!==context.site()){throw new Error("Connection changed or queue is stale. Load shared work again.");}return loadedSite;}
    function clearCollection(){intake=null;collection=null;if(el('wfCollection')){el('wfCollection').innerHTML='';}}
    async function collectFiscal(){
      if(!context.collectFiscalResponses || busy){return;}if(editing){info('Save or discard your edits before collecting responses.');return;}
      var generation=epoch;busyState(true);clearCollection();
      try{
        var s=validSite(),token=await context.graph.getToken(),a=identity(token);
        function check(){if(generation!==epoch||s!==context.site()||a!==actor||identity(token)!==a){throw new Error('Connection or account changed. Load shared work again.');}}
        check();var result=await context.collectFiscalResponses(token,a,check,info);check();
        events=result.events;collection=result;render();renderCollection();
        info(result.candidates.length+' response(s) from '+result.threads+' issued fiscal thread(s); '+result.skipped+' already registered; '+result.excluded.length+' request(s) excluded; '+result.warnings.length+' read/extraction warning(s). Collection made no changes.');
      }catch(e){clearCollection();info('Fiscal collection stopped: '+e.message);}finally{busyState(false);}
    }
    function renderCollection(){
      var result=collection;if(!result){return;}
      var panel=el('wfCollection');panel.innerHTML='<h3>Responses to issued fiscal requests</h3><p>Only existing fiscal requests linked to this router’s recorded fiscal publications are included. Replies before the request was recorded are excluded. Nothing is saved until an assigned reviewer or coordinator confirms it. Each comment stays separate; select its request and division explicitly.</p>';
      panel.innerHTML+=result.candidates.map(function(c,i){return '<article class="options"><strong>'+esc(c.author+' · '+c.date)+'</strong><p>'+link(c.commentLink,'Open exact Teams comment')+'</p><pre style="white-space:pre-wrap;overflow-wrap:anywhere">'+esc(c.body.slice(0,2000))+'</pre>'+c.files.map(function(u,j){return link(u,'Attachment '+(j+1));}).join(' · ')+'<p>'+esc(c.evidence.map(function(e){return [e.source,e.fy,e.costRaw,e.evidence].filter(Boolean).join(' — ');}).join('\n')||'No fiscal amount extracted. Review the source; missing detail is not zero.')+'</p><label>Existing request <select data-intake-request="'+i+'"><option value="">Choose the matching request</option>'+c.requests.map(function(r){return '<option value="'+esc(r.id)+'">'+esc(r.bill+' · '+r.label+' · '+r.assembly+' / '+r.session)+'</option>';}).join('')+'</select></label><label>Division <select data-intake-division="'+i+'"><option value="">Choose request first</option></select></label><button type="button" data-intake-prepare="'+i+'">Prepare response for confirmation</button></article>';}).join('');
      panel.innerHTML+='<details><summary>Excluded requests ('+result.excluded.length+')</summary>'+result.excluded.map(function(r){return '<p>'+esc(r.label+': '+r.reason)+'</p>';}).join('')+'</details><details'+(result.warnings.length?' open':'')+'><summary>Read and extraction warnings ('+result.warnings.length+')</summary>'+result.warnings.map(function(w){return '<p>'+esc(w)+'</p>';}).join('')+'<p>Up to five DOCX/XLSX attachments per reply are extracted, at most 20 MB each. Other files remain linked for manual review. Links can point to editable files; open the response and replace them with the exact reviewed version links before saving.</p></details>';
      panel.querySelectorAll('[data-intake-request]').forEach(function(select){select.addEventListener('change',function(){var i=Number(select.getAttribute('data-intake-request')),r=result.candidates[i].requests.find(function(x){return x.id===select.value;}),d=panel.querySelector('[data-intake-division="'+i+'"]');d.innerHTML='<option value="">Choose the responding division</option>'+(r?r.divisions.map(function(x){return '<option value="'+esc(x.code)+'">'+esc(x.code+(x.canRegister?'':' — assigned staff must confirm'))+'</option>';}).join(''):'');});});
      panel.querySelectorAll('[data-intake-prepare]').forEach(function(b){b.addEventListener('click',function(){
        try{if(busy){return;}if(editing){throw new Error('Save or discard your edits before preparing another response.');}validSite();var i=Number(b.getAttribute('data-intake-prepare')),id=panel.querySelector('[data-intake-request="'+i+'"]').value,division=panel.querySelector('[data-intake-division="'+i+'"]').value,s=all().find(function(r){return r.id===id;});if(!s){throw new Error('Choose the existing request and responding division first.');}
          var prepared=context.prepareFiscalResponse(result.candidates[i],s,division,actor);selected=id;detail(division,prepared);
          el('wfVersion').value=prepared.version;el('wfDocs').value=prepared.links;el('wfNote').value=prepared.note;['Summary','Background','Assumptions','Impact','Sources'].forEach(function(k){el('wf'+k).value=prepared.narrative[k.toLowerCase()]||'';});el('wfFinance').value=prepared.financialText;
          editing=true;el('wfDiscard').hidden=false;el('wfConfirm').checked=false;el('wfSaveAction').textContent='Save response and request coordinator review';info('Prepared only. Verify the selected request/division, source and document versions. Enter confirmed structured fiscal lines for the dashboard; extracted amounts are clues, not approved totals.');if(el('wfDetail').scrollIntoView){el('wfDetail').scrollIntoView({block:'start'});}
        }catch(e){info(e.message);}
      });});
    }
    async function setAutomaticReminders(enabled){
      if(busy||!writable){return;}var generation=epoch;busyState(true);
      try{var s=validSite(),token=await context.graph.getToken(),a=identity(token);if(a!==actor||s!==context.site()||generation!==epoch){throw new Error("Account or site changed. Load shared work again.");}
        var guarded={};["listItems","addListItem","updateListItemFields"].forEach(function(name){guarded[name]=async function(){if(generation!==epoch||s!==context.site()||identity(arguments[0])!==a){throw new Error("Connection or account changed. Reload shared work.");}return context.graph[name].apply(context.graph,arguments);};});
        var result=await LrrWorkflowStore.client(guarded,s).configureReminders(token,enabled);
        if(generation!==epoch||s!==context.site()){throw new Error("Connection changed. Reload the previous site to check its reminder setting.");}
        info(enabled?'Automatic reminder feed connected: '+result.pending+' pending assignment(s). The installed Power Automate flow checks current workflow revisions before sending; its schedule and connections are managed in Power Automate.':'Automatic reminders paused for this shared queue.');
      }catch(e){info('Automatic reminders could not connect: '+e.message);}finally{busyState(false);}
    }
    async function draftReminders(){
      if(busy||!writable){return;}busyState(true);
      try{
        var s=validSite(),token=await context.graph.getToken(),a=identity(token);
        if(a!==actor){throw new Error("Account changed. Reload shared work.");}
        var fresh=await LrrWorkflowStore.client(context.graph,s).read(token);
        var groups=LrrWorkflow.reminders(LrrWorkflowStore.requests(fresh),actor,today());
        for(var i=0;i<groups.length;i++){
          if(s!==context.site()||identity(token)!==actor){throw new Error("Connection changed. Review existing Drafts before continuing.");}
          var group=groups[i],html='<p>A gentle reminder about assigned legislative work due within two calendar days or overdue. Please review the current response and confirmed deadline.</p>';
          group.items.forEach(function(item){html+='<h3>'+esc(item.bill)+' — '+esc(item.label)+'</h3><p>'+esc(item.division+' / '+item.type+' / due '+item.due)+'</p><p>'+esc(item.task)+'</p>'+link(item.thread||item.source,'Open the bill conversation or source request');});
          await context.graph.createDraftMessage(token,[group.to],'Legislative work reminder — '+group.items.length+' request(s)',html);
        }
        info(groups.length+' reminder draft(s) saved in Outlook. Review before sending. This button does not schedule unattended reminders.');
      }catch(e){info('Reminder drafting stopped: '+e.message+' Check Outlook Drafts before retrying.');}
      finally{busyState(false);}
    }
    async function load(){
      if(busy){return;}if(editing){info("Save or discard your edits before refreshing shared work.");return;}var generation=++epoch;busyState(true);writable=false;clearCollection();el('wfDetail').innerHTML='';
      try{var s=site(), token=await context.graph.getToken(), a=identity(token), result=await LrrWorkflowStore.client(context.graph,s).read(token);
        if(generation!==epoch||s!==context.site()||identity(token)!==a){throw new Error("Connection or account changed while loading. Load again.");}
        LrrWorkflowStore.requests(result);events=result;loadedSite=s;actor=a;writable=true;
        var previous=val('wfDivision'), previousSession=val('wfSessionFilter'), divisions=Array.from(new Set(all().flatMap(function(r){return r.config.divisions.map(function(d){return d.code;});}))).sort();
        el('wfDivision').innerHTML='<option value="">All divisions</option>'+divisions.map(function(d){return '<option>'+esc(d)+'</option>';}).join('');el('wfDivision').value=previous;
        var sessions=Array.from(new Set(all().map(function(r){return r.config.session;}))).sort();el('wfSessionFilter').innerHTML='<option value="">All sessions</option>'+sessions.map(function(name){return '<option>'+esc(name)+'</option>';}).join('');el('wfSessionFilter').value=previousSession;
        if(!defaultsUsed){var defaults=context.defaults||{};if(divisions.indexOf(defaults.division)>=0){el('wfDivision').value=defaults.division;}if(sessions.indexOf(defaults.session)>=0){el('wfSessionFilter').value=defaults.session;}defaultsUsed=true;}
        info('Loaded '+all().length+' request(s) at '+new Date().toLocaleTimeString()+'. Signed in as '+actor+'. Due-date reminders below use the confirmed calendar date.');render();
      }catch(e){writable=false;el('wfQueue').textContent='Shared queue unavailable. No workflow changes can be saved.';info(e.message);}finally{busyState(false);}
    }
    async function save(kind,input,draftNotification){
      if(busy){return;}var s;try{s=validSite();}catch(e){info(e.message);return;}
      var generation=epoch, committing=false;
      var old=current(), expected=old?old.revision:'', request=old?old.id:uid();input.revision=expected;
      busyState(true);
      try{var token=await context.graph.getToken(), a=identity(token);if(a!==actor||s!==context.site()){throw new Error("Account or site changed. Load shared work again.");}
        var fresh=await LrrWorkflowStore.client(context.graph,s).read(token), related=fresh.filter(function(e){return e.request===request;});
        if(LrrWorkflow.revision(related)!==expected){throw new Error("Shared work changed. Load again and review the latest version.");}
        var event=LrrWorkflow.plan(related,kind,input,a,{id:uid(),request:request,at:new Date().toISOString()});
        var guarded={listItems:async function(t,siteId,listId,top){if(generation!==epoch||s!==context.site()||identity(t)!==a){throw new Error('Connection or account changed. Reload shared work.');}return context.graph.listItems(t,siteId,listId,top);},addListItem:async function(t,siteId,listId,fields){if(generation!==epoch||s!==context.site()||identity(t)!==a){throw new Error('Connection or account changed. Reload shared work.');}return context.graph.addListItem(t,siteId,listId,fields);},updateListItemFields:async function(t,siteId,listId,id,fields){if(generation!==epoch||s!==context.site()||identity(t)!==a){throw new Error('Connection or account changed. Reload shared work.');}return context.graph.updateListItemFields(t,siteId,listId,id,fields);}};
        committing=true;var result=await LrrWorkflowStore.client(guarded,s).append(token,event);
        if(generation!==epoch||s!==context.site()){throw new Error("Connection changed. The save may have completed on the previous site; reload that site to check.");}
        events=result;selected=request;clearCollection();render();if(context.onSaved){context.onSaved();}
        var notice='Saved '+kind+'. No Teams messages or emails were sent.';
        if(draftNotification && event.kind==='artifact' && event.data.requestReview){
          try{
            var savedState=LrrWorkflowStore.requests(result).find(function(r){return r.id===request;});
            if(!savedState || savedState.conflicts.length || !savedState.appliedEvents.some(function(e){return e.id===event.id;})){throw new Error('Shared history needs review before drafting a handoff.');}
            if(generation!==epoch||s!==context.site()||identity(token)!==a){throw new Error('Account or site changed.');}
            var assignment=savedState.config.divisions.find(function(r){return r.code===event.division;});
            await context.graph.createDraftMessage(token,[assignment.coordinator],'Fiscal response ready for review — '+savedState.config.bill+' / '+event.division,
              '<p>'+esc(savedState.config.label+' · '+event.division+' · '+event.data.version)+'</p><p>The assigned reviewer registered this response and requested coordinator review. Open Work queue, load shared work, and review the current registered version before consolidation. Director and CFO approvals are still required.</p>'+event.data.links.map(function(u,i){return link(u,'Response document '+(i+1));}).join(' · ')+'<p>'+link(savedState.config.thread,'Open bill conversation')+'</p>');
            notice+=' Coordinator notification saved in Outlook Drafts; review and send it.';
          }catch(draftError){notice+=' Response saved and review requested, but the notification draft could not be confirmed: '+draftError.message+' Check Outlook Drafts before preparing another notification.';}
        }
        info(notice+(result.reminderWarning?' '+result.reminderWarning:''));
      }catch(e){if(committing){writable=false;}info(e.message+' Load shared work before retrying; check history if the save result was uncertain.');}finally{busyState(false);}
    }
    function render(preserveDetail){
      var rows=all(), div=val('wfDivision'), type=val('wfType'), stage=val('wfStage'), due=val('wfDue'), session=val('wfSessionFilter'), mine=el('wfMine').checked, now=today();
      rows=rows.filter(function(s){var day=s.config.due?Math.round((Date.parse(s.config.due+'T00:00:00Z')-Date.parse(now+'T00:00:00Z'))/86400000):null;return (!session||s.config.session===session)&&(!mine||s.config.cfo===actor||s.config.submitter===actor||s.config.divisions.some(function(r){return r.coordinator===actor||r.director===actor||r.reviewers.indexOf(actor)>=0;}))&&(!div||s.config.divisions.some(function(r){return r.code===div;}))&&(!type||s.config.type===type)&&(!stage||s.stage===stage||s.config.divisions.some(function(r){return (!div||r.code===div)&&s.divisions[r.code].stage===stage;}))&&(!due||(due==='missing'&&!s.config.due)||(due==='soon'&&day!==null&&day<=2&&!s.submitted));});
      rows.sort(function(a,b){return (a.config.due||'9999').localeCompare(b.config.due||'9999');});
      el('wfQueue').innerHTML=rows.length?rows.map(function(s,i){return '<article class="options"><strong>'+esc(s.config.bill)+' · '+esc(s.config.type)+' · '+esc(s.config.label)+'</strong><p>'+esc(s.config.divisions.map(function(d){return d.code;}).join(', '))+' · '+esc(s.stage)+' · '+esc(LrrWorkflow.due(s,now))+' · Legislative: '+esc(s.legislative)+'</p><button type="button" data-work="'+i+'">Open request</button></article>';}).join(''):'<p>No requests match this view. Create a separate request for each comment or fiscal-data cycle.</p>';
      host.querySelectorAll('[data-work]').forEach(function(b){b.addEventListener('click',function(){if(editing){info('Save or discard your edits before opening another request.');return;}selected=rows[Number(b.getAttribute('data-work'))].id;detail();if(el('wfDetail').scrollIntoView){el('wfDetail').scrollIntoView({block:'start'});}});});
      if(!preserveDetail){if(selected&&current()){detail();}else{el('wfDetail').innerHTML='';}}
    }
    function configForm(s){
      var c=s?s.config:{bill:'',assembly:'91',session:'',label:'',type:'Comment',due:'',dueSource:'',source:'',thread:'',related:'',cfo:'',submitter:actor,divisions:[]};
      return '<details'+(!s?' open':'')+'><summary>'+(s?'Request details and assignments':'New request')+'</summary><p>Owner: '+esc(s?s.owner:actor)+'. Saving assignment changes resets existing approvals. Use signed-in usernames for roles; the broad distribution list is separate.</p>'+field('wfBill','Bill (parent bill for amendments)',c.bill)+field('wfAssembly','General Assembly',c.assembly)+field('wfSession','Session/channel label',c.session)+field('wfLabel','Request label (identify this request/amendment)',c.label)+'<label>Request type <select id="wfKind"><option'+(c.type==='Comment'?' selected':'')+'>Comment</option><option'+(c.type==='Fiscal data'?' selected':'')+'>Fiscal data</option></select></label>'+area('wfBillSummary','Bill title / summary (used for past-response suggestions)',c.summary||'')+field('wfDeadline','Confirmed due date (leave blank if unknown)',c.due,'date')+field('wfDeadlineSource','Deadline source / confirmation',c.dueSource)+field('wfSource','Original request link',c.source)+field('wfThread','Existing Teams bill conversation link',c.thread)+field('wfRelated','Related / predecessor / successor bill references',c.related)+field('wfCfo','CFO signed-in email (fiscal data)',c.cfo)+field('wfSubmitter','LSA submitter signed-in email',c.submitter)+area('wfAssignments','One division per line: CODE | coordinator email | director email | reviewer emails separated by commas',c.divisions.map(function(d){return [d.code,d.coordinator,d.director,d.reviewers.join(', ')].join(' | ');}).join('\n'))+button('wfSaveConfig',s?'Save details and reset approvals':'Create request')+'<p class="hint">For the supplied LSA template, five days is relative to the original request; this queue does not invent its request date or assume a holiday convention.</p></details>';
    }
    function configInput(){return {bill:val('wfBill'),assembly:val('wfAssembly'),session:val('wfSession'),label:val('wfLabel'),summary:val('wfBillSummary'),type:val('wfKind'),due:val('wfDeadline'),dueSource:val('wfDeadlineSource'),source:val('wfSource'),thread:val('wfThread'),related:val('wfRelated'),cfo:val('wfCfo'),submitter:val('wfSubmitter'),divisions:val('wfAssignments').split(/\r?\n/).filter(function(x){return x.trim();}).map(function(line){var a=line.split('|').map(function(x){return x.trim();});if(a.length!==4){throw new Error('Each assignment line needs four fields separated by |.');}return {code:a[0],coordinator:a[1],director:a[2],reviewers:a[3]?a[3].split(',').map(function(x){return x.trim();}):[]};})};}
    function bindConfig(s){bind('wfSaveConfig',function(){try{save(s?'config':'create',{config:configInput()});}catch(e){info(e.message);}});}
    function newRequest(){if(busy||!writable){return;}if(editing){info('Save or discard your edits before creating another request.');return;}selected='';el('wfDetail').innerHTML=configForm(null);bindConfig(null);}
    function actionOptions(s,division) {
      var r=s.config.divisions.find(function(row){return row.code===division;}),v=s.divisions[division],out=[];
      var coordinator=r.coordinator===actor,director=r.director===actor,writer=coordinator||r.reviewers.indexOf(actor)>=0,cfo=s.config.type==='Fiscal data'&&s.config.cfo===actor;
      if(director&&v.coordinator&&v.stage==='Director review'){out.push(['director','Approve this division response']);}
      if(coordinator&&v.stage==='Review requested'){out.push(['coordinator','Consolidate response for director review']);}
      if(cfo&&s.ready&&!s.cfo&&!s.submitted){out.push(['cfo','Approve all division responses as CFO']);}
      if(s.config.type==='Fiscal data'&&s.config.submitter===actor&&s.ready&&s.cfo&&!s.submitted){out.push(['submitted','Record completed LSA submission']);}
      if(writer){out.push(['artifact','Register a draft or revised response']);if(v.stage==='Assigned'){out.push(['start','Start drafting']);}if(v.artifact&&v.stage==='Drafting'){out.push(['review','Request coordinator review']);}}
      if(v.artifact&&(coordinator||director||cfo)){out.push(['changes','Request changes to this response']);}
      out.push(['note','Add an FYI or consultation note']);if(s.owner===actor){out.push(['legislation','Record legislative progress']);}
      return out.map(function(option){return '<option value="'+option[0]+'">'+esc(option[1])+'</option>';}).join('');
    }
    function eventDetails(e,s) {
      var d=e.data||{},out='';
      function paragraph(label,value){if(value){out+='<p><strong>'+esc(label)+':</strong> '+esc(value)+'</p>';}}
      if(e.kind==='create'||e.kind==='config'){
        paragraph('Request',[d.bill,d.type,d.label,d.session].filter(Boolean).join(' · '));paragraph('Bill title / summary',d.summary);paragraph('Related bills',d.related);paragraph('Due date',d.due||'Needs confirmation');paragraph('Deadline source',d.dueSource);
        (d.divisions||[]).forEach(function(r){paragraph(r.code,'Coordinator: '+r.coordinator+'; director: '+r.director+'; reviewers: '+(r.reviewers.join(', ')||'coordinator only'));});paragraph('CFO',d.cfo);paragraph('LSA submitter',d.submitter);
      }else if(e.kind==='artifact'){
        paragraph('Response version',d.version);paragraph('Revision note',d.note);out+=(d.links||[]).map(function(u,i){return link(u,'Response document '+(i+1));}).join(' · ');
        var labels={summary:'Bill summary',background:'Background',assumptions:'Assumptions',impact:'Fiscal / operational impact',sources:'Related data and sources'};Object.keys(labels).forEach(function(key){paragraph(labels[key],(d.narrative||{})[key]);});
        if(d.requestReview){paragraph('Handoff','Coordinator review requested when this response was registered.');}if(d.intake){out+=link(d.intake.commentLink,'Original collected Teams comment');}if(d.financialText){out+='<p><strong>Financial detail</strong></p><pre style="white-space:pre-wrap;overflow-wrap:anywhere">'+esc(d.financialText)+'</pre>';}
      }else{
        paragraph('Intent',d.intent);paragraph('Note',d.note);paragraph('Status',d.status);paragraph('Event date',d.date);paragraph('Submitted at',d.at);if(d.evidence){out+=link(d.evidence,'Open source / submission evidence');}
        if(d.artifact){var version=s.events.find(function(row){return row.id===d.artifact;});paragraph('Response version',version?version.data.version:'Version retained in history');}
        if(e.kind==='cfo'){paragraph('Approved package','Current director-approved responses from every required division at the time of approval.');}
      }
      return out||'<p>Workflow step recorded.</p>';
    }
    function detail(preparedDivision,prepared){
      intake=prepared||null;var s=current();if(!s){return;}var remembered=preparedDivision||(el('wfDetail').getAttribute('data-request')===s.id&&el('wfDetail').getAttribute('data-actor')===actor?val('wfActionDivision'):'');editing=false;el('wfDiscard').hidden=true;var c=s.config;
      var html='<hr><h3>'+esc(c.bill)+' — '+esc(c.label)+'</h3><p>'+esc(c.assembly)+' / '+esc(c.session)+' · '+esc(s.stage)+' · '+esc(LrrWorkflow.due(s,today()))+'</p><p>'+link(c.source,'Original request')+' '+link(c.thread,'Open bill conversation in Teams')+'</p>';
      html+='<p>Legislative progress: '+esc(s.legislative)+' (separate from this request).</p>';
      if(c.related){html+='<p>Related bills: '+esc(c.related)+'</p>';}
      if(s.conflicts.length){html+='<p role="alert">Simultaneous changes were detected. The owner must inspect the history and save the reviewed request details to resolve them; approvals will reset.</p>';}
      if(!s.cfo&&s.events.some(function(e){return e.kind==='cfo';})){html+='<p role="status">Earlier CFO approval is no longer current. Complete the required reviews of the latest response version before CFO approval.</p>';}
      if(s.owner===actor){html+=configForm(s);}
      html+='<section aria-label="Relevant past responses"><h4>Relevant past responses</h4><p>Suggested automatically from saved bill references, keywords and code chapters across sessions. Review the original response against the current bill before reusing it. Linked document contents are not searched.</p><label>Show <select id="wfHistoryScope"><option value="all">All shared responses</option><option value="mine">My responses</option><option value="division">My divisions</option></select></label><div id="wfSuggestions"></div></section><h4>Division responses</h4>';
      c.divisions.forEach(function(r){var v=s.divisions[r.code];html+='<article class="options"><strong>'+esc(r.code)+' · '+esc(v.stage)+'</strong><p>Coordinator: '+esc(r.coordinator)+'<br>Director: '+esc(r.director)+'<br>Reviewers: '+esc(r.reviewers.join(', ')||'Coordinator only')+'</p>';
        if(v.artifact){html+='<p>Response '+esc(v.artifact.data.version)+' · registered by '+esc(v.artifact.actor)+' at '+esc(v.artifact.at)+'</p>'+v.artifact.data.links.map(function(u,i){return link(u,'Document '+(i+1));}).join(' · ');}
        var next=v.stage==='Director review'?r.director:v.stage==='Review requested'?r.coordinator:v.stage==='Division approved'?(s.submitted?'Complete':s.ready?(c.type==='Fiscal data'?(s.cfo?c.submitter+' (record completed LSA submission)':c.cfo+' (CFO review)'):'Complete'):'Waiting for other divisions'):(r.reviewers.join(', ')||r.coordinator);html+='<p>Next: '+esc(next||'Assignment needed')+'</p></article>';});
      if(c.type==='Fiscal data'){var f=LrrWorkflow.financialSummary(s);html+='<h4>Registered fiscal detail</h4><p>Staff-entered response figures; approval stage is shown above. Revenue and expenditure changes are separate. Missing detail is not zero.</p>';Object.keys(f.groups).sort().forEach(function(key){var g=f.groups[key];html+='<p>'+esc(key)+': $'+esc(g.total.toLocaleString())+' known'+(g.unknown?' + '+g.unknown+' unknown line(s)':'')+'</p>';});html+='<p>'+f.missingDivisions+' division(s) without structured fiscal lines. Full analysis may be in the linked written response.</p>'; }
      html+='<h4>Record a workflow action</h4><label>Division <select id="wfActionDivision">'+c.divisions.map(function(r){return '<option>'+esc(r.code)+'</option>';}).join('')+'</select></label><label>Action <select id="wfAction">'+actionOptions(s,c.divisions[0].code)+'</select></label><div id="wfActionFields"></div><label><input type="checkbox" id="wfConfirm" /> I reviewed this action and the current document version(s).</label>'+button('wfSaveAction','Save action')+'<h4>History</h4><p>Intent notes and filename labels do not constitute assignment or approval.</p><div id="wfEvents"></div>';
      el('wfDetail').innerHTML=html;el('wfDetail').setAttribute('data-request',s.id);el('wfDetail').setAttribute('data-actor',actor);if(s.owner===actor){bindConfig(s);}
      el('wfHistoryScope').addEventListener('change',suggestions);suggestions();
      var actionNames={create:'Request created',config:'Assignments or request details updated',artifact:'Response version registered',start:'Drafting started',review:'Review requested',changes:'Changes requested',coordinator:'Coordinator consolidation',director:'Director approval',cfo:'CFO approval',submitted:'Submission recorded',note:'Context note',legislation:'Legislative progress'};
      el('wfEvents').innerHTML=s.events.slice().reverse().map(function(e){return '<details><summary>'+esc(e.at)+' · '+esc(actionNames[e.kind]||e.kind)+(s.ignoredIds.indexOf(e.id)>=0?' [not applied: concurrent edit]':'')+' · '+esc(e.division)+' · '+esc(e.actor)+'</summary>'+eventDetails(e,s)+'</details>';}).join('');
      var assigned=c.divisions.find(function(r){return r.coordinator===actor||r.director===actor||r.reviewers.indexOf(actor)>=0;});if(c.divisions.some(function(r){return r.code===remembered;})){el('wfActionDivision').value=remembered;}else if(assigned){el('wfActionDivision').value=assigned.code;}
      el('wfAction').innerHTML=actionOptions(s,val('wfActionDivision'));if(intake){el('wfAction').value='artifact';}
      var actionKind=val('wfAction'),actionDivision=val('wfActionDivision');function changeAction(){if(editing){el('wfAction').value=actionKind;el('wfActionDivision').value=actionDivision;info('Save or discard your response edits before changing actions.');return;}if(actionDivision!==val('wfActionDivision')){el('wfAction').innerHTML=actionOptions(s,val('wfActionDivision'));}actionKind=val('wfAction');actionDivision=val('wfActionDivision');actionFields();}el('wfAction').addEventListener('change',changeAction);el('wfActionDivision').addEventListener('change',changeAction);actionFields();
      bind('wfSaveAction',function(){try{if(!el('wfConfirm').checked){info('Review and confirm this action first.');return;}save(val('wfAction'),{intake:intake?intake.intake:undefined,requestReview:!!intake,division:val('wfActionDivision'),version:val('wfVersion'),links:val('wfDocs'),note:val('wfNote'),financialText:val('wfFinance'),narrative:{summary:val('wfSummary'),background:val('wfBackground'),assumptions:val('wfAssumptions'),impact:val('wfImpact'),sources:val('wfSources')},intent:val('wfIntent'),evidence:val('wfEvidence'),submittedAt:val('wfSubmittedAt')?new Date(val('wfSubmittedAt')).toISOString():'',status:val('wfLegislative'),date:val('wfEventDate')},!!(el('wfIntakeNotify')&&el('wfIntakeNotify').checked));}catch(e){info(e.message);}});
    }
    function suggestions(){
      var s=current();if(!s){return;}
      var matches=LrrWorkflow.relevantResponses(s,all(),actor,val('wfHistoryScope'));
      el('wfSuggestions').innerHTML=matches.length?matches.map(function(m,i){var c=m.request.config,e=m.event;return '<article class="options"><strong>'+esc(c.bill+' — '+c.label)+'</strong><p>'+esc('Assembly '+c.assembly+' / '+c.session+' · '+m.division+' · '+c.type)+'</p><p>'+esc(e.actor+' · '+e.at+' · response '+e.data.version)+'</p><p>'+esc(m.reasons.join('; '))+'</p><p>'+esc(m.excerpt||'Open the registered document to read this response.')+'</p>'+e.data.links.map(function(u,j){return link(u,'Original response document '+(j+1));}).join(' · ')+'<p>'+link(c.thread,'Original Teams conversation')+'</p><button type="button" data-history="'+i+'">Open original request history</button></article>';}).join(''):'<p>No relevant past responses found in the loaded shared queue. Matches use saved request summaries and response text; older documents may have no searchable text here.</p>';
      el('wfSuggestions').querySelectorAll('[data-history]').forEach(function(b){b.addEventListener('click',function(){if(editing){info('Save or discard your edits before opening another request.');return;}var m=matches[Number(b.getAttribute('data-history'))];selected=m.request.id;detail();var entries=el('wfEvents').querySelectorAll('details'),index=m.request.events.slice().reverse().findIndex(function(e){return e.id===m.event.id;});if(entries[index]){entries[index].open=true;if(entries[index].scrollIntoView){entries[index].scrollIntoView({block:'start'});}}});});
    }
    function actionFields(){
      var kind=val('wfAction'), s=current(), v=s.divisions[val('wfActionDivision')], a=v.artifact?v.artifact.data:{}, n=a.narrative||{}, html='';el('wfConfirm').checked=false;
      if(kind==='artifact'){
        html=(intake?'<p><strong>Confirm the existing request, responding division, exact source/version and fiscal detail. Saving registers this response and requests coordinator review; no approval is inferred.</strong></p><label><input type="checkbox" id="wfIntakeNotify" checked /> Prepare an Outlook notification draft for the assigned coordinator (review and send it separately)</label>':'')+field('wfVersion','Document version label',a.version||'')+area('wfDocs','Exact document-version links (one per line)',(a.links||[]).join('\n'))+area('wfNote','Response / revision description',a.note||'');
        if(s.config.type==='Fiscal data'){html+='<details><summary>Fiscal response detail (or keep full analysis in linked form)</summary>'+area('wfSummary','Bill summary',n.summary)+area('wfBackground','Background / implementation',n.background)+area('wfAssumptions','Assumptions, including reason for no impact',n.assumptions)+area('wfImpact','Fiscal / operational impact, funds, FTE and local impacts',n.impact)+area('wfSources','Related data and sources',n.sources)+area('wfFinance','Optional financial lines — paste seven tab-separated columns: FY2027, Revenue/Expenditure, Increase/Decrease, amount or Unknown, fund, One-time/Ongoing, description',a.financialText||'')+'<p>Do not paste template zeros as confirmed estimates. Revenue and expenditure changes stay separate; these entries do not replace the required written response.</p></details>';}
      }else if(kind==='note'){html='<label>Intent <select id="wfIntent"><option>FYI</option><option>Consultation</option><option>Action requested</option><option>Review requested</option><option>Approval requested</option></select></label>'+area('wfNote','Note','');}
      else if(kind==='changes'){html=area('wfNote','Changes required','');}
      else if(kind==='submitted'){html=field('wfSubmittedAt','Actual completed submission time (local time)','','datetime-local')+field('wfEvidence','Submission receipt or source message link','')+area('wfNote','Documents actually submitted','');}
      else if(kind==='legislation'){html=field('wfLegislative','Legislative status / action','')+field('wfEventDate','Event date','','date')+field('wfEvidence','Official source or source message link','')+'<p>Legislative progress is recorded independently; it does not complete or approve this request.</p>';}
      else {html='<p>Current response: '+esc(v.artifact?v.artifact.data.version:'No registered response')+'. '+(kind==='cfo'?'Requires current director approvals for every required division.':kind==='director'?'Requires coordinator consolidation of this exact version.':'')+'</p>';}
      el('wfActionFields').innerHTML=html;
    }
    if(!context.staff){
      el('wfTeamsSetup').innerHTML='<details><summary>Teams setup</summary><p>This installs the custom staff work queue, with filters and review actions. It is separate from the built-in Lists tab for contacts. Your Teams administrator may need to publish the package first. Copy the connection code into the custom tab setup once.</p>'+button('wfTeamsCode','Show Teams connection code')+'<textarea id="wfTeamsCodeValue" rows="3" readonly hidden aria-label="Teams connection code"></textarea><p><a href="../../teams/Legislative-Work-Queue.zip" download>Download Teams app package</a> · <a href="../../teams-install.html" target="_blank" rel="noopener noreferrer">Installation instructions</a></p></details>';
      bind('wfTeamsCode',function(){try{var s=site();if(typeof LrrStaffConfig==='undefined'){throw new Error('Reload the router to load Teams setup tools.');}el('wfTeamsCodeValue').value=LrrStaffConfig.encode({cloud:context.cloud?context.cloud():'commercial',siteId:s.siteId,auditListId:s.auditListId,session:val('wfSessionFilter'),division:val('wfDivision')});el('wfTeamsCodeValue').hidden=false;el('wfTeamsCodeValue').select();}catch(e){info(e.message);}});
    }
    host.addEventListener('input',function(event){if(el('wfDetail').contains(event.target)&&['wfAction','wfActionDivision','wfConfirm','wfHistoryScope'].indexOf(event.target.id)<0){editing=true;el('wfDiscard').hidden=false;if(el('wfConfirm')){el('wfConfirm').checked=false;}}});
    bind('wfDiscard',function(){editing=false;el('wfDiscard').hidden=true;if(current()){detail();}else{newRequest();}info('Edits discarded. Load shared work to refresh the saved record.');});
    bind('wfCollectFiscal',collectFiscal);bind('wfLoad',load);bind('wfNew',newRequest);bind('wfReminders',draftReminders);bind('wfAutoReminders',function(){setAutomaticReminders(true);});bind('wfPauseReminders',function(){setAutomaticReminders(false);});['wfSessionFilter','wfDivision','wfType','wfStage','wfDue','wfMine'].forEach(function(id){el(id).addEventListener('change',function(){render(true);});});
    return {load:load,collectFiscal:collectFiscal,invalidate:function(){epoch++;editing=false;clearCollection();if(el('wfCollectFiscal')){el('wfCollectFiscal').disabled=true;}el('wfDiscard').hidden=true;writable=false;loadedSite=null;events=[];selected='';actor='';el('wfQueue').innerHTML='';el('wfDetail').innerHTML='';el('wfNew').disabled=true;if(el('wfTeamsCodeValue')){el('wfTeamsCodeValue').value='';el('wfTeamsCodeValue').hidden=true;}info('Connection changed. Load shared work again.');}};
  }
  root.LrrWorkflowUI={mount:mount};
})(typeof self!=="undefined"?self:this);
