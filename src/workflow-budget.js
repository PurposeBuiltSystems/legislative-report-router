/* Current response versions projected from the shared approval history. No tracker writes. */
(function(root){
  "use strict";
  function esc(v){return String(v==null?"":v).replace(/[&<>"']/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c];});}
  function rows(requests){
    var out=[];
    (requests||[]).filter(function(s){return s.config.type==="Fiscal data";}).forEach(function(s){
      s.config.divisions.forEach(function(d){
        var v=s.divisions[d.code], a=v.artifact, lines=a&&a.data.finance||[];
        var approval=s.conflicts.length?"Conflicting changes":s.cfo?"CFO approved":v.director?"Division approved":"Pending review";
        if(!lines.length){out.push({request:s.id,bill:s.config.bill,label:s.config.label,session:s.config.session,division:d.code,fy:"(no FY)",amount:null,kind:"Missing detail",direction:"",approval:approval,version:a?a.data.version:"No response",link:s.config.thread||s.config.source});}
        lines.forEach(function(line){out.push(Object.assign({},line,{request:s.id,bill:s.config.bill,label:s.config.label,session:s.config.session,division:d.code,approval:approval,version:a.data.version,link:(a.data.links||[])[0]||s.config.thread}));});
      });
    });
    return out;
  }
  function html(all,opts){
    opts=opts||{};var selected=(all||[]).filter(function(r){return (!opts.fy||r.fy===opts.fy)&&(!opts.division||r.division===opts.division)&&(!opts.session||r.session===opts.session);});
    var waiting=selected.filter(function(r){return r.approval!=="CFO approved";}).length;
    if(opts.approved){selected=selected.filter(function(r){return r.approval==="CFO approved";});}
    var groups=Object.create(null),unknown=0;
    selected.forEach(function(r){if(r.amount===null){unknown++;return;}var key=[r.fy,r.kind,r.direction].join(" / ");groups[key]=(groups[key]||0)+r.amount;});
    function dollars(n){return "$"+Number(n).toLocaleString("en-US",{maximumFractionDigits:2});}
    var h='<h3>'+esc(opts.approved?"CFO-approved fiscal responses":"Current fiscal responses — all approval stages")+'</h3><p>Read directly from shared work requests and their current response versions. Revised responses remove earlier approvals immediately on refresh. Revenue and expenditure changes stay separate; these totals do not overwrite BillTracker estimates.</p>';
    if(opts.at){h+='<p>Loaded '+esc(opts.at)+'. Refresh the rollup to read later approvals or revisions.</p>';}
    h+='<p>'+selected.length+' financial line(s); '+unknown+' with unknown/missing amounts. '+waiting+' line(s) in this view are awaiting current CFO approval. Each request is separate: review overlapping requests before using totals.</p><table><thead><tr><th>Fiscal year / change</th><th>Known total</th></tr></thead><tbody>';
    Object.keys(groups).sort().forEach(function(k){h+='<tr><td>'+esc(k)+'</td><td>'+esc(dollars(groups[k]))+'</td></tr>';});
    h+='</tbody></table><h4>Responses and approval status</h4><table><thead><tr><th>Bill / request</th><th>Division</th><th>FY</th><th>Change / fund</th><th>Amount</th><th>Current approval</th><th>Version</th></tr></thead><tbody>';
    selected.forEach(function(r){h+='<tr><td>'+esc(r.bill+' / '+r.label)+'</td><td>'+esc(r.division)+'</td><td>'+esc(r.fy)+'</td><td>'+esc([r.kind,r.direction,r.fund,r.frequency].filter(Boolean).join(' / '))+'</td><td>'+esc(r.amount===null?'Unknown':dollars(r.amount))+'</td><td>'+esc(r.approval)+'</td><td>'+(r.link&&/^https:\/\/[^\s]+$/i.test(r.link)?'<a href="'+esc(r.link)+'">'+esc(r.version)+'</a>':esc(r.version))+'</td></tr>';});
    return h+'</tbody></table>'+(selected.length?'':'<p>No responses match this approval view. Pending or revised responses are not counted as CFO approved.</p>');
  }
  var api={rows:rows,html:html};if(typeof module!=="undefined"&&module.exports){module.exports=api;}else{root.LrrWorkflowBudget=api;}
})(typeof self!=="undefined"?self:this);
