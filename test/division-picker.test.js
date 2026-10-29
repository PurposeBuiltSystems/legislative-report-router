"use strict";
const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const {JSDOM}=require('jsdom');
const dom=new JSDOM(fs.readFileSync('src/taskpane/taskpane.html','utf8'));
const document=dom.window.document,el=id=>document.getElementById(id);
const raw=fs.readFileSync('src/taskpane/taskpane.js','utf8'),end=raw.lastIndexOf('})();');
const R=require('../src/routing');
const context={self:{},document,Office:{onReady(){}},
  LrrRouting:R,LrrChapters:require('../src/chapters'),LrrReportGen:require('../src/reportgen'),LrrTeams:require('../src/teams')};
vm.runInNewContext(raw.slice(0,end)+'\nself.api={state,renderItems,refreshStats};\n'+raw.slice(end),context);
const a=context.self.api;
function rule(id,code,name,channel,extra={}) {
  return R.ruleFromSharePoint({DivisionCode:code,DivisionName:name,TeamsTeamId:'team',TeamsChannelId:channel,Priority:1,IsActive:true,...extra},id);
}
a.state.rules=[rule('1','MVD','Motor Vehicle Division','main',{Aliases:'Motor Vehicles; Vehicle Services'}),
  rule('2','mvd','Motor Vehicle Division','second',{Aliases:'Licensing'}),
  rule('3','TDD','Transportation Development','main'),
  rule('4','OLD','Retired Division','old',{IsActive:false}),
  rule('5','FUTURE','Future Division','new',{EffectiveStartDate:'2099-01-01'}),
  rule('6','HTML','<img src=x onerror=alert(1)>','safe')];
a.state.items=[{billNumber:'HF1',distributedTo:['MVD','Unknown Division'],commentRequestedFrom:['TDD'],
  brief:'A bill.',parserWarnings:[],parserConfidence:1,routingStatus:'unmatched',codeChapters:[]}];
const item=a.state.items[0];
a.refreshStats();a.renderItems();
const root=()=>el('distributed-to-0'), options=()=>Array.from(root().querySelectorAll('input[type=checkbox]'));
const change=code=>{const checkbox=options().find(x=>x.value.toLowerCase()===code.toLowerCase());checkbox.checked=!checkbox.checked;checkbox.dispatchEvent(new dom.window.Event('change',{bubbles:true}));};
assert.equal(options().length,3,'one choice per active division; inactive/future routes omitted');
assert.equal(options().filter(x=>x.value.toLowerCase()==='mvd').length,1);
assert.equal(options().find(x=>x.value==='MVD').checked,true);
assert.match(root().querySelector('.needs-setup').textContent,/Unknown Division · Needs setup/);
assert.equal(root().querySelectorAll('img').length,0,'configured names are rendered as text');
root().querySelector('details').open=true;
el('distributed-to-0-search').value='transport';el('distributed-to-0-search').dispatchEvent(new dom.window.Event('input'));
assert.equal(root().querySelectorAll('.division-option:not([hidden])').length,1);
el('confirmBox').checked=true;el('publishGo').disabled=false;
change('TDD');
assert.deepEqual(Array.from(item.distributedTo),['MVD','Unknown Division','TDD']);
assert.deepEqual(Array.from(item.commentRequestedFrom),['TDD'],'distribution edits preserve comment requests');
assert.equal(item.routes.length,3,'multiple MVD destinations still route through one picker choice');
assert.equal(el('confirmBox').checked,false,'changing recipients clears prior publish confirmation');
assert.equal(el('publishGo').disabled,true);
assert.equal(root().querySelector('details').open,true,'multiselect stays open');
assert.equal(el('distributed-to-0-search').value,'transport','search survives selection');
assert.equal(document.activeElement.value,'TDD','keyboard focus survives selection');
change('TDD');assert.deepEqual(Array.from(item.distributedTo),['MVD','Unknown Division']);
assert.ok(item.routes.some(r=>r.divisionCode==='TDD'),'Comment From independently retains its routing');
el('distributed-to-0-search').value='Licensing';el('distributed-to-0-search').dispatchEvent(new dom.window.Event('input'));
assert.equal(root().querySelectorAll('.division-option:not([hidden])').length,1,'aliases from any destination are searchable');
el('distributed-to-0-search').value='nothing matches';el('distributed-to-0-search').dispatchEvent(new dom.window.Event('input'));
assert.match(root().querySelector('[role=status]').textContent,/No matching divisions/);
root().querySelector('details').dispatchEvent(new dom.window.KeyboardEvent('keydown',{key:'Escape',bubbles:true}));
assert.equal(root().querySelector('details').open,false);assert.equal(document.activeElement.id,'distributed-to-0-summary');
root().querySelector('.needs-setup button').click();assert.deepEqual(Array.from(item.distributedTo),['MVD']);
item.distributedTo=['Motor Vehicles'];a.renderItems();assert.equal(options().find(x=>x.value==='MVD').checked,true,'imported names and aliases match configured code');
change('MVD');assert.equal(item.distributedTo.length,0,'unchecking alias selection removes it without adding a second token');
a.state.rules=[];item.distributedTo=['MVD'];a.renderItems();
assert.match(root().textContent,/MVD · Needs setup/);assert.match(root().textContent,/No active divisions/);
assert.equal(options().length,0);assert.deepEqual(Array.from(item.distributedTo),['MVD'],'missing setup never drops imported divisions');
console.log('Division picker: active deduplicated choices, search, aliases, multiselect routing, unknown preservation, keyboard focus and publish confirmation passed.');
dom.window.close();
