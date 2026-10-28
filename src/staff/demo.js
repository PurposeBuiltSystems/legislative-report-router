/* Fictional, in-memory practice data. Never calls Graph or publishes notifications. */
/* global LrrWorkflow, LrrWorkflowStore */
(function(root){
  "use strict";
  async function create(){
    var rows=[],sequence=0,number=0,username="reviewer@example.test";
    var people={reviewer:"reviewer@example.test",coordinator:"coordinator@example.test",director:"director@example.test",cfo:"cfo@example.test",owner:"owner@example.test",submitter:"submitter@example.test"};
    async function seed(bill,type,division,days,stage){
      var events=[],request=crypto.randomUUID(),date=new Date();date.setDate(date.getDate()+days);
      var due=date.getFullYear()+"-"+String(date.getMonth()+1).padStart(2,"0")+"-"+String(date.getDate()).padStart(2,"0");
      var config={bill:bill,assembly:"91",session:"2026",label:"Practice "+(type==="Comment"?"comment":"fiscal")+" request",type:type,due:days===99?"":due,dueSource:days===99?"":"Fictional practice deadline",source:"",thread:"",related:"",cfo:people.cfo,submitter:people.submitter,divisions:[{code:division,coordinator:people.coordinator,director:people.director,reviewers:[people.reviewer]}]};
      function step(kind,data,actor){var s=LrrWorkflow.project(events),e=LrrWorkflow.plan(events,kind,Object.assign({revision:s?s.revision:""},data),actor,{id:crypto.randomUUID(),request:request,at:new Date().toISOString()});e.order=++number;events.push(e);}
      step("create",{config:config},people.owner);
      if(stage!=="Assigned"){
        step("artifact",{division:division,version:"Practice response v1",links:"https://example.invalid/practice-response-v1",narrative:{summary:"Fictional response for practice only."}},people.reviewer);
        step("review",{division:division},people.reviewer);
        if(stage==="Director review"||stage==="CFO review"||stage==="Complete") {step("coordinator",{division:division},people.coordinator);}
        if(stage==="CFO review"||stage==="Complete"){step("director",{division:division},people.director);}
        if(stage==="Changes requested"){step("changes",{division:division,note:"Practice: explain the one-time implementation cost."},people.director);}
      }
      for(var i=0;i<events.length;i++){var packed=await LrrWorkflowStore.encode(events[i]);packed.parts.concat([packed.commit]).forEach(function(fields){rows.push({id:String(++sequence),fields:fields});});}
    }
    await seed("SSB3108","Fiscal data","MVD",1,"Director review");await seed("HF2586","Comment","MVD",-1,"Assigned");await seed("SF2355","Fiscal data","TDD",2,"CFO review");await seed("HF2556","Fiscal data","ASD",4,"Changes requested");await seed("HF2669","Comment","TDD",3,"Complete");await seed("SF2234","Fiscal data","MVD",99,"Assigned");
    return {people:people,graph:{getToken:async function(){return username;},authenticatedAccount:function(token){return token===username?username:null;},listItems:async function(){return JSON.parse(JSON.stringify(rows));},addListItem:async function(t,s,l,fields){rows.push({id:String(++sequence),fields:fields});}},setActor:function(value){username=value;}};
  }
  root.LrrStaffDemo={create:create};
})(typeof self!=="undefined"?self:this);
