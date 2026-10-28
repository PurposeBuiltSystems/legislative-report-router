/* Embedded Teams tab and channel setup. TeamsJS initializes before any MSAL request. */
/* global microsoftTeams, GraphData, LrrWorkflowUI, LrrStaffConfig, LrrStaffDemo, document, window */
(function(root){
  "use strict";
  function el(id){return document.getElementById(id);}
  function status(message){el("staffStatus").textContent=message;}
  function bounded(promise,message){var timer;return Promise.race([promise.finally(function(){clearTimeout(timer);}),new Promise(function(resolve,reject){timer=setTimeout(function(){reject(new Error(message));},12000);})]);}
  function fromUrl(href){var u=new URL(href),code=new URLSearchParams(u.hash.slice(1)).get("lrr");return code?LrrStaffConfig.decode(code):null;}
  function theme(value){document.documentElement.setAttribute("data-theme",["dark","contrast"].indexOf(value)>=0?value:"default");}
  async function configure(){
    el("staffSetup").hidden=false;status("Set up this channel's shared queue.");
    var existing={};try{existing=await bounded(microsoftTeams.pages.config.getConfig(),"Could not read existing tab settings.");var restored=fromUrl(existing.contentUrl);if(restored){el("staffCode").value=LrrStaffConfig.encode(restored);el("staffSession").value=restored.session;el("staffDivision").value=restored.division;}if(existing.suggestedDisplayName){el("staffTabName").value=existing.suggestedDisplayName;}}catch(e){/* a new tab has no settings */}
    function settings(){var v=LrrStaffConfig.decode(el("staffCode").value);v.session=el("staffSession").value.trim();v.division=el("staffDivision").value.trim();return LrrStaffConfig.validate(v);}
    function validity(){try{settings();microsoftTeams.pages.config.setValidityState(!!el("staffTabName").value.trim());el("staffSetupStatus").textContent="Ready. Choose Save in Teams to add this tab.";}catch(e){microsoftTeams.pages.config.setValidityState(false);el("staffSetupStatus").textContent=e.message;}}
    ["staffCode","staffSession","staffDivision","staffTabName"].forEach(function(id){el(id).addEventListener("input",validity);});
    microsoftTeams.pages.config.registerOnSaveHandler(function(save){
      try{var value=settings(),name=el("staffTabName").value.trim();if(!name){throw new Error("Enter a tab name.");}var content=LrrStaffConfig.contentUrl(window.location.href,value);
        bounded(microsoftTeams.pages.config.setConfig({entityId:existing.entityId||crypto.randomUUID(),suggestedDisplayName:name,contentUrl:content,websiteUrl:content}),"Teams could not save this tab.").then(function(){save.notifySuccess();},function(){save.notifyFailure("Teams could not save the tab configuration. Try again.");});
      }catch(e){save.notifyFailure(e.message);}
    });validity();
  }
  async function practice(){
    var demo=await LrrStaffDemo.create(),site={siteId:"practice",auditListId:"practice"};
    el("staffDemo").hidden=false;status("Practice mode. All people, requests, deadlines and approvals below are fictional.");
    var label=document.createElement("label");label.textContent="Practice as ";var select=document.createElement("select");Object.keys(demo.people).forEach(function(key){var option=document.createElement("option");option.value=demo.people[key];option.textContent=key;select.appendChild(option);});label.appendChild(select);el("staffDemo").appendChild(label);
    el("staffQueue").hidden=false;var view=LrrWorkflowUI.mount(el("staffWorkflowHost"),{graph:demo.graph,site:function(){return site;},staff:true,defaults:{session:"2026"}});select.addEventListener("change",async function(){view.invalidate();demo.setActor(select.value);await view.load();});await view.load();
  }
  async function start(){
    if(new URLSearchParams(window.location.search).get("demo")==="1"){await practice();return;}
    try{
      await bounded(microsoftTeams.app.initialize(),"Open this app as a channel tab in Microsoft Teams.");
      var context=await bounded(microsoftTeams.app.getContext(),"Teams context did not load. Reload this tab.");theme((context.app||{}).theme);microsoftTeams.app.registerOnThemeChangeHandler(theme);
      if(new URLSearchParams(window.location.search).get("configure")==="1"){await configure();return;}
      var connection=fromUrl(window.location.href);if(!connection){throw new Error("This tab has no connection settings. Use the tab's Settings to paste a connection code from the router.");}
      GraphData.setCloud(connection.cloud);
      var site={siteId:connection.siteId,auditListId:connection.auditListId};
      var adapter={getToken:GraphData.getWorkflowToken,authenticatedAccount:GraphData.authenticatedAccount,listItems:GraphData.listItems,addListItem:GraphData.addListItem};
      var view=LrrWorkflowUI.mount(el("staffWorkflowHost"),{graph:adapter,site:function(){return site;},staff:true,defaults:{session:connection.session,division:connection.division}});
      el("staffQueue").hidden=false;status("Select Load shared work to sign in and read the shared queue. No Outlook add-in is needed.");el("staffSignOut").hidden=false;
      el("staffSignOut").addEventListener("click",async function(){view.invalidate();await GraphData.signOut();status("Signed out of this queue. Load shared work to sign in again.");});
    }catch(e){status(e.message||"Microsoft Teams is unavailable.");el("staffInstall").hidden=false;}
  }
  root.LrrStaffTab={start:start,fromUrl:fromUrl};
  if(document.readyState==="loading"){document.addEventListener("DOMContentLoaded",start);}else{start();}
})(typeof self!=="undefined"?self:this);
