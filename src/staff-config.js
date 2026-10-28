/* Portable queue connection settings. Contains IDs and default filters, never tokens or workflow records. */
(function(root){
  "use strict";
  var GUID="[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}";
  function validate(value){
    var v=value||{},cloud=v.cloud||"commercial",site=String(v.siteId||""),audit=String(v.auditListId||"");
    if(["commercial","gcc","gcchigh","dod"].indexOf(cloud)<0){throw new Error("Unknown Microsoft cloud.");}
    if(!new RegExp("^[a-z0-9.-]+,"+GUID+","+GUID+"$","i").test(site)||!new RegExp("^"+GUID+"$","i").test(audit)){throw new Error("Copy the Teams connection code from the router's connected Work queue.");}
    var session=String(v.session||"").trim(),division=String(v.division||"").trim().toUpperCase();
    if(session.length>120||division.length>30){throw new Error("Default session or division is too long.");}
    return {v:1,cloud:cloud,siteId:site,auditListId:audit,session:session,division:division};
  }
  function encode(value){return btoa(unescape(encodeURIComponent(JSON.stringify(validate(value))))).replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/,"");}
  function decode(code){
    code=String(code||"").trim();
    if(!/^[A-Za-z0-9_-]{1,3000}$/.test(code)){throw new Error("Paste a valid Teams connection code from the router.");}
    try{return validate(JSON.parse(decodeURIComponent(escape(atob(code.replace(/-/g,"+").replace(/_/g,"/"))))));}
    catch(e){throw new Error("The Teams connection code is invalid. Copy it again from the router.");}
  }
  function contentUrl(base,value){var u=new URL("index.html",base);u.search="";u.hash="lrr="+encode(value);return u.href;}
  var api={validate:validate,encode:encode,decode:decode,contentUrl:contentUrl};
  if(typeof module!=="undefined"&&module.exports){module.exports=api;}else{root.LrrStaffConfig=api;}
})(typeof self!=="undefined"?self:this);
