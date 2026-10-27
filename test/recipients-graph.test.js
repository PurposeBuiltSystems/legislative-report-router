"use strict";
const assert=require("node:assert/strict"),fs=require("node:fs"),vm=require("node:vm");
const calls=[];
const context={self:{},setTimeout,clearTimeout,fetch:async(url,opts)=>{
  calls.push({url,opts});
  return {ok:true,status:200,json:async()=>url.includes("/members")?{value:[{userId:"u",email:"u@agency.gov"}]}:{}};
}};
vm.runInNewContext(fs.readFileSync(require("node:path").join(__dirname,"../src/graph.js"),"utf8"),context);
(async()=>{
  const api=context.self.GraphData;
  assert.equal(api.getRecipientToken,undefined,"matching never introduces an incremental directory permission");
  assert.equal(api.listTeamMembers,undefined,"matching uses existing tag reads only");
  await api.listTeamTagMembers("token","team","tag/id");
  assert.match(calls[0].url,/tags\/tag%2Fid\/members$/);
  await assert.rejects(api.updateRecipientEmails("token","site","list","1",{Emails:"a@agency.gov"}),/no version/);
  assert.equal(calls.length,1,"missing version never overwrites concurrent spreadsheet edits");
  await api.updateRecipientEmails("token","site","list","1",{Emails:"a@agency.gov"},'"etag1"');
  assert.equal(calls[1].opts.headers["If-Match"],'"etag1"');
  assert.equal(calls[1].opts.method,"PATCH");
  assert.deepEqual(JSON.parse(calls[1].opts.body),{Emails:"a@agency.gov"});
  console.log("Shared contact Graph permissions and conditional-write tests passed.");
})().catch(e=>{console.error(e);process.exitCode=1;});
