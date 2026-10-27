"use strict";
const assert=require("node:assert/strict"),fs=require("node:fs"),vm=require("node:vm"),path=require("node:path");
const raw=fs.readFileSync(path.join(__dirname,"../src/taskpane/taskpane.js"),"utf8"),at=raw.lastIndexOf("})();");
const source=raw.slice(0,at)+'\nself.ensureWatchModules=ensureWatchModules;\n'+raw.slice(at);
(async()=>{
 const timers=new Map(),loads=[];let id=0,fail=true;
 const context={self:{},Office:{onReady(){}},URL,setTimeout(fn){timers.set(++id,fn);return id;},clearTimeout(i){timers.delete(i);},
  document:{baseURI:"https://purposebuiltsystems.github.io/legislative-report-router/src/taskpane/taskpane.html",createElement(){return {};},head:{appendChild(script){
   loads.push(script.src);
   if(script.src.includes("watch-window")&&fail){script.onerror();return;}
   if(script.src.includes("watch-window"))context.LrrWatch={};else context.LrrAmendments={};
   script.onload();
  }}}
 };
 vm.runInNewContext(source,context);
 await assert.rejects(context.self.ensureWatchModules(),/could not load/);
 fail=false;await context.self.ensureWatchModules();
 assert.equal(loads.length,3,"cached old HTML can load new modules and recover a failed attempt");
 assert.equal(timers.size,0);
 await context.self.ensureWatchModules();assert.equal(loads.length,3,"already loaded modules are reused");
 assert.ok(loads.every(url=>url.startsWith("https://purposebuiltsystems.github.io/legislative-report-router/src/")));
 console.log("Cached command-page watch module loading and retry tests passed.");
})().catch(e=>{console.error(e);process.exitCode=1;});
