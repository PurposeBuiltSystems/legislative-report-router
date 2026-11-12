'use strict';
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
function load(pca, initialize) {
 const context = {self:{}, msal:{createNestablePublicClientApplication:initialize || (async()=>pca)},setTimeout,clearTimeout,Office:{context:{mailbox:{userProfile:{emailAddress:'sample@example.org'}},roamingSettings:{get(){return false},set(){},saveAsync(fn){fn()}}}}};
 vm.runInNewContext(fs.readFileSync(require('node:path').join(__dirname,'../src/graph.js'),'utf8'),context);
 return context.self.GraphData;
}
(async()=>{
 let popups=0;
 const g=load({getAllAccounts:()=>[],acquireTokenSilent:async()=>({accessToken:'token',account:{username:'sample@example.org'}}),acquireTokenPopup:async()=>{popups++;return {accessToken:'interactive'}}});
 assert.equal(await g.getToken(),'token');assert.equal(await g.currentAccount(),'sample@example.org');
 await g.signOut();assert.equal(await g.currentAccount(),null);assert.equal(await g.getToken(),'interactive');assert.equal(popups,1);
 let finish;let entered; const started=new Promise(r=>entered=r);
 const race=load({getAllAccounts:()=>[],acquireTokenSilent:()=>{entered();return new Promise(r=>finish=r)},acquireTokenPopup:async()=>{throw Error('must not open popup after sign out')}});
 const pending=race.getToken();await started;await race.signOut();finish({accessToken:'stale'});await assert.rejects(pending,/cancelled/);assert.equal(await race.currentAccount(),null);
 let completePopup; let popupEntered;
 const popupStarted = new Promise(r=>popupEntered=r);
 const popupRace=load({getAllAccounts:()=>[],acquireTokenSilent:async()=>{throw Error('interaction required')},acquireTokenPopup:()=>{popupEntered();return new Promise(r=>completePopup=r)}});
 const popupPending=popupRace.getToken(); await popupStarted; await popupRace.signOut(); completePopup({accessToken:'stale popup'});
 await assert.rejects(popupPending,/cancelled/); assert.equal(popupRace.isSignedOut(),true);
 // A fresh pane signs in silently; it never opens a window to do so.
 let quietPopups=0;
 const fresh=load({getAllAccounts:()=>[],acquireTokenSilent:async()=>({accessToken:'t',account:{username:'sample@example.org'}}),acquireTokenPopup:async()=>{quietPopups++;return {accessToken:'x'}}});
 assert.equal(await fresh.currentAccount(),null);
 assert.equal(await fresh.trySilentSignIn(),'sample@example.org');
 assert.equal(await fresh.currentAccount(),'sample@example.org');
 const denied=load({getAllAccounts:()=>[],acquireTokenSilent:async()=>{throw Error('interaction required')},acquireTokenPopup:async()=>{quietPopups++;return {accessToken:'x'}}});
 assert.equal(await denied.trySilentSignIn(),null);
 await fresh.signOut(); assert.equal(await fresh.trySilentSignIn(),null,'sign-out is honoured');
 assert.equal(quietPopups,0,'silent sign-in never opens a window');
 // Simultaneous actions share one sign-in window.
 let shared=0, release;
 const multi=load({getAllAccounts:()=>[],acquireTokenSilent:async()=>{throw Error('interaction required')},acquireTokenPopup:()=>{shared++;return new Promise(r=>release=r)}});
 const both=Promise.all([multi.getToken(),multi.getToken(),multi.getToken()]);
 await new Promise(r=>setTimeout(r,10)); release({accessToken:'one',account:{username:'sample@example.org'}});
 assert.deepEqual(await both,['one','one','one']); assert.equal(shared,1,'one popup for concurrent requests');
 // Transient broker startup failures recover on the same action, and a
 // persistent failure cannot poison the cache for the next action.
 const client={acquireTokenSilent:async()=>({accessToken:'recovered'})};
 let creates=0;
 const startup=load(client,async()=>{creates++;if(creates===1)throw Error('broker starting');return client;});
 assert.equal(await startup.getToken(),'recovered');
 assert.equal(creates,2,'retry initialization once on the same action');
 assert.equal(await startup.getToken(),'recovered');
 assert.equal(creates,2,'successful initialization stays cached');
 let broken=true, attempts=0;
 const recovery=load(client,()=>{attempts++;if(broken)throw Error('broker unavailable');return Promise.resolve(client);});
 await assert.rejects(recovery.getToken(),/broker unavailable/);
 assert.equal(attempts,2,'persistent initialization failure is bounded');
 broken=false;
 assert.equal(await recovery.getToken(),'recovered');
 assert.equal(attempts,3,'a later action starts a fresh initialization');
 // An old rejected promise must not erase the new cloud's successful client.
 let rejectOld, cloudCreates=0;
 const cloudRace=load(client,()=>{cloudCreates++;return cloudCreates===1?new Promise((resolve,reject)=>{rejectOld=reject;}):Promise.resolve(client);});
 const old=cloudRace.getToken();
 cloudRace.setCloud('gcchigh');
 assert.equal(await cloudRace.getToken(),'recovered');
 const rejected=assert.rejects(old,/old broker failed/);
 rejectOld(Error('old broker failed'));await rejected;
 assert.equal(await cloudRace.getToken(),'recovered');
 assert.equal(cloudCreates,2,'late rejection preserves the newer cloud client');
 console.log('Authentication state and sign-out race regressions passed.');
})().catch(e=>{console.error(e);process.exitCode=1});
