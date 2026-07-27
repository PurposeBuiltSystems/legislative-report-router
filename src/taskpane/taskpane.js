/*
 * Legislative Report Router — task pane wiring.
 *
 * Workflow: Parse → Review (edit/exclude/re-route) → Preview → Publish
 * (confirm, idempotent, per-operation status, retry-failures-only) → Audit.
 * State lives in memory + roamingSettings drafts; publication state is
 * durable in the LegislativeAudit SharePoint list (also the dedupe store).
 */
/* global Office, GraphData, LrrParser, LrrRouting, LrrTeams, LrrDocx, document */
(function () {
  "use strict";

  var SETTINGS_KEY = "lrr.settings";
  var DRAFT_KEY = "lrr.draft";
  var RULES_CACHE_KEY = "lrr.rulesCache";

  var state = {
    subject: "",
    reportKey: "",
    items: [],
    rules: [],
    site: null,        // {siteId, routingListId, auditListId}
    results: {},       // idempotencyKey -> {status, error, messageId}
    lastSaved: null,
  };

  function byId(id) { return document.getElementById(id); }

  /** Translate raw Graph/MSAL errors into words a coordinator can act on. */
  function friendly(e) {
    var m = (e && e.message) || String(e);
    if (/403|Authorization_RequestDenied|accessDenied/i.test(m)) {
      return "You don't have permission there yet \u2014 ask IT to give you edit access to the site, then try again.";
    }
    if (/interaction_required|AADSTS|consent|login_required/i.test(m)) {
      return "Sign-in needed \u2014 click the button again and finish the Microsoft sign-in window.";
    }
    if (/Failed to fetch|NetworkError/i.test(m)) {
      return "Can't reach Microsoft 365 right now \u2014 check your connection and try again.";
    }
    if (/404|itemNotFound|not found/i.test(m)) {
      return "Not found \u2014 the site address may be wrong, or the list doesn't exist yet (use \u2461 Create my lists).";
    }
    return m.length > 180 ? m.slice(0, 180) + "\u2026" : m;
  }

  /** Live setup progress — tells a non-technical coordinator exactly where she is. */
  function updateSetupChecklist() {
    var el = byId("setupChecklist");
    if (!el) { return; }
    var st = settings();
    var withChannel = state.rules.filter(function (r) { return r.teamsTeamId && r.teamsChannelId; }).length;
    var withPeople = state.rules.filter(function (r) { return (r.emails || []).length; }).length;
    var distN = (typeof LrrReportGen !== "undefined") ? LrrReportGen.extractEmails(st.distList || "").length : 0;
    var steps = [
      [!!st.siteUrl, "Pick where your lists live", "use \u2460 above"],
      [!!(state.site && state.site.routingListId), "Create & connect your lists", "use \u2461 and Connect"],
      [withChannel > 0, "Add your divisions", state.rules.length ? withChannel + " of " + state.rules.length + " have a Teams channel" : "use \u2462"],
      [withPeople > 0, "Put people on divisions", "quick add or paste a list"],
      [distN > 0, "Paste the report's To: line", distN ? distN + " recipients ready" : "in \u201cLet the add-in write\u2026\u201d"],
      [!!st.lastCheckOk, "Run \u2705 Test my setup", st.lastCheckOk ? "last passed " + new Date(st.lastCheckOk).toLocaleDateString() : "one click, checks everything"],
    ];
    el.innerHTML = steps.map(function (x, i) {
      return '<p class="' + (x[0] ? "done" : "todo") + '">' + (x[0] ? "\u2705" : "\u2b55") + " " +
        (i + 1) + ". " + x[1] + (x[0] ? "" : " \u2014 " + x[2]) + "</p>";
    }).join("");
  }
  function esc(s) { return LrrTeams._internals.esc(s); }

  function setStatus(kind, text) {
    var el = byId("status");
    if (!text) { el.hidden = true; return; }
    el.hidden = false;
    el.className = "status " + kind;
    el.textContent = text;
  }

  function settings() {
    try { return JSON.parse(Office.context.roamingSettings.get(SETTINGS_KEY) || "{}"); }
    catch (e) { return {}; }
  }
  function saveSettings(patch) {
    var s = settings();
    Object.keys(patch).forEach(function (k) { s[k] = patch[k]; });
    Office.context.roamingSettings.set(SETTINGS_KEY, JSON.stringify(s));
    Office.context.roamingSettings.saveAsync(function () {});
    return s;
  }

  function updateConnBanner() {
    var el = byId("connBanner");
    if (!el) { return; }
    var st = settings();
    if (!st.siteUrl) {
      el.className = "conn-banner setup";
      el.innerHTML = "<b>First time?</b> Open \u2699\ufe0f Setup below \u2014 about 10 minutes, one time only. After that this pane is ready on every email automatically.";
    } else if (state.rules.length) {
      el.className = "conn-banner ok";
      el.textContent = "\u2713 Setup saved \u2014 " + state.rules.length + " division route(s) connected. Nothing to redo: just Parse.";
    } else {
      el.className = "conn-banner";
      el.textContent = "Saved setup found \u2014 reconnecting\u2026";
    }
  }

  function saveRulesCache() {
    try {
      var blob = JSON.stringify({ site: state.site, rules: state.rules, at: new Date().toISOString() });
      if (blob.length < 24000) {
        Office.context.roamingSettings.set(RULES_CACHE_KEY, blob);
        Office.context.roamingSettings.saveAsync(function () {});
      }
    } catch (e) { /* cache is best-effort */ }
  }

  function loadRulesCache() {
    try {
      var c = JSON.parse(Office.context.roamingSettings.get(RULES_CACHE_KEY) || "null");
      if (c && c.site && c.rules && c.rules.length) {
        state.site = c.site;
        state.rules = c.rules;
        return true;
      }
    } catch (e) { /* ignore */ }
    return false;
  }

  var autoParsed = false;
  var REPORTED_KEY = "lrr.reported";

  function reportedBills() {
    try { return (JSON.parse(Office.context.roamingSettings.get(REPORTED_KEY) || "{}").bills) || []; }
    catch (e) { return []; }
  }

  function markReported(items) {
    try {
      var bills = reportedBills();
      items.forEach(function (it) {
        if (bills.indexOf(it.billNumber) === -1) { bills.push(it.billNumber); }
      });
      Office.context.roamingSettings.set(REPORTED_KEY, JSON.stringify({
        bills: bills.slice(-800), last: new Date().toISOString(),
      }));
      Office.context.roamingSettings.saveAsync(function () {});
    } catch (e) { /* best-effort */ }
  }

  function updateDistInfo() {
    var el = byId("distInfo");
    if (!el) { return; }
    var n = LrrReportGen.extractEmails(byId("distList").value).length;
    el.textContent = n ? n + " address(es) recognized." : "Paste the To: line from a past report \u2014 names and brackets are fine.";
  }

  function maybeAutoParse() {
    if (autoParsed || state.items.length) { return; }
    if (!/bill\s+report|daily\s+bill/i.test(state.subject || "")) { return; }
    var item = Office.context.mailbox.item;
    if (!item || !item.itemId) { return; } // read mode only
    autoParsed = true;
    setStatus("info", "This looks like a bill report \u2014 parsing it now\u2026");
    parseReport();
  }

  // ---------- screens ----------

  function show(screen) {
    ["overview", "filings", "review", "preview", "publish", "audit"].forEach(function (s) {
      byId("screen-" + s).hidden = s !== screen;
    });
    document.querySelectorAll(".tab").forEach(function (t) {
      t.classList.toggle("active", t.getAttribute("data-screen") === screen);
    });
    if (screen === "preview") { renderPreview(); }
    if (screen === "publish") { renderPublishSummary(); }
  }

  var SETTING_KEYS = ["cloud", "siteUrl", "routingList", "auditList", "trackerList", "commentWindow", "watchTerms", "watchDays", "stateName", "identifiers", "trackedChapters", "sessionName", "distList", "autoDaily", "autoDailyTime"];
  var PROFILE_KEYS = SETTING_KEYS.concat([]);

  Office.onReady(function () {
    var s = settings();
    SETTING_KEYS.forEach(function (k) {
      if (s[k] != null && s[k] !== "") { byId(k).value = s[k]; }
    });
    if (s.cloud) { GraphData.setCloud(s.cloud); }
    if (!s.siteUrl) { byId("settings").setAttribute("open", "open"); }

    var sel = byId("stateName");
    LrrPresets.ALL_STATE_NAMES.forEach(function (n) {
      var o = document.createElement("option");
      o.value = o.textContent = n;
      sel.appendChild(o);
    });
    sel.value = s.stateName || "Iowa";
    if (!s.trackedChapters) { byId("trackedChapters").value = LrrChapters.DEFAULT_TRACKED.join(", "); }
    if (!s.identifiers) { byId("identifiers").value = LrrPresets.presetFor(sel.value).identifiers.join(", "); }
    sel.addEventListener("change", function () {
      var preset = LrrPresets.presetFor(sel.value);
      byId("identifiers").value = preset.identifiers.join(", ");
      saveSettings({ stateName: sel.value, identifiers: byId("identifiers").value });
    });
    byId("profileCopy").addEventListener("click", profileCopy);
    byId("profileApply").addEventListener("click", profileApply);

    document.querySelectorAll(".tab").forEach(function (t) {
      t.addEventListener("click", function () { show(t.getAttribute("data-screen")); });
    });
    byId("parse").addEventListener("click", parseReport);
    byId("saveDraft").addEventListener("click", saveDraft);
    byId("loadDraft").addEventListener("click", loadDraft);
    byId("connectRules").addEventListener("click", function () { connectRules(false).catch(function () {}); });
    byId("createLists").addEventListener("click", createLists);
    byId("siteSearchGo").addEventListener("click", siteSearch);
    byId("siteResults").addEventListener("change", function () {
      if (byId("siteResults").value) {
        byId("siteUrl").value = byId("siteResults").value;
        saveSettings({ siteUrl: byId("siteResults").value });
        state.site = null;
      }
    });

    byId("rbTeam").addEventListener("change", loadChannelsAndTags);
    byId("rbAdd").addEventListener("click", addRoute);
    byId("onlyUnassigned").addEventListener("change", function () { renderItems(); });
    byId("contactPreview").addEventListener("click", contactPreview);
    byId("contactApply").addEventListener("click", contactApply);
    byId("quickAdd").addEventListener("click", quickAddContact);
    byId("wizPrev").addEventListener("click", function () { wizShow(wizStep - 1); });
    byId("wizNext").addEventListener("click", function () { wizShow(wizStep + 1); });
    byId("wizSignIn").addEventListener("click", async function () {
      byId("wizSignIn").disabled = true;
      try {
        await GraphData.getToken();
        byId("wizSignInInfo").textContent = "\u2713 Signed in. Click Next.";
      } catch (e) {
        byId("wizSignInInfo").textContent = friendly(e);
      } finally { byId("wizSignIn").disabled = false; }
    });
    byId("tagCreate").addEventListener("click", createDivisionTag);
    wizShow(1);
    byId("runChecks").addEventListener("click", runChecks);
    byId("copyChecks").addEventListener("click", copyChecks);
    byId("distList").addEventListener("input", updateDistInfo);
    updateDistInfo();
    var st0 = settings();
    if (st0.autoDaily === true || st0.autoDaily === "true") { byId("autoDaily").checked = true; }
    byId("autoDaily").addEventListener("change", function () {
      saveSettings({ autoDaily: byId("autoDaily").checked });
    });
    startAutoDraftTimer();
    byId("lookupTags").addEventListener("click", lookupTags);
    byId("bulkApply").addEventListener("click", bulkApply);
    byId("confirmBox").addEventListener("change", function () {
      byId("publishGo").disabled = !byId("confirmBox").checked;
    });
    byId("publishGo").addEventListener("click", function () { publish(false); });
    byId("retryFailed").addEventListener("click", function () { publish(true); });
    byId("refreshAudit").addEventListener("click", refreshAudit);
    byId("loadFilings").addEventListener("click", loadFilings);
    byId("routeFilings").addEventListener("click", routeFilings);
    SETTING_KEYS.forEach(function (k) {
      byId(k).addEventListener("change", function () {
        var p = {}; p[k] = byId(k).value; saveSettings(p);
        if (k === "cloud") { GraphData.setCloud(byId(k).value); state.site = null; }
        if (k === "siteUrl" || k === "routingList" || k === "auditList" || k === "trackerList") { state.site = null; }
      });
    });

    var item = Office.context.mailbox.item;
    if (item && item.subject && typeof item.subject === "string") {
      state.subject = item.subject; // read mode: plain string
      byId("stSubject").textContent = state.subject;
    } else if (item && item.subject && item.subject.getAsync) {
      item.subject.getAsync(function (r) { // compose mode
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          state.subject = r.value || "(no subject)";
          byId("stSubject").textContent = state.subject;
        }
      });
    }

    // Saved setup: reconnect without being asked. Cached rules make the pane
    // useful instantly; a silent refresh follows. If sign-in needs interaction
    // (first run on a new device), we say so calmly instead of erroring.
    var hadCache = loadRulesCache();
    updateConnBanner();
    updateSetupChecklist();
    if (hadCache) {
      byId("rulesInfo").textContent = state.rules.length + " routing rule(s) from saved setup.";
      maybeAutoParse();
    }
    if (s.siteUrl) {
      setTimeout(function () {
        connectRules(true).then(function () {
          updateConnBanner();
          maybeAutoParse();
        }).catch(function () {
          if (!hadCache) {
            var el = byId("connBanner");
            el.className = "conn-banner setup";
            el.textContent = "Saved setup found \u2014 click \u2699\ufe0f Setup \u2192 \u201cConnect & load routing rules\u201d to sign in on this device (one time).";
          }
        });
      }, 350);
    }
  });

  // ---------- source reading ----------

  function getBodyHtml() {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) { resolve(r.value || ""); }
        else { reject(new Error("Couldn't read the message body: " + (r.error && r.error.message))); }
      });
    });
  }

  function restMessageId() {
    return new Promise(function (resolve, reject) {
      var item = Office.context.mailbox.item;
      if (item.itemId) {
        resolve(Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0));
        return;
      }
      // compose: save the draft to obtain an id
      if (item.saveAsync) {
        item.saveAsync(function (r) {
          if (r.status === Office.AsyncResultStatus.Succeeded) {
            resolve(Office.context.mailbox.convertToRestId(r.value, Office.MailboxEnums.RestVersion.v2_0));
          } else { reject(new Error("Couldn't save the draft to read attachments / send.")); }
        });
      } else { reject(new Error("This item has no id and can't be saved here.")); }
    });
  }

  async function docxText(token) {
    var msgId = await restMessageId();
    var atts = await GraphData.getAttachments(token, msgId);
    var docx = atts.find(function (a) { return LrrDocx.isSupported(a.name, a.contentType, a.size).ok; });
    if (!docx) { throw new Error("No .docx attachment found on this message."); }
    var bytes = await GraphData.getAttachmentBytes(token, msgId, docx.id);
    var extracted = await LrrDocx.extractText(bytes, docx.name, docx.contentType, docx.size);
    return extracted.text;
  }

  // ---------- parse ----------

  async function parseReport() {
    byId("parse").disabled = true;
    try {
      setStatus("work", "Reading the report…");
      var source = byId("source").value;
      var text = "";
      if (source === "body" || source === "both") {
        text += await getBodyHtml();
      }
      if (source === "docx" || source === "both") {
        var token = await GraphData.getToken();
        text = (source === "both" ? text + "\n" : "") + await docxText(token);
      }

      state.reportKey = (state.subject || "report") .toLowerCase().replace(/\s+/g, "-").slice(0, 80);
      var known = state.rules.length
        ? state.rules.reduce(function (acc, r) {
            [r.divisionCode, r.divisionName].concat(r.aliases || []).forEach(function (n) {
              if (n && acc.indexOf(n) === -1) { acc.push(n); }
            });
            return acc;
          }, [])
        : [];
      var ids = byId("identifiers").value.split(",").map(function (t) { return t.trim().toUpperCase(); }).filter(Boolean);
      var res = LrrParser.parseReport(text, { knownDivisions: known, reportId: state.reportKey, identifiers: ids.length ? ids : undefined });
      state.items = res.items;
      state.results = {};
      annotateChapters();

      if (state.rules.length) { LrrRouting.routeAll(state.items, state.rules); }
      refreshStats();
      renderItems();

      var wb = byId("parserWarnings");
      var allWarn = res.warnings.slice();
      if (!state.rules.length) { allWarn.push("Routing rules are not loaded — connect the SharePoint routing list in Settings to match divisions."); }
      wb.hidden = !allWarn.length;
      wb.textContent = allWarn.join(" ");

      setStatus("info", state.items.length + " bill(s) parsed. Review the distribution next.");
      if (state.items.length) { show("review"); }
    } catch (e) {
      setStatus("error", "Parse failed: " + friendly(e));
    } finally {
      byId("parse").disabled = false;
    }
  }

  function trackedList() {
    return LrrChapters.parseTrackedText(byId("trackedChapters").value);
  }

  function annotateChapters() {
    var tracked = trackedList();
    state.items.forEach(function (it) {
      it.codeChapters = LrrChapters.extractChapters(it.brief || "");
      it.trackedChapters = LrrChapters.matchTracked(it.codeChapters, tracked);
    });
  }

  function refreshStats() {
    var summary = LrrRouting.routeAll(state.items, state.rules);
    byId("stBills").textContent = summary.totalItems;
    byId("stMatched").textContent = summary.matchedItems;
    byId("stUnmatched").textContent = summary.unmatchedItems + summary.partialItems;
    byId("stWarnings").textContent = state.items.filter(function (i) { return i.parserWarnings.length; }).length;
    byId("stSaved").textContent = state.lastSaved ? state.lastSaved : "never";
  }

  // ---------- review ----------

  function renderItems() {
    var host = byId("items");
    host.innerHTML = "";
    var bulk = byId("bulkRoute");
    bulk.innerHTML = "";
    state.rules.forEach(function (r) {
      var o = document.createElement("option");
      o.value = r.id;
      o.textContent = r.divisionCode + (r.teamsChannelName ? " → " + r.teamsChannelName : "");
      bulk.appendChild(o);
    });

    state.items.forEach(function (it, idx) {
      var card = document.createElement("div");
      card.className = "item " + it.routingStatus;

      var head = document.createElement("div");
      head.className = "item-head";
      var inc = document.createElement("input");
      inc.type = "checkbox";
      inc.checked = it.routingStatus !== "excluded";
      inc.addEventListener("change", function () {
        it.routingStatus = inc.checked ? "unmatched" : "excluded";
        if (inc.checked) { LrrRouting.routeItem(it, state.rules); }
        refreshStats();
        card.className = "item " + it.routingStatus;
        badge.textContent = statusText(it);
      });
      var bill = document.createElement("input");
      bill.className = "bill";
      bill.value = it.billNumber;
      bill.addEventListener("change", function () { it.billNumber = bill.value.trim(); });
      var badge = document.createElement("span");
      badge.className = "badge";
      badge.textContent = statusText(it);
      var conf = document.createElement("span");
      conf.className = "conf" + (it.parserConfidence < 0.7 ? " low" : "");
      conf.textContent = Math.round(it.parserConfidence * 100) + "%";
      head.appendChild(inc); head.appendChild(bill); head.appendChild(badge); head.appendChild(conf);
      card.appendChild(head);

      var divs = document.createElement("input");
      divs.value = it.distributedTo.join(", ");
      divs.title = "Distributed To (comma/semicolon separated)";
      divs.addEventListener("change", function () {
        it.distributedTo = LrrParser.normalizeDivisions(divs.value);
        LrrRouting.routeItem(it, state.rules);
        refreshStats(); renderItems();
      });
      var cf = document.createElement("input");
      cf.value = it.commentRequestedFrom.join(", ");
      cf.title = "Comment Requested From";
      cf.addEventListener("change", function () {
        it.commentRequestedFrom = LrrParser.normalizeDivisions(cf.value);
        LrrRouting.routeItem(it, state.rules);
        refreshStats(); renderItems();
      });
      var lblRow = document.createElement("div");
      lblRow.className = "row2";
      var l1 = document.createElement("label"); l1.textContent = "Distributed To"; l1.appendChild(divs);
      var l2 = document.createElement("label"); l2.textContent = "Comment From"; l2.appendChild(cf);
      lblRow.appendChild(l1); lblRow.appendChild(l2);
      card.appendChild(lblRow);

      if (it.unknownDivisions && it.unknownDivisions.length) {
        var unk = document.createElement("p");
        unk.className = "unknown";
        unk.textContent = "No active route for: " + it.unknownDivisions.join(", ") + " — pick a route below, fix the division, or exclude this bill.";
        card.appendChild(unk);
      }
      if (it.parserWarnings.length) {
        var pw = document.createElement("p");
        pw.className = "hint warn";
        pw.textContent = "⚠ " + it.parserWarnings.join(" ");
        card.appendChild(pw);
      }

      if (it.codeChapters && it.codeChapters.length) {
        var chap = document.createElement("p");
        chap.className = "routes";
        chap.innerHTML = "Code: " + it.codeChapters.map(function (c) {
          var hit = (it.trackedChapters || []).indexOf(c) !== -1;
          return '<span class="chip' + (hit ? "" : " chip-dim") + '">' + esc(c) + (hit ? " ✓" : "") + "</span>";
        }).join(" ");
        card.appendChild(chap);
      }
      if ((it.routingStatus === "unmatched" || (it.unknownDivisions || []).length) && state.rules.length) {
        var byRule = {};
        LrrChapters.suggestRules(it.codeChapters || [], state.rules).forEach(function (sg) {
          byRule[sg.rule.id] = { rule: sg.rule, why: "ch. " + sg.chapters.join(", ") };
        });
        LrrRouting.suggestByKeywords((it.title || "") + "\n" + (it.brief || ""), state.rules).forEach(function (sg) {
          var why = '\u201c' + sg.keywords.join('\u201d, \u201c') + '\u201d';
          if (byRule[sg.rule.id]) { byRule[sg.rule.id].why += " + " + why; }
          else { byRule[sg.rule.id] = { rule: sg.rule, why: why }; }
        });
        var sugIds = Object.keys(byRule);
        if (sugIds.length) {
          var sp = document.createElement("p");
          sp.className = "routes";
          sp.appendChild(document.createTextNode("Suggested: "));
          sugIds.slice(0, 4).forEach(function (idKey) {
            var sg = byRule[idKey];
            var b = document.createElement("button");
            b.type = "button";
            b.className = "chip-btn";
            b.textContent = sg.rule.divisionCode + " (" + sg.why + ")";
            b.addEventListener("click", function () {
              if (it.distributedTo.indexOf(sg.rule.divisionCode) === -1) { it.distributedTo.push(sg.rule.divisionCode); }
              if (!it.commentRequestedFrom.length) { it.commentRequestedFrom = it.distributedTo.slice(); }
              LrrRouting.routeItem(it, state.rules);
              refreshStats(); renderItems();
            });
            sp.appendChild(b);
          });
          card.appendChild(sp);
        }
        // Teach the router: save a keyword onto a division's rule, apply
        // everywhere in this report, remembered for every future report.
        var teach = document.createElement("div");
        teach.className = "teach";
        var tIn = document.createElement("input");
        tIn.type = "text";
        tIn.placeholder = "keyword or phrase\u2026";
        var tSel = document.createElement("select");
        state.rules.forEach(function (r) {
          var o = document.createElement("option");
          o.value = r.id; o.textContent = r.divisionCode;
          tSel.appendChild(o);
        });
        var tBtn = document.createElement("button");
        tBtn.type = "button";
        tBtn.textContent = "Teach";
        tBtn.title = "Save this keyword to the division's routing rule and route matching bills";
        tBtn.addEventListener("click", function () { teachKeyword(it, tIn.value.trim(), tSel.value, tBtn); });
        teach.appendChild(tIn); teach.appendChild(tSel); teach.appendChild(tBtn);
        // clickable word chips from the title seed the keyword box
        var STOP = { with: 1, from: 1, that: 1, this: 1, certain: 1, relating: 1, regarding: 1, department: 1, state: 1, iowa: 1, bill: 1, act: 1, code: 1 };
        var words = String(it.title || it.brief || "").toLowerCase().replace(/[^a-z0-9' -]/g, " ").split(/\s+/)
          .filter(function (w) { return w.length >= 4 && !STOP[w]; }).slice(0, 6);
        if (words.length) {
          var wc = document.createElement("div");
          wc.className = "teach-words";
          words.forEach(function (w) {
            var wb = document.createElement("button");
            wb.type = "button"; wb.className = "chip-btn"; wb.textContent = w;
            wb.addEventListener("click", function () { tIn.value = tIn.value ? tIn.value + " " + w : w; });
            wc.appendChild(wb);
          });
          teach.appendChild(wc);
        }
        card.appendChild(teach);
      }
      var routes = document.createElement("p");
      routes.className = "routes";
      routes.innerHTML = (it.routes || []).map(function (r) {
        return '<span class="chip">' + esc(r.divisionCode) +
          (r.teamsTagName ? " @" + esc(r.teamsTagName) : ' <span class="chip-warn">(no tag — posts without mention)</span>') + "</span>";
      }).join(" ") || '<span class="chip-warn">No Teams route</span>';
      card.appendChild(routes);

      var brief = document.createElement("textarea");
      brief.rows = 3;
      brief.value = it.brief;
      brief.addEventListener("change", function () { it.brief = brief.value; });
      card.appendChild(brief);

      host.appendChild(card);
    });
  }

  function statusText(it) {
    return { matched: "matched", "partially-matched": "partial", unmatched: "unmatched", excluded: "excluded" }[it.routingStatus] || it.routingStatus;
  }

  async function teachKeyword(sourceItem, keyword, ruleId, btn) {
    if (!keyword) { setStatus("error", "Type a keyword first (or click a word chip)."); return; }
    var rule = state.rules.filter(function (r) { return r.id === ruleId; })[0];
    if (!rule) { return; }
    if (!state.site || !state.site.routingListId) {
      setStatus("error", "Connect the routing list first so the keyword can be saved."); return;
    }
    btn.disabled = true;
    try {
      setStatus("work", 'Saving \u201c' + keyword + '\u201d to ' + rule.divisionCode + "\u2026");
      var token = await GraphData.getToken();
      var merged = (rule.keywords || []).concat([keyword]);
      await GraphData.updateListItemFields(token, state.site.siteId, state.site.routingListId, rule.id,
        { RoutingKeywords: merged.join("; ") });
      rule.keywords = merged;
      saveRulesCache();
      // apply across every unassigned bill in this report
      var applied = 0;
      state.items.forEach(function (it) {
        if (it.routingStatus === "excluded") { return; }
        var hit = LrrRouting.keywordHits((it.title || "") + "\n" + (it.brief || ""), [keyword]).length;
        if (hit && it.distributedTo.indexOf(rule.divisionCode) === -1) {
          it.distributedTo.push(rule.divisionCode);
          if (!it.commentRequestedFrom.length) { it.commentRequestedFrom = [rule.divisionCode]; }
          LrrRouting.routeItem(it, state.rules);
          applied++;
        }
      });
      refreshStats(); renderItems();
      setStatus("info", '\u201c' + keyword + '\u201d saved to ' + rule.divisionCode +
        " \u2014 routed " + applied + " bill(s) now, and every future report remembers it.");
    } catch (e) {
      setStatus("error", "Couldn't save the keyword: " + friendly(e));
      btn.disabled = false;
    }
  }

  function bulkApply() {
    var ruleId = byId("bulkRoute").value;
    var rule = state.rules.find(function (r) { return r.id === ruleId; });
    if (!rule) { return; }
    state.items.forEach(function (it) {
      if (it.routingStatus === "excluded") { return; }
      it.distributedTo = [rule.divisionCode];
      LrrRouting.routeItem(it, state.rules);
    });
    refreshStats(); renderItems();
    setStatus("info", "Route " + rule.divisionCode + " applied to all included bills.");
  }

  // ---------- setup wizard ----------

  var wizStep = 1;
  var WIZ_MAX = 6;

  function wizShow(n) {
    wizStep = Math.max(1, Math.min(WIZ_MAX, n));
    document.querySelectorAll(".wiz-step").forEach(function (el) {
      el.hidden = Number(el.getAttribute("data-wiz")) !== wizStep;
    });
    byId("wizPos").textContent = "Step " + wizStep + " of " + WIZ_MAX;
    byId("wizPrev").disabled = wizStep === 1;
    byId("wizNext").textContent = wizStep === WIZ_MAX ? "Done \u2713" : "Next \u2192";
    if (wizStep === 3) { loadTeamsPicker(); renderRoutesSummary(); }
    if (wizStep === 4 || wizStep === 6) { /* dropdowns already refreshed on connect */ }
    if (wizStep === WIZ_MAX && byId("wizNext").textContent === "Done \u2713") {
      byId("wizNext").onclick = null;
    }
  }

  function renderRoutesSummary() {
    var el = byId("routesSummary");
    if (!el) { return; }
    if (!state.rules.length) { el.innerHTML = "No divisions yet \u2014 add your first below."; return; }
    el.innerHTML = state.rules.map(function (r) {
      var chan = r.teamsChannelId ? (r.teamsChannelName || "channel set") : "\u26a0 no channel";
      var tag = r.teamsTagId ? "\ud83c\udff7 " + (r.teamsTagName || "tag set") : "\u26a0 no tag";
      return "<b>" + r.divisionCode + "</b> \u2192 " + chan + " \u00b7 " + tag;
    }).join("<br>");
  }

  async function createDivisionTag() {
    var teamId = byId("rbTeam").value;
    var code = byId("rbCode").value.trim();
    if (!teamId) { byId("rbInfo").textContent = "Pick a team first."; return; }
    if (!code) { byId("rbInfo").textContent = "Type the division code first (the tag will be named after it)."; return; }
    var name = (byId("rbName").value.trim() || code) + " Legislation";
    byId("tagCreate").disabled = true;
    try {
      setStatus("work", 'Creating tag \u201c' + name + '\u201d\u2026');
      var token = await GraphData.getToken();
      var tag = await GraphData.createTeamTag(token, teamId, name);
      var sel = byId("rbTag");
      var o = document.createElement("option");
      o.value = tag.id; o.textContent = tag.displayName;
      sel.appendChild(o);
      sel.value = tag.id;
      byId("rbInfo").textContent = '\u2713 Tag \u201c' + name + '\u201d created with you as its first member \u2014 add the division\u2019s people to it in Teams (Manage team \u2192 Tags).';
      setStatus("info", "Tag created and selected.");
    } catch (e) {
      byId("rbInfo").textContent = "Couldn't create the tag: " + friendly(e) +
        " (You can also create it in Teams: Manage team \u2192 Tags.)";
      setStatus("error", "Tag creation failed \u2014 details above the button.");
    } finally {
      byId("tagCreate").disabled = false;
    }
  }

  // ---------- org profile ----------

  function profileCopy() {
    var st = settings();
    var out = {};
    PROFILE_KEYS.forEach(function (k) { if (st[k]) { out[k] = st[k]; } });
    var blob = btoa(unescape(encodeURIComponent(JSON.stringify(out))));
    byId("profileBlob").value = blob;
    byId("profileBlob").select();
    try { document.execCommand("copy"); } catch (e) { /* manual copy */ }
    setStatus("info", "Profile code is in the box (and copied). Send it to your team.");
  }

  function profileApply() {
    try {
      var raw = byId("profileBlob").value.trim();
      if (!raw) { setStatus("error", "Paste the profile code first."); return; }
      var p = JSON.parse(decodeURIComponent(escape(atob(raw))));
      var patch = {};
      PROFILE_KEYS.forEach(function (k) { if (p[k] != null) { patch[k] = p[k]; } });
      saveSettings(patch);
      SETTING_KEYS.forEach(function (k) { if (patch[k] != null) { byId(k).value = patch[k]; } });
      if (patch.cloud) { GraphData.setCloud(patch.cloud); }
      state.site = null;
      setStatus("info", "Profile applied — click \"Connect & load routing rules\" to finish.");
    } catch (e) {
      setStatus("error", "That doesn't look like a valid profile code.");
    }
  }

  // ---------- new filings (intraday feed) ----------

  var feedSelection = []; // [{entry, checked}]

  async function loadFilings() {
    byId("loadFilings").disabled = true;
    try {
      setStatus("work", "Checking the newly-filed feed…");
      var preset = LrrPresets.presetFor(byId("stateName").value || "Iowa");
      var entries;
      if (preset.feed === "iowa-rss") {
        var res = await fetch("../../feeds/IowaBills.xml?ts=" + Date.now());
        if (!res.ok) { throw new Error("Feed mirror unavailable (" + res.status + ")"); }
        entries = LrrFeed.parseFeed(await res.text());
      } else {
        var res2 = await fetch("../../feeds/openstates-" + preset.slug + ".json?ts=" + Date.now());
        if (!res2.ok) {
          throw new Error("No mirrored feed for " + preset.state + " yet — add it to states.json in the deployment repo (requires the Open States API key; see the admin guide).");
        }
        entries = LrrFeed.parseOpenStates(await res2.text());
      }
      var terms = byId("watchTerms").value.split(",").map(function (t) { return t.trim(); }).filter(Boolean);
      var days = Number(byId("watchDays").value) || 3;
      var hits = LrrFeed.watchFilter(entries, terms, days);
      if (byId("hideReported").checked) {
        var before = hits.length;
        hits = LrrReportGen.filterNew(hits, reportedBills());
        if (before !== hits.length) {
          setStatus("work", (before - hits.length) + " bill(s) hidden (already in a past report)\u2026");
        }
      }
      if (byId("watchChapters").checked) {
        var tracked = trackedList();
        var inWindow = LrrFeed.watchFilter(entries, [], days);
        inWindow.forEach(function (e) {
          if (hits.indexOf(e) !== -1) { return; }
          var chapters = LrrChapters.extractChapters(e.description || "");
          if (LrrChapters.matchTracked(chapters, tracked).length) { hits.push(e); }
        });
        hits.sort(function (a, b) { return b.pubDate - a.pubDate; });
      }

      feedSelection = hits.map(function (e) { return { entry: e, checked: true }; });
      var host = byId("filingsList");
      host.innerHTML = "";
      hits.forEach(function (e, i) {
        var card = document.createElement("div");
        card.className = "item matched";
        var head = document.createElement("div");
        head.className = "item-head";
        var cb = document.createElement("input");
        cb.type = "checkbox"; cb.checked = true;
        cb.addEventListener("change", function () { feedSelection[i].checked = cb.checked; });
        var b = document.createElement("strong");
        b.textContent = e.bill;
        var when = document.createElement("span");
        when.className = "hint";
        when.textContent = e.pubDate.toLocaleString();
        head.appendChild(cb); head.appendChild(b); head.appendChild(when);
        var p = document.createElement("p");
        p.className = "hint";
        p.textContent = e.description;
        card.appendChild(head); card.appendChild(p);
        host.appendChild(card);
      });
      byId("routeFilings").hidden = !hits.length;
      setStatus("info", hits.length
        ? hits.length + " watched filing(s) in the last " + days + " day(s). Select and route."
        : "No watched filings in the window. (Feed mirror updates every 30 minutes on weekdays.)");
    } catch (e) {
      setStatus("error", "Feed check failed: " + friendly(e));
    } finally {
      byId("loadFilings").disabled = false;
    }
  }

  function routeFilings() {
    var picked = feedSelection.filter(function (f) { return f.checked; });
    if (!picked.length) { setStatus("error", "Nothing selected."); return; }
    if (!state.reportKey) {
      state.reportKey = "new-filings-" + new Date().toISOString().slice(0, 10);
      state.subject = "New filings " + new Date().toLocaleDateString();
    }
    var existing = {};
    state.items.forEach(function (it) { existing[it.billNumber] = true; });
    var added = 0;
    var tracked = trackedList();
    picked.forEach(function (f, i) {
      if (existing[f.entry.bill]) { return; }
      var item = LrrFeed.toLegislativeItem(f.entry, state.reportKey, state.items.length + i);
      item.codeChapters = LrrChapters.extractChapters(item.brief || "");
      item.trackedChapters = LrrChapters.matchTracked(item.codeChapters, tracked);
      if (!item.distributedTo.length && state.rules.length) {
        var why = [];
        var codes = [];
        LrrChapters.suggestRules(item.codeChapters, state.rules).forEach(function (sg) {
          if (codes.indexOf(sg.rule.divisionCode) === -1) { codes.push(sg.rule.divisionCode); }
          why.push(sg.rule.divisionCode + ": ch. " + sg.chapters.join(","));
        });
        LrrRouting.suggestByKeywords(item.brief || "", state.rules).forEach(function (sg) {
          if (codes.indexOf(sg.rule.divisionCode) === -1) { codes.push(sg.rule.divisionCode); }
          why.push(sg.rule.divisionCode + ': "' + sg.keywords.join('", "') + '"');
        });
        if (codes.length) {
          item.distributedTo = codes;
          item.commentRequestedFrom = codes.slice();
          item.parserWarnings = ["Divisions auto-suggested (" + why.join("; ") + ") — verify."];
        }
      }
      if (state.rules.length) { LrrRouting.routeItem(item, state.rules); }
      state.items.push(item);
      added++;
    });
    refreshStats(); renderItems();
    if (byId("optDailyReport")) { byId("optDailyReport").checked = true; }
    var auto = state.items.filter(function (it) { return it.routingStatus === "matched"; }).length;
    setStatus("info", added + " bill(s) added — " + auto + " auto-routed from Code chapters. Verify in Review, then Publish (the Daily Bill Report draft is pre-ticked).");
    show("review");
  }

  // ---------- rules ----------

  async function siteSearch() {
    var q = byId("siteSearch").value.trim();
    if (!q) { setStatus("error", "Type part of the site's name first."); return; }
    byId("siteSearchGo").disabled = true;
    try {
      setStatus("work", "Searching your sites\u2026");
      var token = await GraphData.getToken();
      var sites = await GraphData.searchSites(token, q);
      var sel = byId("siteResults");
      sel.innerHTML = "";
      if (!sites.length) { setStatus("info", "No sites matched \u2014 paste the site URL instead."); sel.hidden = true; return; }
      var opt0 = document.createElement("option");
      opt0.value = ""; opt0.textContent = "Pick a site (" + sites.length + " found)\u2026";
      sel.appendChild(opt0);
      sites.forEach(function (st) {
        var o = document.createElement("option");
        o.value = st.webUrl;
        o.textContent = st.displayName || st.webUrl;
        sel.appendChild(o);
      });
      sel.hidden = false;
      setStatus("info", sites.length + " site(s) found \u2014 pick one.");
    } catch (e) {
      setStatus("error", "Site search failed: " + friendly(e));
    } finally {
      byId("siteSearchGo").disabled = false;
    }
  }

  var teamsCache = null;

  async function loadTeamsPicker() {
    if (teamsCache) { return; }
    try {
      var token = await GraphData.getToken();
      teamsCache = await GraphData.joinedTeams(token);
      var sel = byId("rbTeam");
      sel.innerHTML = "";
      var o0 = document.createElement("option");
      o0.value = ""; o0.textContent = "Pick a team\u2026";
      sel.appendChild(o0);
      teamsCache.forEach(function (t) {
        var o = document.createElement("option");
        o.value = t.id; o.textContent = t.displayName;
        sel.appendChild(o);
      });
    } catch (e) {
      byId("rbInfo").textContent = "Couldn't load your teams: " + friendly(e);
    }
  }

  async function loadChannelsAndTags() {
    var teamId = byId("rbTeam").value;
    var chSel = byId("rbChannel");
    var tagSel = byId("rbTag");
    chSel.innerHTML = ""; tagSel.innerHTML = "";
    if (!teamId) { return; }
    try {
      var token = await GraphData.getToken();
      var channels = await GraphData.listChannels(token, teamId);
      channels.forEach(function (c) {
        var o = document.createElement("option");
        o.value = c.id; o.textContent = c.displayName;
        chSel.appendChild(o);
      });
      var oNone = document.createElement("option");
      oNone.value = ""; oNone.textContent = "(no tag mention)";
      tagSel.appendChild(oNone);
      try {
        var tags = await GraphData.listTeamTags(token, teamId);
        tags.forEach(function (t) {
          var o = document.createElement("option");
          o.value = t.id; o.textContent = t.displayName;
          tagSel.appendChild(o);
        });
        byId("rbInfo").textContent = tags.length ? "" : "This team has no tags yet \u2014 create one in Teams (Manage team > Tags) to @mention a group.";
      } catch (eTags) {
        byId("rbInfo").textContent = "Tags unavailable for this team (posting still works without a mention).";
      }
    } catch (e) {
      byId("rbInfo").textContent = "Couldn't load channels: " + friendly(e);
    }
  }

  async function addRoute() {
    var code = byId("rbCode").value.trim();
    var teamId = byId("rbTeam").value;
    var channelId = byId("rbChannel").value;
    if (!code) { byId("rbInfo").textContent = "Division code is required."; return; }
    if (!teamId || !channelId) { byId("rbInfo").textContent = "Pick a team and channel."; return; }
    byId("rbAdd").disabled = true;
    try {
      setStatus("work", "Saving route " + code + "\u2026");
      var token = await GraphData.getToken();
      if (!state.site) {
        var site = await GraphData.resolveSite(token, byId("siteUrl").value);
        var routingListId = await GraphData.findList(token, site.siteId, byId("routingList").value.trim() || "LegislativeRoutingMatrix");
        state.site = { siteId: site.siteId, routingListId: routingListId };
      }
      if (!state.site.routingListId) {
        state.site.routingListId = await GraphData.findList(token, state.site.siteId, byId("routingList").value.trim() || "LegislativeRoutingMatrix");
      }
      var teamName = (teamsCache || []).filter(function (t) { return t.id === teamId; })[0];
      await GraphData.addListItem(token, state.site.siteId, state.site.routingListId, {
        Title: code,
        DivisionCode: code,
        DivisionName: byId("rbName").value.trim(),
        Emails: byId("rbEmails").value.trim(),
        TeamsTeamId: teamId,
        TeamsChannelId: channelId,
        TeamsChannelName: byId("rbChannel").selectedOptions[0] ? byId("rbChannel").selectedOptions[0].textContent : "",
        TeamsTagId: byId("rbTag").value,
        TeamsTagName: byId("rbTag").selectedOptions[0] && byId("rbTag").value ? byId("rbTag").selectedOptions[0].textContent : "",
        CodeChapters: byId("rbChapters").value.trim(),
        RoutingKeywords: byId("rbKeywords").value.trim(),
        IsActive: true,
        Priority: 1,
        Notes: "Created via the add-in setup" + (teamName ? " (team: " + teamName.displayName + ")" : ""),
      });
      byId("rbInfo").textContent = "Division " + code + " saved." +
        (byId("rbTag").value ? "" : " \u26a0 No tag \u2014 its posts won't notify anyone; use \u2795 Create the tag.");
      byId("rbCode").value = ""; byId("rbName").value = ""; byId("rbEmails").value = ""; byId("rbChapters").value = "";
      await connectRules();
    } catch (e) {
      byId("rbInfo").textContent = "Saving failed: " + friendly(e);
      setStatus("error", "Route not saved \u2014 " + friendly(e));
    } finally {
      byId("rbAdd").disabled = false;
    }
  }

  async function quickAddContact() {
    var email = byId("quickEmail").value.trim().toLowerCase();
    var ruleId = byId("quickDivision").value;
    if (!/^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/.test(email)) { setStatus("error", "That doesn't look like an email address."); return; }
    var rule = state.rules.filter(function (r) { return r.id === ruleId; })[0];
    if (!rule) { setStatus("error", "Connect the routing list first, then pick a division."); return; }
    if ((rule.emails || []).map(function (e) { return e.toLowerCase(); }).indexOf(email) !== -1) {
      setStatus("info", email + " is already on " + rule.divisionCode + "."); return;
    }
    byId("quickAdd").disabled = true;
    try {
      setStatus("work", "Adding " + email + " to " + rule.divisionCode + "\u2026");
      var token = await GraphData.getToken();
      if (!state.site || !state.site.routingListId) { throw new Error("routing list not connected"); }
      var merged = (rule.emails || []).concat([email]);
      await GraphData.updateListItemFields(token, state.site.siteId, state.site.routingListId, rule.id, { Emails: merged.join("; ") });
      rule.emails = merged;
      saveRulesCache();
      byId("quickEmail").value = "";
      updateSetupChecklist();
      setStatus("info", email + " added to " + rule.divisionCode + " (" + merged.length + " recipient(s) now).");
    } catch (e) {
      setStatus("error", "Couldn't add the contact: " + friendly(e));
    } finally {
      byId("quickAdd").disabled = false;
    }
  }

  var contactPlanCache = null;

  function contactPreview() {
    var parsed = LrrContacts.parseContactRows(byId("contactRows").value);
    var plan = LrrContacts.mergePlan(parsed.rows, state.rules);
    contactPlanCache = plan;
    var lines = [];
    plan.forEach(function (pl) {
      lines.push("<b>" + pl.division + "</b>: " +
        (pl.addEmails.length ? "+" + pl.addEmails.length + " address(es)" : "nothing new") +
        (pl.rule ? " (rule exists, " + pl.existingCount + " already)"
                 : " \u2014 <b>new rule will be created</b> (add its Teams channel afterward)"));
    });
    parsed.errors.forEach(function (er) { lines.push('<span style="color:#a4262c">' + er + "</span>"); });
    if (!plan.length && !parsed.errors.length) { lines.push("Nothing to import \u2014 paste rows first."); }
    byId("contactPlan").innerHTML = lines.join("<br>");
    byId("contactApply").disabled = !plan.some(function (pl) { return pl.addEmails.length; });
    if (!state.rules.length) {
      byId("contactPlan").innerHTML += "<br><span style=\"color:#8a6d00\">Connect the routing list first (\u2461/Connect) so imports merge into existing rules.</span>";
    }
  }

  async function contactApply() {
    if (!contactPlanCache) { return; }
    if (!state.site || !state.site.routingListId) {
      setStatus("error", "Connect the SharePoint routing list first (Settings \u2192 Connect)."); return;
    }
    byId("contactApply").disabled = true;
    try {
      var token = await GraphData.getToken();
      var applied = 0, created = 0;
      for (var i = 0; i < contactPlanCache.length; i++) {
        var pl = contactPlanCache[i];
        if (!pl.addEmails.length) { continue; }
        if (pl.rule) {
          var merged = (pl.rule.emails || []).concat(pl.addEmails).join("; ");
          setStatus("work", "Updating " + pl.division + "\u2026");
          await GraphData.updateListItemFields(token, state.site.siteId, state.site.routingListId, pl.rule.id, { Emails: merged });
          applied++;
        } else {
          setStatus("work", "Creating rule for " + pl.division + "\u2026");
          await GraphData.addListItem(token, state.site.siteId, state.site.routingListId, {
            Title: pl.division, DivisionCode: pl.division,
            Emails: pl.addEmails.join("; "),
            IsActive: true, Priority: 1,
            Notes: "Created by contact import \u2014 add a Teams channel in the route builder.",
          });
          created++;
        }
      }
      contactPlanCache = null;
      byId("contactRows").value = "";
      byId("contactPlan").innerHTML = "";
      setStatus("info", "Contact import done: " + applied + " rule(s) updated, " + created + " created. Reloading rules\u2026");
      await connectRules();
    } catch (e) {
      setStatus("error", "Import failed: " + friendly(e));
      byId("contactApply").disabled = false;
    }
  }

  async function createLists() {
    var siteUrl = byId("siteUrl").value.trim();
    if (!siteUrl) { setStatus("error", "Enter the SharePoint site URL first."); return; }
    byId("createLists").disabled = true;
    try {
      setStatus("work", "Checking the site\u2026");
      var token = await GraphData.getToken();
      var site = await GraphData.resolveSite(token, siteUrl);
      var defs = LrrProvision.listDefinitions();
      var wanted = [
        { name: byId("routingList").value.trim() || "LegislativeRoutingMatrix", def: defs.routing, role: "routing" },
        { name: byId("auditList").value.trim() || "LegislativeAudit", def: defs.audit, role: "audit" },
        { name: byId("trackerList").value.trim() || "BillTracker", def: defs.tracker, role: "tracker" },
      ];
      var report = [];
      for (var i = 0; i < wanted.length; i++) {
        var w = wanted[i];
        var exists = true;
        try { await GraphData.findList(token, site.siteId, w.name); }
        catch (e) { exists = false; }
        if (exists) { report.push(w.name + ": already exists"); continue; }
        setStatus("work", "Creating " + w.name + "\u2026");
        var created = await GraphData.createList(token, site.siteId, w.name, w.def);
        report.push(w.name + ": created");
        if (w.role === "routing" && created && created.id) {
          try { await GraphData.addListItem(token, site.siteId, created.id, LrrProvision.sampleRoutingRule()); }
          catch (e2) { /* sample row is best-effort */ }
        }
      }
      setStatus("info", report.join(" \u00b7 ") + " \u2014 connecting\u2026");
      await connectRules();
    } catch (e) {
      setStatus("error", "List setup failed: " + friendly(e) +
        " \u2014 you need edit rights on the site; see the admin guide to create lists manually.");
    } finally {
      byId("createLists").disabled = false;
    }
  }

  async function connectRules() {
    byId("connectRules").disabled = true;
    try {
      setStatus("work", "Loading routing rules from SharePoint…");
      var token = await GraphData.getToken();
      var site = await GraphData.resolveSite(token, byId("siteUrl").value);
      var routingListId = await GraphData.findList(token, site.siteId, byId("routingList").value.trim());
      var auditListId = await GraphData.findList(token, site.siteId, byId("auditList").value.trim());
      var trackerListId = null;
      var trackerName = byId("trackerList").value.trim();
      if (trackerName) {
        try { trackerListId = await GraphData.findList(token, site.siteId, trackerName); }
        catch (e) { setStatus("error", 'Tracker list "' + trackerName + '" not found — status tracking disabled until it exists.'); }
      }
      var raw = await GraphData.listItems(token, site.siteId, routingListId);
      state.rules = raw.map(function (r) { return LrrRouting.ruleFromSharePoint(r.fields || {}, r.id); })
        .filter(function (r) { return r.divisionCode; });
      state.site = { siteId: site.siteId, routingListId: routingListId, auditListId: auditListId, trackerListId: trackerListId };
      byId("rulesInfo").textContent = state.rules.length + " routing rule(s) loaded from " + site.name +
        " (" + state.rules.filter(function (r) { return r.teamsTagId; }).length + " with Teams tags).";
      var qSel = byId("quickDivision");
      qSel.innerHTML = "";
      state.rules.forEach(function (r) {
        var o = document.createElement("option");
        o.value = r.id; o.textContent = r.divisionCode;
        qSel.appendChild(o);
      });
      var tpSel = byId("testPostRoute");
      tpSel.innerHTML = "";
      state.rules.filter(function (r) { return r.teamsTeamId && r.teamsChannelId; }).forEach(function (r) {
        var o = document.createElement("option");
        o.value = r.id;
        o.textContent = r.divisionCode + " \u2192 " + (r.teamsChannelName || r.teamsChannelId.slice(0, 18) + "\u2026");
        tpSel.appendChild(o);
      });
      if (!tpSel.children.length) {
        var o0 = document.createElement("option");
        o0.value = ""; o0.textContent = "(no rules with a Teams channel yet)";
        tpSel.appendChild(o0);
      }
      if (state.items.length) { LrrRouting.routeAll(state.items, state.rules); refreshStats(); renderItems(); }
      saveRulesCache();
      updateConnBanner();
      updateSetupChecklist();
      renderRoutesSummary();
      if (!quiet) { setStatus("info", "Routing connected."); }
    } catch (e) {
      if (!quiet) { setStatus("error", "Routing connection failed: " + friendly(e)); }
      throw e;
    } finally {
      byId("connectRules").disabled = false;
    }
  }

  async function lookupTags() {
    try {
      setStatus("work", "Fetching tags…");
      var token = await GraphData.getToken();
      var tags = await GraphData.listTeamTags(token, byId("lookupTeamId").value.trim());
      var pre = byId("tagResults");
      pre.hidden = false;
      pre.textContent = tags.length
        ? tags.map(function (t) { return t.displayName + "\n  TeamsTagId: " + t.id; }).join("\n")
        : "No tags found on that team.";
      setStatus("info", tags.length + " tag(s) found — copy the IDs into the routing list.");
    } catch (e) {
      setStatus("error", "Tag lookup failed: " + friendly(e));
    }
  }

  // ---------- morning auto-draft ----------
  // Runs only while the pane is open (pin the pane to keep it open).
  // Automates ONLY the draft: Teams posts and Send remain explicit.

  var autoDraftRan = false; // once per pane session

  function todayKey() {
    var d = new Date();
    return d.getFullYear() + "-" + (d.getMonth() + 1) + "-" + d.getDate();
  }

  function startAutoDraftTimer() {
    setInterval(maybeAutoDraft, 60 * 1000);
    setTimeout(maybeAutoDraft, 5000); // catch-up shortly after open
  }

  async function maybeAutoDraft() {
    try {
      var st = settings();
      if (!(st.autoDaily === true || st.autoDaily === "true") || autoDraftRan) { return; }
      if (st.lastAutoDraftDay === todayKey()) { return; } // already drafted today
      var t = String(st.autoDailyTime || byId("autoDailyTime").value || "07:30");
      var parts = t.split(":");
      var now = new Date();
      var due = new Date(now.getFullYear(), now.getMonth(), now.getDate(), +parts[0] || 7, +parts[1] || 30);
      if (now < due) { return; } // not time yet
      if (!LrrReportGen.extractEmails(st.distList || "").length) { return; } // not configured
      if (!state.rules.length) {
        try { await connectRules(true); } catch (e) { return; } // sign-in needed — wait for a manual open
      }

      autoDraftRan = true;
      setStatus("work", "Morning auto-draft: checking for new bills\u2026");

      // pull the feed (same path as the Filings tab)
      var preset = LrrPresets.presetFor(byId("stateName").value || "Iowa");
      var entries;
      if (preset.feed === "iowa-rss") {
        var res = await fetch("../../feeds/IowaBills.xml?ts=" + Date.now());
        if (!res.ok) { throw new Error("feed mirror unavailable"); }
        entries = LrrFeed.parseFeed(await res.text());
      } else {
        var res2 = await fetch("../../feeds/openstates-" + preset.slug + ".json?ts=" + Date.now());
        if (!res2.ok) { throw new Error("feed mirror unavailable"); }
        entries = LrrFeed.parseOpenStates(await res2.text());
      }
      var terms = byId("watchTerms").value.split(",").map(function (x) { return x.trim(); }).filter(Boolean);
      var days = Number(byId("watchDays").value) || 3;
      var hits = LrrFeed.watchFilter(entries, terms, days);
      var tracked = trackedList();
      LrrFeed.watchFilter(entries, [], days).forEach(function (e) {
        if (hits.indexOf(e) !== -1) { return; }
        if (LrrChapters.matchTracked(LrrChapters.extractChapters(e.description || ""), tracked).length) { hits.push(e); }
      });
      hits = LrrReportGen.filterNew(hits, reportedBills());
      if (!hits.length) {
        saveSettings({ lastAutoDraftDay: todayKey() });
        setStatus("info", "Morning check: no new bills since your last report. \u2615");
        return;
      }

      // build items with auto-routing (chapters + keywords)
      state.reportKey = "daily-" + todayKey();
      state.subject = LrrReportGen.subjectFor(new Date());
      var items = [];
      hits.forEach(function (e, i) {
        var item = LrrFeed.toLegislativeItem(e, state.reportKey, i);
        item.codeChapters = LrrChapters.extractChapters(item.brief || "");
        item.trackedChapters = LrrChapters.matchTracked(item.codeChapters, tracked);
        var codes = [];
        LrrChapters.suggestRules(item.codeChapters, state.rules).forEach(function (sg) {
          if (codes.indexOf(sg.rule.divisionCode) === -1) { codes.push(sg.rule.divisionCode); }
        });
        LrrRouting.suggestByKeywords(item.brief || "", state.rules).forEach(function (sg) {
          if (codes.indexOf(sg.rule.divisionCode) === -1) { codes.push(sg.rule.divisionCode); }
        });
        if (codes.length) {
          item.distributedTo = codes;
          item.commentRequestedFrom = codes.slice();
          item.parserWarnings = ["Divisions auto-suggested \u2014 verify before publishing."];
        }
        LrrRouting.routeItem(item, state.rules);
        items.push(item);
      });
      state.items = items;
      state.results = {};
      refreshStats(); renderItems();

      // draft only — never posts, never sends
      var token = await GraphData.getToken();
      var recips = LrrReportGen.extractEmails(st.distList || "");
      var rep = LrrReportGen.buildDailyReport(items, {
        sessionName: st.sessionName || "",
        commentWindow: byId("commentWindow").value,
        preparedBy: ((Office.context.mailbox.userProfile || {}).displayName) || "",
      });
      await GraphData.createDraftMessage(token, recips, rep.subject, rep.html);
      markReported(items);
      saveSettings({ lastAutoDraftDay: todayKey() });
      var unrouted = items.filter(function (it) { return it.routingStatus !== "matched"; }).length;
      setStatus("info", "\u2600\ufe0f Today's Daily Bill Report is drafted \u2014 " + rep.count + " bill(s), " +
        recips.length + " recipients, waiting in your Drafts. " +
        (unrouted ? unrouted + " bill(s) still need a division (Review tab)." : "All bills routed.") +
        " Review here, Publish to Teams, then send the draft.");
      if (byId("optDailyReport")) { byId("optDailyReport").checked = false; } // already drafted
      show("review");
    } catch (e) {
      autoDraftRan = false; // let the next tick retry
    }
  }

  // ---------- setup health check ----------

  var checkLog = [];

  function logCheck(ok, label, detail) {
    checkLog.push({ ok: ok, label: label, detail: detail || "" });
    var p = document.createElement("p");
    p.className = ok === true ? "" : (ok === "warn" ? "warn" : "err");
    p.textContent = (ok === true ? "\u2713 " : ok === "warn" ? "\u26a0 " : "\u2717 ") + label + (detail ? " \u2014 " + detail : "");
    byId("checkResults").appendChild(p);
  }

  async function runChecks() {
    byId("runChecks").disabled = true;
    byId("checkResults").innerHTML = "";
    byId("copyChecks").hidden = true;
    checkLog = [];
    try {
      // 1. sign-in
      var token;
      try {
        token = await GraphData.getToken();
        logCheck(true, "Sign-in", "token acquired for this mailbox");
      } catch (e) {
        logCheck(false, "Sign-in", (e && e.message) || String(e));
        setStatus("error", "Sign-in failed \u2014 fix that first; nothing else can work.");
        return;
      }

      // 2. parser self-test (pure, proves the code loaded intact)
      try {
        var sample = LrrParser.parseReport("HF935\nMVD\nMVD\ntest brief. Successor to HSB171.", {});
        var okParse = sample.items.length === 1 && sample.items[0].billNumber === "HF935" &&
          sample.items[0].referencedBills[0] === "HSB171";
        logCheck(okParse ? true : false, "Parser", okParse ? "sample report parses correctly" : "unexpected parse result");
      } catch (e) { logCheck(false, "Parser", (e && e.message) || String(e)); }

      // 3. site + lists
      var siteOk = false;
      try {
        var site = await GraphData.resolveSite(token, byId("siteUrl").value);
        logCheck(true, "SharePoint site", site.name);
        siteOk = true;
        var names = [byId("routingList").value.trim() || "LegislativeRoutingMatrix",
                     byId("auditList").value.trim() || "LegislativeAudit"];
        var trackerName = byId("trackerList").value.trim();
        if (trackerName) { names.push(trackerName); }
        for (var li = 0; li < names.length; li++) {
          try {
            await GraphData.findList(token, site.siteId, names[li]);
            logCheck(true, 'List "' + names[li] + '"', "found");
          } catch (e) {
            logCheck(false, 'List "' + names[li] + '"', "missing \u2014 use \u2461 Create missing lists");
          }
        }
      } catch (e) {
        logCheck(false, "SharePoint site", (e && e.message) || String(e));
      }

      // 4. rules + pure validation
      if (!state.rules.length && siteOk) {
        try { await connectRules(); } catch (e) { /* reported below */ }
      }
      if (state.rules.length) {
        logCheck(true, "Routing rules", state.rules.length + " loaded");
        LrrRouting.validateRules(state.rules).forEach(function (v) {
          logCheck(v.level === "error" ? false : "warn", "Rule check" + (v.division ? " (" + v.division + ")" : ""), v.message);
        });
      } else {
        logCheck(false, "Routing rules", "none loaded \u2014 connect the site and add routes");
      }

      // 5. live Teams validation: channels + tags per unique team (cap 5)
      var teams = {};
      state.rules.forEach(function (r) { if (r.teamsTeamId) { teams[r.teamsTeamId] = teams[r.teamsTeamId] || []; teams[r.teamsTeamId].push(r); } });
      var teamIds = Object.keys(teams).slice(0, 5);
      for (var ti = 0; ti < teamIds.length; ti++) {
        var tid = teamIds[ti];
        try {
          var channels = await GraphData.listChannels(token, tid);
          var chanIds = channels.map(function (c) { return c.id; });
          var tags = [];
          try { tags = await GraphData.listTeamTags(token, tid); } catch (eT) { /* tags optional */ }
          var tagIds = tags.map(function (t) { return t.id; });
          teams[tid].forEach(function (r) {
            var chanOk = chanIds.indexOf(r.teamsChannelId) !== -1;
            logCheck(chanOk ? true : false, "Teams channel (" + r.divisionCode + ")",
              chanOk ? (r.teamsChannelName || "reachable") : "channel ID not found in that team \u2014 re-pick it in the route builder");
            if (r.teamsTagId) {
              var tagOk = tagIds.indexOf(r.teamsTagId) !== -1;
              logCheck(tagOk ? true : "warn", "Teams tag (" + r.divisionCode + ")",
                tagOk ? (r.teamsTagName || "found") : "tag ID not found \u2014 mention would be dropped; re-pick the tag");
            }
          });
        } catch (e) {
          logCheck(false, "Teams team", "can't reach team for " + teams[tid].map(function (r) { return r.divisionCode; }).join("/") +
            " \u2014 are you a member? (" + friendly(e).slice(0, 120) + ")");
        }
      }

      // 6. feed mirror
      try {
        var preset = LrrPresets.presetFor(byId("stateName").value || "Iowa");
        var feedUrl = preset.feed === "iowa-rss" ? "../../feeds/IowaBills.xml" : "../../feeds/openstates-" + preset.slug + ".json";
        var fr = await fetch(feedUrl + "?ts=" + Date.now());
        if (fr.ok) {
          var body = await fr.text();
          var n = preset.feed === "iowa-rss" ? (body.match(/<item>/g) || []).length : (LrrFeed.parseOpenStates(body) || []).length;
          logCheck(true, "New-filings feed", n + " bills in the mirror");
        } else { logCheck("warn", "New-filings feed", "mirror not available (" + fr.status + ") \u2014 Filings tab won't load, everything else works"); }
      } catch (e) { logCheck("warn", "New-filings feed", "unreachable \u2014 Filings tab won't load, everything else works"); }

      // 7. tracker write test (add + delete a row)
      if (state.site && state.site.trackerListId) {
        try {
          var row = await GraphData.addListItem(token, state.site.siteId, state.site.trackerListId,
            { Title: "SETUP TEST \u2014 safe to ignore", Division: "TEST", Status: "Pending review", ReportKey: "setup-check" });
          await GraphData.deleteListItem(token, state.site.siteId, state.site.trackerListId, row.id);
          logCheck(true, "Tracker write", "test row created and removed");
        } catch (e) { logCheck(false, "Tracker write", (e && e.message || e).slice(0, 140)); }
      }

      // 8. optional live: test post
      if (byId("optTestPost").checked) {
        var rid = byId("testPostRoute").value;
        var rule = state.rules.filter(function (r) { return r.id === rid; })[0];
        if (!rule) { logCheck("warn", "Test post", "pick a route first"); }
        else {
          try {
            var mentionRules = byId("optTestMention").checked ? [rule] : [{}];
            var payload = LrrTeams.buildChannelMessage({
              billNumber: "TEST", title: "",
              brief: "\ud83e\uddea Legislative Report Router setup test \u2014 safe to ignore. Posted by " +
                ((Office.context.mailbox.userProfile || {}).displayName || "the coordinator") + " from the add-in's setup check.",
              distributedTo: [rule.divisionCode], commentRequestedFrom: [], sourceLinks: [],
            }, mentionRules, { template: { commentWindow: "" } });
            var msg = await GraphData.postChannelMessage(token, rule.teamsTeamId, rule.teamsChannelId, payload);
            logCheck(true, "Test post", "posted to " + (rule.teamsChannelName || "the channel") +
              (byId("optTestMention").checked ? " with the tag mention" : "") + " (message " + String(msg.id).slice(0, 12) + "\u2026)");
            await writeAudit(token, { Title: "TEST", ReportKey: "setup-check", IdempotencyKey: "setup-check-" + msg.id,
              TeamId: rule.teamsTeamId, ChannelId: rule.teamsChannelId, TeamsMessageId: msg.id || "",
              Status: "published", Divisions: rule.divisionCode,
              PublishedBy: ((Office.context.mailbox.userProfile || {}).emailAddress) || "", SourceSubject: "Setup check" });
            logCheck(true, "Audit write", "test post recorded in the audit list");
          } catch (e) { logCheck(false, "Test post", (e && e.message || e).slice(0, 200)); }
        }
      }

      // 9. optional live: test email to self
      if (byId("optTestEmail").checked) {
        try {
          var me = ((Office.context.mailbox.userProfile || {}).emailAddress) || "";
          await GraphData.sendMail(token, [me], "Legislative Report Router setup test",
            "<p>\ud83e\uddea This is the add-in's setup-check email. If you can read this, division emails will send. Safe to delete.</p>");
          logCheck(true, "Test email", "sent to " + me);
        } catch (e) { logCheck(false, "Test email", (e && e.message || e).slice(0, 200)); }
      }

      var fails = checkLog.filter(function (c) { return c.ok === false; }).length;
      var warns = checkLog.filter(function (c) { return c.ok === "warn"; }).length;
      if (!fails) { saveSettings({ lastCheckOk: new Date().toISOString() }); updateSetupChecklist(); }
      setStatus(fails ? "error" : "info",
        fails ? fails + " check(s) failed, " + warns + " warning(s) \u2014 fix the \u2717 items above."
              : (warns ? "Setup works \u2014 " + warns + " warning(s) worth a look." : "All checks passed \u2014 you're ready to publish. \ud83c\udf89"));
      byId("copyChecks").hidden = false;
    } finally {
      byId("runChecks").disabled = false;
    }
  }

  function copyChecks() {
    var text = checkLog.map(function (c) {
      return (c.ok === true ? "PASS" : c.ok === "warn" ? "WARN" : "FAIL") + "  " + c.label + (c.detail ? " - " + c.detail : "");
    }).join("\n");
    navigator.clipboard.writeText("Legislative Report Router setup check\n" + new Date().toLocaleString() + "\n\n" + text)
      .then(function () { setStatus("info", "Results copied \u2014 paste into an email if you need help."); })
      .catch(function () { setStatus("error", "Copy failed \u2014 select the results manually."); });
  }

  // ---------- draft save/load ----------

  function saveDraft() {
    var draft = { subject: state.subject, reportKey: state.reportKey, items: state.items, savedAt: new Date().toISOString() };
    Office.context.roamingSettings.set(DRAFT_KEY, JSON.stringify(draft));
    Office.context.roamingSettings.saveAsync(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        state.lastSaved = new Date().toLocaleTimeString();
        refreshStats();
        setStatus("info", "Draft distribution saved.");
      } else { setStatus("error", "Couldn't save the draft."); }
    });
  }

  function loadDraft() {
    try {
      var d = JSON.parse(Office.context.roamingSettings.get(DRAFT_KEY) || "null");
      if (!d) { setStatus("info", "No saved draft."); return; }
      state.items = d.items || [];
      state.reportKey = d.reportKey || "";
      annotateChapters();
      if (state.rules.length) { LrrRouting.routeAll(state.items, state.rules); }
      refreshStats(); renderItems();
      setStatus("info", "Draft from " + new Date(d.savedAt).toLocaleString() + " loaded (" + state.items.length + " bills).");
      show("review");
    } catch (e) { setStatus("error", "Couldn't load the draft."); }
  }

  // ---------- preview ----------

  /** item → [{teamId, channelId, rules[]}] grouped so one channel gets ONE post. */
  function channelGroups(it) {
    var groups = {};
    var order = [];
    (it.routes || []).forEach(function (r) {
      if (!r.teamsTeamId || !r.teamsChannelId) { return; }
      var key = r.teamsTeamId + "/" + r.teamsChannelId;
      if (!groups[key]) { groups[key] = { teamId: r.teamsTeamId, channelId: r.teamsChannelId, rules: [] }; order.push(key); }
      groups[key].rules.push(r);
    });
    return order.map(function (k) { return groups[k]; });
  }

  function included() {
    return state.items.filter(function (i) { return i.routingStatus !== "excluded"; });
  }

  function renderPreview() {
    var host = byId("previewList");
    host.innerHTML = "";
    var opts = { template: { commentWindow: byId("commentWindow").value } };
    included().forEach(function (it) {
      channelGroups(it).forEach(function (g) {
        var payload = LrrTeams.buildChannelMessage(it, g.rules, opts);
        var box = document.createElement("div");
        box.className = "preview-post";
        var meta = document.createElement("p");
        meta.className = "hint";
        meta.textContent = it.billNumber + " → " + g.rules.map(function (r) { return r.teamsChannelName || r.divisionCode; }).join(", ");
        var body = document.createElement("div");
        body.className = "preview-body";
        body.innerHTML = payload.body.content; // template-built, user content escaped
        box.appendChild(meta); box.appendChild(body);
        host.appendChild(box);
      });
    });
    LrrRouting.groupByRule(included()).forEach(function (g) {
      if (!(g.rule.emails || []).length) { return; }
      var mail = LrrTeams.buildDivisionEmail(g.rule, g.items, { template: { commentWindow: byId("commentWindow").value } });
      var box = document.createElement("div");
      box.className = "preview-post email";
      box.innerHTML = '<p class="hint">✉ ' + esc(mail.subject) + " → " + esc(mail.to.join(", ")) + "</p>" +
        '<div class="preview-body">' + mail.html + "</div>";
      host.appendChild(box);
    });
    if (!host.children.length) {
      host.innerHTML = '<p class="hint">Nothing to preview — parse a report and connect routing rules first.</p>';
    }
  }

  // ---------- publish ----------

  function renderPublishSummary() {
    var items = included();
    var posts = 0, tags = {}, recipients = {};
    items.forEach(function (it) {
      channelGroups(it).forEach(function (g) {
        posts++;
        g.rules.forEach(function (r) { if (r.teamsTagId) { tags[r.teamsTagId] = 1; } });
      });
    });
    LrrRouting.groupByRule(items).forEach(function (g) {
      (g.rule.emails || []).forEach(function (e) { recipients[e] = 1; });
    });
    byId("publishSummary").innerHTML =
      "<strong>" + posts + "</strong> Teams post(s) will be created.<br>" +
      "<strong>" + Object.keys(tags).length + "</strong> division tag(s) will be mentioned.<br>" +
      "<strong>" + Object.keys(recipients).length + "</strong> recipient(s) will receive targeted email." +
      (byId("optOriginal").checked ? "<br>The original Outlook message will be sent." : "") +
      (byId("optDailyReport").checked
        ? "<br>The <strong>Daily Bill Report draft</strong> will be created for <strong>" +
          LrrReportGen.extractEmails((settings().distList || "")).length + "</strong> recipients (review in Drafts, then send)."
        : "");
  }

  function logLine(text, kind) {
    var p = document.createElement("p");
    p.className = kind || "";
    p.textContent = text;
    byId("publishLog").appendChild(p);
  }

  async function alreadyPublished(token, key) {
    if (!state.site) { return false; }
    if (state.results[key] && state.results[key].status === "published") { return true; }
    // durable check against the audit list (last 400 records)
    if (!state._auditCache) {
      var items = await GraphData.listItems(token, state.site.siteId, state.site.auditListId, 400);
      state._auditCache = {};
      items.forEach(function (i) {
        var f = i.fields || {};
        if (f.IdempotencyKey && f.Status === "published") { state._auditCache[f.IdempotencyKey] = true; }
      });
    }
    return !!state._auditCache[key];
  }

  async function writeAudit(token, fields) {
    if (!state.site) { return; }
    try { await GraphData.addListItem(token, state.site.siteId, state.site.auditListId, fields); }
    catch (e) { logLine("Audit write failed (operation itself succeeded): " + e.message, "err"); }
  }

  async function publish(retryOnly) {
    if (!state.site) { setStatus("error", "Connect the SharePoint routing/audit lists in Settings first."); return; }
    var doTeams = byId("optTeams").checked;
    var doEmail = byId("optEmail").checked;
    var doOriginal = byId("optOriginal").checked;
    byId("publishGo").disabled = true;
    byId("publishLog").innerHTML = "";
    var failures = 0;

    try {
      var token = await GraphData.getToken();
      var me = Office.context.mailbox.userProfile;
      var opts = { template: { commentWindow: byId("commentWindow").value } };

      if (doTeams) {
        for (var i = 0; i < included().length; i++) {
          var it = included()[i];
          var groups = channelGroups(it);
          if (!groups.length) { logLine(it.billNumber + ": no Teams route — skipped.", "warn"); continue; }
          for (var gI = 0; gI < groups.length; gI++) {
            var g = groups[gI];
            var key = LrrTeams.idempotencyKey(state.reportKey, it.billNumber, g.channelId);
            if (retryOnly && state.results[key] && state.results[key].status === "published") { continue; }
            if (await alreadyPublished(token, key)) {
              logLine(it.billNumber + " → " + g.channelId.slice(0, 12) + "…: already published — skipped (idempotent).", "warn");
              continue;
            }
            try {
              setStatus("work", "Posting " + it.billNumber + "…");
              var payload = LrrTeams.buildChannelMessage(it, g.rules, opts);
              var msg = await GraphData.postChannelMessage(token, g.teamId, g.channelId, payload);
              state.results[key] = { status: "published", messageId: msg.id };
              if (state._auditCache) { state._auditCache[key] = true; }
              logLine("✓ " + it.billNumber + " posted (" + g.rules.map(function (r) { return r.divisionCode; }).join(", ") + ")");
              await writeAudit(token, {
                Title: it.billNumber, ReportKey: state.reportKey, IdempotencyKey: key,
                TeamId: g.teamId, ChannelId: g.channelId, TeamsMessageId: msg.id || "",
                Status: "published", Divisions: (it.distributedTo || []).join("; "),
                PublishedBy: (me && me.emailAddress) || "", SourceSubject: state.subject,
              });
              // Bill tracker: one row per bill × division so the Teams Lists
              // tab shows per-division status and "who's still waiting".
              if (state.site.trackerListId) {
                var due = LrrFeed.addBusinessDays(new Date(), 2);
                for (var tI = 0; tI < g.rules.length; tI++) {
                  try {
                    await GraphData.addListItem(token, state.site.siteId, state.site.trackerListId, {
                      Title: it.billNumber,
                      Division: g.rules[tI].divisionCode,
                      Status: "Pending review",
                      DueDate: due.toISOString().slice(0, 10),
                      BillLink: ((it.sourceLinks || [])[0] || {}).href || "",
                      Brief: String(it.title || it.brief || "").slice(0, 250),
                      ReportKey: state.reportKey,
                    });
                  } catch (e2) { logLine("Tracker row for " + g.rules[tI].divisionCode + " failed: " + e2.message, "warn"); }
                }
              }
            } catch (e) {
              failures++;
              state.results[key] = { status: "failed", error: e.message };
              logLine("✗ " + it.billNumber + ": " + e.message, "err");
              await writeAudit(token, {
                Title: it.billNumber, ReportKey: state.reportKey, IdempotencyKey: key,
                TeamId: g.teamId, ChannelId: g.channelId, Status: "failed",
                Error: String(e.message).slice(0, 500),
                PublishedBy: (me && me.emailAddress) || "", SourceSubject: state.subject,
              });
            }
          }
        }
      }

      if (doEmail && !retryOnly) {
        var mailGroups = LrrRouting.groupByRule(included());
        for (var mI = 0; mI < mailGroups.length; mI++) {
          var mg = mailGroups[mI];
          if (!(mg.rule.emails || []).length) { continue; }
          var mkey = LrrTeams.idempotencyKey(state.reportKey, "EMAIL", mg.rule.id);
          if (await alreadyPublished(token, mkey)) { logLine("✉ " + mg.rule.divisionCode + ": email already sent — skipped.", "warn"); continue; }
          try {
            var mail = LrrTeams.buildDivisionEmail(mg.rule, mg.items, opts);
            await GraphData.sendMail(token, mail.to, mail.subject, mail.html);
            state.results[mkey] = { status: "published" };
            if (state._auditCache) { state._auditCache[mkey] = true; }
            logLine("✓ Email to " + mg.rule.divisionCode + " (" + mail.to.join(", ") + ")");
            await writeAudit(token, {
              Title: "EMAIL " + mg.rule.divisionCode, ReportKey: state.reportKey, IdempotencyKey: mkey,
              Status: "published", EmailRecipients: mail.to.join("; "),
              PublishedBy: (me && me.emailAddress) || "", SourceSubject: state.subject,
            });
          } catch (e) {
            failures++;
            logLine("✗ Email to " + mg.rule.divisionCode + ": " + e.message, "err");
          }
        }
      }

      if (byId("optDailyReport").checked && !retryOnly) {
        try {
          var st2 = settings();
          var recips = LrrReportGen.extractEmails(st2.distList || "");
          if (!recips.length) {
            logLine("\u26a0 Daily Bill Report skipped: no distribution list set (Setup \u2192 Daily Bill Report generation).", "warn");
          } else {
            var rep = LrrReportGen.buildDailyReport(included(), {
              sessionName: st2.sessionName || "",
              commentWindow: byId("commentWindow").value,
              preparedBy: (me && me.displayName) || "",
            });
            await GraphData.createDraftMessage(token, recips, rep.subject, rep.html);
            markReported(included());
            logLine("\u2713 Daily Bill Report draft created — " + rep.count + " bills, " + recips.length +
              " recipients. Review it in your Drafts folder and send.");
          }
        } catch (e) {
          failures++;
          logLine("\u2717 Daily Bill Report draft failed: " + e.message, "err");
        }
      }

      if (doOriginal && !retryOnly) {
        try {
          var msgId = await restMessageId();
          await GraphData.sendDraft(token, msgId);
          logLine("✓ Original report email sent.");
        } catch (e) {
          failures++;
          logLine("✗ Sending the original failed: " + e.message + " (if this is a received message, it was already sent).", "err");
        }
      }

      byId("retryFailed").hidden = failures === 0;
      setStatus(failures ? "error" : "info",
        failures ? failures + " operation(s) failed — successful posts are preserved; use \"Retry failed only\"."
                 : "Publication complete. Full record is in the audit list.");
    } catch (e) {
      setStatus("error", "Publish failed: " + friendly(e));
    } finally {
      byId("publishGo").disabled = !byId("confirmBox").checked;
    }
  }

  // ---------- audit ----------

  async function refreshAudit() {
    if (!state.site) { setStatus("error", "Connect SharePoint in Settings first."); return; }
    try {
      setStatus("work", "Loading audit history…");
      var token = await GraphData.getToken();
      var items = await GraphData.listItems(token, state.site.siteId, state.site.auditListId, 200);
      var host = byId("auditListView");
      host.innerHTML = "";
      items.slice().reverse().slice(0, 100).forEach(function (i) {
        var f = i.fields || {};
        var p = document.createElement("p");
        p.className = "audit-row " + (f.Status === "failed" ? "err" : "");
        p.textContent = [f.Created ? new Date(f.Created).toLocaleString() : "", f.Title, f.Status,
          f.Divisions || f.EmailRecipients || "", f.PublishedBy, f.Error || ""].filter(Boolean).join(" · ");
        host.appendChild(p);
      });
      setStatus("info", "Audit history loaded.");
    } catch (e) {
      setStatus("error", "Audit load failed: " + friendly(e));
    }
  }
})();
