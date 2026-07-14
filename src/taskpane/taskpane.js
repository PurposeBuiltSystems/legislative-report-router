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

  // ---------- screens ----------

  function show(screen) {
    ["overview", "review", "preview", "publish", "audit"].forEach(function (s) {
      byId("screen-" + s).hidden = s !== screen;
    });
    document.querySelectorAll(".tab").forEach(function (t) {
      t.classList.toggle("active", t.getAttribute("data-screen") === screen);
    });
    if (screen === "preview") { renderPreview(); }
    if (screen === "publish") { renderPublishSummary(); }
  }

  Office.onReady(function () {
    var s = settings();
    ["cloud", "siteUrl", "routingList", "auditList", "commentWindow"].forEach(function (k) {
      if (s[k]) { byId(k).value = s[k]; }
    });
    if (s.cloud) { GraphData.setCloud(s.cloud); }
    if (!s.siteUrl) { byId("settings").setAttribute("open", "open"); }

    document.querySelectorAll(".tab").forEach(function (t) {
      t.addEventListener("click", function () { show(t.getAttribute("data-screen")); });
    });
    byId("parse").addEventListener("click", parseReport);
    byId("saveDraft").addEventListener("click", saveDraft);
    byId("loadDraft").addEventListener("click", loadDraft);
    byId("connectRules").addEventListener("click", connectRules);
    byId("lookupTags").addEventListener("click", lookupTags);
    byId("bulkApply").addEventListener("click", bulkApply);
    byId("confirmBox").addEventListener("change", function () {
      byId("publishGo").disabled = !byId("confirmBox").checked;
    });
    byId("publishGo").addEventListener("click", function () { publish(false); });
    byId("retryFailed").addEventListener("click", function () { publish(true); });
    byId("refreshAudit").addEventListener("click", refreshAudit);
    ["cloud", "siteUrl", "routingList", "auditList", "commentWindow"].forEach(function (k) {
      byId(k).addEventListener("change", function () {
        var p = {}; p[k] = byId(k).value; saveSettings(p);
        if (k === "cloud") { GraphData.setCloud(byId(k).value); state.site = null; }
        if (k === "siteUrl" || k === "routingList" || k === "auditList") { state.site = null; }
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
      var res = LrrParser.parseReport(text, { knownDivisions: known, reportId: state.reportKey });
      state.items = res.items;
      state.results = {};

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
      setStatus("error", "Parse failed: " + ((e && e.message) || e));
    } finally {
      byId("parse").disabled = false;
    }
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

  // ---------- rules ----------

  async function connectRules() {
    byId("connectRules").disabled = true;
    try {
      setStatus("work", "Loading routing rules from SharePoint…");
      var token = await GraphData.getToken();
      var site = await GraphData.resolveSite(token, byId("siteUrl").value);
      var routingListId = await GraphData.findList(token, site.siteId, byId("routingList").value.trim());
      var auditListId = await GraphData.findList(token, site.siteId, byId("auditList").value.trim());
      var raw = await GraphData.listItems(token, site.siteId, routingListId);
      state.rules = raw.map(function (r) { return LrrRouting.ruleFromSharePoint(r.fields || {}, r.id); })
        .filter(function (r) { return r.divisionCode; });
      state.site = { siteId: site.siteId, routingListId: routingListId, auditListId: auditListId };
      byId("rulesInfo").textContent = state.rules.length + " routing rule(s) loaded from " + site.name +
        " (" + state.rules.filter(function (r) { return r.teamsTagId; }).length + " with Teams tags).";
      if (state.items.length) { LrrRouting.routeAll(state.items, state.rules); refreshStats(); renderItems(); }
      setStatus("info", "Routing connected.");
    } catch (e) {
      setStatus("error", "Routing connection failed: " + ((e && e.message) || e));
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
      setStatus("error", "Tag lookup failed: " + ((e && e.message) || e));
    }
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
      (byId("optOriginal").checked ? "<br>The original Outlook message will be sent." : "");
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
      setStatus("error", "Publish failed: " + ((e && e.message) || e));
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
      setStatus("error", "Audit load failed: " + ((e && e.message) || e));
    }
  }
})();
