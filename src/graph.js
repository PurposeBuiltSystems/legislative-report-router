/*
 * Legislative Report Router — Microsoft Graph data layer.
 *
 * AUTH: Nested App Authentication (NAA) via MSAL — delegated only, no
 * backend. Cloud endpoints are configurable (never hardcoded commercial):
 * Commercial, GCC (commercial endpoints, gov tenancy), GCC High, DoD.
 *
 * Delegated scopes (see docs/permissions.md):
 *   required: Mail.ReadWrite, Mail.Send, ChannelMessage.Send, Sites.ReadWrite.All
 *   optional: TeamworkTag.Read, Team.ReadBasic.All, Channel.ReadBasic.All
 */
/* global msal */
(function (root) {
  "use strict";

  var CLIENT_ID = "0860a653-ddbd-4455-8bff-affda2a8879f"; // "Legislative Report Router" Entra app

  var CLOUDS = {
    commercial: { graph: "https://graph.microsoft.com", authority: "https://login.microsoftonline.com/common" },
    gcc:        { graph: "https://graph.microsoft.com", authority: "https://login.microsoftonline.com/common" },
    gcchigh:    { graph: "https://graph.microsoft.us",  authority: "https://login.microsoftonline.us/common" },
    dod:        { graph: "https://dod-graph.microsoft.us", authority: "https://login.microsoftonline.us/common" },
  };

  // Must stay in lockstep with the Entra app registration: a scope requested
  // here but not registered makes MSAL prompt for consent that a non-admin
  // cannot grant, which blocks sign-in entirely. Team/Channel.ReadBasic.All
  // back the setup wizard's Team and Channel pickers (/me/joinedTeams,
  // /teams/{id}/channels).
  var SCOPES = ["Mail.ReadWrite", "Mail.Send", "ChannelMessage.Send", "ChannelMessage.Read.All",
    "Sites.ReadWrite.All", "TeamworkTag.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All"];
  // Requested only when the user actually creates a tag (see getTagWriteToken).
  var TAG_WRITE_SCOPE = "TeamworkTag.ReadWrite";
  // Sites.ReadWrite.All covers list ITEMS ("edit or delete documents and list
  // items"); creating a LIST needs Sites.Manage.All ("create or delete
  // document libraries and lists"). Requested only for provisioning, so a
  // tenant that never grants it still gets a fully working add-in.
  var LIST_CREATE_SCOPE = "Sites.Manage.All";

  var cloudKey = "commercial";
  var pcaPromise = null;

  function setCloud(key) {
    if (CLOUDS[key] && key !== cloudKey) { cloudKey = key; pcaPromise = null; }
  }

  function graphBase() { return CLOUDS[cloudKey].graph + "/v1.0"; }

  function getPca() {
    if (!pcaPromise) {
      pcaPromise = msal.createNestablePublicClientApplication({
        auth: { clientId: CLIENT_ID, authority: CLOUDS[cloudKey].authority },
      });
    }
    return pcaPromise;
  }

  /**
   * Sign-in must never hang the pane. If the brokered flow can't complete —
   * a popup opened behind Outlook, was blocked, or the broker never answers —
   * an un-timed await leaves a spinner running forever with nothing to act
   * on. Fail loudly instead, with the two things that actually fix it.
   */
  function withTimeout(promise, ms, message) {
    var timer;
    return Promise.race([
      promise.then(function (v) { clearTimeout(timer); return v; },
                   function (e) { clearTimeout(timer); throw e; }),
      new Promise(function (_, reject) {
        timer = setTimeout(function () { reject(new Error(message)); }, ms);
      }),
    ]);
  }

  async function getToken() {
    var pca = await withTimeout(getPca(), 20000,
      "Sign-in didn't start. Fully quit Outlook (Cmd+Q) and reopen, then try again.");
    try {
      return (await withTimeout(pca.acquireTokenSilent({ scopes: SCOPES }), 20000, "silent timeout")).accessToken;
    } catch (e) {
      var interactive = await withTimeout(
        pca.acquireTokenPopup({ scopes: SCOPES }), 120000,
        "Sign-in didn't finish. A Microsoft sign-in window may have opened behind Outlook — " +
        "check for it (or Mission Control), finish signing in, and click again. If no window " +
        "appeared at all, fully quit Outlook (Cmd+Q), reopen, and retry.");
      return interactive.accessToken;
    }
  }

  /**
   * Token that additionally carries TeamworkTag.ReadWrite, requested only
   * when the user clicks "Create the tag" — so a tenant that hasn't granted
   * it still gets a fully working add-in everywhere else.
   */
  async function getTagWriteToken() {
    var pca = await getPca();
    var scopes = SCOPES.concat([TAG_WRITE_SCOPE]);
    try {
      var silent = await pca.acquireTokenSilent({ scopes: scopes });
      return silent.accessToken;
    } catch (e) {
      var interactive = await pca.acquireTokenPopup({ scopes: scopes });
      return interactive.accessToken;
    }
  }

  /** Token that additionally carries Sites.Manage.All, for creating lists. */
  async function getListCreateToken() {
    var pca = await getPca();
    var scopes = SCOPES.concat([LIST_CREATE_SCOPE]);
    try {
      return (await withTimeout(pca.acquireTokenSilent({ scopes: scopes }), 20000, "silent timeout")).accessToken;
    } catch (e) {
      var interactive = await withTimeout(pca.acquireTokenPopup({ scopes: scopes }), 120000,
        "Sign-in for list creation didn't finish \u2014 look for a Microsoft window behind Outlook.");
      return interactive.accessToken;
    }
  }

  async function graphJson(token, method, path, body) {
    var res = await fetch(graphBase() + path, {
      method: method,
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (res.status === 429 || res.status === 503) {
      // one respectful retry on throttle
      var wait = Number(res.headers.get("Retry-After") || 3) * 1000;
      await new Promise(function (r) { setTimeout(r, Math.min(wait, 15000)); });
      res = await fetch(graphBase() + path, {
        method: method,
        headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
        body: body ? JSON.stringify(body) : undefined,
      });
    }
    if (!res.ok) { throw new Error("Graph " + method + " " + path.split("?")[0] + " -> " + res.status + " " + (await res.text()).slice(0, 400)); }
    return res.status === 204 ? null : res.json();
  }

  // ---------- SharePoint (routing matrix + audit) ----------

  /**
   * "https://x.sharepoint.com/sites/Y" → Graph siteId. Tolerates what
   * people actually paste: URLs copied from a page inside the site
   * (/SitePages/Home.aspx, /Lists/..., /_layouts/...), query strings,
   * and trailing slashes. Access is evaluated as the signed-in user, so
   * org-only / private sites work exactly like public ones.
   */
  async function resolveSite(token, siteUrl) {
    var u = new URL(String(siteUrl).trim());
    var path = u.pathname
      .replace(/\/(SitePages|Lists|Shared%20Documents|Shared Documents|_layouts|Forms)\/.*$/i, "")
      .replace(/\/+$/, "");
    if (!path) { path = "/"; } // tenant root site
    var site = await graphJson(token, "GET", "/sites/" + u.hostname + ":" + path + "?$select=id,displayName");
    return { siteId: site.id, name: site.displayName };
  }

  async function findList(token, siteId, listName) {
    var res = await graphJson(token, "GET", "/sites/" + siteId + "/lists?$select=id,name,displayName&$top=200");
    var hit = (res.value || []).find(function (l) {
      return l.displayName === listName || l.name === listName;
    });
    if (!hit) { throw new Error('List "' + listName + '" not found on the site.'); }
    return hit.id;
  }

  /** All items of a list with fields (paged). */
  async function listItems(token, siteId, listId, top) {
    var out = [];
    var url = "/sites/" + siteId + "/lists/" + listId + "/items?$expand=fields&$top=" + (top || 200);
    var guard = 0;
    while (url && guard++ < 20) {
      var page = await graphJson(token, "GET", url);
      out = out.concat(page.value || []);
      url = page["@odata.nextLink"] ? page["@odata.nextLink"].replace(graphBase(), "") : null;
    }
    return out;
  }

  /** Add one column to an existing list; 409/"already exists" = fine. */
  async function addListColumn(token, siteId, listId, columnDef) {
    try {
      await graphJson(token, "POST", "/sites/" + siteId + "/lists/" + listId + "/columns", columnDef);
      return true;
    } catch (e) {
      if (/already ?exists|409|nameAlreadyExists/i.test(String(e && e.message))) { return false; }
      throw e;
    }
  }

  async function createList(token, siteId, displayName, definition) {
    return graphJson(token, "POST", "/sites/" + siteId + "/lists", {
      displayName: displayName,
      description: definition.description || "",
      columns: definition.columns,
      list: { template: "genericList" },
    });
  }

  async function addListItem(token, siteId, listId, fields) {
    return graphJson(token, "POST", "/sites/" + siteId + "/lists/" + listId + "/items", { fields: fields });
  }

  async function updateListItemFields(token, siteId, listId, itemId, fields) {
    return graphJson(token, "PATCH", "/sites/" + siteId + "/lists/" + listId + "/items/" + itemId + "/fields", fields);
  }

  async function deleteListItem(token, siteId, listId, itemId) {
    return graphJson(token, "DELETE", "/sites/" + siteId + "/lists/" + listId + "/items/" + itemId);
  }

  // ---------- Teams ----------

  async function joinedTeams(token) {
    var res = await graphJson(token, "GET", "/me/joinedTeams?$select=id,displayName");
    return (res.value || []).sort(function (a, b) { return a.displayName.localeCompare(b.displayName); });
  }

  async function listChannels(token, teamId) {
    var res = await graphJson(token, "GET", "/teams/" + teamId + "/channels?$select=id,displayName");
    return (res.value || []).sort(function (a, b) { return a.displayName.localeCompare(b.displayName); });
  }

  async function followedSites(token) {
    try {
      var res = await graphJson(token, "GET", "/me/followedSites?$select=id,displayName,webUrl");
      return res.value || [];
    } catch (e) { return []; }
  }

  async function allSites(token) {
    var res = await graphJson(token, "GET", "/sites?search=*&$select=id,displayName,webUrl&$top=50");
    return res.value || [];
  }

  /** A Team's built-in SharePoint site — every team member has access. */
  async function teamSite(token, groupId) {
    var site = await graphJson(token, "GET", "/groups/" + groupId + "/sites/root?$select=id,displayName,webUrl");
    return { siteId: site.id, name: site.displayName, webUrl: site.webUrl };
  }

  /** The user's own OneDrive as a site (personal site can host the lists). */
  async function myPersonalSite(token) {
    var drive = await graphJson(token, "GET", "/me/drive?$select=webUrl");
    var u = new URL(drive.webUrl); // .../personal/user_x/Documents
    var path = u.pathname.replace(/\/Documents\/?$/i, "");
    var site = await graphJson(token, "GET", "/sites/" + u.hostname + ":" + path + "?$select=id,displayName,webUrl");
    return { siteId: site.id, name: site.displayName || "My OneDrive", webUrl: site.webUrl };
  }

  async function searchSites(token, query) {
    var res = await graphJson(token, "GET", "/sites?search=" + encodeURIComponent(query) + "&$select=id,displayName,webUrl");
    return res.value || [];
  }

  async function postChannelMessage(token, teamId, channelId, payload) {
    return graphJson(token, "POST", "/teams/" + teamId + "/channels/" + channelId + "/messages", payload);
  }

  function oidFromToken(token) {
    try {
      var payload = JSON.parse(atob(token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/")));
      return payload.oid || "";
    } catch (e) { return ""; }
  }

  /** Create a Teams tag with the signed-in coordinator as first member. */
  async function createTeamTag(token, teamId, displayName) {
    var oid = oidFromToken(token);
    if (!oid) { throw new Error("couldn't read your user id from the sign-in token"); }
    return graphJson(token, "POST", "/teams/" + teamId + "/tags", {
      displayName: displayName,
      members: [{ userId: oid }],
    });
  }

  async function listTeamTags(token, teamId) {
    var res = await graphJson(token, "GET", "/teams/" + teamId + "/tags");
    return res.value || [];
  }

  // ---------- mail ----------

  async function sendMail(token, to, subject, html) {
    return graphJson(token, "POST", "/me/sendMail", {
      message: {
        subject: subject,
        body: { contentType: "HTML", content: html },
        toRecipients: (to || []).map(function (a) { return { emailAddress: { address: a } }; }),
      },
      saveToSentItems: true,
    });
  }

  async function createDraftMessage(token, to, subject, html) {
    return graphJson(token, "POST", "/me/messages", {
      subject: subject,
      body: { contentType: "HTML", content: html },
      toRecipients: (to || []).map(function (a) { return { emailAddress: { address: a } }; }),
    });
  }

  async function sendDraft(token, messageId) {
    return graphJson(token, "POST", "/me/messages/" + encodeURIComponent(messageId) + "/send");
  }

  async function getAttachments(token, messageId) {
    var res = await graphJson(token, "GET", "/me/messages/" + encodeURIComponent(messageId) +
      "/attachments?$select=id,name,contentType,size");
    return res.value || [];
  }

  /** All replies to a channel message (fiscal harvest), capped paging. */
  async function listReplies(token, teamId, channelId, messageId) {
    var out = [];
    var url = "/teams/" + teamId + "/channels/" + channelId +
      "/messages/" + messageId + "/replies?$top=50";
    var guard = 0;
    while (url && guard++ < 4) {
      var page = await graphJson(token, "GET", url);
      out = out.concat(page.value || []);
      url = page["@odata.nextLink"] ? page["@odata.nextLink"].replace(graphBase(), "") : null;
    }
    return out;
  }

  /** Resolve a Teams attachment contentUrl (SharePoint webUrl) to a driveItem. */
  async function driveItemFromUrl(token, webUrl) {
    var b64 = btoa(unescape(encodeURIComponent(webUrl)))
      .replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
    return graphJson(token, "GET", "/shares/u!" + b64 + "/driveItem?$select=id,name,size,parentReference");
  }

  /** driveItem content as base64 (feeds the DOCX/XLSX extractors). */
  async function driveItemContentB64(token, driveId, itemId) {
    // NOT /content: that answers 302 to a pre-authenticated *.sharepoint.com
    // URL and browsers drop the Authorization header across a cross-origin
    // redirect. The download URL is already authenticated — fetch it plain.
    var meta = await graphJson(token, "GET", "/drives/" + driveId + "/items/" + itemId +
      "?$select=id,@microsoft.graph.downloadUrl");
    var url = meta && meta["@microsoft.graph.downloadUrl"];
    if (!url) { throw new Error("no download URL for that attachment"); }
    var res = await fetch(url);
    if (!res.ok) { throw new Error("attachment download -> " + res.status); }
    var buf = new Uint8Array(await res.arrayBuffer());
    var bin = "";
    var CHUNK = 0x8000;
    for (var i = 0; i < buf.length; i += CHUNK) {
      bin += String.fromCharCode.apply(null, buf.subarray(i, i + CHUNK));
    }
    return btoa(bin);
  }

  async function getAttachmentBytes(token, messageId, attachmentId) {
    var full = await graphJson(token, "GET", "/me/messages/" + encodeURIComponent(messageId) +
      "/attachments/" + encodeURIComponent(attachmentId));
    return full.contentBytes; // base64
  }

  root.GraphData = {
    setCloud: setCloud,
    getToken: getToken,
    getTagWriteToken: getTagWriteToken,
    getListCreateToken: getListCreateToken,
    resolveSite: resolveSite,
    findList: findList,
    listItems: listItems,
    createList: createList,
    addListColumn: addListColumn,
    addListItem: addListItem,
    updateListItemFields: updateListItemFields,
    deleteListItem: deleteListItem,
    joinedTeams: joinedTeams,
    listChannels: listChannels,
    searchSites: searchSites,
    followedSites: followedSites,
    allSites: allSites,
    myPersonalSite: myPersonalSite,
    teamSite: teamSite,
    postChannelMessage: postChannelMessage,
    listTeamTags: listTeamTags,
    createTeamTag: createTeamTag,
    sendMail: sendMail,
    createDraftMessage: createDraftMessage,
    sendDraft: sendDraft,
    getAttachments: getAttachments,
    getAttachmentBytes: getAttachmentBytes,
    listReplies: listReplies,
    driveItemFromUrl: driveItemFromUrl,
    driveItemContentB64: driveItemContentB64,
    _config: { clientId: CLIENT_ID, clouds: CLOUDS },
  };
})(typeof self !== "undefined" ? self : this);
