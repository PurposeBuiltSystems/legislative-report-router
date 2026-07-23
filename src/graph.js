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

  var SCOPES = ["Mail.ReadWrite", "Mail.Send", "ChannelMessage.Send", "Sites.ReadWrite.All", "TeamworkTag.Read"];

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

  async function getToken() {
    var pca = await getPca();
    try {
      var silent = await pca.acquireTokenSilent({ scopes: SCOPES });
      return silent.accessToken;
    } catch (e) {
      var interactive = await pca.acquireTokenPopup({ scopes: SCOPES });
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

  /** "https://x.sharepoint.com/sites/Y" → Graph siteId. */
  async function resolveSite(token, siteUrl) {
    var u = new URL(String(siteUrl).trim());
    var path = u.pathname.replace(/\/+$/, "");
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

  // ---------- Teams ----------

  async function joinedTeams(token) {
    var res = await graphJson(token, "GET", "/me/joinedTeams?$select=id,displayName");
    return (res.value || []).sort(function (a, b) { return a.displayName.localeCompare(b.displayName); });
  }

  async function listChannels(token, teamId) {
    var res = await graphJson(token, "GET", "/teams/" + teamId + "/channels?$select=id,displayName");
    return (res.value || []).sort(function (a, b) { return a.displayName.localeCompare(b.displayName); });
  }

  async function searchSites(token, query) {
    var res = await graphJson(token, "GET", "/sites?search=" + encodeURIComponent(query) + "&$select=id,displayName,webUrl");
    return res.value || [];
  }

  async function postChannelMessage(token, teamId, channelId, payload) {
    return graphJson(token, "POST", "/teams/" + teamId + "/channels/" + channelId + "/messages", payload);
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

  async function sendDraft(token, messageId) {
    return graphJson(token, "POST", "/me/messages/" + encodeURIComponent(messageId) + "/send");
  }

  async function getAttachments(token, messageId) {
    var res = await graphJson(token, "GET", "/me/messages/" + encodeURIComponent(messageId) +
      "/attachments?$select=id,name,contentType,size");
    return res.value || [];
  }

  async function getAttachmentBytes(token, messageId, attachmentId) {
    var full = await graphJson(token, "GET", "/me/messages/" + encodeURIComponent(messageId) +
      "/attachments/" + encodeURIComponent(attachmentId));
    return full.contentBytes; // base64
  }

  root.GraphData = {
    setCloud: setCloud,
    getToken: getToken,
    resolveSite: resolveSite,
    findList: findList,
    listItems: listItems,
    createList: createList,
    addListItem: addListItem,
    joinedTeams: joinedTeams,
    listChannels: listChannels,
    searchSites: searchSites,
    postChannelMessage: postChannelMessage,
    listTeamTags: listTeamTags,
    sendMail: sendMail,
    sendDraft: sendDraft,
    getAttachments: getAttachments,
    getAttachmentBytes: getAttachmentBytes,
    _config: { clientId: CLIENT_ID, clouds: CLOUDS },
  };
})(typeof self !== "undefined" ? self : this);
