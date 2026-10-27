/* Permission-free shared contacts. A saved, user-confirmed mapping connects
 * existing Teams tag user IDs to supplied emails. No directory lookup or
 * Team/tag membership write is performed. */
(function (root) {
  "use strict";
  var MARKER = "LRR_CONTACT_MATCH_V1";
  function email(value) { return String(value || "").trim().toLowerCase(); }
  function valid(value) { return /^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/.test(value); }
  function unique(values) { return Array.from(new Set(values)); }
  function norm(value) { return String(value || "").trim().replace(/\s+/g, " ").toLowerCase(); }
  function isContactRecord(fields) { return fields && fields.Notes === MARKER; }
  function mappings(rows, rules) {
    var result = new Map();
    (rules || []).forEach(function (rule) {
      (rule.mentionUserIds || []).forEach(function (id, i) {
        var address = email((rule.mentionUserEmails || [])[i]);
        if (rule.teamsTeamId && id && valid(address)) {
          result.set(rule.teamsTeamId + "/" + id, { teamId: rule.teamsTeamId, userId: id, email: address, displayName: address });
        }
      });
    });
    var saved = new Map();
    (rows || []).filter(function (r) { return isContactRecord(r.fields); }).forEach(function (row) {
      var f = row.fields, address = email(f.MentionUserEmails), id = f.MentionUserIds;
      if (!f.TeamsTeamId || !id || !valid(address)) { return; }
      var key = f.TeamsTeamId + "/" + id;
      if (saved.has(key) && saved.get(key).email !== address) {
        throw new Error("Conflicting saved email matches for " + (f.DivisionName || id) + ". Correct the contact rows in the routing list first.");
      }
      saved.set(key, { teamId: f.TeamsTeamId, userId: id, email: address,
        displayName: f.DivisionName || address, id: row.id,
        etag: row.eTag || row["@odata.etag"] || f["@odata.etag"] });
    });
    saved.forEach(function (match, key) { result.set(key, match); });
    return Array.from(result.values());
  }

  // Name matches only suggest a choice in the form. They are never saved or
  // used to address a message until the user explicitly confirms them.
  function suggest(person, people, pasted, rule) {
    if (!norm(person.displayName) || people.filter(function (p) { return norm(p.displayName) === norm(person.displayName); }).length !== 1) { return ""; }
    var codes = [rule.divisionCode, rule.divisionName].concat(rule.aliases || []).map(norm);
    var choices = unique((pasted || []).filter(function (p) {
      return norm(p.name) === norm(person.displayName) && (p.divisions || []).some(function (d) { return codes.indexOf(norm(d)) !== -1; });
    }).map(function (p) { return email(p.email); }).filter(valid));
    return choices.length === 1 ? choices[0] : "";
  }

  async function loadPeople(graph, token, rule) {
    if (!rule.teamsTeamId || !rule.teamsTagId) { throw new Error("Save this division's Team, posting channel and tag first."); }
    var tags = await graph.listTeamTags(token, rule.teamsTeamId);
    if (!tags.some(function (t) { return t.id === rule.teamsTagId; })) { tags.unshift({ id: rule.teamsTagId }); }
    var people = new Map(), selected = [];
    var next = 0;
    async function worker() {
      while (next < tags.length) {
        var tag = tags[next++];
        var members = await graph.listTeamTagMembers(token, rule.teamsTeamId, tag.id);
        members.forEach(function (member) {
          if (!member.userId) { return; }
          var previous = people.get(member.userId);
          if (previous && previous.tenantId && member.tenantId && previous.tenantId !== member.tenantId) {
            throw new Error("A Teams identity is ambiguous across tenants; match it with your Team owner.");
          }
          people.set(member.userId, member);
          if (tag.id === rule.teamsTagId) { selected.push(member.userId); }
        });
      }
    }
    await Promise.all([worker(), worker(), worker()]);
    return { people: Array.from(people.values()), selected: unique(selected) };
  }

  async function reconcile(graph, token, rules, site, matches) {
    var active = rules.filter(function (r) { return r.isActive !== false; });
    var tags = new Map(), groups = new Map(), warnings = [];
    matches = matches || [];
    // All tag reads and required mappings must succeed before any writes.
    for (var i = 0; i < active.length; i++) {
      var rule = active[i], code = norm(rule.divisionCode);
      if (!groups.has(code)) { groups.set(code, []); }
      var addresses = groups.get(code);
      (rule.emails || []).forEach(function (address) {
        var normalized = email(address);
        if (!valid(normalized)) { throw new Error(rule.divisionCode + ": invalid email contact " + address); }
        addresses.push(normalized);
      });
      if (!rule.teamsTeamId || !rule.teamsTagId) { continue; }
      var key = rule.teamsTeamId + "/" + rule.teamsTagId;
      if (!tags.has(key)) { tags.set(key, await graph.listTeamTagMembers(token, rule.teamsTeamId, rule.teamsTagId)); }
      tags.get(key).forEach(function (person) {
        var match = matches.find(function (m) { return m.teamId === rule.teamsTeamId && m.userId === person.userId; });
        if (!match || !valid(email(match.email))) {
          throw new Error(rule.divisionCode + ": match an email for " + (person.displayName || person.userId) + " in Step 4 → Match Teams people to emails. No new permission is needed.");
        }
        addresses.push(email(match.email));
      });
    }
    var plans = active.map(function (r) {
      var addresses = unique(groups.get(norm(r.divisionCode))), ids = [], names = [];
      var covered = tags.get(r.teamsTeamId + "/" + r.teamsTagId) || [];
      addresses.forEach(function (address) {
        var found = matches.filter(function (m) { return m.teamId === r.teamsTeamId && email(m.email) === address; });
        if (!r.teamsTeamId || !r.teamsChannelId || !found.length) {
          warnings.push(r.divisionCode + ": " + address + " receives email only; " +
            (!r.teamsTeamId || !r.teamsChannelId ? "save a Team and channel." : "confirm its Teams match in Step 4. If the person is missing, add them to a tag in this Team first."));
          return;
        }
        found.forEach(function (match) {
          if (covered.some(function (p) { return p.userId === match.userId; }) || ids.indexOf(match.userId) !== -1) { return; }
          ids.push(match.userId); names.push(match.displayName || address);
        });
      });
      return { rule: r, emails: addresses, mentionUserIds: ids, mentionUserEmails: names };
    });
    var changed = 0;
    for (var p = 0; p < plans.length; p++) {
      var plan = plans[p];
      if (unique((plan.rule.emails || []).map(email)).sort().join(";") !== plan.emails.slice().sort().join(";")) {
        await graph.updateRecipientEmails(token, site.siteId, site.routingListId, plan.rule.id, { Emails: plan.emails.join("; ") }, plan.rule.recipientEtag);
        changed++;
      }
    }
    plans.forEach(function (plan) { plan.rule.emails = plan.emails; plan.rule.mentionUserIds = plan.mentionUserIds; plan.rule.mentionUserEmails = plan.mentionUserEmails; });
    return { changed: changed, warnings: unique(warnings), rules: active.length };
  }
  var api = { MARKER: MARKER, isContactRecord: isContactRecord, mappings: mappings, suggest: suggest, loadPeople: loadPeople, reconcile: reconcile };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrRecipients = api; }
})(typeof self !== "undefined" ? self : this);
