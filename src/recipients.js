/* Shared division audience: Teams tags + routing-list/pasted email contacts.
 * Additive reconciliation only; never add Team members or change tag membership.
 */
(function (root) {
  "use strict";
  function email(value) { return String(value || "").trim().toLowerCase(); }
  function valid(value) { return /^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/.test(value); }
  function unique(values) { return Array.from(new Set(values)); }

  async function reconcile(graph, token, rules, site) {
    var active = rules.filter(function (r) { return r.isActive !== false; });
    var teams = new Map(), tags = new Map(), groups = new Map(), warnings = [];
    // Read everything before any writes. An unavailable source is an error,
    // never permission to publish using only the other half of the audience.
    for (var i = 0; i < active.length; i++) {
      var rule = active[i];
      var code = String(rule.divisionCode).trim().toLowerCase();
      if (!groups.has(code)) { groups.set(code, []); }
      var addresses = groups.get(code);
      (rule.emails || []).forEach(function (address) {
        var normalized = email(address);
        if (!valid(normalized)) { throw new Error(rule.divisionCode + ": invalid email contact " + address); }
        addresses.push(normalized);
      });
      if (!rule.teamsTeamId) { continue; }
      if (!teams.has(rule.teamsTeamId)) {
        teams.set(rule.teamsTeamId, await graph.listTeamMembers(token, rule.teamsTeamId));
      }
      if (!rule.teamsTagId) { continue; }
      var key = rule.teamsTeamId + "/" + rule.teamsTagId;
      if (!tags.has(key)) { tags.set(key, await graph.listTeamTagMembers(token, rule.teamsTeamId, rule.teamsTagId)); }
      tags.get(key).forEach(function (tagMember) {
        var member = teams.get(rule.teamsTeamId).find(function (m) {
          return m.userId === tagMember.userId && (!m.tenantId || !tagMember.tenantId || m.tenantId === tagMember.tenantId);
        });
        if (!member || !valid(email(member.email))) {
          throw new Error(rule.divisionCode + ": no email address found for tag member " + (tagMember.displayName || tagMember.userId) + ". Check that person's Team membership.");
        }
        addresses.push(email(member.email));
      });
    }
    var plans = active.map(function (r) {
      var addresses = unique(groups.get(String(r.divisionCode).trim().toLowerCase()));
      var ids = (r.mentionUserIds || []).slice(), names = (r.mentionUserEmails || []).slice();
      var covered = tags.get(r.teamsTeamId + "/" + r.teamsTagId) || [];
      addresses.forEach(function (address) {
        var matches = (teams.get(r.teamsTeamId) || []).filter(function (m) { return email(m.email) === address && m.userId; });
        if (!r.teamsTeamId || !r.teamsChannelId || matches.length !== 1) {
          warnings.push(r.divisionCode + ": " + address + " receives email only; " +
            (!r.teamsTeamId || !r.teamsChannelId ? "save a Team and channel." : "match this email to one Team member (add the person to the Team if needed)."));
          return;
        }
        var member = matches[0];
        if (covered.some(function (m) { return m.userId === member.userId; }) || ids.indexOf(member.userId) !== -1) { return; }
        ids.push(member.userId); names.push(member.displayName || address);
      });
      return { rule: r, emails: addresses, mentionUserIds: ids, mentionUserEmails: names };
    });
    var changed = 0;
    for (var p = 0; p < plans.length; p++) {
      var plan = plans[p];
      var before = unique((plan.rule.emails || []).map(email)).sort().join(";");
      if (before !== plan.emails.slice().sort().join(";")) {
        await graph.updateRecipientEmails(token, site.siteId, site.routingListId, plan.rule.id, { Emails: plan.emails.join("; ") }, plan.rule.recipientEtag);
        changed++;
      }
    }
    // Commit the in-memory audience only after every required write succeeds.
    plans.forEach(function (plan) {
      plan.rule.emails = plan.emails;
      plan.rule.mentionUserIds = plan.mentionUserIds;
      plan.rule.mentionUserEmails = plan.mentionUserEmails;
    });
    return { changed: changed, warnings: unique(warnings), rules: active.length };
  }
  var api = { reconcile: reconcile };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrRecipients = api; }
})(typeof self !== "undefined" ? self : this);
