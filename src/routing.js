/*
 * Legislative Report Router — routing engine (pure logic).
 *
 * Matches parsed legislative items against routing rules loaded from the
 * LegislativeRoutingMatrix SharePoint list (or any RoutingRepository —
 * the shape below is storage-agnostic so Dataverse/SQL can swap in).
 *
 * RoutingRule shape:
 *   { id, divisionCode, divisionName, aliases:[], emails:[],
 *     teamsTeamId, teamsChannelId, teamsChannelName,
 *     teamsTagId, teamsTagName, mentionUserIds:[], mentionUserEmails:[],
 *     isActive, priority, effectiveStartDate, effectiveEndDate, notes }
 */
(function (root) {
  "use strict";

  function norm(s) { return String(s || "").replace(/\s+/g, " ").trim().toLowerCase(); }

  function splitList(s) {
    return String(s || "").split(/[;,]/).map(function (t) { return t.trim(); }).filter(Boolean);
  }

  /** Is the rule active on `date` (defaults now)? */
  function ruleActive(rule, date) {
    if (rule.isActive === false) { return false; }
    var d = date ? new Date(date) : new Date();
    if (rule.effectiveStartDate && d < new Date(rule.effectiveStartDate)) { return false; }
    if (rule.effectiveEndDate && d > new Date(rule.effectiveEndDate)) { return false; }
    return true;
  }

  /** All names a rule answers to. */
  function ruleNames(rule) {
    return [rule.divisionCode, rule.divisionName]
      .concat(rule.aliases || [])
      .map(norm)
      .filter(Boolean);
  }

  /**
   * Find active rules for one division token, highest priority first.
   * Tolerates qualifier suffixes from real reports: "ELT - awareness"
   * matches the ELT rule (full token tried first, then the code part).
   */
  function rulesForDivision(division, rules, date) {
    var candidates = [norm(division)];
    var codePart = norm(String(division || "").split(/\s*[-–]\s*/)[0]);
    if (codePart && candidates.indexOf(codePart) === -1) { candidates.push(codePart); }
    for (var i = 0; i < candidates.length; i++) {
      var d = candidates[i];
      var hits = (rules || [])
        .filter(function (r) { return ruleActive(r, date) && ruleNames(r).indexOf(d) !== -1; })
        .sort(function (a, b) { return (b.priority || 0) - (a.priority || 0); });
      if (hits.length) { return hits; }
    }
    return [];
  }

  /**
   * Route one item: match every division in distributedTo ∪
   * commentRequestedFrom. Returns the item annotated with routingStatus,
   * matchedRoutingRules (rule ids), routes (full rule objects, deduped by
   * channel), and unknownDivisions.
   */
  function routeItem(item, rules, date) {
    var wanted = [];
    (item.distributedTo || []).concat(item.commentRequestedFrom || []).forEach(function (dv) {
      if (dv && wanted.indexOf(dv) === -1) { wanted.push(dv); }
    });

    var routes = [];
    var matchedIds = [];
    var unknown = [];
    wanted.forEach(function (dv) {
      var found = rulesForDivision(dv, rules, date);
      if (!found.length) { unknown.push(dv); return; }
      var top = found[0]; // highest priority wins per division
      if (matchedIds.indexOf(top.id) === -1) {
        matchedIds.push(top.id);
        routes.push(top);
      }
    });

    var status;
    if (!wanted.length) { status = "unmatched"; }
    else if (unknown.length === 0) { status = "matched"; }
    else if (routes.length > 0) { status = "partially-matched"; }
    else { status = "unmatched"; }

    item.matchedRoutingRules = matchedIds;
    item.routes = routes;
    item.unknownDivisions = unknown;
    item.routingStatus = item.routingStatus === "excluded" ? "excluded" : status;
    return item;
  }

  /** Route all items; returns summary counts. */
  function routeAll(items, rules, date) {
    var summary = { totalItems: 0, matchedItems: 0, partialItems: 0, unmatchedItems: 0, excludedItems: 0, lowConfidenceItems: 0 };
    (items || []).forEach(function (it) {
      routeItem(it, rules, date);
      summary.totalItems++;
      if (it.routingStatus === "matched") { summary.matchedItems++; }
      else if (it.routingStatus === "partially-matched") { summary.partialItems++; }
      else if (it.routingStatus === "excluded") { summary.excludedItems++; }
      else { summary.unmatchedItems++; }
      if ((it.parserConfidence || 0) < 0.7) { summary.lowConfidenceItems++; }
    });
    return summary;
  }

  /**
   * Group included, routed items by division rule for consolidated emails:
   * [{rule, items:[...]}] — one entry per distinct rule id.
   */
  function groupByRule(items) {
    var byId = {};
    var order = [];
    (items || []).forEach(function (it) {
      if (it.routingStatus === "excluded") { return; }
      (it.routes || []).forEach(function (r) {
        if (!byId[r.id]) { byId[r.id] = { rule: r, items: [] }; order.push(r.id); }
        byId[r.id].items.push(it);
      });
    });
    return order.map(function (id) { return byId[id]; });
  }

  /** Map a raw SharePoint list item (fields object) to a RoutingRule. */
  function ruleFromSharePoint(fields, id) {
    return {
      id: String(id != null ? id : (fields.id || fields.ID || "")),
      divisionCode: fields.DivisionCode || fields.Title || "",
      divisionName: fields.DivisionName || "",
      aliases: splitList(fields.Aliases),
      emails: splitList(fields.Emails),
      teamsTeamId: fields.TeamsTeamId || "",
      teamsChannelId: fields.TeamsChannelId || "",
      teamsChannelName: fields.TeamsChannelName || "",
      teamsTagId: fields.TeamsTagId || "",
      teamsTagName: fields.TeamsTagName || "",
      codeChapters: splitList(fields.CodeChapters),
      mentionUserIds: splitList(fields.MentionUserIds),
      mentionUserEmails: splitList(fields.MentionUserEmails),
      isActive: fields.IsActive !== false && fields.IsActive !== "No" && fields.IsActive !== "false",
      priority: Number(fields.Priority) || 0,
      effectiveStartDate: fields.EffectiveStartDate || "",
      effectiveEndDate: fields.EffectiveEndDate || "",
      notes: fields.Notes || "",
    };
  }

  var api = {
    ruleActive: ruleActive,
    rulesForDivision: rulesForDivision,
    routeItem: routeItem,
    routeAll: routeAll,
    groupByRule: groupByRule,
    ruleFromSharePoint: ruleFromSharePoint,
    splitList: splitList,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrRouting = api; }
})(typeof self !== "undefined" ? self : this);
