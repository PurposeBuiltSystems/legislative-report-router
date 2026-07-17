/*
 * Legislative Report Router — per-state presets (pure data + lookup).
 *
 * Each preset: bill identifier prefixes (drives the parser's boundary
 * detection) and the Open States jurisdiction name (drives the feed
 * mirror). DEFAULT_IDS covers the ~35 states that use the common
 * HB/SB pattern; states with distinctive numbering get overrides.
 * Everything remains editable in Settings after a preset is applied —
 * presets are starting points, not constraints.
 */
(function (root) {
  "use strict";

  var DEFAULT_IDS = ["HB", "SB", "HR", "SR", "HJR", "SJR", "HCR", "SCR", "HM", "SM"];

  var STATES = {
    "Iowa":          { ids: ["HF", "SF", "HSB", "SSB", "HJR", "SJR", "HCR", "SCR", "HR", "SR"], feed: "iowa-rss" },
    "Minnesota":     { ids: ["HF", "SF"] },
    "California":    { ids: ["AB", "SB", "ACA", "SCA", "AJR", "SJR", "ACR", "SCR", "AR", "SR"] },
    "Wisconsin":     { ids: ["AB", "SB", "AJR", "SJR", "AR", "SR"] },
    "Nevada":        { ids: ["AB", "SB", "AJR", "SJR", "ACR", "SCR"] },
    "New York":      { ids: ["A", "S"] },
    "New Jersey":    { ids: ["A", "S", "AJR", "SJR", "ACR", "SCR", "AR"] },
    "Massachusetts": { ids: ["H", "S", "HD", "SD"] },
    "Maryland":      { ids: ["HB", "SB", "HJ", "SJ"] },
    "Virginia":      { ids: ["HB", "SB", "HJ", "SJ", "HR", "SR"] },
    "North Dakota":  { ids: ["HB", "SB", "HCR", "SCR", "HMR", "SMR"] },
    "United States": { ids: ["HR", "S", "HJRES", "SJRES", "HCONRES", "SCONRES", "HRES", "SRES"], jurisdiction: "United States" },
  };

  var ALL_STATE_NAMES = [
    "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado",
    "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho",
    "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana",
    "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota",
    "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada",
    "New Hampshire", "New Jersey", "New Mexico", "New York",
    "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon",
    "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota",
    "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington",
    "West Virginia", "Wisconsin", "Wyoming", "United States",
  ];

  function presetFor(stateName) {
    var s = STATES[stateName] || {};
    return {
      state: stateName,
      identifiers: (s.ids || DEFAULT_IDS).slice(),
      feed: s.feed || "openstates",
      jurisdiction: s.jurisdiction || stateName,
      slug: String(stateName).toLowerCase().replace(/\s+/g, "-"),
    };
  }

  var api = { presetFor: presetFor, ALL_STATE_NAMES: ALL_STATE_NAMES, DEFAULT_IDS: DEFAULT_IDS };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrPresets = api; }
})(typeof self !== "undefined" ? self : this);
