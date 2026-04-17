#!/usr/bin/env node
/**
 * Parse official Microsoft Office 365 control ID xlsx files
 * and output a unified command-map.json for the SuperQAT add-in.
 */
var XLSX = require("xlsx");
var fs = require("fs");
var path = require("path");

var files = {
  word: "/tmp/wordcontrols.xlsx",
  excel: "/tmp/excelcontrols.xlsx",
  powerpoint: "/tmp/powerpointcontrols.xlsx"
};

var actionableTypes = new Set([
  "button", "toggleButton", "splitButton", "checkBox"
]);

var allCommands = {};   // idMso -> { name, type, tab, group, apps: [] }
var appCounts = {};

Object.keys(files).forEach(function(app) {
  var wb = XLSX.readFile(files[app]);
  var ws = wb.Sheets[wb.SheetNames[0]];
  var rows = XLSX.utils.sheet_to_json(ws);
  var count = 0;

  rows.forEach(function(row) {
    var name = (row["Control Name"] || "").trim();
    var type = (row["Control Type"] || "").trim().toLowerCase();
    var tab = (row["Tab"] || "").trim();
    var group = (row["Group/Context Menu Name"] || "").trim();

    if (!name || !actionableTypes.has(type)) return;

    count++;
    if (!allCommands[name]) {
      allCommands[name] = {
        name: name,
        type: type,
        tab: tab,
        group: group,
        apps: []
      };
    }
    if (allCommands[name].apps.indexOf(app) === -1) {
      allCommands[name].apps.push(app);
    }
  });

  appCounts[app] = count;
  console.log(app + ": " + count + " actionable controls");
});

// Sort by name
var sorted = Object.values(allCommands).sort(function(a, b) {
  return a.name.localeCompare(b.name);
});

console.log("\nTotal unique actionable commands: " + sorted.length);
console.log("Common to all 3 apps: " + sorted.filter(function(c) { return c.apps.length === 3; }).length);

// Group by tab for categorization
var tabs = {};
sorted.forEach(function(c) {
  var t = c.tab || "Other";
  if (!tabs[t]) tabs[t] = [];
  tabs[t].push(c.name);
});

console.log("\nTabs found: " + Object.keys(tabs).length);
Object.keys(tabs).sort().forEach(function(t) {
  console.log("  " + t + ": " + tabs[t].length + " commands");
});

// Write the full command map
var dataDir = path.join(__dirname, "..", "src", "data");
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

// Full map: { idMso: { type, tab, group, apps } }
fs.writeFileSync(
  path.join(dataDir, "command-map.json"),
  JSON.stringify(sorted, null, 2) + "\n"
);
console.log("\nWrote src/data/command-map.json (" + sorted.length + " commands)");

// Simplified list for the taskpane: just names grouped by tab
var byTab = {};
sorted.forEach(function(c) {
  var t = c.tab || "Other";
  if (!byTab[t]) byTab[t] = [];
  byTab[t].push({ name: c.name, type: c.type, apps: c.apps });
});
fs.writeFileSync(
  path.join(dataDir, "commands-by-tab.json"),
  JSON.stringify(byTab, null, 2) + "\n"
);
console.log("Wrote src/data/commands-by-tab.json");

// App-specific lists
["word", "excel", "powerpoint"].forEach(function(app) {
  var appCmds = sorted.filter(function(c) { return c.apps.indexOf(app) !== -1; });
  fs.writeFileSync(
    path.join(dataDir, app + "-commands.json"),
    JSON.stringify(appCmds.map(function(c) { return c.name; }), null, 2) + "\n"
  );
  console.log("Wrote src/data/" + app + "-commands.json (" + appCmds.length + " commands)");
});
