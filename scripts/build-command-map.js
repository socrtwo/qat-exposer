#!/usr/bin/env node
/**
 * Processes the deduplicated Office command list.
 * 1. Reads the TSV data
 * 2. Cleans labels (removes & accelerators)
 * 3. Outputs src/data/command-map.json
 * 4. Prints commands NOT yet in ALL_COMMANDS
 */

var fs = require("fs");
var path = require("path");

// Read the raw TSV
var raw = fs.readFileSync(path.join(__dirname, "raw-deduplicated.tsv"), "utf8");
var lines = raw.split("\n");

var commands = [];
lines.forEach(function (line) {
  line = line.trim();
  if (!line || line.startsWith("ID")) return; // skip header

  var tab = line.indexOf("\t");
  if (tab === -1) return;

  var id = parseInt(line.substring(0, tab).trim(), 10);
  var rawLabel = line.substring(tab + 1).trim();
  if (isNaN(id) || !rawLabel) return;

  // Clean: remove & accelerator, remove surrounding quotes, trim trailing ...
  var clean = rawLabel
    .replace(/^"/, "").replace(/"$/, "")
    .replace(/&/g, "")
    .replace(/\.\.\.$/,"")
    .trim();

  commands.push({ id: id, raw: rawLabel, label: clean });
});

// Sort by ID
commands.sort(function(a, b) { return a.id - b.id; });

// Write the mapping file
var dataDir = path.join(__dirname, "..", "src", "data");
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

var mapObj = {};
commands.forEach(function(c) { mapObj[c.id] = c.label; });
fs.writeFileSync(
  path.join(dataDir, "command-map.json"),
  JSON.stringify(mapObj, null, 2) + "\n"
);

console.log("Wrote " + commands.length + " commands to src/data/command-map.json");

// Now read taskpane.js to find existing ALL_COMMANDS labels
var taskpane = fs.readFileSync(
  path.join(__dirname, "..", "src", "taskpane", "taskpane.js"), "utf8"
);

// Extract existing labels from ALL_COMMANDS
var existingLabels = new Set();
var labelRegex = /\["([^"]+)"/g;
var m;
while ((m = labelRegex.exec(taskpane)) !== null) {
  existingLabels.add(m[1].toLowerCase());
}

// Find commands not yet in ALL_COMMANDS
var missing = [];
commands.forEach(function(c) {
  if (!existingLabels.has(c.label.toLowerCase())) {
    missing.push(c);
  }
});

console.log("\nExisting ALL_COMMANDS labels: " + existingLabels.size);
console.log("New commands to add: " + missing.length);
console.log("\n--- NEW COMMANDS ---");
missing.forEach(function(c) {
  console.log(c.id + "\t" + c.label);
});

// Write missing commands to a file for reference
fs.writeFileSync(
  path.join(dataDir, "new-commands.json"),
  JSON.stringify(missing, null, 2) + "\n"
);
