#!/usr/bin/env node
'use strict';

/**
 * set-host.js — Switch manifest URLs between localhost and a remote domain.
 *
 * Usage:
 *   node scripts/set-host.js                      # interactive prompt
 *   node scripts/set-host.js local                 # https://localhost:3000
 *   node scripts/set-host.js https://superqat.app  # remote domain
 */

const fs = require('fs');
const path = require('path');
const readline = require('readline');

const MANIFESTS_DIR = path.join(__dirname, '..', 'manifests');
const MANIFEST_FILES = ['word-manifest.xml', 'excel-manifest.xml', 'powerpoint-manifest.xml'];

// Matches any https://... URL up to the next quote or closing tag
const URL_PATTERN = /https:\/\/[^"<\s]+/g;

function rewriteManifests(baseUrl) {
  // Ensure no trailing slash
  baseUrl = baseUrl.replace(/\/+$/, '');

  MANIFEST_FILES.forEach(function (filename) {
    const filePath = path.join(MANIFESTS_DIR, filename);
    if (!fs.existsSync(filePath)) {
      console.log('  Skipped (not found): ' + filename);
      return;
    }
    let xml = fs.readFileSync(filePath, 'utf8');

    // Replace all localhost:3000 or any previous host with the new base
    xml = xml.replace(/https:\/\/localhost:3000/g, baseUrl);
    // Also catch any previous remote domain that was set
    xml = xml.replace(/https:\/\/[a-zA-Z0-9._-]+(?::\d+)?(?=\/assets\/|\/taskpane\.|\/commands\.)/g, baseUrl);

    fs.writeFileSync(filePath, xml, 'utf8');
    console.log('  Updated: ' + filename);
  });
}

function run() {
  var arg = process.argv[2];

  if (arg === 'local') {
    console.log('Setting manifests to https://localhost:3000');
    rewriteManifests('https://localhost:3000');
    return;
  }

  if (arg && arg.startsWith('https://')) {
    console.log('Setting manifests to ' + arg);
    rewriteManifests(arg);
    return;
  }

  if (arg) {
    // Assume it's a bare domain, add https://
    var url = 'https://' + arg;
    console.log('Setting manifests to ' + url);
    rewriteManifests(url);
    return;
  }

  // Interactive prompt
  var rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  console.log('');
  console.log('SuperQAT — Set hosting domain');
  console.log('');
  console.log('  1) Local development server (https://localhost:3000)');
  console.log('  2) Remote domain (e.g. https://superqat.app)');
  console.log('');
  rl.question('Choice [1/2]: ', function (choice) {
    if (choice.trim() === '2') {
      rl.question('Enter domain (e.g. superqat.app): ', function (domain) {
        domain = domain.trim().replace(/^https?:\/\//, '').replace(/\/+$/, '');
        var url = 'https://' + domain;
        console.log('Setting manifests to ' + url);
        rewriteManifests(url);
        rl.close();
      });
    } else {
      console.log('Setting manifests to https://localhost:3000');
      rewriteManifests('https://localhost:3000');
      rl.close();
    }
  });
}

run();
