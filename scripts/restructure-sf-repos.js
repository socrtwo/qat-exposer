#!/usr/bin/env node
'use strict';

/**
 * Restructure SF-migrated GitHub repos.
 *
 * For each repo:
 * 1. Clone the GitHub repo
 * 2. Download the REAL zip from SourceForge (not the corrupt GitHub copy)
 * 3. Extract it into proper directory structure
 * 4. Move old corrupt files to releases/
 * 5. Add LICENSE, .gitignore if missing
 * 6. Commit and push
 *
 * Usage: GITHUB_TOKEN=ghp_xxx node scripts/restructure-sf-repos.js
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');
const https = require('https');
const http = require('http');

const TOKEN = process.env.GITHUB_TOKEN;
const OWNER = process.env.GITHUB_OWNER || 'socrtwo';

if (!TOKEN) {
  console.error('Set GITHUB_TOKEN environment variable');
  process.exit(1);
}

// Each repo with its SF project name and the best file to download from SF
const SF_REPOS = [
  { repo: 'autoscrshotanno-SF', sfProject: 'autoscrshotanno', sfFile: 'screenshot-annotate.zip' },
  { repo: 'catalog-of-life-SF', sfProject: 'catalog-of-life', sfFile: 'Catalogue-of-Life-Converter-1.0.zip' },
  { repo: 'corruptexcelrec-SF', sfProject: 'corruptexcelrec', sfFile: 's2_tools_for_excel_recovery_4.0.2_source_adware_removed.zip' },
  { repo: 'crrptoffcxtrctr-SF', sfProject: 'crrptoffcxtrctr', sfFile: 'corrupt_office_2007_extractor_delphi_7_source_code.zip' },
  { repo: 'datarecoverfree-SF', sfProject: 'datarecoverfree', sfFile: 'freeware_site_script_2.0.zip' },
  { repo: 'fasterposter-SF', sfProject: 'fasterposter', sfFile: 'fasterposter.com_11_29_2011.zip' },
  { repo: 'ged2wiki-SF', sfProject: 'ged2wiki', sfFile: 'gedcom2wiki_1.0.zip' },
  { repo: 'godskingsheroes-SF', sfProject: 'godskingsheroes', sfFile: 'famous family trees.zip' },
  { repo: 'qatindex-SF', sfProject: 'qatindex', sfFile: 'excel-powerpoint-qat-index.zip' },
  { repo: 'quickwordrecovr-SF', sfProject: 'quickwordrecovr', sfFile: 'savvy_docx_recovery_version_3.0_source.zip' },
  { repo: 'savvyoffice-SF', sfProject: 'savvyoffice', sfFile: 'Savvy_Repair_for_Microsoft_Office_v1.0.22_source.zip' },
  { repo: 'vistaprevrsrcvr-SF', sfProject: 'vistaprevrsrcvr', sfFile: 'previous_version_file_explorer_source_2.0.zip' },
  { repo: 'whereyoubin-SF', sfProject: 'whereyoubin', sfFile: 'wherehaveibeen_3.0.zip' },
  { repo: 'wordrecovery-SF', sfProject: 'wordrecovery', sfFile: 'Version 3.0.5-alpha-source.zip' },
  { repo: 'xmltrncatorfixr-SF', sfProject: 'xmltrncatorfixr', sfFile: 'xml_truncator_fixer_source.zip' },
];

function run(cmd, opts = {}) {
  console.log('  $ ' + cmd.substring(0, 120) + (cmd.length > 120 ? '...' : ''));
  return execSync(cmd, { stdio: 'pipe', timeout: 300000, ...opts }).toString().trim();
}

/**
 * Download a file following redirects (SourceForge uses many).
 * Returns a Buffer of the file content.
 */
function downloadFromSF(sfProject, sfFile) {
  // SourceForge direct download URL format
  const url = `https://sourceforge.net/projects/${sfProject}/files/${encodeURIComponent(sfFile)}/download`;
  console.log('  Downloading from SF: ' + sfFile);
  return followRedirects(url, 15);
}

function followRedirects(url, maxRedirects) {
  return new Promise((resolve, reject) => {
    if (maxRedirects <= 0) return reject(new Error('Too many redirects'));
    const mod = url.startsWith('https') ? https : http;
    const req = mod.get(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; SF2GH-Migrator/1.0)',
        'Accept': '*/*',
      },
      timeout: 60000,
    }, (res) => {
      // Follow redirects
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        let next = res.headers.location;
        if (next.startsWith('/')) {
          const parsed = new URL(url);
          next = parsed.protocol + '//' + parsed.host + next;
        }
        res.resume();
        return resolve(followRedirects(next, maxRedirects - 1));
      }
      if (res.statusCode !== 200) {
        res.resume();
        return reject(new Error(`HTTP ${res.statusCode}`));
      }
      // Check content type — if it's HTML, we got a mirror page not the file
      const ct = (res.headers['content-type'] || '').toLowerCase();
      if (ct.includes('text/html')) {
        // Consume the HTML and look for the actual download link
        let html = '';
        res.on('data', (d) => { html += d; });
        res.on('end', () => {
          // SF mirror pages have a direct link in a meta refresh or JS redirect
          const match = html.match(/https?:\/\/[^"'\s]+\/download[^"'\s]*/i) ||
                        html.match(/url=([^"'\s;]+)/i);
          if (match) {
            const directUrl = match[1] || match[0];
            console.log('  Following mirror redirect...');
            resolve(followRedirects(directUrl, maxRedirects - 1));
          } else {
            reject(new Error('Got HTML instead of file — download redirect failed'));
          }
        });
        return;
      }
      // Collect the binary data
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => resolve(Buffer.concat(chunks)));
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('Download timeout')); });
  });
}

function flattenSingleSubdir(dir) {
  const entries = fs.readdirSync(dir).filter(e => e !== '.git');
  if (entries.length === 1) {
    const child = path.join(dir, entries[0]);
    if (fs.statSync(child).isDirectory()) {
      console.log('    Flattening: ' + entries[0] + '/');
      const childEntries = fs.readdirSync(child);
      for (const e of childEntries) {
        const src = path.join(child, e);
        const dst = path.join(dir, e);
        if (!fs.existsSync(dst)) fs.renameSync(src, dst);
      }
      try { fs.rmSync(child, { recursive: true }); } catch (_) {}
    }
  }
}

const GITIGNORE = `# OS files
.DS_Store
Thumbs.db
desktop.ini

# IDE
.idea/
.vscode/
*.swp

# Build
*.o
*.obj
`;

const LICENSE_MIT = `MIT License

Copyright (c) ${new Date().getFullYear()} Paul D Pruitt

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
`;

async function processRepo(entry) {
  const { repo, sfProject, sfFile } = entry;
  console.log(`\n=== Processing ${repo} ===`);

  const tmpDir = path.join(os.tmpdir(), 'sf-restructure-' + repo);
  const extractDir = path.join(os.tmpdir(), 'sf-extract-' + repo);
  if (fs.existsSync(tmpDir)) fs.rmSync(tmpDir, { recursive: true });
  if (fs.existsSync(extractDir)) fs.rmSync(extractDir, { recursive: true });

  try {
    // Clone the GitHub repo
    const cloneUrl = `https://${TOKEN}@github.com/${OWNER}/${repo}.git`;
    console.log('  Cloning...');
    run(`git clone "${cloneUrl}" "${tmpDir}"`);

    const allFiles = fs.readdirSync(tmpDir).filter(f => f !== '.git');
    console.log('  Current files: ' + (allFiles.join(', ') || '(empty)'));

    // Check if already restructured
    if (allFiles.includes('src') || allFiles.includes('releases')) {
      console.log('  Already restructured, skipping.');
      return;
    }

    // Download the REAL zip from SourceForge
    let zipBuffer;
    try {
      zipBuffer = await downloadFromSF(sfProject, sfFile);
    } catch (dlErr) {
      console.log('  Download failed: ' + dlErr.message);
      return;
    }

    // Verify it's actually a zip (first 2 bytes = PK = 0x50 0x4B)
    if (zipBuffer.length < 4 || zipBuffer[0] !== 0x50 || zipBuffer[1] !== 0x4B) {
      console.log('  Downloaded file is not a valid zip (got ' + zipBuffer.length + ' bytes, starts with: ' +
        zipBuffer.slice(0, 4).toString('hex') + '). Skipping.');
      return;
    }
    console.log('  Downloaded ' + (zipBuffer.length / 1024).toFixed(0) + ' KB — valid zip.');

    // Save the real zip
    const zipPath = path.join(os.tmpdir(), 'sf-download-' + sfFile.replace(/[^a-zA-Z0-9._-]/g, '_'));
    fs.writeFileSync(zipPath, zipBuffer);

    // Extract
    fs.mkdirSync(extractDir, { recursive: true });
    try {
      run(`unzip -o -q "${zipPath}" -d "${extractDir}"`);
    } catch (unzipErr) {
      console.log('  Unzip failed: ' + unzipErr.message.split('\n')[0]);
      fs.unlinkSync(zipPath);
      return;
    }
    fs.unlinkSync(zipPath);

    flattenSingleSubdir(extractDir);

    const extracted = fs.readdirSync(extractDir);
    console.log('  Extracted ' + extracted.length + ' item(s).');

    if (extracted.length === 0) {
      console.log('  Nothing extracted, skipping.');
      return;
    }

    // Move old corrupt files to releases/
    const releasesDir = path.join(tmpDir, 'releases');
    fs.mkdirSync(releasesDir, { recursive: true });
    for (const f of allFiles) {
      if (f === 'README.md' || f === '.gitignore' || f === 'LICENSE') continue;
      const src = path.join(tmpDir, f);
      const dst = path.join(releasesDir, f);
      try { fs.renameSync(src, dst); } catch (_) {}
    }

    // Copy extracted files to repo root
    for (const item of extracted) {
      const src = path.join(extractDir, item);
      const dst = path.join(tmpDir, item);
      if (!fs.existsSync(dst)) {
        if (fs.statSync(src).isDirectory()) {
          run(`cp -r "${src}" "${dst}"`);
        } else {
          fs.copyFileSync(src, dst);
        }
      }
    }

    // Add .gitignore if missing
    if (!fs.existsSync(path.join(tmpDir, '.gitignore'))) {
      fs.writeFileSync(path.join(tmpDir, '.gitignore'), GITIGNORE);
      console.log('  Added .gitignore');
    }

    // Add LICENSE if missing
    if (!fs.existsSync(path.join(tmpDir, 'LICENSE')) && !fs.existsSync(path.join(tmpDir, 'LICENSE.md'))) {
      fs.writeFileSync(path.join(tmpDir, 'LICENSE'), LICENSE_MIT);
      console.log('  Added LICENSE');
    }

    // Update README
    if (!fs.existsSync(path.join(tmpDir, 'README.md'))) {
      fs.writeFileSync(path.join(tmpDir, 'README.md'),
        `# ${sfProject}\n\nMigrated from SourceForge via SF2GH Migrator.\n\nOriginal: https://sourceforge.net/projects/${sfProject}/\n`);
    }

    // Commit and push
    run('git config user.name "SF2GH Migrator"', { cwd: tmpDir });
    run('git config user.email "sf2gh@localhost"', { cwd: tmpDir });
    run('git add -A', { cwd: tmpDir });

    const status = run('git status --porcelain', { cwd: tmpDir });
    if (!status) {
      console.log('  No changes to commit.');
      return;
    }

    run('git commit -m "Restructure: extract source from SF release archives"', { cwd: tmpDir });
    console.log('  Pushing...');
    run('git push origin main', { cwd: tmpDir });
    console.log('  DONE!');

  } catch (err) {
    console.error('  ERROR: ' + err.message.split('\n')[0]);
  } finally {
    if (fs.existsSync(tmpDir)) try { fs.rmSync(tmpDir, { recursive: true }); } catch (_) {}
    if (fs.existsSync(extractDir)) try { fs.rmSync(extractDir, { recursive: true }); } catch (_) {}
  }
}

async function main() {
  console.log('SF Repo Restructuring Script');
  console.log('Owner: ' + OWNER);
  console.log('Repos: ' + SF_REPOS.length);
  console.log('');

  for (const entry of SF_REPOS) {
    await processRepo(entry);
  }

  console.log('\n=== ALL DONE ===');
}

main().catch(console.error);
