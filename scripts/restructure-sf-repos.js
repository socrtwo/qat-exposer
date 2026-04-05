#!/usr/bin/env node
'use strict';

/**
 * Restructure SF-migrated GitHub repos.
 *
 * For each repo:
 * 1. Clone it
 * 2. Find the best source archive (zip/tar)
 * 3. Extract it into proper directory structure
 * 4. Move remaining archives to releases/
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

const TOKEN = process.env.GITHUB_TOKEN;
const OWNER = process.env.GITHUB_OWNER || 'socrtwo';

if (!TOKEN) {
  console.error('Set GITHUB_TOKEN environment variable');
  process.exit(1);
}

// All SF-migrated repos with their known best source archive (if any).
// Empty repos and those with no archives are skipped automatically.
const SF_REPOS = [
  { name: 'autoscrshotanno-SF', bestArchive: 'screenshot-annotate.zip' },
  { name: 'catalog-of-life-SF', bestArchive: 'Catalogue-of-Life-Converter-1.0.zip' },
  { name: 'coffice2txt-SF', bestArchive: null },  // empty repo
  { name: 'corruptexcelrec-SF', bestArchive: 's2_tools_for_excel_recovery_4.0.2_source_adware_removed.zip' },
  { name: 'crrptoffcxtrctr-SF', bestArchive: 'corrupt_office_2007_extractor_delphi_7_source_code.zip' },
  { name: 'damageddocx2txt-SF', bestArchive: '__none_but_has_source__' },  // .pl and .txt already extracted
  { name: 'datarecoverfree-SF', bestArchive: 'freeware_site_script_2.0.zip' },
  { name: 'excel2ged-SF', bestArchive: null },  // empty repo
  { name: 'excelrcvryaddin-SF', bestArchive: null },  // empty repo
  { name: 'fasterposter-SF', bestArchive: 'fasterposter.com_11_29_2011.zip' },
  { name: 'ged2wiki-SF', bestArchive: 'gedcom2wiki_1.0.zip' },
  { name: 'genealogyoflife-SF', bestArchive: null },  // empty repo
  { name: 'godskingsheroes-SF', bestArchive: 'famous_family_trees.zip' },  // data repo (.ged files)
  { name: 'oorecovery-SF', bestArchive: null },  // no archive, just readme
  { name: 'pptxrecovery-SF', bestArchive: null },  // no archive, just readme
  { name: 'qatindex-SF', bestArchive: 'excel-powerpoint-qat-index.zip' },
  { name: 'quickwordrecovr-SF', bestArchive: 'savvy_docx_recovery_version_3.0_source.zip' },
  { name: 'saveofficedata-SF', bestArchive: null },  // empty repo
  { name: 'savvyoffice-SF', bestArchive: 'Savvy_Repair_for_Microsoft_Office_v1.0.22_source.zip' },
  { name: 'shiftf3-SF', bestArchive: null },  // empty repo
  { name: 'vistaprevrsrcvr-SF', bestArchive: 'previous_version_file_explorer_source_2.0.zip' },
  { name: 'whereyoubin-SF', bestArchive: 'wherehaveibeen_3.0.zip' },
  { name: 'wordrecovery-SF', bestArchive: 'Version_3.0.5-alpha-source.zip' },
  { name: 'xmltrncatorfixr-SF', bestArchive: 'xml_truncator_fixer_source.zip' },
];

function run(cmd, opts = {}) {
  console.log('  $ ' + cmd);
  return execSync(cmd, { stdio: 'pipe', timeout: 120000, ...opts }).toString().trim();
}

function findBestArchive(files) {
  // Prefer: files with "source" in name > latest version zip > any zip > any tar
  const archives = files.filter(f => /\.(zip|tar\.gz|tgz|tar\.bz2|tar|7z)$/i.test(f));
  if (archives.length === 0) return null;

  // Priority 1: has "source" in name
  const sourceArchive = archives.find(f => /source/i.test(f));
  if (sourceArchive) return sourceArchive;

  // Priority 2: largest version number or most recent
  // Sort by version-like numbers descending
  const sorted = archives.slice().sort((a, b) => {
    const va = (a.match(/(\d+[\d.]*)/g) || []).join('.');
    const vb = (b.match(/(\d+[\d.]*)/g) || []).join('.');
    return vb.localeCompare(va, undefined, { numeric: true });
  });

  return sorted[0];
}

function extractArchive(archivePath, destDir) {
  const ext = archivePath.toLowerCase();
  fs.mkdirSync(destDir, { recursive: true });

  if (ext.endsWith('.zip')) {
    run(`unzip -o -q "${archivePath}" -d "${destDir}"`);
  } else if (ext.endsWith('.tar.gz') || ext.endsWith('.tgz')) {
    run(`tar xzf "${archivePath}" -C "${destDir}"`);
  } else if (ext.endsWith('.tar.bz2')) {
    run(`tar xjf "${archivePath}" -C "${destDir}"`);
  } else if (ext.endsWith('.tar')) {
    run(`tar xf "${archivePath}" -C "${destDir}"`);
  } else if (ext.endsWith('.7z')) {
    try { run(`7z x "${archivePath}" -o"${destDir}" -y`); } catch (_) {
      console.log('    7z not available, skipping .7z file');
      return false;
    }
  }
  return true;
}

function flattenSingleSubdir(dir) {
  const entries = fs.readdirSync(dir).filter(e => e !== '.git');
  if (entries.length === 1) {
    const child = path.join(dir, entries[0]);
    if (fs.statSync(child).isDirectory()) {
      console.log('    Flattening single subdirectory: ' + entries[0]);
      const childEntries = fs.readdirSync(child);
      for (const e of childEntries) {
        const src = path.join(child, e);
        const dst = path.join(dir, e);
        if (!fs.existsSync(dst)) {
          fs.renameSync(src, dst);
        }
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
*.swo

# Build
*.o
*.obj
*.exe
*.dll
*.so
*.dylib
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

async function processRepo(repoEntry) {
  const repoName = repoEntry.name;
  const knownBestArchive = repoEntry.bestArchive;

  console.log(`\n=== Processing ${repoName} ===`);

  if (!knownBestArchive) {
    console.log('  No source archive known — skipping (empty or no archives).');
    return;
  }

  const tmpDir = path.join(os.tmpdir(), 'sf-restructure-' + repoName);
  if (fs.existsSync(tmpDir)) fs.rmSync(tmpDir, { recursive: true });

  try {
    // Clone
    const cloneUrl = `https://${TOKEN}@github.com/${OWNER}/${repoName}.git`;
    console.log('  Cloning...');
    run(`git clone "${cloneUrl}" "${tmpDir}"`);

    // List files
    const allFiles = fs.readdirSync(tmpDir).filter(f => f !== '.git');
    console.log('  Files: ' + allFiles.join(', '));

    if (allFiles.length === 0) {
      console.log('  Empty repo, skipping.');
      return;
    }

    // Check if already restructured (has src/ or releases/ directory)
    if (allFiles.includes('src') || allFiles.includes('releases')) {
      console.log('  Already restructured, skipping.');
      return;
    }

    // Separate archives from non-archives
    const archiveExts = /\.(zip|tar\.gz|tgz|tar\.bz2|tar|7z)$/i;
    const archives = allFiles.filter(f => archiveExts.test(f));

    if (archives.length === 0) {
      console.log('  No archives to extract. Adding .gitignore and LICENSE if missing...');
      // Still add standard files
    }

    // Use the pre-identified best archive
    const bestArchive = archives.includes(knownBestArchive) ? knownBestArchive : findBestArchive(archives);
    console.log('  Best archive: ' + (bestArchive || 'none'));

    // Create releases/ for other archives
    const releasesDir = path.join(tmpDir, 'releases');
    fs.mkdirSync(releasesDir, { recursive: true });

    // Move all archives to releases/
    for (const arch of archives) {
      const src = path.join(tmpDir, arch);
      const dst = path.join(releasesDir, arch);
      fs.renameSync(src, dst);
    }

    // Extract the best archive into a temp extraction dir
    if (bestArchive) {
      const extractTmp = path.join(os.tmpdir(), 'sf-extract-' + repoName);
      if (fs.existsSync(extractTmp)) fs.rmSync(extractTmp, { recursive: true });

      console.log('  Extracting: ' + bestArchive);
      const ok = extractArchive(path.join(releasesDir, bestArchive), extractTmp);
      if (ok) {
        flattenSingleSubdir(extractTmp);

        // Copy extracted files to repo root
        const extracted = fs.readdirSync(extractTmp);
        for (const item of extracted) {
          const src = path.join(extractTmp, item);
          const dst = path.join(tmpDir, item);
          if (!fs.existsSync(dst)) {
            if (fs.statSync(src).isDirectory()) {
              run(`cp -r "${src}" "${dst}"`);
            } else {
              fs.copyFileSync(src, dst);
            }
          }
        }
        console.log('  Extracted ' + extracted.length + ' item(s) to repo root.');
      }

      if (fs.existsSync(extractTmp)) fs.rmSync(extractTmp, { recursive: true });
    }

    // Add .gitignore if missing
    if (!fs.existsSync(path.join(tmpDir, '.gitignore'))) {
      fs.writeFileSync(path.join(tmpDir, '.gitignore'), GITIGNORE);
      console.log('  Added .gitignore');
    }

    // Add LICENSE if missing
    if (!fs.existsSync(path.join(tmpDir, 'LICENSE')) && !fs.existsSync(path.join(tmpDir, 'LICENSE.md'))) {
      fs.writeFileSync(path.join(tmpDir, 'LICENSE'), LICENSE_MIT);
      console.log('  Added LICENSE (MIT)');
    }

    // Git config, add, commit, push
    run('git config user.name "SF2GH Migrator"', { cwd: tmpDir });
    run('git config user.email "sf2gh@localhost"', { cwd: tmpDir });
    run('git add -A', { cwd: tmpDir });

    try {
      const status = run('git status --porcelain', { cwd: tmpDir });
      if (!status) {
        console.log('  No changes to commit.');
        return;
      }
    } catch (_) {}

    run('git commit -m "Restructure: extract source from archives, organize files\n\nExtracted source code from release archives into proper directory\nstructure. Archives moved to releases/ folder. Added .gitignore\nand LICENSE."', { cwd: tmpDir });

    console.log('  Pushing...');
    run('git push origin main', { cwd: tmpDir });
    console.log('  Done!');

  } catch (err) {
    console.error('  ERROR: ' + err.message);
  } finally {
    if (fs.existsSync(tmpDir)) {
      try { fs.rmSync(tmpDir, { recursive: true }); } catch (_) {}
    }
  }
}

async function main() {
  console.log('SF Repo Restructuring Script');
  console.log('Owner: ' + OWNER);
  console.log('Repos to process: ' + SF_REPOS.length);

  const toProcess = SF_REPOS.filter(r => r.bestArchive);
  console.log('Repos with archives to restructure: ' + toProcess.length);
  console.log('Repos skipped (empty/no archive): ' + (SF_REPOS.length - toProcess.length));

  for (const repo of SF_REPOS) {
    await processRepo(repo);
  }

  console.log('\n=== ALL DONE ===');
}

main().catch(console.error);
