# SuperQAT — Office Quick Access Toolbar Add-in

<!--PAGES_LINK_BANNER-->
> 🌐 **Live page:** [https://socrtwo.github.io/qat-exposer/](https://socrtwo.github.io/qat-exposer/)  
> 📦 **Releases:** [github.com/socrtwo/qat-exposer/releases](https://github.com/socrtwo/qat-exposer/releases)
<!--/PAGES_LINK_BANNER-->

An Office Add-in for **Word**, **Excel**, and **PowerPoint** that puts **all 2,199 Office 365 commands** into a single searchable dropdown. Pick any command, click Run, and it executes on your current content.

## Download

Grab the latest release from [GitHub Releases](https://github.com/socrtwo/qat-exposer/releases):

| File | What it is |
|------|-----------|
| **SuperQAT Setup.exe** | Windows desktop app (installer) |
| **SuperQAT.dmg** | Mac desktop app |
| **SuperQAT-web.zip** | Web add-in (host on any HTTPS server) |
| **SuperQAT-manifests.zip** | Office manifests for sideloading |
| **SuperQAT-VBA-source.zip** | VBA source for COM add-ins (Word .dotm, Excel .xlam, PowerPoint .ppam) |

## What it does

Opens a task pane with a **search box**, **category filter**, and a scrollable list of every Office 365 ribbon command. There are two versions:

### Web Add-in (Office.js)
- **2,199 commands** across Word (1,334), Excel (1,080), and PowerPoint (769)
- **~80 commands execute directly** via Office.js (marked with ✓) — bold, italic, font sizes, colors, styles, alignment, breaks, tables, etc.
- Other commands show **ribbon guidance** — tells you exactly which tab and button to click
- Auto-filters to only show commands available in the current app
- Works in desktop Office, Office on the web, and as a standalone Electron app

### COM Add-ins (VBA)
- **Execute ALL commands directly** — no ribbon guidance needed
- Uses `Application.CommandBars.ExecuteMso()` to trigger any built-in ribbon command
- Desktop only (Windows and Mac)
- See [com-addin/README.md](com-addin/README.md) for installation instructions

## Platform support

| Platform | Web Add-in | COM Add-in |
|----------|-----------|------------|
| **Windows desktop** | ✓ Electron .exe or sideload | ✓ .dotm / .xlam / .ppam |
| **Mac desktop** | ✓ Electron .dmg or sideload | ✓ .dotm / .xlam / .ppam |
| **Office on the Web** | ✓ Sideload manifest | ✗ Not supported |
| **iOS / Android** | ✓ Capacitor wrapper | ✗ Not supported |

## Quick start

### Prerequisites

- [Node.js](https://nodejs.org/) 18+
- Microsoft 365 subscription (desktop) or free account (Office on Web)

### 1. Install

```bash
cd qat-exposer
npm install
```

### 2. Start dev server

```bash
npm start
```

Starts a local HTTPS server at `https://localhost:3000`.

### 3. Sideload the add-in

**Desktop (Windows/Mac):**
```bash
npm run sideload:word       # or :excel or :powerpoint
```

**Office on the Web:**
1. Open Word/Excel/PowerPoint at office.com
2. Insert > Office Add-ins > Upload My Add-in
3. Upload the manifest from `manifests/` (word, excel, or powerpoint)

### 4. Use it

Click the **SuperQAT** tab in the ribbon, then **Open SuperQAT**. Type in the search box to filter commands. Select a command and click **Run** (or double-click).

## Building releases

### Windows .exe

```bash
npm run electron:build:win
```

### Mac .dmg

```bash
npm run electron:build:mac
```

### Both at once

```bash
npm run electron:build:all
```

### COM Add-ins (.dotm, .xlam, .ppam)

On a Windows PC with Office installed:

```powershell
cd com-addin
.\build-all.ps1
```

Output goes to `com-addin/build/`. See [com-addin/README.md](com-addin/README.md) for details.

### GitHub Releases (automated)

Tag a commit and push — GitHub Actions builds everything:

```bash
git tag v3.0.2
git push origin v3.0.2
```

## Hosting modes

### Local server (development)

```bash
npm start
```

Manifests point to `https://localhost:3000`. The dev server must be running.

### Remote domain (production)

```bash
npm run build:prod
npm run set-host https://yourdomain.app
```

Upload the `dist/` folder to your HTTPS host. To switch back to local:

```bash
npm run set-host local
```

## Project structure

```
qat-exposer/
├── manifests/                 # Office Add-in manifests (Word, Excel, PowerPoint)
├── src/
│   ├── assets/                # Icons (16–512px, .ico)
│   ├── commands/              # Ribbon command stubs
│   ├── data/                  # 2,199 commands from official Microsoft 365 control IDs
│   │   ├── command-map.json   # Full command data (name, type, tab, group, apps)
│   │   ├── commands-slim.json # Compact version for webpack bundling (64KB)
│   │   └── *-commands.json    # Per-app command lists
│   └── taskpane/              # Main UI (search, filter, command list)
├── com-addin/                 # VBA COM add-ins for Word, Excel, PowerPoint
│   ├── word/                  # 1,334 commands
│   ├── excel/                 # 1,080 commands
│   ├── powerpoint/            # 769 commands
│   ├── build-all.ps1          # One-click PowerShell build script
│   └── README.md              # Installation instructions
├── electron/                  # Electron desktop app wrapper
├── scripts/                   # Build and hosting utilities
├── .github/workflows/         # CI/CD (builds .exe, .dmg, web zip on tag)
├── webpack.config.js          # Dev build
├── webpack.prod.js            # Production build
├── electron-builder.yml       # Windows/Mac packaging config
├── capacitor.config.ts        # iOS/Android config
└── package.json
```

## Command data source

All 2,199 commands come from Microsoft's official control ID files:
[OfficeDev/office-fluent-ui-command-identifiers](https://github.com/OfficeDev/office-fluent-ui-command-identifiers)
(Microsoft 365 Current Channel)

## License

MIT