# SuperQAT — Office Quick Access Toolbar Add-in

An Office Web Add-in for **Word**, **Excel**, and **PowerPoint** that puts **all 649 Quick Access Toolbar commands** into a single dropdown. Pick any command, click Run, and it executes on your current content.

## What it does

Opens a task pane with one scrollable dropdown containing every command you'd find in File > Options > Quick Access Toolbar > All Commands. Commands that Office.js can execute directly (bold, italic, font sizes, colors, styles, alignment, highlights, underline variants, insert table, breaks, etc.) run immediately. Commands requiring the native ribbon or keyboard shortcuts show a toast with guidance.

## Platform support

| Platform | Status |
|----------|--------|
| **Windows desktop (exe)** | Full support — Electron installer |
| **Mac desktop (dmg)** | Full support — Electron dmg |
| **Office on the Web** | Full support — sideload manifest |
| **iOS / Android** | Capacitor wrapper or web sideload |

## Quick start

### Prerequisites

- [Node.js](https://nodejs.org/) 18+
- Microsoft 365 subscription (desktop) or free account (Office on Web)

### 1. Install

```bash
cd office-quick-access-addon
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

**Manual sideload (Windows):**
1. Copy manifest to `%LocalAppData%\Microsoft\Office\16.0\Wef\`
2. Open Office app > Insert > My Add-ins > Shared Folder

### 4. Use it

Click the **SuperQAT** tab in the ribbon, then **Open SuperQAT**. Browse or scroll the dropdown, select a command, click **Run** (or double-click).

## Building releases

### Windows .exe

```bash
npm run electron:build:win
```

Output: `release/SuperQAT Setup.exe` (NSIS installer)

### Mac .dmg

```bash
npm run electron:build:mac
```

Output: `release/SuperQAT.dmg`

### Both at once

```bash
npm run electron:build:all
```

## Hosting modes

### Local server (development)

```bash
npm start
```

Manifests point to `https://localhost:3000`. The dev server must be running.

### Remote domain (production)

1. Build: `npm run build:prod`
2. Upload the `dist/` folder to your HTTPS host (e.g. superqat.app)
3. Point manifests to your domain:

```bash
npm run set-host https://superqat.app
```

Or interactively:
```bash
npm run set-host
```

To switch back to local:
```bash
npm run set-host local
```

### Mobile (iOS/Android via Capacitor)

```bash
npm run build:prod
npm run cap:add:ios       # or cap:add:android
npm run cap:sync
```

Then open the native project in Xcode / Android Studio and build.

## Project structure

```
office-quick-access-addon/
├── manifests/
│   ├── word-manifest.xml
│   ├── excel-manifest.xml
│   └── powerpoint-manifest.xml
├── src/
│   ├── assets/              # Icons (16, 32, 64, 80, 128 px)
│   ├── commands/
│   │   ├── commands.js      # Ribbon command stubs
│   │   └── commands.html
│   └── taskpane/
│       ├── taskpane.html    # Dropdown UI
│       ├── taskpane.css     # Styles
│       └── taskpane.js      # 649 commands + Office.js handlers
├── electron/
│   └── main.js              # Electron main process
├── scripts/
│   └── set-host.js          # Switch manifests between local/remote
├── webpack.config.js        # Dev build
├── webpack.prod.js          # Production build
├── electron-builder.yml     # Windows/Mac packaging config
├── capacitor.config.ts      # iOS/Android config
├── package.json
└── README.md
```

## Deploying to production

1. `npm run build:prod`
2. Upload `dist/` to an HTTPS host
3. `npm run set-host https://yourdomain.app`
4. Distribute manifests via:
   - **Microsoft 365 Admin Center** (org-wide)
   - **AppSource** (public marketplace)
   - **SharePoint App Catalog**

## License

MIT
