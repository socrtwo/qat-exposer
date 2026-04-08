'use strict';

const { app, BrowserWindow, shell } = require('electron');
const path = require('path');
const { execSync } = require('child_process');
const http = require('http');
const fs = require('fs');

const PORT = 3000;
let mainWindow;
let serverProcess;

// Serve the built dist/ folder over HTTPS-like local server
// Office.js requires the page to be served over HTTPS or localhost
function startStaticServer() {
  const distDir = path.join(__dirname, '..', 'dist');
  if (!fs.existsSync(distDir)) {
    // Build first if dist doesn't exist
    execSync('npm run build', { cwd: path.join(__dirname, '..'), stdio: 'inherit' });
  }

  return new Promise((resolve, reject) => {
    const handler = (req, res) => {
      let urlPath = req.url === '/' ? '/taskpane.html' : req.url;
      // Strip query strings
      urlPath = urlPath.split('?')[0];
      const filePath = path.join(distDir, urlPath);
      const ext = path.extname(filePath);
      const mimeTypes = {
        '.html': 'text/html', '.js': 'application/javascript',
        '.css': 'text/css', '.png': 'image/png', '.svg': 'image/svg+xml',
        '.json': 'application/json', '.ico': 'image/x-icon',
      };
      fs.readFile(filePath, (err, data) => {
        if (err) {
          // Fall back to taskpane.html for SPA routing
          fs.readFile(path.join(distDir, 'taskpane.html'), (err2, fallback) => {
            if (err2) { res.writeHead(404); res.end('Not found'); return; }
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(fallback);
          });
          return;
        }
        res.writeHead(200, { 'Content-Type': mimeTypes[ext] || 'application/octet-stream' });
        res.end(data);
      });
    };

    const server = http.createServer(handler);
    server.listen(PORT, () => {
      serverProcess = server;
      resolve(PORT);
    });
    server.on('error', (err) => {
      if (err.code === 'EADDRINUSE') {
        // Port already in use, just use it
        resolve(PORT);
      } else {
        reject(err);
      }
    });
  });
}

function createWindow(port) {
  mainWindow = new BrowserWindow({
    width: 420,
    height: 700,
    minWidth: 320,
    minHeight: 500,
    title: 'SuperQAT',
    backgroundColor: '#fafafa',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadURL(`http://localhost:${port}/taskpane.html`);

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('https://') || url.startsWith('http://')) {
      shell.openExternal(url);
    }
    return { action: 'deny' };
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
    if (serverProcess) serverProcess.close();
  });
}

app.whenReady().then(() => {
  startStaticServer().then(createWindow).catch((err) => {
    const { dialog } = require('electron');
    dialog.showErrorBox('Server Error', `Could not start server: ${err.message}`);
    app.quit();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    startStaticServer().then(createWindow);
  }
});
