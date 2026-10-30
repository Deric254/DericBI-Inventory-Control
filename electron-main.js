/**
 * St. Anne Mission Hospital — ICT Command Centre
 * Electron Main Process
 */

const { app, BrowserWindow, ipcMain, shell, dialog } = require('electron');
const { spawn, execFileSync }                 = require('child_process');
const path = require('path');
const http = require('http');
const fs   = require('fs');

// ── Paths ─────────────────────────────────────────────────────────────────────
const IS_PACKED = app.isPackaged;

function resPath(...parts) {
  return IS_PACKED
    ? path.join(process.resourcesPath, ...parts)
    : path.join(__dirname, ...parts);
}

const PYTHON_DIR = resPath('python');   // bundled Python runtime
const SRC_DIR    = resPath('src');      // main.py + ui.html
const ASSETS_DIR = resPath('assets');  // seed files

// Data directory:
// - Installed (NSIS): AppData/Roaming/ICT Command Centre (writable, survives updates)
// - Dev (start.bat):  project/data-dev
function getDataDir() {
  if (!IS_PACKED) return path.join(__dirname, 'data-dev');
  try { return path.join(app.getPath('userData'), '..', 'ICT Command Centre'); }
  catch (_) {
    try { return path.join(app.getPath('home'), 'AppData', 'Roaming', 'ICT Command Centre'); }
    catch (_) { return path.join(path.dirname(app.getPath('exe')), 'data'); }
  }
}
const DATA_DIR = getDataDir();

fs.mkdirSync(DATA_DIR, { recursive: true });

// Seed data files on first run
['ICT_MASTER.xlsx', 'auth.json', 'logo.png'].forEach(name => {
  const dst = path.join(DATA_DIR, name);
  const src = path.join(ASSETS_DIR, name);
  if (!fs.existsSync(dst) && fs.existsSync(src)) fs.copyFileSync(src, dst);
});

// ── Python discovery ──────────────────────────────────────────────────────────
function findPython() {
  const candidates = [
    // 1. Explicitly set by start.bat via env var (most reliable in dev)
    process.env.ICT_PYTHON,
    // 2. Bundled Python relative to app (production or if start.bat was used)
    path.join(PYTHON_DIR, 'python.exe'),
    path.join(PYTHON_DIR, 'python'),
    path.join(PYTHON_DIR, 'python3'),
    // 3. System Python fallback
    'python', 'python3', 'py',
  ].filter(Boolean);

  for (const c of candidates) {
    if (path.isAbsolute(c) && !fs.existsSync(c)) continue;
    try {
      const out = execFileSync(c, ['-c', 'import sys;print(sys.version_info.major)'], {
        timeout: 4000, stdio: ['ignore', 'pipe', 'pipe'],
      }).toString().trim();
      if (parseInt(out, 10) >= 3) {
        console.log(`[py] Using Python: ${c}`);
        return c;
      }
    } catch (_) { /* try next */ }
  }

  console.warn('[py] No working Python found — will fail at spawn');
  return 'python';
}

// ── Python backend ────────────────────────────────────────────────────────────
let pyProcess  = null;
let serverPort = null;
let mainWindow = null;

function startPython(onReady) {
  const exe    = findPython();
  const script = path.join(SRC_DIR, 'main.py');

  pyProcess = spawn(exe, [script, '--electron'], {
    cwd: SRC_DIR,
    env: {
      ...process.env,
      ICT_DATA_DIR:            DATA_DIR,
      ICT_BUNDLE_DIR:          SRC_DIR,
      PYTHONUNBUFFERED:        '1',
      PYTHONDONTWRITEBYTECODE: '1',
    },
    stdio: ['ignore', 'pipe', 'pipe'],
  });

  pyProcess.stdout.on('data', data => {
    const text = data.toString();
    const m = text.match(/ELECTRON_PORT=(\d+)/);
    if (m && !serverPort) {
      serverPort = parseInt(m[1], 10);
      console.log(`[py] Ready on port ${serverPort}`);
      onReady(serverPort);
    }
  });

  pyProcess.stderr.on('data', d => process.stderr.write('[py] ' + d));
  pyProcess.on('close', code => { console.log(`[py] Exited (${code})`); pyProcess = null; });
}

function killPython() {
  if (pyProcess) { pyProcess.kill(); pyProcess = null; }
}

// ── IPC ───────────────────────────────────────────────────────────────────────
ipcMain.handle('py-call', (_e, method, args) => new Promise(resolve => {
  if (!serverPort) return resolve(JSON.stringify({ ok: false, error: 'Backend not ready' }));
  const body = JSON.stringify({ args: args || [] });
  const req = http.request({
    hostname: '127.0.0.1', port: serverPort,
    path: `/api/${encodeURIComponent(method)}`, method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
  }, res => {
    let raw = '';
    res.on('data', c => raw += c);
    res.on('end', () => resolve(raw));
  });
  req.on('error', e => resolve(JSON.stringify({ ok: false, error: e.message })));
  req.write(body); req.end();
}));

ipcMain.handle('open-file', (_e, filePath) =>
  shell.openPath(filePath).then(() => ({ ok: true })).catch(e => ({ ok: false, error: e.message }))
);

ipcMain.on('refocus-window', () => {
  if (!mainWindow) return;
  // Blur then focus replicates exactly what minimize/restore does at OS level
  mainWindow.blur();
  setTimeout(() => {
    if (mainWindow) {
      mainWindow.focus();
      setTimeout(() => {
        if (mainWindow) mainWindow.webContents.focus();
      }, 50);
    }
  }, 80);
});

// ── Window ────────────────────────────────────────────────────────────────────
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440, height: 880, minWidth: 1024, minHeight: 680,
    backgroundColor: '#0d1117',
    title: 'St. Anne Mission Hospital — ICT Command Centre',
    icon: resPath('assets', 'icon.ico'),
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  mainWindow.setMenuBarVisibility(false);
  mainWindow.on('ready-to-show', () => mainWindow.show());
  mainWindow.on('closed', () => { mainWindow = null; killPython(); });

  // Every time the page finishes loading (including after logout state reset),
  // blur then focus the window so the webview gets OS-level keyboard focus.
  // This is exactly what minimize/restore does — we replicate it automatically.
  mainWindow.webContents.on('did-finish-load', () => {
    if (mainWindow) {
      mainWindow.blur();
      setTimeout(() => {
        if (mainWindow) {
          mainWindow.focus();
          mainWindow.webContents.focus();
        }
      }, 80);
    }
  });
}

// ── Auto-update (GitHub Releases) ─────────────────────────────────────────────
// Installed builds only. Never touches DATA_DIR except to copy a snapshot of it.
// Every failure path is non-fatal: the app must always start and run normally.
const UPDATE_CHECK_DELAY_MS    = 15 * 1000;
const UPDATE_CHECK_INTERVAL_MS = 4 * 60 * 60 * 1000;

function showBox(opts) {
  return (mainWindow && !mainWindow.isDestroyed())
    ? dialog.showMessageBox(mainWindow, opts)
    : dialog.showMessageBox(opts);
}

// Copy live data to <data>/backups/pre-update/<ts>/ (kept: newest 3).
// Sub-folder is invisible to the in-app backup list and its 5-file rotation.
function snapshotBeforeUpdate() {
  try {
    const root = path.join(DATA_DIR, 'backups', 'pre-update');
    const ts   = new Date().toISOString().replace(/[:.]/g, '-');
    const dir  = path.join(root, ts);
    fs.mkdirSync(dir, { recursive: true });
    ['ICT_MASTER.xlsx', 'auth.json', 'config.json'].forEach(name => {
      const src = path.join(DATA_DIR, name);
      if (fs.existsSync(src)) fs.copyFileSync(src, path.join(dir, name));
    });
    fs.readdirSync(root).sort().slice(0, -3).forEach(old =>
      fs.rmSync(path.join(root, old), { recursive: true, force: true }));
  } catch (e) { console.warn('[update] snapshot skipped:', e.message); }
}

function setupAutoUpdate() {
  if (!IS_PACKED) return;                       // never in dev / start.bat

  let autoUpdater;
  try { ({ autoUpdater } = require('electron-updater')); }
  catch (e) { console.warn('[update] updater unavailable:', e.message); return; }

  autoUpdater.autoDownload         = false;     // user decides
  autoUpdater.autoInstallOnAppQuit = true;      // "Later" still installs on next close
  autoUpdater.allowPrerelease      = false;

  let prompted = false;                         // one prompt per session
  let userDownloading = false;

  autoUpdater.on('update-available', async info => {
    if (prompted) return;
    prompted = true;
    const { response } = await showBox({
      type: 'info', title: 'Update available', noLink: true,
      message: `Version ${info.version} is available`,
      detail: `You are on version ${app.getVersion()}.\n\nYour inventory data and PIN are not affected by updates.`,
      buttons: ['Download', 'Later'], defaultId: 0, cancelId: 1,
    });
    if (response !== 0) return;
    userDownloading = true;
    autoUpdater.downloadUpdate().catch(e => console.warn('[update] download:', e.message));
  });

  autoUpdater.on('download-progress', p => {
    if (mainWindow && !mainWindow.isDestroyed()) mainWindow.setProgressBar(p.percent / 100);
  });

  autoUpdater.on('update-downloaded', async info => {
    userDownloading = false;
    if (mainWindow && !mainWindow.isDestroyed()) mainWindow.setProgressBar(-1);
    snapshotBeforeUpdate();
    const { response } = await showBox({
      type: 'info', title: 'Update ready', noLink: true,
      message: `Version ${info.version} is ready to install`,
      detail: 'The app will close, update, and reopen automatically.\nChoose Later to install when you next close the app.',
      buttons: ['Restart and install', 'Later'], defaultId: 0, cancelId: 1,
    });
    if (response !== 0) return;
    snapshotBeforeUpdate();                     // capture anything changed since download
    killPython();                               // release file locks before installer runs
    setImmediate(() => autoUpdater.quitAndInstall(true, true));
  });

  autoUpdater.on('error', e => {
    console.warn('[update] error:', e && e.message);
    if (mainWindow && !mainWindow.isDestroyed()) mainWindow.setProgressBar(-1);
    if (userDownloading) {                      // only speak up if the user asked for it
      userDownloading = false;
      dialog.showErrorBox('Update failed',
        'The update could not be downloaded. Your data is safe and the current version keeps working.\nPlease try again later.');
    }
  });

  const check = () => autoUpdater.checkForUpdates().catch(e => console.warn('[update] check:', e.message));
  setTimeout(check, UPDATE_CHECK_DELAY_MS);
  setInterval(check, UPDATE_CHECK_INTERVAL_MS);
}

// ── Boot ──────────────────────────────────────────────────────────────────────
app.whenReady().then(() => {
  createWindow();
  setupAutoUpdate();
  startPython(port => {
    const tryLoad = () => {
      http.get(`http://127.0.0.1:${port}/`, res => {
        if (res.statusCode === 200 && mainWindow) mainWindow.loadURL(`http://127.0.0.1:${port}/`);
      }).on('error', () => setTimeout(tryLoad, 300));
    };
    tryLoad();
  });
});

app.on('window-all-closed', () => { killPython(); if (process.platform !== 'darwin') app.quit(); });
app.on('before-quit', killPython);
app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
