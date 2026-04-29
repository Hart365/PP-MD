const path = require('node:path');
const { app, BrowserWindow, dialog, shell, ipcMain } = require('electron');

// Ensure Windows uses our App User Model ID so the taskbar shows the app icon
// rather than the generic Electron icon.
if (process.platform === 'win32') {
  app.setAppUserModelId('com.ppmd.desktop');
}

/**
 * Get the current platform identifier
 */
function getPlatformName() {
  switch (process.platform) {
    case 'win32':
      return 'windows';
    case 'darwin':
      return 'macos';
    case 'linux':
      return 'linux';
    default:
      return 'windows';
  }
}

/**
 * Get the current architecture identifier
 */
function getArchitectureName() {
  const arch = process.arch;
  return arch === 'arm64' ? 'arm64' : 'x64';
}

/**
 * Detect whether the app is running as portable or installer
 * For Windows: check if running from AppData or Program Files
 * For macOS: check if running from /Applications
 * For Linux: assume portable (AppImage)
 */
function getInstallationType() {
  const exePath = app.getPath('exe').toLowerCase();

  if (process.platform === 'win32') {
    // If in AppData/Local/Programs, it's likely an installer
    if (exePath.includes('appdata')) {
      return 'installer';
    }
    return 'portable';
  } else if (process.platform === 'darwin') {
    // macOS: check if in Applications folder
    if (exePath.includes('/applications/')) {
      return 'installer';
    }
    return 'portable';
  }

  // Linux AppImage is considered portable
  return 'portable';
}

/**
 * Set up IPC handlers for app info
 */
function setupIpcHandlers() {
  ipcMain.handle('get-app-info', async () => {
    return {
      version: app.getVersion(),
      platform: getPlatformName(),
      architecture: getArchitectureName(),
      installType: getInstallationType(),
    };
  });
}

// Ensure Windows uses our App User Model ID so the taskbar shows the app icon
// rather than the generic Electron icon.
if (process.platform === 'win32') {
  app.setAppUserModelId('com.ppmd.desktop');
}

function createMainWindow() {
  const mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1024,
    minHeight: 720,
    autoHideMenuBar: true,
    backgroundColor: '#f8f9fa',
    icon: path.join(__dirname, '..', 'build', 'icon.ico'),
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      devTools: true,
      preload: path.join(__dirname, 'preload.cjs'),
    },
  });

  const isDev = !app.isPackaged;

  if (isDev) {
    mainWindow.loadURL('http://localhost:5173');
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }

  // Inject app version into window after loading
  mainWindow.webContents.on('did-finish-load', () => {
    mainWindow.webContents.executeJavaScript(`
      window.__PPMD_VERSION__ = '${app.getVersion()}';
    `);
  });

  mainWindow.webContents.on('did-fail-load', (_event, errorCode, errorDescription, validatedURL) => {
    if (!isDev) {
      dialog.showErrorBox(
        'PP-MD failed to load',
        `Could not load application UI.\n\nURL: ${validatedURL}\nError: ${errorCode} - ${errorDescription}`,
      );
    }
  });

  mainWindow.webContents.on('render-process-gone', (_event, details) => {
    if (!isDev) {
      dialog.showErrorBox('PP-MD renderer crashed', `Reason: ${details.reason}`);
    }
  });

  // Keep external links in the default browser.
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
}

app.whenReady().then(() => {
  setupIpcHandlers();
  createMainWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
