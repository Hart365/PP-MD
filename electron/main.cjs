const path = require('node:path');
const { app, BrowserWindow, dialog, shell } = require('electron');

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
    },
  });

  const isDev = !app.isPackaged;

  if (isDev) {
    mainWindow.loadURL('http://localhost:5173');
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }

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
