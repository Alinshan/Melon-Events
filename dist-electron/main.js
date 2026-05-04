"use strict";
const { app, BrowserWindow, ipcMain, shell, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { autoUpdater } = require('electron-updater');
// Configure AutoUpdater
autoUpdater.autoDownload = false; // Let user decide
autoUpdater.autoInstallOnAppQuit = true;
let win = null;
let isManualCheck = false;
const isDev = process.env.NODE_ENV === 'development';
const DATA_PATH = path.join(app.getPath('userData'), 'melonevents_data.json');
const EXPORT_DIR = path.join(app.getPath('documents'), 'MelonEvents');
if (!fs.existsSync(EXPORT_DIR)) {
    fs.mkdirSync(EXPORT_DIR, { recursive: true });
}
function createWindow() {
    win = new BrowserWindow({
        width: 1280,
        height: 800,
        minWidth: 1000,
        minHeight: 600,
        title: 'Melon Events',
        icon: path.join(__dirname, '..', 'public', 'icon.png'),
        autoHideMenuBar: true,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true,
        },
    });
    if (isDev) {
        win.loadURL('http://localhost:4444');
    }
    else {
        win.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
    }
}
app.whenReady().then(() => {
    createWindow();
    if (!isDev) {
        autoUpdater.checkForUpdates();
    }
});
app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
        win = null;
    }
});
app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0)
        createWindow();
});
// --- IPC Data Handlers ---
ipcMain.handle('get-data', () => {
    if (fs.existsSync(DATA_PATH)) {
        return JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8'));
    }
    return {
        settings: { platePrice: 10, glassPrice: 5, plateCleaningPrice: 2, glassCleaningPrice: 0.40 },
        rentals: [],
        salaries: []
    };
});
ipcMain.handle('save-data', (_, data) => {
    fs.writeFileSync(DATA_PATH, JSON.stringify(data, null, 2));
    return true;
});
ipcMain.handle('export-excel', async (_, payload) => {
    const { filePath } = await dialog.showSaveDialog(win, {
        title: 'Save Excel Report',
        defaultPath: path.join(app.getPath('downloads'), payload.fileName),
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
    });
    if (filePath) {
        fs.writeFileSync(filePath, Buffer.from(payload.buffer));
        shell.showItemInFolder(filePath);
        return filePath;
    }
    return null;
});
ipcMain.handle('save-pdf', async (_, payload) => {
    const { filePath } = await dialog.showSaveDialog(win, {
        title: 'Save PDF Invoice',
        defaultPath: path.join(app.getPath('downloads'), payload.fileName),
        filters: [{ name: 'PDF Files', extensions: ['pdf'] }]
    });
    if (filePath) {
        fs.writeFileSync(filePath, Buffer.from(payload.buffer));
        shell.openPath(filePath);
        return filePath;
    }
    return null;
});
ipcMain.handle('get-version', () => app.getVersion());
ipcMain.handle('open-external', (_, url) => {
    shell.openExternal(url);
});
// --- AutoUpdate Logic ---
ipcMain.handle('check-for-updates', async () => {
    if (isDev)
        return { status: 'dev', message: 'Update check skipped in development mode.' };
    try {
        isManualCheck = true;
        const result = await autoUpdater.checkForUpdates();
        return { status: 'checking', result };
    }
    catch (error) {
        isManualCheck = false;
        return { status: 'error', message: error.message };
    }
});
autoUpdater.on('update-available', (info) => {
    dialog.showMessageBox(win, {
        type: 'info',
        title: 'Update Available',
        message: `A new version (${info.version}) is available. Do you want to download it now?`,
        buttons: ['Download', 'Later']
    }).then((result) => {
        isManualCheck = false;
        if (result.response === 0) {
            autoUpdater.downloadUpdate();
        }
    });
});
autoUpdater.on('update-downloaded', (info) => {
    dialog.showMessageBox(win, {
        type: 'info',
        title: 'Update Ready',
        message: `Version ${info.version} has been downloaded and is ready to install.`,
        buttons: ['Install and Relaunch', 'Later']
    }).then((result) => {
        if (result.response === 0) {
            autoUpdater.quitAndInstall();
        }
    });
});
autoUpdater.on('update-not-available', () => {
    if (isManualCheck) {
        dialog.showMessageBox(win, {
            type: 'info',
            title: 'Up to Date',
            message: 'You are already using the latest version of Melon Events.',
            buttons: ['OK']
        });
        isManualCheck = false;
    }
});
autoUpdater.on('error', (err) => {
    console.error('Update error:', err);
    isManualCheck = false;
});
