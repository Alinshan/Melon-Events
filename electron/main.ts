const { app, BrowserWindow, ipcMain, shell, dialog } = require('electron')
const path = require('path')
const fs = require('fs')

let win: any = null

const isDev = process.env.NODE_ENV === 'development'

const DATA_PATH = path.join(app.getPath('userData'), 'melonevents_data.json')
const EXPORT_DIR = path.join(app.getPath('documents'), 'MelonEvents')

if (!fs.existsSync(EXPORT_DIR)) {
  fs.mkdirSync(EXPORT_DIR, { recursive: true })
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
  })

  if (isDev) {
    win.loadURL('http://localhost:4444')
  } else {
    win.loadFile(path.join(__dirname, '..', 'dist', 'index.html'))
  }
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
    win = null
  }
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow()
})

// --- IPC Data Handlers ---
ipcMain.handle('get-data', () => {
  if (fs.existsSync(DATA_PATH)) {
    return JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8'))
  }
  return {
    settings: { platePrice: 10, glassPrice: 5, plateCleaningPrice: 2, glassCleaningPrice: 1 },
    rentals: [],
    salaries: []
  }
})

ipcMain.handle('save-data', (_: any, data: any) => {
  fs.writeFileSync(DATA_PATH, JSON.stringify(data, null, 2))
  return true
})

ipcMain.handle('export-excel', async (_: any, payload: { fileName: string; buffer: number[] }) => {
  const { filePath } = await dialog.showSaveDialog(win, {
    title: 'Save Excel Report',
    defaultPath: path.join(app.getPath('downloads'), payload.fileName),
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
  })

  if (filePath) {
    fs.writeFileSync(filePath, Buffer.from(payload.buffer))
    shell.showItemInFolder(filePath)
    return filePath
  }
  return null
})

ipcMain.handle('save-pdf', async (_: any, payload: { fileName: string; buffer: number[] }) => {
  const { filePath } = await dialog.showSaveDialog(win, {
    title: 'Save PDF Invoice',
    defaultPath: path.join(app.getPath('downloads'), payload.fileName),
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }]
  })

  if (filePath) {
    fs.writeFileSync(filePath, Buffer.from(payload.buffer))
    shell.openPath(filePath)
    return filePath
  }
  return null
})
ipcMain.handle('get-version', () => app.getVersion())

ipcMain.handle('open-external', (_: any, url: string) => {
  shell.openExternal(url)
})
