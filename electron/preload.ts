const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('electronAPI', {
  getData: () => ipcRenderer.invoke('get-data'),
  saveData: (data: any) => ipcRenderer.invoke('save-data', data),
  exportExcel: (payload: { fileName: string; buffer: number[] }) =>
    ipcRenderer.invoke('export-excel', payload),
  savePDF: (payload: { fileName: string; buffer: number[] }) =>
    ipcRenderer.invoke('save-pdf', payload),
})
