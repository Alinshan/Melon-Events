"use strict";
const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('electronAPI', {
    getData: () => ipcRenderer.invoke('get-data'),
    saveData: (data) => ipcRenderer.invoke('save-data', data),
    exportExcel: (payload) => ipcRenderer.invoke('export-excel', payload),
    savePDF: (payload) => ipcRenderer.invoke('save-pdf', payload),
    openExternal: (url) => ipcRenderer.invoke('open-external', url),
    getVersion: () => ipcRenderer.invoke('get-version'),
    checkForUpdates: () => ipcRenderer.invoke('check-for-updates'),
});
