/// <reference types="vite/client" />

interface Window {
  electronAPI: {
    getData: () => Promise<any>
    saveData: (data: any) => Promise<boolean>
    exportExcel: (payload: { fileName: string; buffer: Uint8Array }) => Promise<string>
    savePDF: (payload: { fileName: string; buffer: Uint8Array }) => Promise<string>
    onMessage: (callback: any) => void
  }
}
