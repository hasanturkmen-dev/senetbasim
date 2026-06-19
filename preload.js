const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
    senetBas: (veri, koordinatlar) => ipcRenderer.invoke('senet-bas', veri, koordinatlar),
    teslimFisiBas: (veri) => ipcRenderer.invoke('teslim-fisi-bas', veri),
    gecmisGetir: () => ipcRenderer.invoke('gecmis-getir') // Yeni Eklenen Satır
});