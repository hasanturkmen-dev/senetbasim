const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
    senetBas: (veri, koordinatlar) => ipcRenderer.invoke('senet-bas', veri, koordinatlar),
    teslimFisiBas: (veri) => ipcRenderer.invoke('teslim-fisi-bas', veri),
    gecmisGetir: () => ipcRenderer.invoke('gecmis-getir'),
    ortaklariGetir: () => ipcRenderer.invoke('ortaklari-getir'),
    excelOku: () => ipcRenderer.invoke('excel-oku'),
    excelKaydet: (veri) => ipcRenderer.invoke('excel-kaydet', veri),
    manuelOrtakEkle: (veri) => ipcRenderer.invoke('manuel-ortak-ekle', veri),
    gecmistenYazdir: (islemId) => ipcRenderer.invoke('gecmisten-yazdir', islemId),
    
    // --- YAZICI YENİ KANALLARI ---
    getPrinters: () => ipcRenderer.invoke('get-printers'),
    getPrintSettings: () => ipcRenderer.invoke('get-print-settings'),
    savePrintSettings: (settings) => ipcRenderer.invoke('save-print-settings', settings)
});