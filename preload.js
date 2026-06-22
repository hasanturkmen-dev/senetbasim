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
    gecmistenTeslimFisi: (islemId, printerName) => ipcRenderer.invoke('gecmisten-teslim-fisi', islemId, printerName),
    getPrinters: () => ipcRenderer.invoke('get-printers'),
    getPrintSettings: () => ipcRenderer.invoke('get-print-settings'),
    savePrintSettings: (settings) => ipcRenderer.invoke('save-print-settings', settings),
    // --- YENİ EKLENEN EXCEL DIŞA AKTARMA KANALI ---
    excelDisaAktar: (veri, dosyaAdi) => ipcRenderer.invoke('excel-disa-aktar', veri, dosyaAdi),
    // --- YENİ: GEÇMİŞ İŞLEM DETAYINI GETİRME KANALI ---
    islemDetayGetir: (islemId) => ipcRenderer.invoke('islem-detay-getir', islemId),
});