const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const { PDFDocument, rgb } = require('pdf-lib');
const fontkit = require('@pdf-lib/fontkit');
const xlsx = require('xlsx');
const db = require('./db.js');
const { print } = require('pdf-to-printer');

// --- HASSAS METİN KAYDIRMA (TEXT WRAP) ZEKI MOTORU ---
function drawWrappedText(page, text, x, y, maxWidth, font, fontSize, lineSpacing = 13) {
    const words = text.split(' ');
    let line = '';
    let currentY = y;

    for (let n = 0; n < words.length; n++) {
        let testLine = line + words[n] + ' ';
        let testWidth = font.widthOfTextAtSize(testLine, fontSize);
        if (testWidth > maxWidth && n > 0) {
            page.drawText(line.trim(), { x, y: currentY, size: fontSize, font, color: rgb(0, 0, 0) });
            line = words[n] + ' ';
            currentY -= lineSpacing;
        } else {
            line = testLine;
        }
    }
    page.drawText(line.trim(), { x, y: currentY, size: fontSize, font, color: rgb(0, 0, 0) });
}

function ayEkle(baslangicTarihi, eklenecekAy) {
    const basTarih = new Date(baslangicTarihi);
    const basGun = basTarih.getDate();
    const basAy = basTarih.getMonth();
    const basYil = basTarih.getFullYear();
    const baslangicAydakiGunSayisi = new Date(basYil, basAy + 1, 0).getDate();
    const baslangicAySonuMu = (basGun === baslangicAydakiGunSayisi);
    let yeniTarih = new Date(basYil, basAy + eklenecekAy, 1);
    const yeniYil = yeniTarih.getFullYear();
    const yeniAy = yeniTarih.getMonth();
    const aydakiGunSayisi = new Date(yeniYil, yeniAy + 1, 0).getDate();
    let sonGun = baslangicAySonuMu ? aydakiGunSayisi : Math.min(basGun, aydakiGunSayisi);
    return new Date(yeniYil, yeniAy, sonGun);
}

function sayiyiYaziyaCevir(hamTutar) {
    if (!hamTutar) return "";
    let sayi = hamTutar.toString().replace(/\./g, '').replace(',', '.');
    if (isNaN(sayi) || sayi == 0) return "SIFIR";
    const birler = ["", "BİR", "İKİ", "ÜÇ", "DÖRT", "BEŞ", "ALTI", "YEDİ", "SEKİZ", "DOKUZ"];
    const onlar = ["", "ON", "YİRMİ", "OTUZ", "KIRK", "ELLİ", "ALTMIŞ", "YETMİŞ", "SEKSEN", "DOKSAN"];
    const binler = ["", "BİN", "MİLYON", "MİLYAR"];
    let str = parseFloat(sayi).toFixed(2).toString();
    let parcalar = str.split('.');
    let tl = parcalar[0]; let kurus = parcalar[1];
    function ucluOku(n) {
        let yuz = Math.floor(n / 100); let on = Math.floor((n % 100) / 10); let bir = n % 10;
        let t = "";
        if (yuz == 1) t += "YÜZ"; else if (yuz > 1) t += birler[yuz] + "YÜZ";
        t += onlar[on]; t += birler[bir]; return t;
    }
    let yaziTL = "";
    let grupSayisi = Math.ceil(tl.length / 3);
    tl = tl.padStart(grupSayisi * 3, '0');
    for (let i = 0; i < grupSayisi; i++) {
        let grupDegeri = parseInt(tl.substring(i * 3, i * 3 + 3));
        let grupAdi = binler[grupSayisi - 1 - i];
        if (grupDegeri > 0) {
            if (grupAdi === "BİN" && grupDegeri === 1) yaziTL += "BİN"; 
            else yaziTL += ucluOku(grupDegeri) + grupAdi; 
        }
    }
    if (yaziTL === "") yaziTL = "SIFIR";
    let yaziKurus = "";
    if (parseInt(kurus) > 0) yaziKurus = ", " + ucluOku(parseInt(kurus)) + " KURUŞ"; 
    return "# " + yaziTL + yaziKurus + " #";
}

// --- MODERN TESLİM FİŞİ ÇİZİM MOTORU (Ortak Kullanım İçin) ---
async function teslimFisiSayfasiOlustur(pdfDoc, ozelFont, fisVerisi) {
    const a4Width = 595.28; const a4Height = 841.89; // Tam A4 Formatı
    const page = pdfDoc.addPage([a4Width, a4Height]);

    let logoImage = null;
    try { 
        if (fs.existsSync(path.join(app.getAppPath(), 'logo.png'))) {
            logoImage = await pdfDoc.embedPng(fs.readFileSync(path.join(app.getAppPath(), 'logo.png'))); 
        } 
    } catch (e) {}

    if (logoImage) {
        const scaled = logoImage.scaleToFit(80, 65);
        page.drawImage(logoImage, { x: 30, y: 815 - scaled.height, width: scaled.width, height: scaled.height });
    }

    page.drawText('S.S. TÜM OPTİSYENLER VE GÖZLÜKÇÜLER TEMİN TEVZİ KOOPERATİFİ', { x: 120, y: 795, size: 12, font: ozelFont, color: rgb(0.05, 0.1, 0.15) });
    page.drawText('Fevzipaşa Mahallesi 847 Sokak No: 4/A Konak - İZMİR', { x: 120, y: 780, size: 9, font: ozelFont, color: rgb(0.3, 0.3, 0.3) });
    page.drawText('Tel: 0232 489 89 62  |  Fax: 0232 489 39 62', { x: 120, y: 767, size: 9, font: ozelFont, color: rgb(0.3, 0.3, 0.3) });
    page.drawText('Mersis No: 0875 0359 0980 00 15  |  Tic. Sicil: 155588  |  Kemeraltı V.D.', { x: 120, y: 754, size: 9, font: ozelFont, color: rgb(0.3, 0.3, 0.3) });

    page.drawRectangle({ x: 395, y: 705, width: 170, height: 26, color: rgb(0.05, 0.1, 0.15) });
    page.drawText('SENET ALINDI BELGESİ', { x: 415, y: 714, size: 11, font: ozelFont, color: rgb(1, 1, 1) });
    
    const duzTarih = fisVerisi.duzenlemeTarihi.split('-').reverse().join('.');
    page.drawText(`Düzenleme Tarihi: ${duzTarih}`, { x: 435, y: 690, size: 9, font: ozelFont, color: rgb(0.3, 0.3, 0.3) });

    let currentY = 650; const startX = 30; const rowHeight = 22;
    const cols = [
        { title: 'SIRA', w: 40, x: startX }, { title: 'DÜZENLEME', w: 85, x: startX + 40 },
        { title: 'TİCARİ ÜNVAN', w: 225, x: startX + 125 }, { title: 'VADE', w: 85, x: startX + 350 },
        { title: 'TUTAR', w: 100, x: startX + 435 }
    ];

    page.drawRectangle({ x: startX, y: currentY, width: 535, height: rowHeight, color: rgb(0.95, 0.96, 0.98), borderColor: rgb(0.7, 0.7, 0.7), borderWidth: 1 });
    cols.forEach(col => {
        page.drawText(col.title, { x: col.x + 8, y: currentY + 7, size: 9, font: ozelFont, color: rgb(0.2, 0.2, 0.2) });
        if(col.x > startX) page.drawLine({ start: { x: col.x, y: currentY }, end: { x: col.x, y: currentY + rowHeight }, thickness: 1, color: rgb(0.7, 0.7, 0.7) });
    });
    currentY -= rowHeight;

    const senetTutariFloat = parseFloat(String(fisVerisi.tutar).replace(/\./g, '').replace(',', '.'));
    const formatliSenetTutari = new Intl.NumberFormat('tr-TR', { minimumFractionDigits: 2 }).format(senetTutariFloat) + ' TL';
    const formatliToplam = new Intl.NumberFormat('tr-TR', { minimumFractionDigits: 2 }).format(senetTutariFloat * fisVerisi.taksitSayisi) + ' TL';

    for (let i = 0; i < 15; i++) {
        page.drawRectangle({ x: startX, y: currentY, width: 535, height: rowHeight, borderColor: rgb(0.7, 0.7, 0.7), borderWidth: 1 });
        cols.forEach(col => {
            if(col.x > startX) page.drawLine({ start: { x: col.x, y: currentY }, end: { x: col.x, y: currentY + rowHeight }, thickness: 1, color: rgb(0.7, 0.7, 0.7) });
        });

        page.drawText((i + 1).toString(), { x: cols[0].x + 15, y: currentY + 7, size: 9, font: ozelFont });

        if (i < fisVerisi.taksitSayisi) {
            const tTarih = ayEkle(fisVerisi.vade, i);
            const formatliVade = `${String(tTarih.getDate()).padStart(2, '0')}.${String(tTarih.getMonth() + 1).padStart(2, '0')}.${tTarih.getFullYear()}`;

            page.drawText(duzTarih, { x: cols[1].x + 8, y: currentY + 7, size: 9, font: ozelFont });
            page.drawText(formatliVade, { x: cols[3].x + 8, y: currentY + 7, size: 9, font: ozelFont });
            page.drawText(formatliSenetTutari, { x: cols[4].x + 10, y: currentY + 7, size: 9, font: ozelFont });

            let unvanFontSize = fisVerisi.unvan.length > 45 ? 6.5 : 8;
            let unvanY = fisVerisi.unvan.length > 45 ? currentY + 13 : currentY + 7;
            let unvanLineSpacing = fisVerisi.unvan.length > 45 ? 8 : 11;

            drawWrappedText(page, fisVerisi.unvan, cols[2].x + 6, unvanY, 210, ozelFont, unvanFontSize, unvanLineSpacing);
        }
        currentY -= rowHeight;
    }

    currentY -= 40;
    page.drawText('TESLİM EDEN', { x: 70, y: currentY + 15, size: 9, font: ozelFont, color: rgb(0.4, 0.4, 0.4) });
    page.drawText('KAŞE - İMZA', { x: 70, y: currentY, size: 10, font: ozelFont, color: rgb(0.1, 0.1, 0.1) });

    page.drawText('TESLİM ALAN', { x: 230, y: currentY + 15, size: 9, font: ozelFont, color: rgb(0.4, 0.4, 0.4) });
    page.drawText('KAŞE - İMZA', { x: 230, y: currentY, size: 10, font: ozelFont, color: rgb(0.1, 0.1, 0.1) });

    page.drawRectangle({ x: 415, y: currentY - 10, width: 150, height: 40, color: rgb(0.95, 0.96, 0.98), borderColor: rgb(0.7, 0.7, 0.7), borderWidth: 1, borderRadius: 4 });
    page.drawText('GENEL TOPLAM', { x: 450, y: currentY + 16, size: 8, font: ozelFont, color: rgb(0.4, 0.4, 0.4) });
    page.drawText(formatliToplam, { x: 435, y: currentY - 2, size: 12, font: ozelFont, color: rgb(0.05, 0.1, 0.15) });
}

// --- HAYALET YAZDIRMA (GHOST PRINTING) YARDIMCISI ---
async function hayaletYazdirVeYokEt(pdfDoc, printerName, dosyaOnEki) {
    const tempDir = app.getPath('temp');
    const dosyaYolu = path.join(tempDir, `${dosyaOnEki}_${Date.now()}_${Math.random().toString(36).substr(2, 5)}.pdf`);
    
    fs.writeFileSync(dosyaYolu, await pdfDoc.save());
    
    if (printerName) {
        // await KALDIRILDI! Emri ver ve arkana bakmadan devam et (Hızı uçuran sihirli dokunuş)
        print(dosyaYolu, { printer: printerName }).catch((e) => { 
            console.error("Yazdırma Hatası:", e); 
        });
    }
    
    // Arka planda yazdırma sürerken Windows dosyayı kilitlemesin diye imha süresini 15 saniyeye çıkarttık
    setTimeout(() => { try { fs.unlinkSync(dosyaYolu); } catch(err){} }, 15000);
}

function createWindow () {
  const mainWindow = new BrowserWindow({
    width: 1280, height: 800, title: "Senet Basım Programı",
    webPreferences: { preload: path.join(__dirname, 'preload.js'), nodeIntegration: false, contextIsolation: true }
  });
  mainWindow.loadFile('index.html');
}

app.whenReady().then(() => { createWindow(); app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); }); });
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); }); 

// ============================================================
// IPC ARKA PLAN MOTORLARI (MERMİ GİBİ HIZLI ÇOK SAYFALI SİSTEM)
// ============================================================

ipcMain.handle('islem-detay-getir', async (event, islemId) => {
    return new Promise((resolve) => {
        const sql = `SELECT s.*, o.ticari_unvan FROM senet_gecmisi s LEFT JOIN ortaklar o ON s.ortak_id = o.id WHERE s.id = ?`;
        db.get(sql, [islemId], (err, row) => {
            if (err || !row) resolve({ success: false, error: "Kayıt bulunamadı." });
            else resolve({ success: true, data: row });
        });
    });
});

// SIFIRDAN YENİ SENET KESME MOTORU
ipcMain.handle('senet-bas', async (event, veri, koordinatlar) => {
    return new Promise(async (resolve) => {
        try {
            const senetTutariFloat = parseFloat(veri.tutar.replace(/\./g, '').replace(',', '.'));
            const toplamTutarFloat = senetTutariFloat * veri.taksitSayisi; 
            const koordinatlarJSON = JSON.stringify(koordinatlar);

            db.run(`INSERT INTO senet_gecmisi (ortak_id, kampanya_id, toplam_tutar, taksit_sayisi, baslangic_vadesi, koordinatlar, senet_koku, tekil_tutar) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`, 
            [veri.ortakId || 1, 1, toplamTutarFloat, veri.taksitSayisi, veri.vade, koordinatlarJSON, veri.senetNo, veri.tutar], async function(err) {
                if (err) return resolve({ success: false, error: err.message });
                
                const islemId = this.lastID;
                const paperWidth = 210 * 2.83465; const paperHeight = 148 * 2.83465; 

                let senetKok = veri.senetNo; let baslangicNo = 1;
                const match = senetKok.match(/^(.*?)(\d+)$/);
                if (match) { senetKok = match[1]; baslangicNo = parseInt(match[2]); } 
                else if (senetKok.trim() !== '' && !senetKok.endsWith('-')) { senetKok += "-"; }

                const guncelTutarFormatli = veri.tutar ? `# ${veri.tutar} #` : '';
                const guncelTutarYazi = sayiyiYaziyaCevir(veri.tutar);

                // YENİ: Performans Optimizasyonu -> PDF ve Fontları SADECE 1 KERE YÜKLE
                const pdfDoc = await PDFDocument.create(); 
                pdfDoc.registerFontkit(fontkit);
                // app.getAppPath() kullanarak paketlenme sonrasında da dosyanın yerini şak diye bulmasını sağlıyoruz
                const ozelFont = await pdfDoc.embedFont(fs.readFileSync(path.join(app.getAppPath(), 'font.ttf')));

                for (let i = 0; i < veri.taksitSayisi; i++) {
                    // Her taksit için AYNı PDF DOSYASINA yeni bir sayfa ekliyoruz
                    const page = pdfDoc.addPage([paperWidth, paperHeight]);
                    
                    const taksitTarihi = ayEkle(veri.vade, i);
                    const kisaTarih = `${String(taksitTarihi.getDate()).padStart(2, '0')}.${String(taksitTarihi.getMonth() + 1).padStart(2, '0')}.${taksitTarihi.getFullYear()}`;
                    const aylar = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"];
                    const uzunTarih = `${taksitTarihi.getDate()} ${aylar[taksitTarihi.getMonth()]} ${taksitTarihi.getFullYear()}`;
                    const guncelSenetNo = (veri.senetNo.trim() === '') ? '' : (senetKok + (baslangicNo + i));

                    for (const [id, data] of Object.entries(koordinatlar)) {
                        let yazi = data.text;
                        if (id === 'prev_vade') yazi = kisaTarih;
                        if (id === 'prev_vade_yazi') yazi = uzunTarih;
                        if (id === 'prev_senetno') yazi = guncelSenetNo;
                        if (id === 'prev_tutar') yazi = guncelTutarFormatli;
                        if (id === 'prev_tutar_yazi') yazi = guncelTutarYazi;

                        let xPoint = data.x * paperWidth; let yPoint = ((1 - data.y) * paperHeight) - 10;
                        let customSize = data.fontSize || 11; let isBold = data.isBold || false;

                        if (id === 'prev_unvan' || id === 'prev_adres') {
                            drawWrappedText(page, yazi, xPoint, yPoint, 230, ozelFont, customSize);
                        } else {
                            page.drawText(yazi, { x: xPoint, y: yPoint, size: customSize, font: ozelFont, color: rgb(0,0,0) });
                            if (isBold) {
                                page.drawText(yazi, { x: xPoint + 0.3, y: yPoint, size: customSize, font: ozelFont, color: rgb(0,0,0) });
                                page.drawText(yazi, { x: xPoint, y: yPoint + 0.1, size: customSize, font: ozelFont, color: rgb(0,0,0) });
                            }
                        }
                    }
                }
                
                // YENİ: Döngü bitti, içi tamamen dolu tek bir PDF var. Şimdi bunu 1 KERE yazıcıya atıyoruz.
                await hayaletYazdirVeYokEt(pdfDoc, veri.printerName, 'Yeni_Senet');
                resolve({ success: true, senetGecmisId: islemId });
            });
        } catch (err) { resolve({ success: false, error: err.message }); }
    });
});

// GEÇMİŞTEN TEKRAR SENET BASMA MOTORU
ipcMain.handle('gecmisten-yazdir', async (event, islemId, printerName) => {
    return new Promise((resolve) => {
        db.get("SELECT * FROM senet_gecmisi WHERE id = ?", [islemId], async (err, row) => {
            if (err || !row) return resolve({ success: false, error: "Kayıt bulunamadı." });
            try {
                const koordinatlar = JSON.parse(row.koordinatlar);
                const paperWidth = 210 * 2.83465; const paperHeight = 148 * 2.83465; 

                let senetKok = row.senet_koku; let baslangicNo = 1;
                const match = senetKok.match(/^(.*?)(\d+)$/);
                if (match) { senetKok = match[1]; baslangicNo = parseInt(match[2]); } 
                else if (senetKok.trim() !== '' && !senetKok.endsWith('-')) { senetKok += "-"; }

                const guncelTutarFormatli = row.tekil_tutar ? `# ${row.tekil_tutar} #` : '';
                const guncelTutarYazi = sayiyiYaziyaCevir(row.tekil_tutar);

                // YENİ: Performans Optimizasyonu -> PDF ve Fontları SADECE 1 KERE YÜKLE
                const pdfDoc = await PDFDocument.create(); 
                pdfDoc.registerFontkit(fontkit);
                const ozelFont = await pdfDoc.embedFont(fs.readFileSync(path.join(__dirname, 'font.ttf')));

                for (let i = 0; i < row.taksit_sayisi; i++) {
                    const page = pdfDoc.addPage([paperWidth, paperHeight]);
                    
                    const taksitTarihi = ayEkle(row.baslangic_vadesi, i);
                    const kisaTarih = `${String(taksitTarihi.getDate()).padStart(2, '0')}.${String(taksitTarihi.getMonth() + 1).padStart(2, '0')}.${taksitTarihi.getFullYear()}`;
                    const aylar = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"];
                    const uzunTarih = `${taksitTarihi.getDate()} ${aylar[taksitTarihi.getMonth()]} ${taksitTarihi.getFullYear()}`;
                    const guncelSenetNo = (row.senet_koku.trim() === '') ? '' : (senetKok + (baslangicNo + i));

                    for (const [id, data] of Object.entries(koordinatlar)) {
                        let yazi = data.text;
                        if (id === 'prev_vade') yazi = kisaTarih;
                        if (id === 'prev_vade_yazi') yazi = uzunTarih;
                        if (id === 'prev_senetno') yazi = guncelSenetNo;
                        if (id === 'prev_tutar') yazi = guncelTutarFormatli;
                        if (id === 'prev_tutar_yazi') yazi = guncelTutarYazi;

                        let xPoint = data.x * paperWidth; let yPoint = ((1 - data.y) * paperHeight) - 10;
                        let customSize = data.fontSize || 11; let isBold = data.isBold || false;

                        if (id === 'prev_unvan' || id === 'prev_adres') {
                            drawWrappedText(page, yazi, xPoint, yPoint, 240, ozelFont, customSize);
                        } else {
                            page.drawText(yazi, { x: xPoint, y: yPoint, size: customSize, font: ozelFont, color: rgb(0,0,0) });
                            if (isBold) {
                                page.drawText(yazi, { x: xPoint + 0.3, y: yPoint, size: customSize, font: ozelFont, color: rgb(0,0,0) });
                                page.drawText(yazi, { x: xPoint, y: yPoint + 0.1, size: customSize, font: ozelFont, color: rgb(0,0,0) });
                            }
                        }
                    }
                }
                
                // YENİ: Döngü bitti, yazıcıya tek emir
                await hayaletYazdirVeYokEt(pdfDoc, printerName, 'Gecmis_Senet');
                resolve({ success: true });
            } catch (err) { resolve({ success: false, error: err.message }); }
        });
    });
});

// SIFIRDAN VEYA GEÇMİŞTEN GELEN TESLİM FİŞİNİ BASMA MOTORU
ipcMain.handle('teslim-fisi-bas', async (event, veri) => {
    return new Promise(async (resolve) => {
        try {
            const pdfDoc = await PDFDocument.create(); pdfDoc.registerFontkit(fontkit);
            const ozelFont = await pdfDoc.embedFont(fs.readFileSync(path.join(__dirname, 'font.ttf')));
            
            await teslimFisiSayfasiOlustur(pdfDoc, ozelFont, veri);
            await hayaletYazdirVeYokEt(pdfDoc, veri.printerName, 'Teslim_Fisi');

            if(veri.veritabaninaKaydet) {
                db.run(`INSERT INTO teslim_fisleri (senet_gecmis_id, duzenleme_tarihi) VALUES (?, ?)`, [veri.senetGecmisId, veri.duzenlemeTarihi], () => resolve({ success: true }));
            } else {
                resolve({ success: true });
            }
        } catch (err) { resolve({ success: false, error: err.message }); }
    });
});

// GEÇMİŞ LİSTESİNİ ÇEKEN MOTOR
ipcMain.handle('gecmis-getir', async () => {
    return new Promise((resolve) => {
        const sql = `
            SELECT s.id as islem_id, s.toplam_tutar, s.taksit_sayisi, s.baslangic_vadesi,
                   t.id as teslim_fisi_id, t.duzenleme_tarihi, o.ticari_unvan, s.tekil_tutar, s.senet_koku
            FROM senet_gecmisi s
            LEFT JOIN ortaklar o ON s.ortak_id = o.id
            LEFT JOIN teslim_fisleri t ON s.id = t.senet_gecmis_id
            ORDER BY s.id DESC
        `;
        db.all(sql, [], (err, rows) => {
            if (err) resolve({ success: false, error: err.message });
            else resolve({ success: true, data: rows });
        });
    });
});

ipcMain.handle('ortaklari-getir', async () => { return new Promise((resolve) => { db.all("SELECT * FROM ortaklar ORDER BY ticari_unvan ASC", [], (err, rows) => { resolve({ success: true, data: rows }); }); }); });
ipcMain.handle('excel-oku', async () => { const { canceled, filePaths } = await dialog.showOpenDialog({ properties: ['openFile'], filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }] }); if (canceled) return { success: false }; try { const workbook = xlsx.readFile(filePaths[0]); const data = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: '', raw: false }); const clean = data.map(r => { if(r['Vergi No']) r['Vergi No'] = String(r['Vergi No']).replace(/\.0$/, '').trim(); if(r['T.C. Kimlik No']) r['T.C. Kimlik No'] = String(r['T.C. Kimlik No']).replace(/\.0$/, '').trim(); return r; }); return { success: true, data: clean }; } catch(e){ return { success: false, error: e.message }; } });
ipcMain.handle('excel-kaydet', async (e, data) => { return new Promise((resolve) => { db.serialize(() => { db.run('BEGIN TRANSACTION'); const stmt = db.prepare("INSERT INTO ortaklar (ticari_unvan, adi, soyadi, ili, vergi_dairesi, vergi_no, tc_kimlik, adresi) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"); data.forEach(r => stmt.run(r['Ticari Unvanı']||'', r['Adı']||'', r['Soyadı']||'', r['İli']||'', r['Vergi Dairesi']||'', r['Vergi No']||'', r['T.C. Kimlik No']||'', r['Adresi']||'')); stmt.finalize(); db.run('COMMIT', () => resolve({ success: true })); }); }); });
ipcMain.handle('manuel-ortak-ekle', async (e, d) => { return new Promise((resolve) => { db.run("INSERT INTO ortaklar (ticari_unvan, adi, soyadi, ili, vergi_dairesi, vergi_no, tc_kimlik, adresi) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", [d.unvan, d.ad, d.soyad, d.il, d.vd, d.vkn, d.tc, d.adres], () => resolve({ success: true })); }); });

// --- EXCEL DIŞA AKTARMA MOTORU ---
ipcMain.handle('excel-disa-aktar', async (event, data, varsayilanIsim) => {
    try {
        const { canceled, filePath } = await dialog.showSaveDialog({
            defaultPath: varsayilanIsim + '.xlsx',
            filters: [{ name: 'Excel Dosyası', extensions: ['xlsx'] }]
        });
        if (canceled || !filePath) return { success: false, canceled: true };

        const workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.json_to_sheet(data);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Rapor');
        xlsx.writeFile(workbook, filePath);
        return { success: true, filePath };
    } catch (e) { return { success: false, error: e.message }; }
});

// --- YAZICI SÜRÜCÜ AYARLARI ---
ipcMain.handle('get-printers', async () => { return await BrowserWindow.getAllWindows()[0].webContents.getPrintersAsync(); });
ipcMain.handle('save-print-settings', (event, settings) => { 
    const settingsPath = path.join(app.getPath('userData'), 'print_settings.json');
    fs.writeFileSync(settingsPath, JSON.stringify(settings)); 
    return { success: true }; 
});

ipcMain.handle('get-print-settings', () => { 
    const settingsPath = path.join(app.getPath('userData'), 'print_settings.json');
    if (fs.existsSync(settingsPath)) { 
        return JSON.parse(fs.readFileSync(settingsPath)); 
    } 
    return { printerName: '' }; 
});