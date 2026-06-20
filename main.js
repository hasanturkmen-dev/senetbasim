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
// IPC ARKA PLAN MOTORLARI
// ============================================================

ipcMain.handle('senet-bas', async (event, veri, koordinatlar) => {
    return new Promise(async (resolve) => {
        try {
            const senetTutariFloat = parseFloat(veri.tutar.replace(/\./g, '').replace(',', '.'));
            const toplamTutarFloat = senetTutariFloat * veri.taksitSayisi; 
            const koordinatlarJSON = JSON.stringify(koordinatlar);

            // 1. Önce Veritabanına Kaydet
            db.run(`INSERT INTO senet_gecmisi (ortak_id, kampanya_id, toplam_tutar, taksit_sayisi, baslangic_vadesi, koordinatlar, senet_koku, tekil_tutar) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`, 
            [veri.ortakId || 1, 1, toplamTutarFloat, veri.taksitSayisi, veri.vade, koordinatlarJSON, veri.senetNo, veri.tutar], async function(err) {
                if (err) return resolve({ success: false, error: err.message });
                
                const islemId = this.lastID;
                const paperWidth = 210 * 2.83465; 
                const paperHeight = 148 * 2.83465; // A5 Formatı

                let senetKok = veri.senetNo; let baslangicNo = 1;
                const match = senetKok.match(/^(.*?)(\d+)$/);
                if (match) { senetKok = match[1]; baslangicNo = parseInt(match[2]); } 
                else if (senetKok.trim() !== '' && !senetKok.endsWith('-')) { senetKok += "-"; }

                const guncelTutarFormatli = veri.tutar ? `# ${veri.tutar} #` : '';
                const guncelTutarYazi = sayiyiYaziyaCevir(veri.tutar);

                const ciktiKlasoru = path.join(__dirname, 'Ciktilar', `Islem_${islemId}`);
                if (!fs.existsSync(ciktiKlasoru)) fs.mkdirSync(ciktiKlasoru, { recursive: true });

                // 2. Arka Planda PDF'leri Üret
                for (let i = 0; i < veri.taksitSayisi; i++) {
                    const pdfDoc = await PDFDocument.create();
                    pdfDoc.registerFontkit(fontkit);
                    const ozelFont = await pdfDoc.embedFont(fs.readFileSync(path.join(__dirname, 'font.ttf')));
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

                        let xPoint = data.x * paperWidth;
                        let yPoint = ((1 - data.y) * paperHeight) - 10;
                        
                        // ÖN YÜZDEN GELEN DİNAMİK TASARIM AYARLARI
                        let customSize = data.fontSize || 11;
                        let isBold = data.isBold || false;

                        // UZUN METİNLERİ AŞAĞI KAYDIRMA (TEXT-WRAP) ZEKASI
                        if (id === 'prev_unvan') {
                            drawWrappedText(page, yazi, xPoint, yPoint, 230, ozelFont, customSize);
                        } else if (id === 'prev_adres') {
                            drawWrappedText(page, yazi, xPoint, yPoint, 230, ozelFont, customSize);
                        } else {
                            // NORMAL YAZIM
                            page.drawText(yazi, { x: xPoint, y: yPoint, size: customSize, font: ozelFont, color: rgb(0,0,0) });
                            
                            // EFSANE 'FAKE BOLD' ZEKASI (Türkçe karakterleri bozmadan kalınlaştırma)
                            if (isBold) {
                                page.drawText(yazi, { x: xPoint + 0.3, y: yPoint, size: customSize, font: ozelFont, color: rgb(0,0,0) });
                                page.drawText(yazi, { x: xPoint, y: yPoint + 0.1, size: customSize, font: ozelFont, color: rgb(0,0,0) });
                            }
                        }
                    }
                    
                    const dosyaYolu = path.join(ciktiKlasoru, `Senet_${guncelSenetNo || 'Isimsiz'}_Taksit_${i+1}.pdf`);
                    fs.writeFileSync(dosyaYolu, await pdfDoc.save());

                    // 3. SEÇİLİ YAZICIYA ANINDA GÖNDER!
                    if (veri.printerName) {
                        try {
                            await print(dosyaYolu, { printer: veri.printerName });
                        } catch (printErr) {
                            console.error("Yazdırma hatası:", printErr);
                        }
                    }
                }

                resolve({ success: true, senetGecmisId: islemId });
            });
        } catch (err) { resolve({ success: false, error: err.message }); }
    });
});

ipcMain.handle('gecmisten-yazdir', async (event, islemId) => {
    return new Promise((resolve) => {
        db.get("SELECT * FROM senet_gecmisi WHERE id = ?", [islemId], async (err, row) => {
            if (err || !row) return resolve({ success: false, error: "Kayıt bulunamadı." });

            try {
                const koordinatlar = JSON.parse(row.koordinatlar);
                const paperWidth = 210 * 2.83465; 
                const paperHeight = 148 * 2.83465; // Tam 148 mm A5 formatı
                
                const ciktiKlasoru = path.join(__dirname, 'Ciktilar', `Islem_${islemId}`);
                if (!fs.existsSync(ciktiKlasoru)) fs.mkdirSync(ciktiKlasoru, { recursive: true });

                let senetKok = row.senet_koku; let baslangicNo = 1;
                const match = senetKok.match(/^(.*?)(\d+)$/);
                if (match) { senetKok = match[1]; baslangicNo = parseInt(match[2]); } 
                else if (senetKok.trim() !== '' && !senetKok.endsWith('-')) { senetKok += "-"; }

                const guncelTutarFormatli = row.tekil_tutar ? `# ${row.tekil_tutar} #` : '';
                const guncelTutarYazi = sayiyiYaziyaCevir(row.tekil_tutar);

                for (let i = 0; i < row.taksit_sayisi; i++) {
                    const pdfDoc = await PDFDocument.create();
                    pdfDoc.registerFontkit(fontkit);
                    const ozelFont = await pdfDoc.embedFont(fs.readFileSync(path.join(__dirname, 'font.ttf')));
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

                        let xPoint = data.x * paperWidth;
                        let yPoint = ((1 - data.y) * paperHeight) - 10;

                        // ÖN YÜZDEN GELEN DİNAMİK TASARIM AYARLARI
                        let customSize = data.fontSize || 11;
                        let isBold = data.isBold || false;

                        // UZUN METİNLERİ AŞAĞI KAYDIRMA (TEXT-WRAP) ZEKASI
                        if (id === 'prev_unvan') {
                            drawWrappedText(page, yazi, xPoint, yPoint, 240, ozelFont, customSize);
                        } else if (id === 'prev_adres') {
                            drawWrappedText(page, yazi, xPoint, yPoint, 240, ozelFont, customSize);
                        } else {
                            page.drawText(yazi, { x: xPoint, y: yPoint, size: customSize, font: ozelFont, color: rgb(0,0,0) });
                            
                            // EFSANE 'FAKE BOLD' ZEKASI 
                            if (isBold) {
                                page.drawText(yazi, { x: xPoint + 0.3, y: yPoint, size: customSize, font: ozelFont, color: rgb(0,0,0) });
                                page.drawText(yazi, { x: xPoint, y: yPoint + 0.1, size: customSize, font: ozelFont, color: rgb(0,0,0) });
                            }
                        }
                    }
                    fs.writeFileSync(path.join(ciktiKlasoru, `Senet_${guncelSenetNo || 'Isimsiz'}_Taksit_${i+1}.pdf`), await pdfDoc.save());
                }
                
                // Kayıtlı yazıcı varsa ona yazdır
                const settingsPath = path.join(__dirname, 'print_settings.json');
                if (fs.existsSync(settingsPath)) {
                    const settings = JSON.parse(fs.readFileSync(settingsPath));
                    if (settings.printerName) {
                        const files = fs.readdirSync(ciktiKlasoru).filter(f => f.endsWith('.pdf'));
                        for (const file of files) {
                            try {
                                await print(path.join(ciktiKlasoru, file), { printer: settings.printerName });
                            } catch (e) {
                                console.error("Geçmişten yazdırma hatası:", e);
                            }
                        }
                    } else {
                        shell.openPath(ciktiKlasoru); // Yazıcı yoksa klasörü aç
                    }
                } else {
                    shell.openPath(ciktiKlasoru); // Yazıcı ayarı yoksa klasörü aç
                }
                
                resolve({ success: true });
            } catch (err) { resolve({ success: false, error: err.message }); }
        });
    });
});

ipcMain.handle('teslim-fisi-bas', async (event, veri) => {
    return new Promise(async (resolve) => {
        try {
            const a4Width = 595.28; const a4Height = 841.89; // Tam A4 Formatı
            const pdfDoc = await PDFDocument.create(); pdfDoc.registerFontkit(fontkit);
            const ozelFont = await pdfDoc.embedFont(fs.readFileSync(path.join(__dirname, 'font.ttf')));
            const page = pdfDoc.addPage([a4Width, a4Height]);

            let logoImage = null;
            try { if (fs.existsSync(path.join(__dirname, 'logo.png'))) logoImage = await pdfDoc.embedPng(fs.readFileSync(path.join(__dirname, 'logo.png'))); } catch (e) {}

            if (logoImage) page.drawImage(logoImage, { x: 30, y: 760, width: 60, height: 50 });

            page.drawText('S.S. Tüm Optisyenler ve Gözlükçüler Kooperatifi', { x: 100, y: 800, size: 12, font: ozelFont });
            page.drawText('Fevzipaşa Mahallesi 847 Sokak No: 4/A Konak - İZMİR', { x: 100, y: 765, size: 9, font: ozelFont });

            page.drawRectangle({ x: 380, y: 775, width: 170, height: 35, borderWidth: 1, borderColor: rgb(0,0,0) });
            page.drawText('SENET ALINDI BELGESİ', { x: 395, y: 787, size: 11, font: ozelFont });

            let currentY = 680; const startX = 30; const rowHeight = 20;
            const cols = [
                { title: 'SIRA', w: 40, x: startX }, { title: 'DÜZENLEME', w: 90, x: startX + 40 },
                { title: 'TİCARİ ÜNVAN', w: 215, x: startX + 130 }, { title: 'VADE', w: 90, x: startX + 345 },
                { title: 'TUTAR', w: 95, x: startX + 435 }
            ];

            cols.forEach(col => {
                page.drawRectangle({ x: col.x, y: currentY, width: col.w, height: rowHeight, borderWidth: 1, borderColor: rgb(0,0,0), color: rgb(0.9, 0.9, 0.9) });
                page.drawText(col.title, { x: col.x + 5, y: currentY + 6, size: 10, font: ozelFont });
            });
            currentY -= rowHeight;

            const senetTutariFloat = parseFloat(veri.tutar.replace(/\./g, '').replace(',', '.'));
            const formatliSenetTutari = new Intl.NumberFormat('tr-TR', { minimumFractionDigits: 2 }).format(senetTutariFloat);
            const formatliToplam = new Intl.NumberFormat('tr-TR', { minimumFractionDigits: 2 }).format(senetTutariFloat * veri.taksitSayisi);

            // 4. MADDE: TESLİM FİŞİNDE TAM 15 ADET SABIT SATIR ÇİZİMİ
            for (let i = 0; i < 15; i++) {
                cols.forEach(col => { page.drawRectangle({ x: col.x, y: currentY, width: col.w, height: rowHeight, borderWidth: 1, borderColor: rgb(0,0,0) }); });

                page.drawText((i + 1).toString(), { x: cols[0].x + 15, y: currentY + 6, size: 9, font: ozelFont });

                if (i < veri.taksitSayisi) {
                    const tTarih = ayEkle(veri.vade, i);
                    const formatliVade = `${String(tTarih.getDate()).padStart(2, '0')}.${String(tTarih.getMonth() + 1).padStart(2, '0')}.${tTarih.getFullYear()}`;
                    const duzTarih = veri.duzenlemeTarihi.split('-').reverse().join('.');

                    page.drawText(duzTarih, { x: cols[1].x + 10, y: currentY + 6, size: 9, font: ozelFont });
                    page.drawText(formatliVade, { x: cols[3].x + 10, y: currentY + 6, size: 9, font: ozelFont });
                    page.drawText(formatliSenetTutari, { x: cols[4].x + 15, y: currentY + 6, size: 9, font: ozelFont });

                    drawWrappedText(page, veri.unvan, cols[2].x + 5, currentY + 6, 205, ozelFont, 8, 0);
                }
                currentY -= rowHeight;
            }

            currentY -= 20;
            page.drawText('KAŞE - İMZA', { x: 75, y: currentY - 15, size: 10, font: ozelFont });
            page.drawRectangle({ x: 330, y: currentY - 20, width: 80, height: 30, borderWidth: 1, borderColor: rgb(0,0,0) });
            page.drawText('Toplam Tutar', { x: 335, y: currentY - 10, size: 10, font: ozelFont });
            page.drawRectangle({ x: 410, y: currentY - 20, width: 120, height: 30, borderWidth: 1, borderColor: rgb(0,0,0) });
            page.drawText(formatliToplam + " TL", { x: 420, y: currentY - 10, size: 11, font: ozelFont });

            const dosyaYolu = path.join(__dirname, `Teslim_Fisi_SenetID_${veri.senetGecmisId}.pdf`);
            fs.writeFileSync(dosyaYolu, await pdfDoc.save());
            
            // Eğer yazıcı seçiliyse anında teslim fişini de bas
            if (veri.printerName) {
                try {
                    await print(dosyaYolu, { printer: veri.printerName });
                } catch (e) {
                    console.error("Teslim Fişi yazdırma hatası:", e);
                }
            }

            db.run(`INSERT INTO teslim_fisleri (senet_gecmis_id, duzenleme_tarihi) VALUES (?, ?)`, [veri.senetGecmisId, veri.duzenlemeTarihi], () => resolve({ success: true }));
        } catch (err) { resolve({ success: false, error: err.message }); }
    });
});

ipcMain.handle('gecmis-getir', async () => {
    return new Promise((resolve) => {
        const sql = `
            SELECT s.id as islem_id, s.toplam_tutar, s.taksit_sayisi, s.baslangic_vadesi,
                   t.id as teslim_fisi_id, t.duzenleme_tarihi, o.ticari_unvan
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

// --- YAZICI LİSTELEME VE AYAR HAFIZASI ---
ipcMain.handle('get-printers', async () => {
    return await BrowserWindow.getAllWindows()[0].webContents.getPrintersAsync();
});

ipcMain.handle('save-print-settings', (event, settings) => {
    fs.writeFileSync(path.join(__dirname, 'print_settings.json'), JSON.stringify(settings));
    return { success: true };
});

ipcMain.handle('get-print-settings', () => {
    if (fs.existsSync(path.join(__dirname, 'print_settings.json'))) {
        return JSON.parse(fs.readFileSync(path.join(__dirname, 'print_settings.json')));
    }
    return { printerName: '' };
});