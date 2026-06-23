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
// --- Orijinal Matbaa Tasarımlı Teslim Fişi PDF Motoru ---
async function teslimFisiSayfasiOlustur(pdfDoc, ozelFont, fisVerisi) {
    const a4Width = 595.28; const a4Height = 841.89; // Tam A4 Formatı
    const page = pdfDoc.addPage([a4Width, a4Height]);

    let logoImage = null;
    try { 
        if (fs.existsSync(path.join(app.getAppPath(), 'logo.png'))) {
            logoImage = await pdfDoc.embedPng(fs.readFileSync(path.join(app.getAppPath(), 'logo.png'))); 
        } 
    } catch (e) {}

    // Sol Üst - Logo ve Kurumsal Kimlik (PDF Tasarımına Göre)
    if (logoImage) {
        const scaled = logoImage.scaleToFit(80, 65);
        page.drawImage(logoImage, { x: 30, y: 815 - scaled.height, width: scaled.width, height: scaled.height });
    }

    page.drawText('S.S. Tüm Optisyenler ve Gözlükcüler', { x: 120, y: 805, size: 12, font: ozelFont, color: rgb(0.1, 0.1, 0.1) });
    page.drawText('Temin Tevzi Kooperatifi', { x: 120, y: 790, size: 12, font: ozelFont, color: rgb(0.1, 0.1, 0.1) });
    
    page.drawText('Fevzipaşa Mahallesi 847 Sokak No: 4/A Konak - İZMİR', { x: 120, y: 775, size: 9, font: ozelFont, color: rgb(0.3, 0.3, 0.3) });
    page.drawText('Tel: 0232 489 89 62  |  Fax: 0232 489 39 62', { x: 120, y: 762, size: 9, font: ozelFont, color: rgb(0.3, 0.3, 0.3) });
    page.drawText('muhasebe@gozkoop.com  |  www.gozkoop.com', { x: 120, y: 749, size: 9, font: ozelFont, color: rgb(0.3, 0.3, 0.3) });
    page.drawText('Mersis No: 0875 0359 0980 00 15  |  Ticaret Sicil No: 155588  |  Kemeraltı V.D. 8750359098', { x: 120, y: 736, size: 9, font: ozelFont, color: rgb(0.3, 0.3, 0.3) });

    // Orta Başlık Çizgisi
    const title = 'SENET ALINDI BELGESİ';
    const titleWidth = ozelFont.widthOfTextAtSize(title, 14);
    page.drawRectangle({ x: (a4Width - titleWidth) / 2 - 20, y: 678, width: titleWidth + 40, height: 28, color: rgb(0.95, 0.96, 0.98), borderColor: rgb(0.8, 0.85, 0.9), borderWidth: 1 });
    page.drawText(title, { x: (a4Width - titleWidth) / 2, y: 688, size: 14, font: ozelFont, color: rgb(0.1, 0.15, 0.2) });
    
    // Sağ Üst Tarih Alanı (Dinamik)
    const ustTarihFormatli = fisVerisi.ustTarih.split('-').reverse().join('.');
    page.drawText(`Tarih: ${ustTarihFormatli}`, { x: 460, y: 688, size: 11, font: ozelFont, color: rgb(0.2, 0.25, 0.3) });

    // Dinamik Tablo Çizimi
    let currentY = 640; const startX = 30; const rowHeight = 22;
    const cols = fisVerisi.kolonlar; 

    // Tablo Başlıkları
    page.drawRectangle({ x: startX, y: currentY, width: 535, height: rowHeight, color: rgb(0.95, 0.96, 0.98), borderColor: rgb(0.7, 0.7, 0.7), borderWidth: 1 });
    cols.forEach(col => {
        page.drawText(col.title, { x: col.x + 8, y: currentY + 7, size: 9, font: ozelFont, color: rgb(0.2, 0.2, 0.2) });
        if(col.x > startX) page.drawLine({ start: { x: col.x, y: currentY }, end: { x: col.x, y: currentY + rowHeight }, thickness: 1, color: rgb(0.7, 0.7, 0.7) });
    });
    currentY -= rowHeight;

    // A4 Satırları
    for (let i = 0; i < 15; i++) {
        page.drawRectangle({ x: startX, y: currentY, width: 535, height: rowHeight, borderColor: rgb(0.7, 0.7, 0.7), borderWidth: 1 });
        cols.forEach(col => {
            if(col.x > startX) page.drawLine({ start: { x: col.x, y: currentY }, end: { x: col.x, y: currentY + rowHeight }, thickness: 1, color: rgb(0.7, 0.7, 0.7) });
        });

        if (i < fisVerisi.taksitSayisi && fisVerisi.satirlar[i]) {
            const satirVerisi = fisVerisi.satirlar[i];
            cols.forEach((col, index) => {
                let hucreYazisi = satirVerisi[index] || '';
                let hucreX = col.x + 6; let hucreY = currentY + 7; let hucreFontSize = 9;

                if (col.id === 'unvan') {
                    hucreFontSize = hucreYazisi.length > 45 ? 6.5 : 8;
                    hucreY = hucreYazisi.length > 45 ? currentY + 13 : currentY + 7;
                    let wrapSpacing = hucreYazisi.length > 45 ? 8 : 11;
                    drawWrappedText(page, hucreYazisi, hucreX, hucreY, col.w - 10, ozelFont, hucreFontSize, wrapSpacing);
                } else {
                    page.drawText(hucreYazisi, { x: hucreX, y: hucreY, size: hucreFontSize, font: ozelFont, color: rgb(0.1, 0.1, 0.1) });
                }
            });
        } else {
            page.drawText((i + 1).toString(), { x: cols[0].x + 15, y: currentY + 7, size: 9, font: ozelFont });
        }
        currentY -= rowHeight;
    }

    // Kaşe - İmza ve Toplam Tutar (Smooth Matbaa Stili)
    currentY -= 40;
    
    // Teslim Eden Bölümü
    page.drawText('TESLİM EDEN', { x: 75, y: currentY + 15, size: 9, font: ozelFont, color: rgb(0.4, 0.45, 0.5) });
    page.drawLine({ start: { x: 50, y: currentY + 8 }, end: { x: 155, y: currentY + 8 }, thickness: 0.5, color: rgb(0.7, 0.75, 0.8) });
    page.drawText('KAŞE - İMZA', { x: 75, y: currentY - 5, size: 10, font: ozelFont, color: rgb(0.15, 0.2, 0.25) });

    // Teslim Alan Bölümü
    page.drawText('TESLİM ALAN', { x: 235, y: currentY + 15, size: 9, font: ozelFont, color: rgb(0.4, 0.45, 0.5) });
    page.drawLine({ start: { x: 210, y: currentY + 8 }, end: { x: 315, y: currentY + 8 }, thickness: 0.5, color: rgb(0.7, 0.75, 0.8) });
    page.drawText('KAŞE - İMZA', { x: 235, y: currentY - 5, size: 10, font: ozelFont, color: rgb(0.15, 0.2, 0.25) });

    // Toplam Tutar (Premium Yumuşak Kutu)
    const formatliToplam = new Intl.NumberFormat('tr-TR', { minimumFractionDigits: 2 }).format(fisVerisi.toplamTutarFloat) + ' TL';
    page.drawRectangle({ x: 375, y: currentY - 15, width: 190, height: 45, color: rgb(0.96, 0.97, 0.98), borderColor: rgb(0.7, 0.75, 0.8), borderWidth: 1 });
    page.drawText('GENEL TOPLAM', { x: 390, y: currentY + 16, size: 8, font: ozelFont, color: rgb(0.4, 0.45, 0.5) });
    page.drawText(formatliToplam, { x: 390, y: currentY - 4, size: 15, font: ozelFont, color: rgb(0.05, 0.15, 0.25) });
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