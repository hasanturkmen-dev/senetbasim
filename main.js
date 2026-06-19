const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { PDFDocument, rgb } = require('pdf-lib');
const fontkit = require('@pdf-lib/fontkit');
const db = require('./db.js'); // SQLite Veritabanı bağlantımız

// --- KUSURSUZ TARİH FONKSİYONU ---
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

    let sonGun;
    if (baslangicAySonuMu) {
        sonGun = aydakiGunSayisi;
    } else {
        sonGun = Math.min(basGun, aydakiGunSayisi);
    }

    return new Date(yeniYil, yeniAy, sonGun);
}

// --- RAKAMI YAZIYA ÇEVİREN ARKA PLAN MOTORU ---
function sayiyiYaziyaCevir(hamTutar) {
    if (!hamTutar) return "";
    let sayi = hamTutar.toString().replace(/\./g, '').replace(',', '.');
    if (isNaN(sayi) || sayi == 0) return "SIFIR";
    
    const birler = ["", "BİR", "İKİ", "ÜÇ", "DÖRT", "BEŞ", "ALTI", "YEDİ", "SEKİZ", "DOKUZ"];
    const onlar = ["", "ON", "YİRMİ", "OTUZ", "KIRK", "ELLİ", "ALTMIŞ", "YETMİŞ", "SEKSEN", "DOKSAN"];
    const binler = ["", "BİN", "MİLYON", "MİLYAR"];
    
    let str = parseFloat(sayi).toFixed(2).toString();
    let parcalar = str.split('.');
    let tl = parcalar[0];
    let kurus = parcalar[1];
    
    function ucluOku(n) {
        let yuz = Math.floor(n / 100);
        let on = Math.floor((n % 100) / 10);
        let bir = n % 10;
        let t = "";
        if (yuz == 1) t += "YÜZ"; else if (yuz > 1) t += birler[yuz] + "YÜZ";
        t += onlar[on]; t += birler[bir]; 
        return t;
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
    width: 1280,
    height: 800,
    title: "Senet Sistemi",
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    }
  });
  mainWindow.loadFile('index.html');
}

app.whenReady().then(() => {
  createWindow();
  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
}); 

// ============================================================
// IPC ARKA PLAN MOTORLARI (PROMISE YAPISI İLE GÜVENLİ HALE GETİRİLDİ)
// ============================================================

// 1. SENET BASIMI VE VERİTABANI GEÇMİŞ KAYDI DİNLEYİCİSİ
ipcMain.handle('senet-bas', async (event, veri, koordinatlar) => {
    return new Promise(async (resolve) => {
        try {
            const paperWidth = 210 * 2.83465; 
            const paperHeight = 145 * 2.83465;

            // Senet No Kökünü Bul (Örn: "GRT-1" veya sadece "GRT" girilirse ona göre artırır)
            let senetKok = veri.senetNo;
            let baslangicNo = 1;
            const match = veri.senetNo.match(/^(.*?)(\d+)$/);
            if (match) {
                senetKok = match[1];
                baslangicNo = parseInt(match[2]);
            } else if (veri.senetNo.trim() !== '' && !veri.senetNo.endsWith('-')) {
                senetKok += "-";
            }

            // Tutarı arka planda tekrar hesapla (Ön yüz hatalarına karşı garanti)
            const guncelTutarFormatli = veri.tutar ? `# ${veri.tutar} #` : '';
            const guncelTutarYazi = sayiyiYaziyaCevir(veri.tutar);

            for (let i = 0; i < veri.taksitSayisi; i++) {
                const pdfDoc = await PDFDocument.create();
                pdfDoc.registerFontkit(fontkit);

                const fontYolu = path.join(__dirname, 'font.ttf');
                const fontBytes = fs.readFileSync(fontYolu);
                const ozelFont = await pdfDoc.embedFont(fontBytes);

                const page = pdfDoc.addPage([paperWidth, paperHeight]);
                
                const taksitTarihi = ayEkle(veri.vade, i);
                const kisaTarih = `${String(taksitTarihi.getDate()).padStart(2, '0')}.${String(taksitTarihi.getMonth() + 1).padStart(2, '0')}.${taksitTarihi.getFullYear()}`;
                const aylar = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"];
                const uzunTarih = `${taksitTarihi.getDate()} ${aylar[taksitTarihi.getMonth()]} ${taksitTarihi.getFullYear()}`;

                const guncelSenetNo = (veri.senetNo.trim() === '') ? '' : (senetKok + (baslangicNo + i));

                for (const [id, data] of Object.entries(koordinatlar)) {
                    let yazi = data.text;
                    
                    // Ön yüzdeki veriyi arka plan matematiğiyle ez
                    if (id === 'prev_vade') yazi = kisaTarih;
                    if (id === 'prev_vade_yazi') yazi = uzunTarih;
                    if (id === 'prev_senetno') yazi = guncelSenetNo;
                    if (id === 'prev_tutar') yazi = guncelTutarFormatli;
                    if (id === 'prev_tutar_yazi') yazi = guncelTutarYazi;

                    const xPoint = data.x * paperWidth;
                    const yPoint = (1 - data.y) * paperHeight;

                    page.drawText(yazi, {
                        x: xPoint,
                        y: yPoint - 10,
                        size: 11,
                        font: ozelFont,
                        color: rgb(0, 0, 0),
                    });
                }

                const pdfBytes = await pdfDoc.save();
                const dosyaYolu = path.join(__dirname, `Senet_${guncelSenetNo || 'Isimsiz'}_Taksit_${i+1}.pdf`);
                fs.writeFileSync(dosyaYolu, pdfBytes);
            }

            // Veritabanına TOPLAM tutarı kaydet
            const senetTutariFloat = parseFloat(veri.tutar.replace(/\./g, '').replace(',', '.'));
            const toplamTutarFloat = senetTutariFloat * veri.taksitSayisi; 
            
            const sql = `INSERT INTO senet_gecmisi (ortak_id, kampanya_id, toplam_tutar, taksit_sayisi, baslangic_vadesi) 
                         VALUES (?, ?, ?, ?, ?)`;
                         
            db.run(sql, [1, 1, toplamTutarFloat, veri.taksitSayisi, veri.vade], function(err) {
                if (err) {
                    console.error("Veritabanı kayıt hatası:", err.message);
                    resolve({ success: false, error: err.message });
                } else {
                    console.log(`Veri başarıyla kaydedildi. Kayıt ID: ${this.lastID}`);
                    resolve({ success: true, senetGecmisId: this.lastID });
                }
            });

        } catch (err) {
            resolve({ success: false, error: err.message });
        }
    });
});

// 2. YENİ 'teslim-fisi-bas' DİNLEYİCİSİ (A4 FORMATI VE DOĞRU HESAPLAMA)
ipcMain.handle('teslim-fisi-bas', async (event, veri, ayarlar) => {
    return new Promise(async (resolve) => {
        try {
            const a4Width = 595.28;
            const a4Height = 841.89;

            const pdfDoc = await PDFDocument.create();
            pdfDoc.registerFontkit(fontkit);

            const fontYolu = path.join(__dirname, 'font.ttf');
            const fontBytes = fs.readFileSync(fontYolu);
            const ozelFont = await pdfDoc.embedFont(fontBytes);

            let logoImage = null;
            try {
                const logoYolu = path.join(__dirname, 'logo.png');
                if (fs.existsSync(logoYolu)) {
                    const logoBytes = fs.readFileSync(logoYolu);
                    logoImage = await pdfDoc.embedPng(logoBytes);
                }
            } catch (err) {
                console.log("Logo yüklenemedi, metin ile devam edilecek.");
            }

            const page = pdfDoc.addPage([a4Width, a4Height]);
            
            const boyutlar = {
                baslik: ayarlar?.baslikBoyut || 12,
                adres: ayarlar?.adresBoyut || 9,
                tabloBaslik: ayarlar?.tabloBaslikBoyut || 10,
                tabloIcerik: ayarlar?.tabloIcerikBoyut || 9
            };

            // --- HEADER BÖLÜMÜ ---
            if (logoImage) {
                page.drawImage(logoImage, { x: 30, y: 760, width: 60, height: 50 });
            }

            page.drawText('S.S. Tüm Optisyenler ve Gözlükçüler', { x: 100, y: 800, size: boyutlar.baslik, font: ozelFont });
            page.drawText('Temin Tevzi Kooperatifi', { x: 130, y: 785, size: boyutlar.baslik, font: ozelFont });
            
            page.drawText('Fevzipaşa Mahallesi 847 Sokak No: 4/A Konak - İZMİR', { x: 100, y: 765, size: boyutlar.adres, font: ozelFont });
            page.drawText('Tel: 0232 489 89 62 - Fax: 0232 489 39 62', { x: 125, y: 753, size: boyutlar.adres, font: ozelFont });
            page.drawText('bilgi@gozkoop.com - www.gozkoop.com', { x: 125, y: 741, size: boyutlar.adres, font: ozelFont });
            page.drawText('Ticaret Sicil No: 155588 - Kemeraltı V.D. 8750359098', { x: 110, y: 729, size: boyutlar.adres, font: ozelFont });

            // SENET ALINDI BELGESİ KUTUSU VE TARİH
            page.drawRectangle({ x: 380, y: 775, width: 170, height: 35, borderWidth: 1, borderColor: rgb(0,0,0) });
            page.drawText('SENET ALINDI BELGESİ', { x: 395, y: 787, size: 12, font: ozelFont });

            page.drawText('Tarih', { x: 380, y: 745, size: boyutlar.tabloBaslik, font: ozelFont });
            page.drawRectangle({ x: 420, y: 740, width: 25, height: 20, borderWidth: 1, borderColor: rgb(0,0,0) });
            page.drawRectangle({ x: 450, y: 740, width: 25, height: 20, borderWidth: 1, borderColor: rgb(0,0,0) });
            page.drawRectangle({ x: 480, y: 740, width: 40, height: 20, borderWidth: 1, borderColor: rgb(0,0,0) });
            
            const dTarih = veri.duzenlemeTarihi.split('-');
            if(dTarih.length === 3) {
                page.drawText(dTarih[2], { x: 425, y: 746, size: 10, font: ozelFont });
                page.drawText(dTarih[1], { x: 455, y: 746, size: 10, font: ozelFont });
                page.drawText(dTarih[0], { x: 485, y: 746, size: 10, font: ozelFont });
            }

            // --- TABLO ÇİZİMİ ---
            let currentY = 680;
            const startX = 30;
            const rowHeight = 20;
            
            const cols = [
                { title: 'SIRA', w: 40, x: startX },
                { title: 'DÜZENLEME', w: 90, x: startX + 40 },
                { title: 'TİCARİ ÜNVAN', w: 215, x: startX + 130 },
                { title: 'VADE', w: 90, x: startX + 345 },
                { title: 'TUTAR', w: 95, x: startX + 435 }
            ];

            cols.forEach(col => {
                page.drawRectangle({ x: col.x, y: currentY, width: col.w, height: rowHeight, borderWidth: 1, borderColor: rgb(0,0,0), color: rgb(0.9, 0.9, 0.9) });
                page.drawText(col.title, { x: col.x + 5, y: currentY + 6, size: boyutlar.tabloBaslik, font: ozelFont });
            });

            currentY -= rowHeight;

            // HESAPLAMA: Her satıra Tekil Senet Tutarı, en alta Toplam Tutar
            const senetTutariFloat = parseFloat(veri.tutar.replace(/\./g, '').replace(',', '.'));
            const toplamTutarFloat = senetTutariFloat * veri.taksitSayisi;
            
            const formatliSenetTutari = new Intl.NumberFormat('tr-TR', { minimumFractionDigits: 2 }).format(senetTutariFloat);
            const formatliToplam = new Intl.NumberFormat('tr-TR', { minimumFractionDigits: 2 }).format(toplamTutarFloat);

            for (let i = 0; i < veri.taksitSayisi; i++) {
                const tTarih = ayEkle(veri.vade, i);
                const formatliVade = `${String(tTarih.getDate()).padStart(2, '0')}.${String(tTarih.getMonth() + 1).padStart(2, '0')}.${tTarih.getFullYear()}`;
                const duzTarih = veri.duzenlemeTarihi.split('-').reverse().join('.');

                cols.forEach(col => {
                    page.drawRectangle({ x: col.x, y: currentY, width: col.w, height: rowHeight, borderWidth: 1, borderColor: rgb(0,0,0) });
                });

                page.drawText((i+1).toString(), { x: cols[0].x + 15, y: currentY + 6, size: boyutlar.tabloIcerik, font: ozelFont });
                page.drawText(duzTarih, { x: cols[1].x + 10, y: currentY + 6, size: boyutlar.tabloIcerik, font: ozelFont });
                
                let fontSize = boyutlar.tabloIcerik;
                if(veri.unvan.length > 40) fontSize = fontSize - 2;
                page.drawText(veri.unvan, { x: cols[2].x + 5, y: currentY + 6, size: fontSize, font: ozelFont });

                page.drawText(formatliVade, { x: cols[3].x + 10, y: currentY + 6, size: boyutlar.tabloIcerik, font: ozelFont });
                
                // Her satıra formda girilen tekil senet tutarı basılır
                page.drawText(formatliSenetTutari, { x: cols[4].x + 15, y: currentY + 6, size: boyutlar.tabloIcerik, font: ozelFont });

                currentY -= rowHeight;
            }

            // --- KAŞE İMZA VE TOPLAM BÖLÜMÜ ---
            currentY -= 40;
            
            page.drawLine({ start: { x: 40, y: currentY }, end: { x: 180, y: currentY }, thickness: 1 });
            page.drawText('KAŞE - İMZA', { x: 75, y: currentY - 15, size: 10, font: ozelFont });

            page.drawRectangle({ x: 330, y: currentY - 20, width: 80, height: 30, borderWidth: 1, borderColor: rgb(0,0,0) });
            page.drawText('Toplam Tutar', { x: 335, y: currentY - 10, size: 10, font: ozelFont });
            
            page.drawRectangle({ x: 410, y: currentY - 20, width: 120, height: 30, borderWidth: 1, borderColor: rgb(0,0,0) });
            // En alta formda girilen tutar X taksit sayısı basılır
            page.drawText(formatliToplam + " TL", { x: 420, y: currentY - 10, size: 11, font: ozelFont });

            const pdfBytes = await pdfDoc.save();
            const dosyaYolu = path.join(__dirname, `Teslim_Fisi_SenetID_${veri.senetGecmisId}.pdf`);
            fs.writeFileSync(dosyaYolu, pdfBytes);

            // Veritabanı İlişkilendirmesi
            const sqlFis = `INSERT INTO teslim_fisleri (senet_gecmis_id, duzenleme_tarihi) VALUES (?, ?)`;
            db.run(sqlFis, [veri.senetGecmisId, veri.duzenlemeTarihi], function(err) {
                if (err) resolve({ success: false, error: err.message });
                else resolve({ success: true });
            });

        } catch (err) {
            resolve({ success: false, error: err.message });
        }
    });
});
// 3. GEÇMİŞ İŞLEMLERİ VERİTABANINDAN GETİRME
ipcMain.handle('gecmis-getir', async () => {
    return new Promise((resolve) => {
        const sql = `
            SELECT 
                s.id as islem_id, 
                s.toplam_tutar, 
                s.taksit_sayisi, 
                s.baslangic_vadesi,
                t.id as teslim_fisi_id,
                t.duzenleme_tarihi
            FROM senet_gecmisi s
            LEFT JOIN teslim_fisleri t ON s.id = t.senet_gecmis_id
            ORDER BY s.id DESC
        `;
        
        db.all(sql, [], (err, rows) => {
            if (err) {
                resolve({ success: false, error: err.message });
            } else {
                resolve({ success: true, data: rows });
            }
        });
    });
});