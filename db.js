const sqlite3 = require('sqlite3').verbose();
const path = require('path');

// Veritabanı dosyasının yolu (Uygulama klasöründe 'senet_sistemi.db' adında bir dosya oluşturacak)
const dbPath = path.join(__dirname, 'senet_sistemi.db');

// Veritabanı bağlantısını başlat
const db = new sqlite3.Database(dbPath, (err) => {
    if (err) {
        console.error("Veritabanına bağlanırken hata oluştu:", err.message);
    } else {
        console.log("SQLite veritabanına başarıyla bağlanıldı. 📦");
        tablolariOlustur();
    }
});

// Tabloları oluşturan fonksiyon
function tablolariOlustur() {
    db.serialize(() => {
        
        // 1. ORTAKLAR TABLOSU
        // Fuarda CSV'den aktardığın veya yeni eklediğin ortakların duracağı yer.
        db.run(`CREATE TABLE IF NOT EXISTS ortaklar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unvan TEXT NOT NULL,
            il TEXT,
            vergi_dairesi TEXT,
            kimlik_no TEXT,
            adres TEXT,
            eklenme_tarihi DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);

        // 2. KAMPANYALAR TABLOSU
        // Hangi fuar veya hangi anlaşma? (Örn: "2026 İstanbul Optik Fuarı Kampanyası")
        db.run(`CREATE TABLE IF NOT EXISTS kampanyalar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kampanya_adi TEXT NOT NULL,
            aktif_mi INTEGER DEFAULT 1,
            olusturulma_tarihi DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);

        // 3. BASILAN SENETLERİN GEÇMİŞİ
        // Sistemin kalbi. Hangi ortak, hangi kampanyadan, toplam ne kadarlık ve kaç taksit senet imzaladı?
        db.run(`CREATE TABLE IF NOT EXISTS senet_gecmisi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ortak_id INTEGER,
            kampanya_id INTEGER,
            toplam_tutar REAL NOT NULL,
            taksit_sayisi INTEGER NOT NULL,
            baslangic_vadesi TEXT,
            basim_tarihi DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(ortak_id) REFERENCES ortaklar(id),
            FOREIGN KEY(kampanya_id) REFERENCES kampanyalar(id)
        )`);
        // 4. TESLİM FİŞLERİ TABLOSU (Senet Geçmişi ile Birebir/Çoka Bir İlişkili)
        db.run(`CREATE TABLE IF NOT EXISTS teslim_fisleri (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            senet_gecmis_id INTEGER,
            duzenleme_tarihi TEXT,
            olusturma_tarihi DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(senet_gecmis_id) REFERENCES senet_gecmisi(id)
        )`);

        console.log("Veritabanı tabloları kontrol edildi ve hazır.");
    });
}

// Diğer dosyalardan bu veritabanına erişebilmek için dışarı aktarıyoruz
module.exports = db;