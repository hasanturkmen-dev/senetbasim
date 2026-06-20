const sqlite3 = require('sqlite3').verbose();
const path = require('path');

const dbPath = path.join(__dirname, 'database.sqlite');
const db = new sqlite3.Database(dbPath);

db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS senet_gecmisi (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ortak_id INTEGER,
        kampanya_id INTEGER,
        toplam_tutar REAL,
        taksit_sayisi INTEGER,
        baslangic_vadesi TEXT,
        olusturma_tarihi DATETIME DEFAULT CURRENT_TIMESTAMP
    )`);
    db.run(`ALTER TABLE senet_gecmisi ADD COLUMN koordinatlar TEXT`, () => {});
    db.run(`ALTER TABLE senet_gecmisi ADD COLUMN senet_koku TEXT`, () => {});
    db.run(`ALTER TABLE senet_gecmisi ADD COLUMN tekil_tutar TEXT`, () => {});

    db.run(`CREATE TABLE IF NOT EXISTS teslim_fisleri (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        senet_gecmis_id INTEGER,
        duzenleme_tarihi TEXT,
        olusturma_tarihi DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(senet_gecmis_id) REFERENCES senet_gecmisi(id)
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS ortaklar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ticari_unvan TEXT,
        adi TEXT,
        soyadi TEXT,
        ili TEXT,
        vergi_dairesi TEXT,
        vergi_no TEXT,
        tc_kimlik TEXT,
        adresi TEXT
    )`);
});

module.exports = db;