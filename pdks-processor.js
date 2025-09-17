const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Excel dosyasını oku
const workbook = XLSX.readFile('PDKS DENEME-2.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Tüm veriyi array olarak al
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

console.log('Excel dosyası yüklendi. Toplam satır:', data.length);

// Vardiya tanımları
const VARDİYALAR = {
    V1: { kod: 'V1', plan_bas: '08:30', plan_bit: '16:30', giris_penceresi: { bas: '06:00', bit: '12:00' } },
    V2: { kod: 'V2', plan_bas: '16:30', plan_bit: '00:30', giris_penceresi: { bas: '14:00', bit: '20:00' } },
    V3: { kod: 'V3', plan_bas: '00:30', plan_bit: '08:30', giris_penceresi: { bas: '22:00', bit: '04:00' } }
};

// Saat string'ini dakikaya çevir
function saatToDakika(saatStr) {
    if (!saatStr || saatStr === '' || saatStr === '..:..') return null;
    
    // Tarih + saat formatını parse et (örn: "01.09.2025 08:28")
    const tarihSaatMatch = saatStr.match(/(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2})/);
    if (tarihSaatMatch) {
        const [, gun, ay, yil, saat, dakika] = tarihSaatMatch;
        const tarih = new Date(yil, ay - 1, gun, parseInt(saat), parseInt(dakika));
        return { tarih, saat: parseInt(saat), dakika: parseInt(dakika) };
    }
    
    // Sadece saat formatını parse et (örn: "08:28")
    const saatMatch = saatStr.match(/(\d{1,2}):(\d{2})/);
    if (saatMatch) {
        const [, saat, dakika] = saatMatch;
        return { tarih: null, saat: parseInt(saat), dakika: parseInt(dakika) };
    }
    
    return null;
}

// Saat string'ini normalize et
function normalizeSaat(saatStr) {
    if (!saatStr) return '';
    const match = saatStr.match(/(\d{1,2}):(\d{1,2})/);
    if (match) {
        const saat = match[1].padStart(2, '0');
        const dakika = match[2].padStart(2, '0');
        return `${saat}:${dakika}`;
    }
    return saatStr;
}

// Dakikayı HH:mm formatına çevir
function dakikaToHHMM(dakika) {
    if (dakika < 0) return '00:00';
    const saat = Math.floor(dakika / 60);
    const dk = dakika % 60;
    return `${saat.toString().padStart(2, '0')}:${dk.toString().padStart(2, '0')}`;
}

// Giriş saatine göre vardiya belirle
function vardiyaBelirle(girisSaat) {
    if (!girisSaat) return null;
    
    const { saat, dakika } = girisSaat;
    const toplamDakika = saat * 60 + dakika;
    
    // V1: 06:00-12:00 (360-720 dakika)
    if (toplamDakika >= 360 && toplamDakika < 720) {
        return VARDİYALAR.V1;
    }
    
    // V2: 14:00-20:00 (840-1200 dakika)
    if (toplamDakika >= 840 && toplamDakika < 1200) {
        return VARDİYALAR.V2;
    }
    
    // V3: 22:00-04:00 (1320-1440 dakika veya 0-240 dakika)
    if (toplamDakika >= 1320 || toplamDakika < 240) {
        return VARDİYALAR.V3;
    }
    
    // Hiçbir pencereye düşmüyorsa, en yakın vardiyayı bul
    const mesafeler = [
        { vardiya: VARDİYALAR.V1, mesafe: Math.min(Math.abs(toplamDakika - 360), Math.abs(toplamDakika - 720)) },
        { vardiya: VARDİYALAR.V2, mesafe: Math.min(Math.abs(toplamDakika - 840), Math.abs(toplamDakika - 1200)) },
        { vardiya: VARDİYALAR.V3, mesafe: Math.min(Math.abs(toplamDakika - 1320), Math.abs(toplamDakika - 240)) }
    ];
    
    mesafeler.sort((a, b) => a.mesafe - b.mesafe);
    return mesafeler[0].vardiya;
}

// Personel bloklarını ayır
function personelBloklariniAyir(data) {
    const bloklar = [];
    let mevcutBlok = null;
    
    // 5. satırdan başla (index 4)
    for (let i = 4; i < data.length; i++) {
        const row = data[i];
        const personelAdi = row[1] || '';
        const giris = row[4] || '';
        const cikis = row[6] || '';
        const icDis = row[9] || '';
        
        // Personel adı varsa yeni blok başlat
        if (personelAdi && personelAdi.trim() !== '') {
            // Önceki bloku kaydet
            if (mevcutBlok) {
                bloklar.push(mevcutBlok);
            }
            
            // Yeni blok başlat
            mevcutBlok = {
                personel: personelAdi.trim(),
                kayitlar: []
            };
        }
        
        // Mevcut blok varsa ve giriş/çıkış bilgisi varsa kayıt ekle
        if (mevcutBlok && (giris || cikis)) {
            // TOPLAM satırlarını atla
            if (giris.includes('TOPLAM:') || cikis.includes('TOPLAM:')) {
                continue;
            }
            
            mevcutBlok.kayitlar.push({
                giris: giris,
                cikis: cikis,
                ic_dis: icDis
            });
        }
    }
    
    // Son bloku kaydet
    if (mevcutBlok) {
        bloklar.push(mevcutBlok);
    }
    
    return bloklar;
}

// Ana işleme fonksiyonu
function verileriIsle() {
    const bloklar = personelBloklariniAyir(data);
    const sonuclar = [];
    const hatalar = [];
    
    console.log(`\n${bloklar.length} personel bloğu bulundu:`);
    bloklar.forEach((blok, index) => {
        console.log(`${index + 1}. ${blok.personel} - ${blok.kayitlar.length} kayıt`);
    });
    
    bloklar.forEach(blok => {
        const { personel, kayitlar } = blok;
        
        // Aynı gün için kayıtları grupla
        const gunlukKayitlar = {};
        
        kayitlar.forEach(kayit => {
            const girisParsed = saatToDakika(kayit.giris);
            const cikisParsed = saatToDakika(kayit.cikis);
            
            // Tarih belirleme - giriş veya çıkıştan
            let tarihKey = null;
            if (girisParsed && girisParsed.tarih) {
                tarihKey = girisParsed.tarih.toISOString().split('T')[0];
            } else if (cikisParsed && cikisParsed.tarih) {
                tarihKey = cikisParsed.tarih.toISOString().split('T')[0];
            }
            
            if (!tarihKey) {
                hatalar.push({
                    personel,
                    kayit,
                    hata: 'Giriş veya çıkış tarihi parse edilemedi'
                });
                return;
            }
            
            if (!gunlukKayitlar[tarihKey]) {
                gunlukKayitlar[tarihKey] = {
                    girisler: [],
                    cikislar: [],
                    ic_dis: kayit.ic_dis
                };
            }
            
            if (girisParsed) {
                gunlukKayitlar[tarihKey].girisler.push(girisParsed);
            }
            if (cikisParsed) {
                gunlukKayitlar[tarihKey].cikislar.push(cikisParsed);
            }
        });
        
        // Her gün için hesaplama yap
        Object.entries(gunlukKayitlar).forEach(([tarih, gunlukData]) => {
            const { girisler, cikislar, ic_dis } = gunlukData;
            
            // İlk giriş ve son çıkış
            const ilkGiris = girisler.length > 0 ? girisler[0] : null;
            const sonCikis = cikislar.length > 0 ? cikislar[cikislar.length - 1] : null;
            
            if (!ilkGiris || !sonCikis) {
                sonuclar.push({
                    personel,
                    tarih,
                    ic_dis,
                    vardiya: null,
                    gercek: {
                        gir: ilkGiris ? normalizeSaat(`${ilkGiris.saat}:${ilkGiris.dakika}`) : '',
                        cik: sonCikis ? normalizeSaat(`${sonCikis.saat}:${sonCikis.dakika}`) : ''
                    },
                    calisma_dk: 0,
                    fm_dk: 0,
                    fm_hhmm: '00:00',
                    durum: 'eksik',
                    not: 'Giriş veya çıkış eksik'
                });
                return;
            }
            
            // Vardiya belirle
            const vardiya = vardiyaBelirle(ilkGiris);
            if (!vardiya) {
                sonuclar.push({
                    personel,
                    tarih,
                    ic_dis,
                    vardiya: null,
                    gercek: {
                        gir: normalizeSaat(`${ilkGiris.saat}:${ilkGiris.dakika}`),
                        cik: normalizeSaat(`${sonCikis.saat}:${sonCikis.dakika}`)
                    },
                    calisma_dk: 0,
                    fm_dk: 0,
                    fm_hhmm: '00:00',
                    durum: 'geçersiz',
                    not: 'Vardiya belirlenemedi'
                });
                return;
            }
            
            // Çalışma süresi hesapla (mola düş)
            const girisDakika = ilkGiris.saat * 60 + ilkGiris.dakika;
            const cikisDakika = sonCikis.saat * 60 + sonCikis.dakika;
            
            // Gece taşması kontrolü
            let calismaDakika = cikisDakika - girisDakika;
            if (calismaDakika < 0) {
                calismaDakika += 24 * 60; // Ertesi güne taşma
            }
            
            // 30 dakika mola düş
            calismaDakika = Math.max(0, calismaDakika - 30);
            
            // Fazla mesai hesapla
            const planliCikisDakika = vardiya.kod === 'V1' ? 16 * 60 + 30 : 
                                    vardiya.kod === 'V2' ? 24 * 60 + 30 : 
                                    8 * 60 + 30; // V3 için ertesi gün 08:30
            
            let fmDakika = 0;
            if (vardiya.kod === 'V2' || vardiya.kod === 'V3') {
                // Gece vardiyaları için özel hesaplama
                if (cikisDakika > planliCikisDakika) {
                    fmDakika = Math.max(0, cikisDakika - planliCikisDakika - 20);
                }
            } else {
                // Gündüz vardiyası
                if (cikisDakika > planliCikisDakika) {
                    fmDakika = Math.max(0, cikisDakika - planliCikisDakika - 20);
                }
            }
            
            sonuclar.push({
                personel,
                tarih,
                ic_dis,
                vardiya: {
                    kod: vardiya.kod,
                    plan_bas: vardiya.plan_bas,
                    plan_bit: vardiya.plan_bit
                },
                gercek: {
                    gir: normalizeSaat(`${ilkGiris.saat}:${ilkGiris.dakika}`),
                    cik: normalizeSaat(`${sonCikis.saat}:${sonCikis.dakika}`)
                },
                calisma_dk: calismaDakika,
                fm_dk: fmDakika,
                fm_hhmm: dakikaToHHMM(fmDakika),
                durum: 'ok',
                not: ''
            });
        });
    });
    
    return { sonuclar, hatalar };
}

// Çıktı dosyalarını oluştur
function ciktiDosyalariniOlustur(sonuclar, hatalar) {
    // out klasörünü oluştur
    if (!fs.existsSync('out')) {
        fs.mkdirSync('out');
    }
    
    // JSON çıktısı
    const jsonCikti = {
        islem_tarihi: new Date().toISOString(),
        toplam_kayit: sonuclar.length,
        hata_sayisi: hatalar.length,
        veriler: sonuclar
    };
    
    fs.writeFileSync('out/daily.json', JSON.stringify(jsonCikti, null, 2), 'utf8');
    
    // CSV çıktısı
    const csvBaslik = 'personel,tarih,ic_dis,vardiya_kod,vardiya_plan_bas,vardiya_plan_bit,gercek_gir,gercek_cik,calisma_dk,fm_dk,fm_hhmm,durum,not';
    const csvSatirlar = sonuclar.map(kayit => [
        kayit.personel,
        kayit.tarih,
        kayit.ic_dis,
        kayit.vardiya ? kayit.vardiya.kod : '',
        kayit.vardiya ? kayit.vardiya.plan_bas : '',
        kayit.vardiya ? kayit.vardiya.plan_bit : '',
        kayit.gercek.gir,
        kayit.gercek.cik,
        kayit.calisma_dk,
        kayit.fm_dk,
        kayit.fm_hhmm,
        kayit.durum,
        kayit.not
    ].map(alan => `"${alan}"`).join(','));
    
    const csvIcerik = [csvBaslik, ...csvSatirlar].join('\n');
    fs.writeFileSync('out/daily.csv', csvIcerik, 'utf8');
    
    console.log('\n=== ÇIKTI DOSYALARI OLUŞTURULDU ===');
    console.log('out/daily.json - JSON formatında detaylı veri');
    console.log('out/daily.csv - CSV formatında tablo verisi');
    console.log(`\nToplam işlenen kayıt: ${sonuclar.length}`);
    console.log(`Hata sayısı: ${hatalar.length}`);
    
    if (hatalar.length > 0) {
        console.log('\n=== HATALAR ===');
        hatalar.forEach((hata, index) => {
            console.log(`${index + 1}. ${hata.personel}: ${hata.hata}`);
        });
    }
}

// Ana işlemi başlat
const { sonuclar, hatalar } = verileriIsle();
ciktiDosyalariniOlustur(sonuclar, hatalar);