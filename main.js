const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

// Cache sorunlarını önlemek için
app.commandLine.appendSwitch('--disable-gpu-sandbox');
app.commandLine.appendSwitch('--disable-software-rasterizer');
app.commandLine.appendSwitch('--disable-gpu');
app.commandLine.appendSwitch('--no-sandbox');

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
            webSecurity: false,
            allowRunningInsecureContent: true
        },
        icon: path.join(__dirname, 'assets/icon.png'), // İsteğe bağlı icon
        title: 'PDKS İşleme Sistemi',
        show: false
    });

    mainWindow.loadFile('renderer/index.html');

    // Pencere yüklendiğinde göster
    mainWindow.once('ready-to-show', () => {
        mainWindow.show();
    });

    // Geliştirme modunda DevTools'u aç
    if (process.env.NODE_ENV === 'development') {
        mainWindow.webContents.openDevTools();
    }
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});

// Dosya seçme dialogu
ipcMain.handle('select-file', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        title: 'PDKS Excel Dosyasını Seçin',
        filters: [
            { name: 'Excel Dosyaları', extensions: ['xlsx', 'xls'] },
            { name: 'Tüm Dosyalar', extensions: ['*'] }
        ],
        properties: ['openFile']
    });

    if (!result.canceled && result.filePaths.length > 0) {
        return result.filePaths[0];
    }
    return null;
});

// PDKS işleme fonksiyonu
ipcMain.handle('process-pdks', async (event, filePath) => {
    try {
        // Excel dosyasını oku
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

        // Vardiya tanımları
        const VARDİYALAR = {
            V1: { kod: 'V1', plan_bas: '08:30', plan_bit: '16:30', giris_penceresi: { bas: '06:00', bit: '12:00' } },
            V2: { kod: 'V2', plan_bas: '16:30', plan_bit: '00:30', giris_penceresi: { bas: '14:00', bit: '20:00' } },
            V3: { kod: 'V3', plan_bas: '00:30', plan_bit: '08:30', giris_penceresi: { bas: '22:00', bit: '04:00' } }
        };

        // Saat string'ini dakikaya çevir
        function saatToDakika(saatStr) {
            if (!saatStr || saatStr === '' || saatStr === '..:..') return null;
            
            const tarihSaatMatch = saatStr.match(/(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2})/);
            if (tarihSaatMatch) {
                const [, gun, ay, yil, saat, dakika] = tarihSaatMatch;
                const tarih = new Date(yil, ay - 1, gun, parseInt(saat), parseInt(dakika));
                return { tarih, saat: parseInt(saat), dakika: parseInt(dakika) };
            }
            
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
            
            if (toplamDakika >= 360 && toplamDakika < 720) {
                return VARDİYALAR.V1;
            }
            
            if (toplamDakika >= 840 && toplamDakika < 1200) {
                return VARDİYALAR.V2;
            }
            
            if (toplamDakika >= 1320 || toplamDakika < 240) {
                return VARDİYALAR.V3;
            }
            
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
            
            for (let i = 4; i < data.length; i++) {
                const row = data[i];
                const personelAdi = row[1] || '';
                const giris = row[4] || '';
                const cikis = row[6] || '';
                const icDis = row[9] || '';
                
                if (personelAdi && personelAdi.trim() !== '') {
                    if (mevcutBlok) {
                        bloklar.push(mevcutBlok);
                    }
                    
                    mevcutBlok = {
                        personel: personelAdi.trim(),
                        kayitlar: []
                    };
                }
                
                if (mevcutBlok && (giris || cikis)) {
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
            
            if (mevcutBlok) {
                bloklar.push(mevcutBlok);
            }
            
            return bloklar;
        }

        // Ana işleme
        const bloklar = personelBloklariniAyir(data);
        const sonuclar = [];
        const hatalar = [];
        
        bloklar.forEach(blok => {
            const { personel, kayitlar } = blok;
            const gunlukKayitlar = {};
            
            kayitlar.forEach(kayit => {
                const girisParsed = saatToDakika(kayit.giris);
                const cikisParsed = saatToDakika(kayit.cikis);
                
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
            
            Object.entries(gunlukKayitlar).forEach(([tarih, gunlukData]) => {
                const { girisler, cikislar, ic_dis } = gunlukData;
                
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
                
                const girisDakika = ilkGiris.saat * 60 + ilkGiris.dakika;
                const cikisDakika = sonCikis.saat * 60 + sonCikis.dakika;
                
                let calismaDakika = cikisDakika - girisDakika;
                if (calismaDakika < 0) {
                    calismaDakika += 24 * 60;
                }
                
                calismaDakika = Math.max(0, calismaDakika - 30);
                
                const planliCikisDakika = vardiya.kod === 'V1' ? 16 * 60 + 30 : 
                                        vardiya.kod === 'V2' ? 24 * 60 + 30 : 
                                        8 * 60 + 30;
                
                let fmDakika = 0;
                if (vardiya.kod === 'V2' || vardiya.kod === 'V3') {
                    if (cikisDakika > planliCikisDakika) {
                        fmDakika = Math.max(0, cikisDakika - planliCikisDakika - 20);
                    }
                } else {
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

        // Çıktı dosyalarını oluştur
        const outputDir = path.join(path.dirname(filePath), 'out');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir);
        }
        
        const jsonCikti = {
            islem_tarihi: new Date().toISOString(),
            dosya_yolu: filePath,
            toplam_kayit: sonuclar.length,
            hata_sayisi: hatalar.length,
            veriler: sonuclar
        };
        
        const jsonPath = path.join(outputDir, 'daily.json');
        const csvPath = path.join(outputDir, 'daily.csv');
        
        fs.writeFileSync(jsonPath, JSON.stringify(jsonCikti, null, 2), 'utf8');
        
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
        fs.writeFileSync(csvPath, csvIcerik, 'utf8');
        
        return {
            success: true,
            message: 'İşlem başarıyla tamamlandı!',
            data: {
                toplam_kayit: sonuclar.length,
                hata_sayisi: hatalar.length,
                personel_sayisi: bloklar.length,
                json_path: jsonPath,
                csv_path: csvPath
            },
            hatalar: hatalar
        };
        
    } catch (error) {
        return {
            success: false,
            message: 'İşlem sırasında hata oluştu: ' + error.message,
            error: error
        };
    }
});

// Rapor oluşturma
ipcMain.handle('generate-report', async (event, { type, data }) => {
    try {
        const reportPath = await generateExcelReport(type, data);
        return { success: true, filePath: reportPath, filename: path.basename(reportPath) };
    } catch (error) {
        console.error('Rapor oluşturma hatası:', error);
        return { success: false, message: error.message };
    }
});

// Excel rapor oluşturma fonksiyonu
async function generateExcelReport(reportType, data) {
    const outDir = path.join(__dirname, 'out');
    if (!fs.existsSync(outDir)) {
        fs.mkdirSync(outDir, { recursive: true });
    }
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    let filename = '';
    let reportData = [];
    
    switch (reportType) {
        case 'personel-ozet':
            filename = `Personel_Ozet_Raporu_${timestamp}.xlsx`;
            reportData = generatePersonelOzetReport(data);
            break;
        case 'vardiya-analiz':
            filename = `Vardiya_Analiz_Raporu_${timestamp}.xlsx`;
            reportData = generateVardiyaAnalizReport(data);
            break;
        case 'fm-raporu':
            filename = `Fazla_Mesai_Raporu_${timestamp}.xlsx`;
            reportData = generateFMReport(data);
            break;
        case 'gunluk-ozet':
            filename = `Gunluk_Ozet_Raporu_${timestamp}.xlsx`;
            reportData = generateGunlukOzetReport(data);
            break;
        case 'hata-raporu':
            filename = `Hata_Analiz_Raporu_${timestamp}.xlsx`;
            reportData = generateHataReport(data);
            break;
        case 'tum-veriler':
            filename = `Tum_Veriler_${timestamp}.xlsx`;
            reportData = generateTumVerilerReport(data);
            break;
        default:
            throw new Error('Bilinmeyen rapor türü');
    }
    
    const filePath = path.join(outDir, filename);
    
    // Excel dosyası oluştur
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(reportData);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Rapor');
    XLSX.writeFile(workbook, filePath);
    
    return filePath;
}

// Personel özet raporu
function generatePersonelOzetReport(data) {
    const personelMap = new Map();
    
    data.forEach(item => {
        if (!personelMap.has(item.personel)) {
            personelMap.set(item.personel, {
                personel: item.personel,
                toplam_gun: 0,
                toplam_calisma: 0,
                toplam_fm: 0,
                ortalama_calisma: 0,
                ortalama_fm: 0
            });
        }
        
        const personel = personelMap.get(item.personel);
        personel.toplam_gun++;
        personel.toplam_calisma += item.calisma_dk;
        personel.toplam_fm += parseFMToMinutes(item.fm_hhmm);
    });
    
    return Array.from(personelMap.values()).map(p => ({
        'Personel': p.personel,
        'Toplam Gün': p.toplam_gun,
        'Toplam Çalışma (dk)': p.toplam_calisma,
        'Toplam FM (dk)': p.toplam_fm,
        'Ortalama Çalışma (saat)': Math.round(p.toplam_calisma / p.toplam_gun / 60 * 100) / 100,
        'Ortalama FM (saat)': Math.round(p.toplam_fm / p.toplam_gun / 60 * 100) / 100
    }));
}

// Vardiya analiz raporu
function generateVardiyaAnalizReport(data) {
    const vardiyaMap = new Map();
    
    data.forEach(item => {
        const vardiya = item.vardiya ? item.vardiya.kod : 'Belirsiz';
        if (!vardiyaMap.has(vardiya)) {
            vardiyaMap.set(vardiya, {
                vardiya: vardiya,
                personel_sayisi: new Set(),
                toplam_calisma: 0,
                toplam_fm: 0,
                kayit_sayisi: 0
            });
        }
        
        const v = vardiyaMap.get(vardiya);
        v.personel_sayisi.add(item.personel);
        v.toplam_calisma += item.calisma_dk;
        v.toplam_fm += parseFMToMinutes(item.fm_hhmm);
        v.kayit_sayisi++;
    });
    
    return Array.from(vardiyaMap.values()).map(v => ({
        'Vardiya': v.vardiya,
        'Personel Sayısı': v.personel_sayisi.size,
        'Toplam Kayıt': v.kayit_sayisi,
        'Toplam Çalışma (saat)': Math.round(v.toplam_calisma / 60 * 100) / 100,
        'Toplam FM (saat)': Math.round(v.toplam_fm / 60 * 100) / 100,
        'Ortalama Çalışma (saat)': Math.round(v.toplam_calisma / v.kayit_sayisi / 60 * 100) / 100,
        'Ortalama FM (saat)': Math.round(v.toplam_fm / v.kayit_sayisi / 60 * 100) / 100
    }));
}

// Fazla mesai raporu
function generateFMReport(data) {
    return data
        .filter(item => parseFMToMinutes(item.fm_hhmm) > 0)
        .map(item => ({
            'Personel': item.personel,
            'Tarih': item.tarih,
            'Vardiya': item.vardiya ? item.vardiya.kod : '',
            'Çalışma Süresi': item.calisma_dk,
            'Fazla Mesai': item.fm_hhmm,
            'FM (dakika)': parseFMToMinutes(item.fm_hhmm),
            'Durum': item.durum
        }))
        .sort((a, b) => parseFMToMinutes(b['Fazla Mesai']) - parseFMToMinutes(a['Fazla Mesai']));
}

// Günlük özet raporu
function generateGunlukOzetReport(data) {
    const gunlukMap = new Map();
    
    data.forEach(item => {
        if (!gunlukMap.has(item.tarih)) {
            gunlukMap.set(item.tarih, {
                tarih: item.tarih,
                personel_sayisi: new Set(),
                toplam_calisma: 0,
                toplam_fm: 0,
                kayit_sayisi: 0
            });
        }
        
        const gun = gunlukMap.get(item.tarih);
        gun.personel_sayisi.add(item.personel);
        gun.toplam_calisma += item.calisma_dk;
        gun.toplam_fm += parseFMToMinutes(item.fm_hhmm);
        gun.kayit_sayisi++;
    });
    
    return Array.from(gunlukMap.values())
        .sort((a, b) => new Date(a.tarih) - new Date(b.tarih))
        .map(g => ({
            'Tarih': g.tarih,
            'Personel Sayısı': g.personel_sayisi.size,
            'Toplam Kayıt': g.kayit_sayisi,
            'Toplam Çalışma (saat)': Math.round(g.toplam_calisma / 60 * 100) / 100,
            'Toplam FM (saat)': Math.round(g.toplam_fm / 60 * 100) / 100
        }));
}

// Hata raporu
function generateHataReport(data) {
    return data
        .filter(item => item.durum !== 'ok')
        .map(item => ({
            'Personel': item.personel,
            'Tarih': item.tarih,
            'Vardiya': item.vardiya ? item.vardiya.kod : '',
            'Giriş': item.gercek ? item.gercek.gir : '',
            'Çıkış': item.gercek ? item.gercek.cik : '',
            'Durum': item.durum,
            'Not': item.not || ''
        }));
}

// Tüm veriler raporu
function generateTumVerilerReport(data) {
    return data.map(item => ({
        'Personel': item.personel,
        'Tarih': item.tarih,
        'İç/Dış': item.ic_dis || '',
        'Vardiya': item.vardiya ? item.vardiya.kod : '',
        'Plan Başlangıç': item.vardiya ? item.vardiya.plan_bas : '',
        'Plan Bitiş': item.vardiya ? item.vardiya.plan_bit : '',
        'Giriş': item.gercek ? item.gercek.gir : '',
        'Çıkış': item.gercek ? item.gercek.cik : '',
        'Çalışma (dk)': item.calisma_dk,
        'FM (dk)': item.fm_dk,
        'FM (saat)': item.fm_hhmm,
        'Durum': item.durum,
        'Not': item.not || ''
    }));
}

// FM string'ini dakikaya çevir
function parseFMToMinutes(fmString) {
    if (!fmString || fmString === '00:00') return 0;
    const [hours, minutes] = fmString.split(':').map(Number);
    return hours * 60 + minutes;
}
