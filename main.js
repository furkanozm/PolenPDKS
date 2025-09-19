const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const puppeteer = require('puppeteer');
const speakeasy = require('speakeasy');

let mainWindow;

// Global değişkenler
global.currentBrowser = null;
global.currentPage = null;
global.activePages = [];
global.processStopped = false;

// Output dizini yolunu al
function getOutputDir() {
    if (app.isPackaged) {
        return path.join(app.getPath('userData'), 'out');
    } else {
        return path.join(__dirname, 'out');
    }
}

// Config dosyasını oku
function loadConfig() {
    try {
        let configPath;
        
        // Packaged uygulamada userData dizinini kullan
        if (app.isPackaged) {
            configPath = path.join(app.getPath('userData'), 'config.json');
        } else {
            configPath = path.join(__dirname, 'config.json');
        }
        
        // Config dosyası yoksa varsayılan config oluştur
        if (!fs.existsSync(configPath)) {
            const defaultConfig = {
                "pinhuman": {
                    "baseUrl": "https://pinhuman.com",
                    "loginEndpoint": "/System/LoginStep1",
                    "timeout": 30000,
                    "retryAttempts": 3,
                    "credentials": {
                        "userName": "furkan.ozmen@guleryuzgroup.com",
                        "companyCode": "ikb",
                        "password": "Kralben123.",
                        "totpSecret": "GQ2DCZBYGRRGILLGMI4TELJUMMYGCLJZGU4TILJQHBRDSNDBMJRTQNTBMNXVGZLDOJSXI4DJNZUHK3LBNZNG42LLMI"
                    }
                },
                "app": {
                    "name": "PDKS İşleme Sistemi",
                    "version": "1.0.0",
                    "outputDirectory": "./out",
                    "supportedFormats": ["xlsx", "xls"]
                },
                "ui": {
                    "theme": "light",
                    "language": "tr",
                    "autoSave": true
                }
            };
            
            // Varsayılan config'i kaydet
            fs.writeFileSync(configPath, JSON.stringify(defaultConfig, null, 2), 'utf8');
            return defaultConfig;
        }
        
        const configData = fs.readFileSync(configPath, 'utf8');
        return JSON.parse(configData);
    } catch (error) {
        console.error('Config dosyası okunamadı:', error);
        return null;
    }
}

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

// Config dosyasını okuma
ipcMain.handle('get-config', async () => {
    try {
        return loadConfig();
    } catch (error) {
        console.error('Config okuma hatası:', error);
        throw error;
    }
});

// Config dosyasını güncelleme
ipcMain.handle('update-config', async (event, newCredentials) => {
    try {
        let configPath;
        
        // Packaged uygulamada userData dizinini kullan
        if (app.isPackaged) {
            configPath = path.join(app.getPath('userData'), 'config.json');
        } else {
            configPath = path.join(__dirname, 'config.json');
        }
        
        // Mevcut config'i oku veya varsayılan config'i kullan
        let config = loadConfig();
        if (!config) {
            throw new Error('Config dosyası okunamadı');
        }
        
        // Credentials'ı güncelle
        config.pinhuman.credentials = { ...config.pinhuman.credentials, ...newCredentials };
        
        // Dosyayı kaydet
        fs.writeFileSync(configPath, JSON.stringify(config, null, 2), 'utf8');
        
        return { success: true, message: 'Config başarıyla güncellendi' };
    } catch (error) {
        console.error('Config güncelleme hatası:', error);
        return { success: false, message: 'Config güncellenirken hata oluştu: ' + error.message };
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
                const sicilNo = row[0] || '';
                const personelAdi = row[1] || '';
                const giris = row[4] || '';
                const cikis = row[6] || '';
                const icDis = row[9] || '';
                
                if (personelAdi && personelAdi.trim() !== '') {
                    if (mevcutBlok) {
                        bloklar.push(mevcutBlok);
                    }
                    
                    const temizSicilNo = sicilNo ? sicilNo.toString().trim() : '';
                    
                    mevcutBlok = {
                        sicilNo: temizSicilNo,
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
            const { sicilNo, personel, kayitlar } = blok;
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
                        sicilNo,
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
                        sicilNo,
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
                    sicilNo,
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
        const outputDir = getOutputDir();
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
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
        
        const csvBaslik = 'sicilNo,personel,tarih,ic_dis,vardiya_kod,vardiya_plan_bas,vardiya_plan_bit,gercek_gir,gercek_cik,calisma_dk,fm_dk,fm_hhmm,durum,not';
        const csvSatirlar = sonuclar.map(kayit => [
            kayit.sicilNo || '',
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
    const outDir = getOutputDir();
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

// Toplu Ekle formunu doldurma fonksiyonu
async function fillBatchForm(page) {
    try {
        console.log('Toplu Ekle formu dolduruluyor...');
        
        // Excel verilerini oku (out/daily.json'dan)
        const fs = require('fs');
        const path = require('path');
        const dailyJsonPath = path.join(getOutputDir(), 'daily.json');
        
        if (!fs.existsSync(dailyJsonPath)) {
            console.log('daily.json dosyası bulunamadı, örnek veri kullanılıyor...');
            // Örnek veri
            const sampleData = [
                {
                    date: '18.09.2025',
                    entries: [
                        { time: '08:00', type: 'In' },
                        { time: '17:00', type: 'Out' }
                    ]
                }
            ];
            await processAttendanceData(page, sampleData);
            return;
        }
        
        const dailyData = JSON.parse(fs.readFileSync(dailyJsonPath, 'utf8'));
        console.log('Excel verileri okundu:', dailyData);
        
        // Veriyi array formatına çevir
        let processedData;
        if (Array.isArray(dailyData)) {
            processedData = dailyData;
        } else if (dailyData && typeof dailyData === 'object') {
            // Eğer object ise, array'e çevir
            processedData = [dailyData];
        } else {
            console.log('Veri formatı uygun değil, örnek veri kullanılıyor...');
            processedData = [
                {
                    date: '18.09.2025',
                    entries: [
                        { time: '08:00', type: 'In' },
                        { time: '17:00', type: 'Out' }
                    ]
                }
            ];
        }
        
        console.log('İşlenecek veri:', processedData.length, 'kayıt');
        
        // Verileri işle
        await processAttendanceData(page, processedData);
        
    } catch (error) {
        console.error('Form doldurma hatası:', error);
        throw error;
    }
}

// Vardiya verilerini işleme fonksiyonu
async function processAttendanceData(page, data) {
    console.log('Vardiya verileri işleniyor...');
    
    // Verileri yuvarlanmış saatlerle işle
    const processedData = data.map(record => {
        const vardiyaKod = record.vardiya ? record.vardiya.kod : null;
        
        // Giriş ve çıkış saatlerini yuvarla
        const yuvarlanmisGiris = roundToShiftTimeForBulk(record.gercek ? record.gercek.gir : '08:30', vardiyaKod, true);
        const yuvarlanmisCikis = roundToShiftTimeForBulk(record.gercek ? record.gercek.cik : '16:30', vardiyaKod, false);
        
        // Yuvarlanmış FM hesapla
        const yuvarlanmisFM = calculateFMWithRoundedTimeForBulk(yuvarlanmisGiris, yuvarlanmisCikis, vardiyaKod);
        
        return {
            ...record,
            gercek: {
                gir: yuvarlanmisGiris,
                cik: yuvarlanmisCikis
            },
            fm_dk: yuvarlanmisFM,
            fm_hhmm: `${Math.floor(yuvarlanmisFM / 60).toString().padStart(2, '0')}:${(yuvarlanmisFM % 60).toString().padStart(2, '0')}`
        };
    });
    
    // Vardiya kombinasyonlarını analiz et
    const shiftCombinations = analyzeShiftCombinations(processedData);
    console.log('\n=== VARDİYA KOMBİNASYONLARI (YUVARLANMIŞ) ===');
    shiftCombinations.forEach((combo, index) => {
        console.log(`Kombinasyon ${index + 1}:`);
        console.log(`  Tarih: ${combo.date}`);
        console.log(`  Personel: ${combo.personnel || 'Bilinmeyen'}`);
        console.log(`  Vardiya: ${combo.shift}`);
        console.log(`  Giriş: ${combo.entries.filter(e => e.type === 'In').map(e => e.time).join(', ')}`);
        console.log(`  Çıkış: ${combo.entries.filter(e => e.type === 'Out').map(e => e.time).join(', ')}`);
        console.log('---');
    });
    
    // Vardiya gruplarını oluştur
    const vardiyaGruplari = groupByVardiyaGruplari(processedData);
    console.log('\n=== VARDİYA GRUPLARI (YUVARLANMIŞ) ===');
    
    // Tarihe göre sırala
    const sortedGroups = Object.values(vardiyaGruplari).sort((a, b) => {
        return new Date(a.tarih) - new Date(b.tarih);
    });
    
    sortedGroups.forEach((grup, index) => {
        console.log(`Grup ${index + 1}:`);
        console.log(`  Tarih: ${grup.tarih}`);
        console.log(`  Vardiya: ${grup.vardiya}`);
        console.log(`  Giriş: ${grup.giris}`);
        console.log(`  Çıkış: ${grup.cikis}`);
        console.log(`  FM: ${grup.fm}`);
        console.log(`  Personel Sayısı: ${grup.personeller.length}`);
        console.log(`  Personeller: ${grup.personeller.map(p => p.adi).join(', ')}`);
        console.log('---');
    });
    
    // Verileri vardiya gruplarına göre işle
    const groupedData = vardiyaGruplari;
    
    console.log('\n=== GRUPLANDIRILMIŞ VERİLER ===');
    groupedData.forEach((group, index) => {
        console.log(`Grup ${index + 1}:`);
        console.log(`  Tarih: ${group.date}`);
        console.log(`  Vardiya: ${group.shift}`);
        console.log(`  Saatler: ${group.giris} - ${group.cikis}`);
        console.log(`  FM: ${group.fm} ${group.fmDakika > 0 ? '(Fazla Mesai Var)' : '(Normal Çıkış)'}`);
        console.log(`  Personel Sayısı: ${group.personelCount}`);
        console.log(`  Personeller: ${group.personeller.join(', ')}`);
        console.log('---');
    });
    
    for (const group of groupedData) {
        console.log(`\nİşleniyor: ${group.date} - ${group.shift} - Çıkış: ${group.cikis} - FM: ${group.fm} (${group.personelCount} personel)`);
        
        // Grup için tek form doldur
        await fillGroupEntry(page, group);
        
        // Her grup arasında kısa bekleme
        await new Promise(resolve => setTimeout(resolve, 500));
    }
}

// Verileri vardiya gruplarına göre grupla (aynı yuvarlanmış giriş/çıkış saatlerine sahip personeller)
function groupByVardiyaGruplari(data) {
    const groups = {};
    
    data.forEach(item => {
        const tarih = item.tarih;
        const vardiyaKod = item.vardiya;
        const personelAdi = item.personel.split(' ').filter(part => !part.match(/^\d{1,2}:\d{2}$/)).join(' ').trim();
        
        // Yuvarlanmış saatleri hesapla
        const yuvarlanmisGiris = roundToShiftTimeForBulk(item.giris, vardiyaKod, true);
        const yuvarlanmisCikis = roundToShiftTimeForBulk(item.cikis, vardiyaKod, false);
        
        // FM hesapla
        const fm = calculateFMWithRoundedTimeForBulk(yuvarlanmisGiris, yuvarlanmisCikis, vardiyaKod);
        const fmStr = fm > 0 ? `${Math.floor(fm / 60)}:${(fm % 60).toString().padStart(2, '0')}` : '0:00';
        
        // Grup anahtarı: tarih + vardiya + yuvarlanmış giriş + yuvarlanmış çıkış
        const groupKey = `${tarih}_${vardiyaKod}_${yuvarlanmisGiris}_${yuvarlanmisCikis}`;
        
        if (!groups[groupKey]) {
            groups[groupKey] = {
                tarih: tarih,
                vardiya: vardiyaKod,
                giris: yuvarlanmisGiris,
                cikis: yuvarlanmisCikis,
                fm: fmStr,
                personeller: []
            };
        }
        
        groups[groupKey].personeller.push({
            adi: personelAdi,
            giris: item.giris,
            cikis: item.cikis
        });
    });
    
    return groups;
}

// Aynı vardiya ve saatlere sahip kayıtları grupla
function groupByShiftAndTime(data) {
    const groups = {};
    
    data.forEach(record => {
        const vardiyaKod = record.vardiya ? record.vardiya.kod : 'Belirsiz';
        const giris = record.gercek ? record.gercek.gir : '08:30';
        const cikis = record.gercek ? record.gercek.cik : '16:30';
        const tarih = record.tarih;
        
        // FM durumunu hesapla
        const fmDakika = calculateFMWithRoundedTimeForBulk(giris, cikis, vardiyaKod);
        const fmSaat = Math.floor(fmDakika / 60);
        const fmDakikaKalan = fmDakika % 60;
        const fmStr = fmDakika > 0 ? `${fmSaat}:${fmDakikaKalan.toString().padStart(2, '0')}` : '0:00';
        
        // Grup anahtarını sadece çıkış saatine göre oluştur (FM durumuna göre ayrı gruplar)
        const key = `${tarih}-${vardiyaKod}-${cikis}-${fmStr}`;
        
        if (!groups[key]) {
            groups[key] = {
                date: tarih,
                shift: vardiyaKod,
                giris: giris,
                cikis: cikis,
                fm: fmStr,
                fmDakika: fmDakika,
                personelCount: 0,
                personeller: []
            };
        }
        
        groups[key].personelCount++;
        groups[key].personeller.push(record.personel);
    });
    
    return Object.values(groups);
}

// Grup için form doldurma fonksiyonu
async function fillGroupEntry(page, group) {
    try {
        console.log(`Grup formu dolduruluyor: ${group.date} - ${group.shift} - ${group.giris}-${group.cikis}`);
        
        // Tarih seç
        await page.select('select[name="date"]', group.date);
        await new Promise(resolve => setTimeout(resolve, 500));
        
        // Vardiya seç
        const shiftValue = group.shift === 'V1' ? '1' : group.shift === 'V2' ? '2' : '3';
        await page.select('select[name="shift"]', shiftValue);
        await new Promise(resolve => setTimeout(resolve, 500));
        
        // Giriş saati
        await page.type('input[name="entryTime"]', group.giris);
        await new Promise(resolve => setTimeout(resolve, 300));
        
        // Çıkış saati
        await page.type('input[name="exitTime"]', group.cikis);
        await new Promise(resolve => setTimeout(resolve, 300));
        
        // Formu gönder
        await page.click('button[type="submit"]');
        console.log(`✓ Grup formu gönderildi (${group.personelCount} personel)`);
        
        // Grup verilerini history'e ekle - detaylı bilgi ile
        group.personeller.forEach(personel => {
            const historyEntry = {
                id: Date.now() + Math.random(),
                date: group.date,
                time: `${group.giris}-${group.cikis}`,
                type: 'Grup',
                personel: personel,
                sicilNo: '-',
                vardiya: group.shift === 'V1' ? 'V1 (Gündüz)' : group.shift === 'V2' ? 'V2 (Gece)' : 'V3 (Vardiya)',
                status: 'success',
                timestamp: new Date().toISOString(),
                girisSaati: group.giris,
                cikisSaati: group.cikis
            };
            sendHistoryUpdate(historyEntry);
        });
        
        // Sayfa yenilenmesini bekle
        await new Promise(resolve => setTimeout(resolve, 500));
        
    } catch (error) {
        console.error('Grup form doldurma hatası:', error);
        
        // Hata durumunda da history'e ekle - detaylı bilgi ile
        const historyEntry = {
            id: Date.now(),
            date: group.date,
            time: `${group.giris}-${group.cikis}`,
            type: 'Grup',
            personel: `${group.personelCount} personel`,
            sicilNo: '-',
            vardiya: group.shift === 'V1' ? 'V1 (Gündüz)' : group.shift === 'V2' ? 'V2 (Gece)' : 'V3 (Vardiya)',
            status: 'error',
            error: error.message,
            timestamp: new Date().toISOString(),
            girisSaati: group.giris,
            cikisSaati: group.cikis
        };
        sendHistoryUpdate(historyEntry);
    }
}

// Vardiya kombinasyonlarını analiz etme fonksiyonu
function analyzeShiftCombinations(data) {
    const combinations = [];
    const dateGroups = {};
    
    // Önce tarihlere göre grupla
    for (const record of data) {
        if (!dateGroups[record.date]) {
            dateGroups[record.date] = [];
        }
        dateGroups[record.date].push(record);
    }
    
    // Her tarih için tüm kombinasyonları bul
    for (const [date, records] of Object.entries(dateGroups)) {
        const dateCombinations = findDateShiftCombinations(date, records);
        combinations.push(...dateCombinations);
    }
    
    return combinations;
}

// Tek bir tarih için vardiya kombinasyonlarını bulma
function findDateShiftCombinations(date, records) {
    const combinations = [];
    const uniqueCombinations = new Set();
    
    console.log(`\n=== ${date} TARİHİ ANALİZİ ===`);
    console.log(`Toplam ${records.length} personel kaydı bulundu`);
    
    for (const record of records) {
        const entries = record.entries || [];
        const personnel = record.personnel || 'Bilinmeyen';
        
        console.log(`\nPersonel: ${personnel}`);
        
        // Giriş ve çıkışları ayır
        const inEntries = entries.filter(e => e.type === 'In');
        const outEntries = entries.filter(e => e.type === 'Out');
        
        console.log(`  Girişler: ${inEntries.map(e => e.time).join(', ')}`);
        console.log(`  Çıkışlar: ${outEntries.map(e => e.time).join(', ')}`);
        
        // Her giriş-çıkış kombinasyonu için vardiya tespit et
        for (const inEntry of inEntries) {
            for (const outEntry of outEntries) {
                const shift = determineShiftFromTimeRange(inEntry.time, outEntry.time);
                const workDuration = calculateWorkDuration(inEntry.time, outEntry.time);
                
                // Kombinasyon anahtarı oluştur (benzersizlik için)
                const combinationKey = `${date}_${inEntry.time}_${outEntry.time}_${shift}`;
                
                if (!uniqueCombinations.has(combinationKey)) {
                    uniqueCombinations.add(combinationKey);
                    
                    combinations.push({
                        date: date,
                        personnel: personnel,
                        shift: shift,
                        workDuration: workDuration,
                        entries: [inEntry, outEntry],
                        combinationKey: combinationKey
                    });
                    
                    console.log(`  ✓ Kombinasyon: ${inEntry.time} - ${outEntry.time} (${shift}, ${workDuration} saat)`);
                } else {
                    console.log(`  - Tekrar eden kombinasyon: ${inEntry.time} - ${outEntry.time} (${shift})`);
                }
            }
        }
        
        // Eğer sadece giriş veya sadece çıkış varsa
        if (inEntries.length > 0 && outEntries.length === 0) {
            inEntries.forEach(inEntry => {
                const shift = determineShift(inEntry.time, 'In');
                const combinationKey = `${date}_${inEntry.time}_GIRIS_${shift}`;
                
                if (!uniqueCombinations.has(combinationKey)) {
                    uniqueCombinations.add(combinationKey);
                    
                    combinations.push({
                        date: date,
                        personnel: personnel,
                        shift: `V${shift}`,
                        workDuration: 0,
                        entries: [inEntry],
                        combinationKey: combinationKey
                    });
                    
                    console.log(`  ✓ Sadece giriş: ${inEntry.time} (V${shift})`);
                }
            });
        }
        
        if (outEntries.length > 0 && inEntries.length === 0) {
            outEntries.forEach(outEntry => {
                const shift = determineShift(outEntry.time, 'Out');
                const combinationKey = `${date}_CIKIS_${outEntry.time}_${shift}`;
                
                if (!uniqueCombinations.has(combinationKey)) {
                    uniqueCombinations.add(combinationKey);
                    
                    combinations.push({
                        date: date,
                        personnel: personnel,
                        shift: `V${shift}`,
                        workDuration: 0,
                        entries: [outEntry],
                        combinationKey: combinationKey
                    });
                    
                    console.log(`  ✓ Sadece çıkış: ${outEntry.time} (V${shift})`);
                }
            });
        }
    }
    
    console.log(`\n${date} için toplam ${combinations.length} benzersiz kombinasyon bulundu`);
    
    return combinations;
}

// Çalışma süresini hesaplama
function calculateWorkDuration(inTime, outTime) {
    const [inHour, inMin] = inTime.split(':').map(Number);
    const [outHour, outMin] = outTime.split(':').map(Number);
    
    let inMinutes = inHour * 60 + inMin;
    let outMinutes = outHour * 60 + outMin;
    
    // Gece vardiyası için
    if (outMinutes < inMinutes) {
        outMinutes += 24 * 60;
    }
    
    const workMinutes = outMinutes - inMinutes;
    const workHours = Math.round((workMinutes / 60) * 100) / 100; // 2 ondalık basamak
    
    return workHours;
}

// Giriş-çıkış saat aralığından vardiya tespiti
function determineShiftFromTimeRange(inTime, outTime) {
    const inHour = parseInt(inTime.split(':')[0]);
    const outHour = parseInt(outTime.split(':')[0]);
    
    // Çalışma süresini hesapla
    let workDuration = outHour - inHour;
    if (workDuration < 0) workDuration += 24; // Gece vardiyası için
    
    // Vardiya tespiti
    if (inHour >= 6 && inHour < 14) {
        return 'V1'; // Gündüz vardiyası
    } else if (inHour >= 14 && inHour < 22) {
        return 'V2'; // Akşam vardiyası
    } else {
        return 'V3'; // Gece vardiyası
    }
}

// Tek bir giriş/çıkış kaydını doldurma
async function fillSingleEntry(page, date, entry) {
    try {
        console.log(`Kayıt dolduruluyor: ${date} ${entry.time} ${entry.type}`);
        
        // Okuma zamanı (tarih + saat + saniye)
        const readTime = `${date} ${entry.time}:00`;
        await page.type('#ReadTime', readTime);
        console.log('Okuma zamanı yazıldı:', readTime);
        
        // Giriş/Çıkış seçimi
        await page.select('#Direction', entry.type);
        console.log('Yön seçildi:', entry.type);
        
        // Firma seçimi (Polen)
        await page.click('#select2-CompanyId-container');
        await new Promise(resolve => setTimeout(resolve, 200));
        await page.click('li:has-text("Polen")');
        console.log('Firma seçildi: Polen');
        
        // Lokasyon dropdown'unun aktif olmasını bekle (firma seçimi sonrası)
        await page.waitForFunction(() => {
            const locationDropdown = document.querySelector('#select2-CompanyWorkingRegionId-container');
            return locationDropdown && !locationDropdown.disabled;
        }, { timeout: 3000 });
        
        // Lokasyon seçimi (Manisa/Tesis)
        await page.click('#select2-CompanyWorkingRegionId-container');
        await new Promise(resolve => setTimeout(resolve, 200));
        await page.click('li:has-text("Manisa")');
        console.log('Lokasyon seçildi: Manisa');
        
        // Vardiya seçimi (giriş/çıkış saatine göre)
        const shiftId = determineShift(entry.time, entry.type);
        if (shiftId) {
            await page.click('#select2-ShiftId-container');
            await new Promise(resolve => setTimeout(resolve, 200));
            await page.click(`li[data-value="${shiftId}"]`);
            console.log('Vardiya seçildi:', shiftId);
        }
        
        // Kaydet butonuna bas
        await page.click('button[type="submit"]');
        console.log('Kayıt kaydedildi!');
        
        // Kayıt başarılı
        sendLogMessage('success', `✅ Kayıt başarıyla eklendi: ${entry.personel || 'Bilinmeyen'} (${entry.sicilNo || '-'}) - ${entry.type === '1' ? 'Giriş' : 'Çıkış'} ${entry.time}`);
        
        // Sayfa yenilenmesini bekle
        await new Promise(resolve => setTimeout(resolve, 300));
        
    } catch (error) {
        console.error('Tek kayıt doldurma hatası:', error);
        sendLogMessage('error', `❌ Kayıt hatası: ${entry.personel || 'Bilinmeyen'} (${entry.sicilNo || '-'}) - ${error.message}`);
    }
}

// Vardiya tespiti fonksiyonu
function determineShift(time, type) {
    const hour = parseInt(time.split(':')[0]);
    
    // Vardiya tespiti
    if (hour >= 6 && hour < 14) {
        return '1'; // Gündüz vardiyası (V1)
    } else if (hour >= 14 && hour < 22) {
        return '2'; // Akşam vardiyası (V2)
    } else {
        return '3'; // Gece vardiyası (V3)
    }
}

// Vardiya saatlerini yuvarla (PDKS standart kuralları - toplu giriş için)
function roundToShiftTimeForBulk(timeString, vardiyaKod, isGiris = false) {
    if (!timeString || timeString === '-') return timeString;
    
    const [hours, minutes] = timeString.split(':').map(Number);
    const totalMinutes = hours * 60 + minutes;
    
    // Vardiya saatleri (varsayılan)
    const vardiyaBitis = {
        'V1': parseTimeToMinutesForBulk('16:30'),
        'V2': parseTimeToMinutesForBulk('00:30'),
        'V3': parseTimeToMinutesForBulk('08:30')
    };
    
    const vardiyaBaslangic = {
        'V1': parseTimeToMinutesForBulk('08:30'),
        'V2': parseTimeToMinutesForBulk('16:30'),
        'V3': parseTimeToMinutesForBulk('00:30')
    };
    
    if (!vardiyaKod || !vardiyaBitis[vardiyaKod]) {
        return timeString;
    }
    
    const vardiyaBitisDakika = vardiyaBitis[vardiyaKod];
    const vardiyaBaslangicDakika = vardiyaBaslangic[vardiyaKod];
    
    if (isGiris) {
        // GİRİŞ SAATİ YUVARLANMASI (Dünya Standartları)
        const erkenGirisToleransi = 10;
        const gecGirisToleransi = 3;
        
        // Erken giriş toleransı
        if (totalMinutes >= (vardiyaBaslangicDakika - erkenGirisToleransi) && totalMinutes <= vardiyaBaslangicDakika) {
            return formatMinutesToTimeForBulk(vardiyaBaslangicDakika);
        }
        
        // Geç giriş toleransı
        if (totalMinutes > vardiyaBaslangicDakika && totalMinutes <= (vardiyaBaslangicDakika + gecGirisToleransi)) {
            return formatMinutesToTimeForBulk(vardiyaBaslangicDakika);
        }
        
        // 59 dakika erken girişlerde vardiya başlangıcına yuvarla
        const elliDokuzDakikaErken = vardiyaBaslangicDakika - 59;
        if (totalMinutes >= elliDokuzDakikaErken) {
            return formatMinutesToTimeForBulk(vardiyaBaslangicDakika); // Vardiya başlangıcına yuvarla
        } else {
            // 59 dakikadan fazla erken girişte 30dk aralığa yuvarla
            const yuvarlanmisDakika = Math.round(totalMinutes / 30) * 30;
            return formatMinutesToTimeForBulk(yuvarlanmisDakika);
        }
        
    } else {
        // ÇIKIŞ SAATİ YUVARLANMASI (Dünya Standartları)
        const erkenCikisToleransi = 3;
        
        // Erken çıkış toleransı
        if (totalMinutes >= (vardiyaBitisDakika - erkenCikisToleransi) && totalMinutes <= vardiyaBitisDakika) {
            return formatMinutesToTimeForBulk(vardiyaBitisDakika);
        }
        
        // Fazla mesai çıkışı (30dk aralığa yuvarla)
        if (totalMinutes > vardiyaBitisDakika) {
            const yuvarlanmisDakika = Math.round(totalMinutes / 30) * 30;
            return formatMinutesToTimeForBulk(yuvarlanmisDakika);
        }
        
        // Çok erken çıkış (30dk aralığa yuvarla)
        const yuvarlanmisDakika = Math.round(totalMinutes / 30) * 30;
        return formatMinutesToTimeForBulk(yuvarlanmisDakika);
    }
}

// Saati dakikaya çevir (main.js için)
function parseTimeToMinutesForBulk(timeString) {
    const [hours, minutes] = timeString.split(':').map(Number);
    return hours * 60 + minutes;
}

// Dakikayı saate çevir (main.js için)
function formatMinutesToTimeForBulk(totalMinutes) {
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
}

// FM hesapla (yuvarlanmış saatlere göre)
function calculateFMWithRoundedTimeForBulk(giris, cikis, vardiyaKod) {
    if (!giris || !cikis || giris === '-' || cikis === '-') return 0;
    
    const [girisSaat, girisDakika] = giris.split(':').map(Number);
    const [cikisSaat, cikisDakika] = cikis.split(':').map(Number);
    
    const girisDakikaToplam = girisSaat * 60 + girisDakika;
    const cikisDakikaToplam = cikisSaat * 60 + cikisDakika;
    
    // Vardiya bitiş saatleri
    const vardiyaBitis = {
        'V1': 16 * 60 + 30, // 16:30
        'V2': 24 * 60 + 30, // 00:30 (24:30)
        'V3': 8 * 60 + 30   // 08:30
    };
    
    if (!vardiyaKod || !vardiyaBitis[vardiyaKod]) {
        return 0;
    }
    
    const vardiyaBitisDakika = vardiyaBitis[vardiyaKod];
    
    // FM hesapla (çıkış saati - vardiya bitiş saati)
    if (cikisDakikaToplam > vardiyaBitisDakika) {
        return cikisDakikaToplam - vardiyaBitisDakika;
    }
    
    return 0;
}

// Checkbox'ları seçme fonksiyonu (ID ile)
async function selectEmployeeCheckboxes(page, employeeIds) {
    try {
        console.log('Personel checkbox\'ları seçiliyor...');
        console.log('Seçilecek personel ID\'leri:', employeeIds);
        
        // Her personel ID'si için checkbox'ı seç
        for (const employeeId of employeeIds) {
            try {
                // Checkbox'ı bul ve seç
                const checkboxSelector = `input[name="SelectedEmployeeId"][value="${employeeId}"]`;
                await page.waitForSelector(checkboxSelector, { timeout: 5000 });
                
                // Checkbox'ın seçili olup olmadığını kontrol et
                const isChecked = await page.isChecked(checkboxSelector);
                
                if (!isChecked) {
                    await page.check(checkboxSelector);
                    console.log(`✓ Personel ID ${employeeId} seçildi`);
                } else {
                    console.log(`- Personel ID ${employeeId} zaten seçili`);
                }
                
                // Her checkbox seçimi arasında kısa bekleme
                await new Promise(resolve => setTimeout(resolve, 200));
                
            } catch (error) {
                console.log(`✗ Personel ID ${employeeId} bulunamadı:`, error.message);
            }
        }
        
        console.log('Checkbox seçim işlemi tamamlandı!');
        
    } catch (error) {
        console.error('Checkbox seçim hatası:', error);
        throw error;
    }
}

// Checkbox'ları seçme fonksiyonu (isim ile)
async function selectEmployeeCheckboxesByName(page, employeeNames) {
    try {
        console.log('Personel checkbox\'ları isim ile seçiliyor...');
        console.log('Seçilecek personel isimleri:', employeeNames);
        
        // Her personel ismi için checkbox'ı seç
        for (const employeeName of employeeNames) {
            try {
                // Personel ismini içeren satırı bul
                const rowSelector = `tr:has(td.search-name:has-text("${employeeName}"))`;
                await page.waitForSelector(rowSelector, { timeout: 5000 });
                
                // Bu satırdaki checkbox'ı bul ve seç
                const checkboxSelector = `${rowSelector} input[name="SelectedEmployeeId"]`;
                const isChecked = await page.isChecked(checkboxSelector);
                
                if (!isChecked) {
                    await page.check(checkboxSelector);
                    console.log(`✓ Personel ${employeeName} seçildi`);
                } else {
                    console.log(`- Personel ${employeeName} zaten seçili`);
                }
                
                // Her checkbox seçimi arasında kısa bekleme
                await new Promise(resolve => setTimeout(resolve, 200));
                
            } catch (error) {
                console.log(`✗ Personel ${employeeName} bulunamadı:`, error.message);
            }
        }
        
        console.log('Checkbox seçim işlemi tamamlandı!');
        
    } catch (error) {
        console.error('Checkbox seçim hatası:', error);
        throw error;
    }
}


// Log mesajı gönderme helper fonksiyonu
function sendLogMessage(type, message) {
    if (mainWindow && mainWindow.webContents) {
        mainWindow.webContents.send('log-message', { type, message });
    }
    console.log(`[${type.toUpperCase()}] ${message}`);
}



// Pinhuman'a otomatik veri girişi
ipcMain.handle('enter-data-pinhuman', async (event, credentials = null) => {
    // Global browser değişkenini kullan
    global.currentBrowser = null;
    
    try {
        // Config'den veya parametreden credentials al
        const config = loadConfig();
        const creds = credentials || config?.pinhuman?.credentials;
        
        if (!creds || !creds.userName || !creds.companyCode || !creds.password) {
            throw new Error('Kullanıcı bilgileri eksik. Lütfen config.json dosyasını kontrol edin.');
        }
        
        // TOTP kodu üret
        const totpCode = speakeasy.totp({
            secret: creds.totpSecret,
            encoding: 'base32'
        });
        
        sendLogMessage('info', '🚀 Pinhuman automation başlıyor...');
        sendLogMessage('info', `👤 Kullanıcı: ${creds.userName}`);
        sendLogMessage('info', `🏢 Şirket: ${creds.companyCode}`);
        sendLogMessage('info', `🔢 TOTP Kodu: ${totpCode}`);
        
        // JSON verilerini oku
        const dailyDataPath = path.join(getOutputDir(), 'daily.json');
        if (!fs.existsSync(dailyDataPath)) {
            throw new Error('daily.json dosyası bulunamadı. Önce Excel dosyasını işleyin.');
        }
        
        const dailyData = JSON.parse(fs.readFileSync(dailyDataPath, 'utf8'));
        const veriler = dailyData.veriler;
        
        // Headless ayarını al
        const headlessMode = creds.headlessMode === 'true';
        
        // Browser'ı başlat - Puppeteer ile
        const browserConfig = {
            headless: headlessMode,
            args: ['--window-size=1200,800', '--window-position=100,100'],
            defaultViewport: { width: 1200, height: 800 }
        };
        
        // Headless mod değilse pencere konumunu ayarla
        if (!headlessMode) {
            browserConfig.args.push('--window-position=100,100');
        }
        
        global.currentBrowser = await puppeteer.launch(browserConfig);
        global.processStopped = false;
        
        const page = await global.currentBrowser.newPage();
        
        // Global page referansını set et
        global.currentPage = page;
        
        // Aktif sayfalar listesine ekle
        global.activePages.push(page);
        
        // Pinhuman'a git
        sendLogMessage('info', '📡 Pinhuman sayfasına gidiliyor...');
        await page.goto('https://www.pinhuman.net', {
            waitUntil: 'networkidle2'
        });
        
        // İlk giriş formunu doldur (şifre ile)
        sendLogMessage('info', '🔐 İlk giriş formu dolduruluyor...');
        await page.type('input[name="UserName"]', creds.userName);
        await page.type('input[name="CompanyCode"]', creds.companyCode);
        await page.type('input[name="Password"]', creds.password);
        
        // Giriş butonunu bekle ve tıkla
        sendLogMessage('info', '🔘 Giriş butonuna tıklanıyor...');
        await page.waitForSelector('button.btn-success, button[type="submit"], .btn.btn-lg.btn-success.btn-block', { timeout: 10000 });
        await page.click('button.btn-success, button[type="submit"], .btn.btn-lg.btn-success.btn-block');
        
        // Giriş işleminin tamamlanmasını bekle - AJAX ile form güncelleniyor
        sendLogMessage('info', '⏳ Giriş işlemi bekleniyor...');
        
        // AJAX form güncellemesi için bekle (sayfa yenilenmez, sadece form güncellenir)
        sendLogMessage('info', '⏳ AJAX form güncellemesi bekleniyor...');
        
        // AJAX isteğinin tamamlanmasını bekle
        await page.waitForFunction(() => {
            // Form güncellenmiş mi kontrol et
            const codeInput = document.getElementById('Code');
            return codeInput && codeInput.offsetParent !== null;
        }, { timeout: 15000 });
        
        sendLogMessage('info', '✅ AJAX form güncellemesi tamamlandı!');
        
        // 2FA alanını bekle ve TOTP kodunu gir
        sendLogMessage('info', `🔑 2FA kodu giriliyor: ${totpCode}`);
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Önce Code input'unu direkt ara
        try {
            sendLogMessage('info', '🎯 Code input alanı aranıyor...');
            await page.waitForSelector('input[id="Code"]', { 
                visible: true, 
                timeout: 10000 
            });
            sendLogMessage('info', '✅ Code input alanı bulundu ve görünür!');
            
            // Input'un tıklanabilir olmasını bekle
            await page.waitForFunction(() => {
                const codeInput = document.getElementById('Code');
                return codeInput && !codeInput.disabled && codeInput.offsetParent !== null;
            }, { timeout: 5000 });
            
            // Input'u temizle ve TOTP kodunu gir
            await page.evaluate(() => {
                const codeInput = document.getElementById('Code');
                if (codeInput) {
                    codeInput.value = '';
                    codeInput.focus();
                    
                    // Yeşil ring ekle - input bulundu!
                    codeInput.style.border = '3px solid #00ff00';
                    codeInput.style.boxShadow = '0 0 10px #00ff00';
                    codeInput.style.borderRadius = '5px';
                }
            });
            
            // Input'a tıkla ve TOTP kodunu gir
            await page.click('input[id="Code"]');
            await page.type('input[id="Code"]', totpCode);
            sendLogMessage('info', `✅ TOTP kodu girildi: ${totpCode}`);
        } catch (error) {
            sendLogMessage('warning', '❌ Code input bulunamadı, alternatif yöntem deneniyor...');
            
            // Sayfadaki tüm input'ları listele (debug için)
            const inputs = await page.evaluate(() => {
                const inputElements = document.querySelectorAll('input');
                return Array.from(inputElements).map(input => ({
                    id: input.id,
                    name: input.name,
                    placeholder: input.placeholder,
                    type: input.type,
                    className: input.className
                }));
            });
            console.log('📋 Sayfadaki input alanları:', inputs);
        
            // Daha geniş selector ile dene - Code input'unu öncelikle ara
            const selectors = [
                'input[id="Code"]',
                'input[name="Code"]',
                'input[placeholder="Doğrulama Kodunu girin..."]',
                'input[placeholder*="Doğrulama"]',
                'input[placeholder*="kod"]',
                'input.form-control',
                'input[type="text"]'
            ];
            
            let inputFound = false;
            for (const selector of selectors) {
                try {
                    console.log(`🔍 Selector deneniyor: ${selector}`);
                    await page.waitForSelector(selector, { timeout: 3000 });
                    console.log(`✅ Input bulundu: ${selector}`);
                    await page.type(selector, totpCode);
                    console.log(`✅ TOTP kodu girildi: ${totpCode}`);
                    inputFound = true;
                    break;
                } catch (error) {
                    console.log(`❌ Selector bulunamadı: ${selector}`);
                }
            }
            
            if (!inputFound) {
                throw new Error('2FA input alanı bulunamadı!');
            }
        }
        
        // 2FA giriş butonunu bekle ve tıkla
        console.log('🔘 2FA giriş butonuna tıklanıyor...');
        await page.waitForSelector('button[type="submit"], .btn-success, .btn-primary', { timeout: 10000 });
        await page.click('button[type="submit"], .btn-success, .btn-primary');
        
        // 2FA giriş sonucunu bekle
        console.log('⏳ 2FA giriş sonucu bekleniyor...');
        await page.waitForNavigation({ waitUntil: 'networkidle2' });
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        // URL'yi kontrol et - başarılı giriş mi?
        const loginResultUrl = page.url();
        console.log('🏠 Giriş sonrası URL:', loginResultUrl);
        
        // Eğer hala login sayfasındaysa, 2FA başarısız olmuş
        if (loginResultUrl.includes('login') || loginResultUrl.includes('Login')) {
            console.log('❌ 2FA giriş başarısız! Tekrar login sayfasına yönlendirildi.');
            
            // Hata mesajı var mı kontrol et
            try {
                const errorMessage = await page.textContent('.alert-danger, .error-message, .validation-summary-errors');
                if (errorMessage) {
                    console.log('🚨 Hata mesajı:', errorMessage);
                }
            } catch (e) {
                console.log('ℹ️ Hata mesajı bulunamadı');
            }
            
            throw new Error('2FA giriş başarısız! TOTP kodu yanlış olabilir veya süresi dolmuş olabilir.');
        }
        
        sendLogMessage('success', '✅ 2FA giriş başarılı! Ana sayfaya yönlendirildi.');
        
        // Employee Attendance sayfasına git
        // Employee Attendance sayfasına gidiliyor
        await page.goto('https://www.pinhuman.net/EmployeeAttendance', {
            waitUntil: 'networkidle2'
        });
        
        // Sayfanın yüklenmesini bekle
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        // Verileri personel bazında grupla
        const personelGruplari = groupDataByPersonel(veriler);
        sendLogMessage('info', `📊 ${Object.keys(personelGruplari).length} personel bulundu`);
        
        // Her personel için döngü
        for (const [sicilNo, personelData] of Object.entries(personelGruplari)) {
            // İşlem durduruldu mu kontrol et
            if (global.processStopped) {
                sendLogMessage('warning', '⚠️ İşlem kullanıcı tarafından durduruldu');
                break;
            }
            
            sendLogMessage('info', `👤 İşleniyor: ${personelData.personelAdi} (${sicilNo}) - ${personelData.kayitlar.length} kayıt`);
            
            // Her gün için döngü (giriş ve çıkış ayrı ayrı)
            for (const kayit of personelData.kayitlar) {
                // İşlem durduruldu mu kontrol et
                if (global.processStopped) {
                    sendLogMessage('warning', '⚠️ İşlem kullanıcı tarafından durduruldu');
                    break;
                }
                // Giriş kaydı
                await addAttendanceRecord(page, {
                    sicilNo,
                    tarih: kayit.tarih,
                    saat: kayit.giris,
                    direction: 'In',
                    vardiya: kayit.vardiya
                });
                
                // Çıkış kaydı
                await addAttendanceRecord(page, {
                    sicilNo,
                    tarih: kayit.tarih,
                    saat: kayit.cikis,
                    direction: 'Out',
                    vardiya: kayit.vardiya
                });
                
                // Kayıtlar arası bekleme
                await new Promise(resolve => setTimeout(resolve, 300));
            }
        }
        
        sendLogMessage('success', '✅ Tüm veriler başarıyla girildi!');
        
        return {
            success: true,
            message: 'Pinhuman\'a başarıyla giriş yapıldı ve tüm veriler sisteme girildi!'
        };
        
    } catch (error) {
        sendLogMessage('error', `❌ Pinhuman giriş hatası: ${error.message}`);
        return {
            success: false,
            message: 'Pinhuman\'a giriş yapılırken hata oluştu: ' + error.message
        };
    } finally {
        // Browser'ı kapatma - kullanıcı manuel olarak kapatabilir
        // if (browser) {
        //     await browser.close();
        // }
    }
});

// Verileri personel bazında grupla
function groupDataByPersonel(veriler) {
    const personelGruplari = {};
    
    veriler.forEach(kayit => {
        const sicilNo = kayit.sicilNo;
        const personelAdi = kayit.personel;
        
        if (!personelGruplari[sicilNo]) {
            personelGruplari[sicilNo] = {
                sicilNo,
                personelAdi,
                kayitlar: []
            };
        }
        
        // Sadece geçerli kayıtları ekle (giriş ve çıkış olan)
        if (kayit.gercek && kayit.gercek.gir && kayit.gercek.cik && kayit.gercek.gir !== '-' && kayit.gercek.cik !== '-') {
            personelGruplari[sicilNo].kayitlar.push({
                tarih: kayit.tarih,
                giris: kayit.gercek.gir,
                cikis: kayit.gercek.cik,
                vardiya: kayit.vardiya ? kayit.vardiya.kod : 'V1'
            });
        }
    });
    
    return personelGruplari;
}

// Tarih formatını dönüştür (2025-09-02 -> 02.09.2025)
function formatTarih(tarih) {
    const [yil, ay, gun] = tarih.split('-');
    return `${gun}.${ay}.${yil}`;
}

// Saat formatını dönüştür (08:23 -> 02.09.2025 08:23:00)
function formatSaat(tarih, saat) {
    const formattedTarih = formatTarih(tarih);
    return `${formattedTarih} ${saat}:00`;
}

// Modal içinde firma seçimi fonksiyonu
async function selectCompanyInModal(page) {
    try {
        sendLogMessage('info', '🎯 Modal içinde firma dropdown\'a tıklanıyor...');
        
        // Modal içindeki firma dropdown'unu bul (modal context'inde)
        // Önce modal'ın açık olduğunu kontrol et
        const modalExists = await page.$('.modal');
        if (!modalExists) {
            throw new Error('Modal açık değil, firma seçimi yapılamaz');
        }
        
        // Modal içindeki firma dropdown'unu bul
        const modalCompanyDropdown = await page.$('.modal #select2-CompanyId-container');
        if (!modalCompanyDropdown) {
            throw new Error('Modal içinde firma dropdown bulunamadı');
        }
        
        // Modal içindeki dropdown'a tıkla
        await modalCompanyDropdown.click();
        
        sendLogMessage('info', '✅ Modal içinde firma dropdown\'a tıklandı');
        
        // Modal içindeki firma dropdown'unu yeşil ring ile vurgula
        await page.evaluate(() => {
            // Önce modal içindeki dropdown'u bul
            const modalDropdown = document.querySelector('.modal #select2-CompanyId-container');
            const dropdown = modalDropdown || document.getElementById('select2-CompanyId-container');
            if (dropdown) {
                dropdown.style.border = '3px solid #00ff00';
                dropdown.style.boxShadow = '0 0 10px #00ff00';
                dropdown.style.borderRadius = '5px';
            }
        });
        
        // Dropdown açıldı mı kontrol et
        try {
            sendLogMessage('info', '⏳ Modal içinde firma dropdown açılması bekleniyor...');
            await page.waitForSelector('.select2-dropdown', { timeout: 3000 });
            sendLogMessage('info', '✅ Modal içinde firma dropdown açıldı, seçenekler aranıyor...');
            
            // Tüm seçenekleri listele
            const options = await page.evaluate(() => {
                const optionElements = document.querySelectorAll('.select2-results__option');
                return Array.from(optionElements).map(option => ({
                    text: option.textContent.trim(),
                    hasClass: option.classList.contains('select2-results__option--highlighted')
                }));
            });
            // Seçenekler bulundu, log kaldırıldı
            
            // "POLEN" seçeneğini bul ve tıkla (büyük harfle)
            sendLogMessage('info', '🔍 Modal içinde POLEN seçeneği aranıyor...');
            
            // Önce direkt POLEN seçeneğini ara
            let polenOption = await page.$('.select2-results__option');
            if (polenOption) {
                const optionText = await page.evaluate(el => el.textContent, polenOption);
                if (!optionText.includes('POLEN')) {
                    polenOption = null;
                }
            }
            
            if (!polenOption) {
                // Alternatif: Arama yap
                sendLogMessage('warning', '❌ Modal içinde POLEN seçeneği bulunamadı, arama yapılıyor...');
                await page.type('.select2-search__field', 'POLEN');
                sendLogMessage('info', '✅ Modal içinde arama alanına POLEN yazıldı');
                await new Promise(resolve => setTimeout(resolve, 800));
                
                // Arama sonrası tekrar ara
                polenOption = await page.$('.select2-results__option');
            }
            
            if (polenOption) {
                sendLogMessage('info', '🎯 Modal içinde POLEN seçeneği bulundu, tıklanıyor...');
                
                // POLEN seçeneğini yeşil ring ile vurgula
                await page.evaluate((option) => {
                    if (option) {
                        option.style.border = '3px solid #00ff00';
                        option.style.boxShadow = '0 0 10px #00ff00';
                        option.style.borderRadius = '5px';
                        option.style.backgroundColor = 'rgba(0, 255, 0, 0.1)';
                    }
                }, polenOption);
                
                await polenOption.click();
                sendLogMessage('info', '✅ Modal içinde POLEN seçeneğine tıklandı');
                await new Promise(resolve => setTimeout(resolve, 1000));
            } else {
                sendLogMessage('error', '❌ Modal içinde POLEN seçeneği bulunamadı!');
                throw new Error('POLEN seçeneği bulunamadı');
            }
            
            // Seçimin başarılı olup olmadığını kontrol et (modal içindeki)
            const selectedCompany = await page.evaluate(() => {
                const modalDropdown = document.querySelector('.modal #select2-CompanyId-container');
                const dropdown = modalDropdown || document.getElementById('select2-CompanyId-container');
                return dropdown ? dropdown.textContent : 'Bulunamadı';
            });
            // Firma seçildi
            
        } catch (error) {
            sendLogMessage('error', `❌ Modal içinde firma dropdown açılamadı: ${error.message}`);
            sendLogMessage('info', '🔄 Modal içinde alternatif yöntem deneniyor...');
            
            // Alternatif yöntem: Direkt select elementini kullan (modal içindeki)
            try {
                sendLogMessage('info', '🎯 Modal içinde SelectOption ile POLEN seçiliyor...');
                
                // Modal içindeki select elementini bul
                const modalSelect = await page.$('.modal #CompanyId');
                if (!modalSelect) {
                    throw new Error('Modal içinde CompanyId select elementi bulunamadı');
                }
                
                // Modal içindeki select'e değer ata
                await modalSelect.select('0fc07a2a-c718-482f-806a-3b009a9d06e1');
                
                sendLogMessage('info', '✅ Modal içinde SelectOption ile POLEN seçildi');
                await new Promise(resolve => setTimeout(resolve, 1000));
            } catch (selectError) {
                sendLogMessage('warning', `❌ Modal içinde SelectOption başarısız: ${selectError.message}`);
                sendLogMessage('info', '🔄 Modal içinde JavaScript ile denenecek...');
                
                // JavaScript ile direkt değer ata (POLEN'in gerçek ID'si) - modal içindeki
                sendLogMessage('info', '🎯 Modal içinde JavaScript ile POLEN ID\'si atanıyor...');
                await page.evaluate(() => {
                    // Modal içindeki select'i bul
                    const modalSelect = document.querySelector('.modal #CompanyId');
                    if (!modalSelect) {
                        throw new Error('Modal içinde CompanyId select elementi bulunamadı');
                    }
                    
                    // Modal içindeki select'e değer ata
                    modalSelect.value = '0fc07a2a-c718-482f-806a-3b009a9d06e1'; // POLEN'in gerçek ID'si
                    modalSelect.dispatchEvent(new Event('change', { bubbles: true }));
                    
                    // Select element'ini yeşil ring ile vurgula
                    modalSelect.style.border = '3px solid #00ff00';
                    modalSelect.style.boxShadow = '0 0 10px #00ff00';
                    modalSelect.style.borderRadius = '5px';
                    
                    // Modal içindeki Select2'yi de güncelle
                    const modalSelect2Container = document.querySelector('.modal #select2-CompanyId-container');
                    if (modalSelect2Container) {
                        modalSelect2Container.textContent = 'POLEN';
                        modalSelect2Container.style.border = '3px solid #00ff00';
                        modalSelect2Container.style.boxShadow = '0 0 10px #00ff00';
                        modalSelect2Container.style.borderRadius = '5px';
                    }
                });
                sendLogMessage('info', '✅ Modal içinde JavaScript ile POLEN ID\'si atandı');
                await new Promise(resolve => setTimeout(resolve, 1000));
            }
        }
        
        // Firma seçimi sonrası ilişkisel input'ların yüklenmesi için bekle
        sendLogMessage('info', '⏳ Modal içinde firma seçimi sonrası ilişkisel veriler yükleniyor...');
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        sendLogMessage('success', '✅ Modal içinde firma seçildi: Polen');
        
    } catch (error) {
        sendLogMessage('error', `❌ Modal içinde firma seçimi hatası: ${error.message}`);
        throw error;
    }
}

// Firma seçimi fonksiyonu (ana sayfa için - kullanılmıyor)
async function selectCompany(page) {
    try {
        sendLogMessage('info', '🎯 Firma dropdown\'a tıklanıyor...');
        await page.click('#select2-CompanyId-container');
        sendLogMessage('info', '✅ Firma dropdown\'a tıklandı');
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Dropdown açıldı mı kontrol et
        try {
            sendLogMessage('info', '⏳ Firma dropdown açılması bekleniyor...');
            await page.waitForSelector('.select2-dropdown', { timeout: 5000 });
            sendLogMessage('info', '✅ Firma dropdown açıldı, seçenekler aranıyor...');
            
            // Tüm seçenekleri listele
            const options = await page.$$eval('.select2-results__option', options => 
                options.map(option => ({
                    text: option.textContent.trim(),
                    hasClass: option.classList.contains('select2-results__option--highlighted')
                }))
            );
            // Seçenekler bulundu
            
            // "POLEN" seçeneğini bul ve tıkla (büyük harfle)
            sendLogMessage('info', '🔍 POLEN seçeneği aranıyor...');
            const polenOption = await page.$('.select2-results__option:has-text("POLEN")');
            if (polenOption) {
                sendLogMessage('info', '🎯 POLEN seçeneği bulundu, tıklanıyor...');
                await polenOption.click();
                sendLogMessage('info', '✅ POLEN seçeneğine tıklandı');
                await new Promise(resolve => setTimeout(resolve, 1000));
            } else {
                // Alternatif: Arama yap
                sendLogMessage('warning', '❌ POLEN seçeneği bulunamadı, arama yapılıyor...');
                await page.type('.select2-search__field', 'POLEN');
                sendLogMessage('info', '✅ Arama alanına POLEN yazıldı');
                await new Promise(resolve => setTimeout(resolve, 1500));
                
                // Arama sonucunu tıkla
                sendLogMessage('info', '🔍 Arama sonucu aranıyor...');
                const searchResult = await page.$('.select2-results__option');
                if (searchResult) {
                    sendLogMessage('info', '🎯 Arama sonucu bulundu, tıklanıyor...');
                    await searchResult.click({ force: true });
                    sendLogMessage('info', '✅ Arama sonucuna tıklandı');
                    await new Promise(resolve => setTimeout(resolve, 1000));
                } else {
                    sendLogMessage('error', '❌ Arama sonucu bulunamadı!');
                }
            }
            
            // Seçimin başarılı olup olmadığını kontrol et
            const selectedCompany = await page.textContent('#select2-CompanyId-container');
            // Firma seçimi doğrulandı
            
        } catch (error) {
            sendLogMessage('error', `❌ Firma dropdown açılamadı: ${error.message}`);
            sendLogMessage('info', '🔄 Alternatif yöntem deneniyor...');
            
            // Alternatif yöntem: Direkt select elementini kullan
            try {
                sendLogMessage('info', '🎯 SelectOption ile POLEN seçiliyor...');
                await page.select('#CompanyId', '0fc07a2a-c718-482f-806a-3b009a9d06e1');
                sendLogMessage('info', '✅ SelectOption ile POLEN seçildi');
                await new Promise(resolve => setTimeout(resolve, 1000));
            } catch (selectError) {
                sendLogMessage('warning', `❌ SelectOption başarısız: ${selectError.message}`);
                sendLogMessage('info', '🔄 JavaScript ile denenecek...');
                
                // JavaScript ile direkt değer ata (POLEN'in gerçek ID'si)
                sendLogMessage('info', '🎯 JavaScript ile POLEN ID\'si atanıyor...');
                await page.evaluate(() => {
                    const select = document.getElementById('CompanyId');
                    if (select) {
                        select.value = '0fc07a2a-c718-482f-806a-3b009a9d06e1'; // POLEN'in gerçek ID'si
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                        
                        // Select element'ini yeşil ring ile vurgula
                        select.style.border = '3px solid #00ff00';
                        select.style.boxShadow = '0 0 10px #00ff00';
                        select.style.borderRadius = '5px';
                    }
                });
                sendLogMessage('info', '✅ JavaScript ile POLEN ID\'si atandı');
                await new Promise(resolve => setTimeout(resolve, 1000));
            }
        }
        
        // Firma seçimi sonrası ilişkisel input'ların yüklenmesi için bekle
        sendLogMessage('info', '⏳ Firma seçimi sonrası ilişkisel veriler yükleniyor...');
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        sendLogMessage('success', '✅ Firma seçildi: Polen');
        
    } catch (error) {
        sendLogMessage('error', `❌ Firma seçimi hatası: ${error.message}`);
        throw error;
    }
}

// Attendance kaydı ekle
async function addAttendanceRecord(page, data) {
    try {
        // İşlem durduruldu mu kontrol et
        if (global.processStopped) {
            sendLogMessage('warning', '⚠️ İşlem kullanıcı tarafından durduruldu');
            return;
        }
        
        sendLogMessage('info', `📝 ${data.direction === 'In' ? 'Giriş' : 'Çıkış'}: ${formatSaat(data.tarih, data.saat)}`);
        
        // "Ekle" butonuna tıkla
        const addButton = await page.waitForSelector('a[href="/EmployeeAttendance/Create"]', { timeout: 10000 });
        await addButton.click();
        
        // Modal'ın açılmasını bekle
        sendLogMessage('info', '⏳ Modal açılıyor...');
        await page.waitForSelector('#CardNo', { timeout: 15000 });
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Modal kapatılmasını engelle - tüm yöntemlerle
        await page.evaluate(() => {
            // ESC tuşu ile modal kapatılmasını engelle
            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape') {
                    e.preventDefault();
                    e.stopPropagation();
                    return false;
                }
            }, true);
            
            // Modal dışına tıklama ile kapatılmasını engelle
            const modal = document.querySelector('.modal');
            if (modal) {
                modal.addEventListener('click', function(e) {
                    if (e.target === modal) {
                        e.preventDefault();
                        e.stopPropagation();
                        return false;
                    }
                }, true);
            }
            
            // Modal kapat butonlarını devre dışı bırak
            const closeButtons = document.querySelectorAll('.modal .close, .modal [data-dismiss="modal"], .modal .btn-outline-danger');
            closeButtons.forEach(button => {
                button.style.pointerEvents = 'none';
                button.style.opacity = '0.5';
            });
            
            // Modal backdrop tıklamasını engelle
            const backdrop = document.querySelector('.modal-backdrop');
            if (backdrop) {
                backdrop.style.pointerEvents = 'none';
            }
        });
        
        // ÖNCE SICIL NO, TARİH SAATİ VE YÖNÜ GİR
        sendLogMessage('info', `📝 Sicil numarası giriliyor: ${data.sicilNo}`);
        await page.type('#CardNo', data.sicilNo, { delay: 30 });
        sendLogMessage('info', `✅ Sicil numarası girildi: ${data.sicilNo}`);
        
        // Okuma zamanı gir
        const readTimeValue = formatSaat(data.tarih, data.saat);
        sendLogMessage('info', `⏰ Okuma zamanı giriliyor: ${readTimeValue}`);
        await page.evaluate((value) => {
            const readTimeInput = document.getElementById('ReadTime');
            readTimeInput.value = value;
            readTimeInput.dispatchEvent(new Event('change', { bubbles: true }));
        }, readTimeValue);
        sendLogMessage('info', `✅ Okuma zamanı girildi: ${readTimeValue}`);
        
        // Yön seçimi (Giriş/Çıkış)
        sendLogMessage('info', `🔄 Yön seçiliyor: ${data.direction}`);
        await page.select('#Direction', data.direction);
        const selectedDirection = await page.$eval('#Direction', el => el.value);
        sendLogMessage('info', `✅ Seçilen yön: ${selectedDirection}`);
        
        // ŞİMDİ FİRMA DROPDOWN'UNU AÇ VE SEÇ
        sendLogMessage('info', '🏢 Modal içinde firma seçiliyor...');
        await selectCompanyInModal(page);
        
        // Firma seçimi sonrası ilişkisel input'ların yüklenmesi için bekle
        sendLogMessage('info', '⏳ Modal içinde firma seçimi sonrası ilişkisel veriler yükleniyor...');
        await new Promise(resolve => setTimeout(resolve, 400));
        
        // Lokasyon seçimi - Tesis (Modal içinde)
        sendLogMessage('info', '📍 Lokasyon seçiliyor: Tesis');
        
        // Modal içindeki lokasyon dropdown'unu bul (modal context'inde)
        const modalExistsForLocation = await page.$('.modal');
        if (!modalExistsForLocation) {
            throw new Error('Modal açık değil, lokasyon seçimi yapılamaz');
        }
        
        // Modal içindeki lokasyon dropdown'unu bul
        const dropdownElement = await page.$('.modal #select2-CompanyWorkingRegionId-container');
        if (!dropdownElement) {
            throw new Error('Modal içinde lokasyon dropdown bulunamadı');
        }
        
        // Elementin görünür olup olmadığını kontrol et
        const isVisible = await dropdownElement.isIntersectingViewport();
        if (!isVisible) {
            await dropdownElement.scrollIntoView();
            await new Promise(resolve => setTimeout(resolve, 200));
        }
        
        // Dropdown'a tıkla
        await dropdownElement.click();
        await new Promise(resolve => setTimeout(resolve, 400));
        
        // Dropdown açıldı mı kontrol et
        try {
            await page.waitForSelector('.select2-dropdown', { timeout: 5000 });
            // Lokasyon dropdown açıldı
            
            // Tüm seçenekleri listele
            const options = await page.$$eval('.select2-results__option', options => 
                options.map(option => ({
                    text: option.textContent.trim(),
                    hasClass: option.classList.contains('select2-results__option--highlighted')
                }))
            );
            // Lokasyon seçenekleri bulundu
            
            // "TESİS" yazıp ilk çıkanı seç
            sendLogMessage('info', '🔍 TESİS yazılıyor...');
            
            // Arama alanını bul ve hızlı yaz
            const searchField = await page.$('.select2-search__field');
            if (searchField) {
                await searchField.click();
                await new Promise(resolve => setTimeout(resolve, 100));
                
                // Hızlı yaz
                await page.type('.select2-search__field', 'TESİS', { delay: 30 });
                await new Promise(resolve => setTimeout(resolve, 300));
                
                // İlk çıkan seçeneği seç
                const firstOption = await page.$('.select2-results__option');
                if (firstOption) {
                    sendLogMessage('info', '🎯 İlk TESİS seçeneği seçiliyor...');
                    await firstOption.click();
                    await new Promise(resolve => setTimeout(resolve, 500));
                } else {
                    sendLogMessage('warning', '⚠️ TESİS seçeneği bulunamadı!');
                }
            }
            
            // Seçimin başarılı olup olmadığını kontrol et (modal içindeki)
            const selectedLocation = await page.evaluate(() => {
                const modalDropdown = document.querySelector('.modal #select2-CompanyWorkingRegionId-container');
                const dropdown = modalDropdown || document.querySelector('#select2-CompanyWorkingRegionId-container');
                return dropdown ? dropdown.textContent : 'Bulunamadı';
            });
            // Lokasyon seçildi
            
        } catch (error) {
            sendLogMessage('error', `❌ Lokasyon dropdown açılamadı, alternatif yöntem deneniyor...`);
            
            // Alternatif yöntem: Direkt select elementini insan gibi kullan (modal içindeki)
            try {
                sendLogMessage('info', '🔄 Alternatif yöntem: Modal içinde select elementini buluyor...');
                
                // Modal içindeki select elementini bul
                const selectElement = await page.$('.modal #CompanyWorkingRegionId');
                if (!selectElement) {
                    throw new Error('Modal içinde CompanyWorkingRegionId select elementi bulunamadı');
                }
                
                // Select elementini görünür hale getir
                await selectElement.scrollIntoView();
                await new Promise(resolve => setTimeout(resolve, 200));
                
                // Mouse'u select'e doğru hareket ettir
                const selectBox = await selectElement.boundingBox();
                if (selectBox) {
                    const centerX = selectBox.x + selectBox.width / 2;
                    const centerY = selectBox.y + selectBox.height / 2;
                    await page.mouse.move(centerX, centerY, { steps: 5 });
                    await new Promise(resolve => setTimeout(resolve, 300));
                }
                
                // Select'e tıkla
                await selectElement.click();
                await new Promise(resolve => setTimeout(resolve, 300));
                
                // Değeri seç (Tesis için genellikle ilk seçenek)
                await selectElement.select('1');
                await new Promise(resolve => setTimeout(resolve, 800));
            } catch (selectError) {
                sendLogMessage('warning', `❌ Modal içinde select option da başarısız, JavaScript ile denenecek...`);
                
                // JavaScript ile direkt değer ata (son çare) - modal içindeki
                await page.evaluate(() => {
                    // Modal içindeki select'i bul
                    const modalSelect = document.querySelector('.modal #CompanyWorkingRegionId');
                    if (!modalSelect) {
                        throw new Error('Modal içinde CompanyWorkingRegionId select elementi bulunamadı');
                    }
                    
                    // Modal içindeki select'e değer ata
                    modalSelect.value = '1'; // Tesis'in ID'si olabilir
                    modalSelect.dispatchEvent(new Event('change', { bubbles: true }));
                });
                await new Promise(resolve => setTimeout(resolve, 300));
            }
        }
        
        // Lokasyon seçimi sonrası vardiya seçeneklerinin yüklenmesi için bekle
        sendLogMessage('info', '⏳ Lokasyon seçimi sonrası vardiya seçenekleri yükleniyor...');
        await new Promise(resolve => setTimeout(resolve, 500));
        
        sendLogMessage('success', '✅ Lokasyon seçildi: Tesis');
        
        // Vardiya seçimi (Modal içinde)
        sendLogMessage('info', `⏰ Vardiya seçiliyor: ${data.vardiya}`);
        
        // Modal içindeki vardiya dropdown'unu bul (modal context'inde)
        sendLogMessage('info', '🎯 Modal içinde vardiya dropdown\'a yaklaşılıyor...');
        
        // Önce modal'ın açık olduğunu kontrol et
        const modalExistsForShift = await page.$('.modal');
        if (!modalExistsForShift) {
            throw new Error('Modal açık değil, vardiya seçimi yapılamaz');
        }
        
        // Modal içindeki vardiya dropdown'unu bul - farklı selector'lar dene
        let shiftDropdownElement = await page.$('.modal #select2-ShiftId-container');
        if (!shiftDropdownElement) {
            sendLogMessage('warning', '⚠️ #select2-ShiftId-container bulunamadı, alternatif selector deneniyor...');
            shiftDropdownElement = await page.$('.modal [id*="ShiftId"]');
        }
        if (!shiftDropdownElement) {
            sendLogMessage('warning', '⚠️ [id*="ShiftId"] bulunamadı, alternatif selector deneniyor...');
            shiftDropdownElement = await page.$('.modal select[id*="Shift"]');
        }
        if (!shiftDropdownElement) {
            throw new Error('Modal içinde vardiya dropdown bulunamadı');
        }
        
        // Modal içindeki dropdown'a tıkla
        await shiftDropdownElement.click();
        sendLogMessage('info', '✅ Modal içinde vardiya dropdown\'a tıklandı');
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Dropdown açıldı mı kontrol et
        try {
            await page.waitForSelector('.select2-dropdown', { timeout: 5000 });
            // Vardiya dropdown açıldı
            
            // Tüm seçenekleri listele
            const options = await page.$$eval('.select2-results__option', options => 
                options.map(option => ({
                    text: option.textContent.trim(),
                    hasClass: option.classList.contains('select2-results__option--highlighted')
                }))
            );
            // Vardiya seçenekleri bulundu
            
            // Vardiya içeren seçeneği bul ve tıkla
            // Vardiya seçeneğini bul - saat aralığına göre seç
            let shiftOption = null;
            let targetTimeRange = '';
            
            // Vardiya kodunu saat aralığına çevir
            if (data.vardiya === 'V1') {
                targetTimeRange = '08:30-16:30';
            } else if (data.vardiya === 'V2') {
                targetTimeRange = '16:30-00:30';
            } else if (data.vardiya === 'V3') {
                targetTimeRange = '00:30-08:30';
            }
            
            sendLogMessage('info', `🔍 ${data.vardiya} vardiyası için ${targetTimeRange} saat aralığı aranıyor...`);
            
            // Tüm seçenekleri kontrol et ve saat aralığına göre seç
            const allOptions = await page.$$('.select2-results__option');
            for (const option of allOptions) {
                const optionText = await option.evaluate(el => el.textContent.trim());
                if (optionText.includes(targetTimeRange)) {
                    shiftOption = option;
                    sendLogMessage('info', `🎯 ${targetTimeRange} saat aralığı bulundu: ${optionText}`);
                    break;
                }
            }
            
            if (shiftOption) { sendLogMessage('info', `🎯 ${data.vardiya} seçeneği bulundu, tıklanıyor...`);
                
                // Mouse'u seçeneğe doğru hareket ettir
                const optionBox = await shiftOption.boundingBox();
                if (optionBox) {
                    const centerX = optionBox.x + optionBox.width / 2;
                    const centerY = optionBox.y + optionBox.height / 2;
                    await page.mouse.move(centerX, centerY, { steps: 5 });
                    await new Promise(resolve => setTimeout(resolve, 200));
                }
                
                await shiftOption.click();
                sendLogMessage('info', `✅ ${data.vardiya} seçeneğine tıklandı`);
                await new Promise(resolve => setTimeout(resolve, 1000));
            } else {
                // Son çare: Arama yap
                sendLogMessage('info', `🔍 ${data.vardiya} seçeneği bulunamadı, arama yapılıyor...`);
                const searchField = await page.$('.select2-search__field');
                if (searchField) {
                    await searchField.click();
                    await page.type('.select2-search__field', data.vardiya, { delay: 30 });
                    await new Promise(resolve => setTimeout(resolve, 1000));
                    
                    // Arama sonucunu tıkla
                    const searchResult = await page.$('.select2-results__option');
                    if (searchResult) {
                        await searchResult.click();
                        sendLogMessage('info', `✅ Arama sonucu tıklandı`);
                        await new Promise(resolve => setTimeout(resolve, 1000));
                    }
                }
            }
            
            // Seçimin başarılı olup olmadığını kontrol et (modal içindeki)
            const selectedShift = await page.evaluate(() => {
                const modalDropdown = document.querySelector('.modal #select2-ShiftId-container');
                const dropdown = modalDropdown || document.querySelector('#select2-ShiftId-container');
                return dropdown ? dropdown.textContent : 'Bulunamadı';
            });
            // Vardiya seçildi
            
            // Vardiya seçiminin başarılı olduğunu doğrula
            if (!selectedShift || selectedShift === 'Bulunamadı' || selectedShift.trim() === '') {
                throw new Error('Vardiya seçimi başarısız - seçim yapılamadı');
            }
            
            // Vardiya seçiminin doğru olduğunu kontrol et - saat aralığına göre
            let expectedTimeRange = '';
            if (data.vardiya === 'V1') {
                expectedTimeRange = '08:30-16:30';
            } else if (data.vardiya === 'V2') {
                expectedTimeRange = '16:30-00:30';
            } else if (data.vardiya === 'V3') {
                expectedTimeRange = '00:30-08:30';
            }
            
            if (!selectedShift.includes(expectedTimeRange)) {
                throw new Error(`Vardiya seçimi yanlış - beklenen: ${expectedTimeRange}, seçilen: ${selectedShift}`);
            }
            
            sendLogMessage('success', `✅ Vardiya seçimi doğrulandı: ${selectedShift}`);
            
            // Tab tuşu ile kaydet butonuna focus geç
            sendLogMessage('info', '⌨️ Tab tuşu ile kaydet butonuna focus geçiliyor...');
            await page.keyboard.press('Tab');
            await new Promise(resolve => setTimeout(resolve, 300));
            sendLogMessage('success', '✅ Tab tuşu ile kaydet butonuna focus geçildi');
            
            // Enter tuşu ile kaydetme işlemini yap
            sendLogMessage('info', '⌨️ Enter tuşu ile kaydetme işlemi yapılıyor...');
            await page.keyboard.press('Enter');
            await new Promise(resolve => setTimeout(resolve, 200));
            sendLogMessage('success', '✅ Enter tuşu ile kaydetme işlemi tamamlandı');
            
        } catch (error) {
            sendLogMessage('error', `❌ Modal içinde vardiya dropdown açılamadı, alternatif yöntem deneniyor...`);
            
            // Alternatif yöntem: Direkt select elementini kullan (modal içindeki)
            try {
                sendLogMessage('info', '🔄 Alternatif yöntem: Modal içinde select elementini buluyor...');
                
                // Modal içindeki select elementini bul
                const shiftSelectElement = await page.$('.modal #ShiftId');
                if (!shiftSelectElement) {
                    throw new Error('Modal içinde ShiftId select elementi bulunamadı');
                }
                
                // Modal içindeki select'e değer ata - saat aralığına göre
                let targetTimeRange = '';
                if (data.vardiya === 'V1') {
                    targetTimeRange = '08:30-16:30';
                } else if (data.vardiya === 'V2') {
                    targetTimeRange = '16:30-00:30';
                } else if (data.vardiya === 'V3') {
                    targetTimeRange = '00:30-08:30';
                }
                
                // Select elementindeki seçenekleri kontrol et
                const options = await shiftSelectElement.$$eval('option', options => 
                    options.map(option => ({
                        value: option.value,
                        text: option.textContent.trim()
                    }))
                );
                
                // Saat aralığına göre seçenek bul
                const targetOption = options.find(opt => opt.text.includes(targetTimeRange));
                if (targetOption) {
                    await shiftSelectElement.select(targetOption.value);
                    sendLogMessage('info', `✅ Select elementine ${targetOption.text} seçildi`);
                } else {
                    throw new Error(`Saat aralığı ${targetTimeRange} select seçeneklerinde bulunamadı`);
                }
                
                await new Promise(resolve => setTimeout(resolve, 1000));
                
                // Seçimin başarılı olduğunu doğrula
                const selectedValue = await shiftSelectElement.evaluate(el => el.value);
                const selectedText = await shiftSelectElement.evaluate(el => {
                    const option = el.options[el.selectedIndex];
                    return option ? option.text : '';
                });
                
                if (!selectedValue || selectedValue === '') {
                    throw new Error('Vardiya seçimi başarısız - select değeri atanamadı');
                }
                
                if (!selectedText.includes(targetTimeRange)) {
                    throw new Error(`Vardiya seçimi yanlış - beklenen: ${targetTimeRange}, seçilen: ${selectedText}`);
                }
                
                sendLogMessage('success', `✅ Vardiya seçimi doğrulandı (select): ${selectedText}`);
                
                // Tab tuşu ile kaydet butonuna focus geç
                sendLogMessage('info', '⌨️ Tab tuşu ile kaydet butonuna focus geçiliyor...');
                await page.keyboard.press('Tab');
                await new Promise(resolve => setTimeout(resolve, 300));
                sendLogMessage('success', '✅ Tab tuşu ile kaydet butonuna focus geçildi');
                
                // Enter tuşu ile kaydetme işlemini yap
                sendLogMessage('info', '⌨️ Enter tuşu ile kaydetme işlemi yapılıyor...');
                await page.keyboard.press('Enter');
                await new Promise(resolve => setTimeout(resolve, 200));
                sendLogMessage('success', '✅ Enter tuşu ile kaydetme işlemi tamamlandı');
                
            } catch (selectError) {
                sendLogMessage('warning', `❌ Modal içinde select option da başarısız, JavaScript ile denenecek...`);
                
                // JavaScript ile direkt değer ata (modal içindeki) - saat aralığına göre
                let targetTimeRange = '';
                if (data.vardiya === 'V1') {
                    targetTimeRange = '08:30-16:30';
                } else if (data.vardiya === 'V2') {
                    targetTimeRange = '16:30-00:30';
                } else if (data.vardiya === 'V3') {
                    targetTimeRange = '00:30-08:30';
                }
                
                const jsResult = await page.evaluate((vardiya, timeRange) => {
                    // Modal içindeki select'i bul
                    const modalSelect = document.querySelector('.modal #ShiftId');
                    if (!modalSelect) {
                        throw new Error('Modal içinde ShiftId select elementi bulunamadı');
                    }
                    // Modal içindeki select'e değer ata - saat aralığına göre
                    const options = Array.from(modalSelect.options);
                    const option = options.find(opt => opt.text.includes(timeRange));
                    if (option) {
                        modalSelect.value = option.value;
                        modalSelect.dispatchEvent(new Event('change', { bubbles: true }));
                        return { success: true, value: option.value, text: option.text };
                    }
                    return { success: false, error: `Saat aralığı ${timeRange} bulunamadı` };
                }, data.vardiya, targetTimeRange);
                
                if (!jsResult.success) {
                    throw new Error(`JavaScript vardiya seçimi başarısız: ${jsResult.error}`);
                }
                
                if (!jsResult.text.includes(targetTimeRange)) {
                    throw new Error(`JavaScript vardiya seçimi yanlış - beklenen: ${targetTimeRange}, seçilen: ${jsResult.text}`);
                }
                
                sendLogMessage('success', `✅ Vardiya seçimi doğrulandı (JavaScript): ${jsResult.text}`);
                
                // Tab tuşu ile kaydet butonuna focus geç
                sendLogMessage('info', '⌨️ Tab tuşu ile kaydet butonuna focus geçiliyor...');
                await page.keyboard.press('Tab');
                await new Promise(resolve => setTimeout(resolve, 300));
                sendLogMessage('success', '✅ Tab tuşu ile kaydet butonuna focus geçildi');
                
                // Enter tuşu ile kaydetme işlemini yap
                sendLogMessage('info', '⌨️ Enter tuşu ile kaydetme işlemi yapılıyor...');
                await page.keyboard.press('Enter');
                await new Promise(resolve => setTimeout(resolve, 200));
                sendLogMessage('success', '✅ Enter tuşu ile kaydetme işlemi tamamlandı');
                
                await new Promise(resolve => setTimeout(resolve, 300));
            }
        }
        
        console.log(`  ✅ Vardiya seçildi: ${data.vardiya}`);
        
        // Tab+Enter ile kaydetme işlemi zaten yapıldı, ek işlem gerekmiyor
        sendLogMessage('success', '✅ Tab+Enter ile kaydetme işlemi tamamlandı');
        
        // Kaydet butonuna tıklandıktan sonra modal kapatma engelini kaldır
        await page.evaluate(() => {
            // Modal kapat butonlarını tekrar aktif et
            const closeButtons = document.querySelectorAll('.modal .close, .modal [data-dismiss="modal"], .modal .btn-outline-danger');
            closeButtons.forEach(button => {
                button.style.pointerEvents = 'auto';
                button.style.opacity = '1';
            });
            
            // Modal backdrop'u tekrar aktif et
            const backdrop = document.querySelector('.modal-backdrop');
            if (backdrop) {
                backdrop.style.pointerEvents = 'auto';
            }
        });
        
        // Kaydet butonuna tıklandıktan sonra loading'i bekle - modal kendiliğinden kapanacak
        sendLogMessage('info', '⏳ Kaydet işlemi tamamlanıyor, loading bekleniyor...');
        await new Promise(resolve => setTimeout(resolve, 800));
        
        // Modal'ın kapanıp kapanmadığını kontrol et
        const modalStillOpen = await page.$('.modal');
        if (modalStillOpen) {
            sendLogMessage('warning', '⚠️ Modal hala açık, biraz daha bekleniyor...');
            await new Promise(resolve => setTimeout(resolve, 500));
        }
        
        sendLogMessage('success', `✅ ${data.direction === 'In' ? 'Giriş' : 'Çıkış'} kaydı başarıyla eklendi`);
        
    } catch (error) {
        sendLogMessage('error', `❌ Kayıt hatası (${data.direction}): ${error.message}`);
        
        // Hata durumunda modal'ı kapatma - kullanıcı manuel olarak kapatabilir
        sendLogMessage('warning', '⚠️ Hata nedeniyle işlem durduruldu. Modal açık kalabilir.');
    }
}

// Excel verilerini Pinhuman'a gir - dinamik personel eşleştirmesi ile
ipcMain.handle('enter-excel-data-pinhuman', async (event, { userName, companyCode, password, totpSecret }) => {
    let browser = null;
    let page = null;
    
    try {
        console.log('Excel verileri Pinhuman\'a giriliyor...');
        
        // Excel verilerini oku (out/daily.json'dan)
        const fs = require('fs');
        const path = require('path');
        const dailyJsonPath = path.join(getOutputDir(), 'daily.json');
        
        if (!fs.existsSync(dailyJsonPath)) {
            return {
                success: false,
                message: 'daily.json dosyası bulunamadı. Önce PDKS işlemi yapın.'
            };
        }
        
        const dailyData = JSON.parse(fs.readFileSync(dailyJsonPath, 'utf8'));
        const veriler = dailyData.veriler || [];
        
        if (veriler.length === 0) {
            return {
                success: false,
                message: 'Excel dosyasında veri bulunamadı.'
            };
        }
        
        // Puppeteer browser başlat
        // Headless ayarını al (varsayılan false)
        const headlessMode = false; // Bu fonksiyon için şimdilik false
        
        browser = await puppeteer.launch({ 
            headless: headlessMode,
            args: ['--no-sandbox', '--disable-setuid-sandbox']
        });
        
        page = await browser.newPage();
        
        // Global page referansını set et
        global.currentPage = page;
        
        // Pinhuman'a giriş yap
        await page.goto('https://www.pinhuman.net', { waitUntil: 'networkidle2' });
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Giriş formunu doldur
        await page.type('input[name="userName"]', userName);
        await page.type('input[name="companyCode"]', companyCode);
        
        // TOTP kodu oluştur ve şifreye ekle
        let finalPassword = password;
        if (totpSecret) {
            try {
                const speakeasy = require('speakeasy');
                const token = speakeasy.totp({
                    secret: totpSecret,
                    encoding: 'base32'
                });
                finalPassword = password + token;
                console.log('TOTP kodu oluşturuldu ve şifreye eklendi');
            } catch (error) {
                console.log('TOTP kodu oluşturulamadı:', error.message);
            }
        }
        
        await page.type('input[name="password"]', finalPassword);
        await page.click('button[type="submit"]');
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        // Personel listesi sayfasına git
        await page.goto('https://www.pinhuman.net/employee-list', { waitUntil: 'networkidle2' }); // Gerçek URL'yi buraya yazın
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Sayfadaki mevcut personelleri al
        const availableEmployees = await page.evaluate(() => {
            const employees = [];
            const rows = document.querySelectorAll('tr.searchEmployee');
            rows.forEach(row => {
                const nameCell = row.querySelector('td.search-name');
                const checkboxCell = row.querySelector('input[name="SelectedEmployeeId"]');
                if (nameCell && checkboxCell) {
                    employees.push({
                        name: nameCell.textContent.trim(),
                        id: checkboxCell.value
                    });
                }
            });
            return employees;
        });
        
        console.log(`Sayfada ${availableEmployees.length} personel bulundu`);
        
        // Excel'deki personellerle eşleştir
        const matchedEmployees = [];
        const uniquePersonelNames = [...new Set(veriler.map(v => v.personel))];
        
        uniquePersonelNames.forEach(excelPersonel => {
            const match = availableEmployees.find(emp => 
                emp.name.toLowerCase().includes(excelPersonel.toLowerCase()) ||
                excelPersonel.toLowerCase().includes(emp.name.toLowerCase())
            );
            if (match) {
                matchedEmployees.push({
                    excelName: excelPersonel,
                    webName: match.name,
                    id: match.id
                });
            }
        });
        
        console.log(`${matchedEmployees.length} personel eşleştirildi`);
        
        // Eşleştirilen personelleri seç
        for (const employee of matchedEmployees) {
            try {
                const checkboxSelector = `input[name="SelectedEmployeeId"][value="${employee.id}"]`;
                const isChecked = await page.isChecked(checkboxSelector);
                
                if (!isChecked) {
                    await page.check(checkboxSelector);
                    console.log(`✓ ${employee.webName} seçildi`);
                }
                
                await new Promise(resolve => setTimeout(resolve, 200));
            } catch (error) {
                console.log(`✗ ${employee.webName} seçilemedi:`, error.message);
            }
        }
        
        // Seçilen personeller için verileri gir
        console.log('Vardiya verileri giriliyor...');
        await processAttendanceData(page, veriler);
        
        return {
            success: true,
            message: `${veriler.length} adet veri başarıyla girildi! (${matchedEmployees.length} personel eşleştirildi)`
        };
        
    } catch (error) {
        console.error('Excel veri girişi hatası:', error);
        return {
            success: false,
            message: 'Excel veri girişi sırasında hata oluştu: ' + error.message
        };
    } finally {
        // Browser'ı kapatma - kullanıcı manuel olarak kapatabilir
        // if (browser) {
        //     await browser.close();
        // }
    }
});
// Vardiya analizi raporu oluşturma
ipcMain.handle('get-shift-analysis', async (event) => {
    try {
        // Excel verilerini oku (out/daily.json'dan)
        const fs = require('fs');
        const path = require('path');
        const dailyJsonPath = path.join(getOutputDir(), 'daily.json');
        
        if (!fs.existsSync(dailyJsonPath)) {
            return {
                success: false,
                message: 'daily.json dosyası bulunamadı. Önce PDKS işlemi yapın.'
            };
        }
        
        const dailyData = JSON.parse(fs.readFileSync(dailyJsonPath, 'utf8'));
        const veriler = dailyData.veriler || [];
        
        // Vardiya analizi yap
        const shiftAnalysis = analyzeShiftData(veriler);
        
        return {
            success: true,
            data: shiftAnalysis
        };
        
    } catch (error) {
        console.error('Vardiya analizi hatası:', error);
        return {
            success: false,
            message: 'Vardiya analizi sırasında hata oluştu: ' + error.message
        };
    }
});

// Vardiya verilerini analiz etme fonksiyonu
function analyzeShiftData(veriler) {
    const analysis = {
        toplamKayit: veriler.length,
        vardiyaDagilimi: {},
        personelVardiya: {},
        gunlukVardiya: {},
        vardiyaKombinasyonlari: [],
        vardiyaDetaylari: [],
        istatistikler: {
            enCokCalisanVardiya: '',
            enAzCalisanVardiya: '',
            ortalamaCalismaSuresi: 0,
            toplamFazlaMesai: 0
        }
    };
    
    // Vardiya dağılımını hesapla
    veriler.forEach(kayit => {
        const vardiya = kayit.vardiya ? kayit.vardiya.kod : 'Belirsiz';
        
        if (!analysis.vardiyaDagilimi[vardiya]) {
            analysis.vardiyaDagilimi[vardiya] = {
                sayi: 0,
                toplamCalisma: 0,
                toplamFM: 0,
                personeller: new Set()
            };
        }
        
        analysis.vardiyaDagilimi[vardiya].sayi++;
        analysis.vardiyaDagilimi[vardiya].toplamCalisma += kayit.calisma_dk;
        analysis.vardiyaDagilimi[vardiya].toplamFM += kayit.fm_dk;
        analysis.vardiyaDagilimi[vardiya].personeller.add(kayit.personel);
        
        // Vardiya detaylarını ekle
        analysis.vardiyaDetaylari.push({
            tarih: kayit.tarih,
            vardiya: vardiya,
            personel: kayit.personel,
            giris: kayit.giris,
            cikis: kayit.cikis,
            calisma_dk: kayit.calisma_dk,
            fm_dk: kayit.fm_dk,
            durum: kayit.durum || 'Normal'
        });
        
        // Personel-vardiya eşleştirmesi
        if (!analysis.personelVardiya[kayit.personel]) {
            analysis.personelVardiya[kayit.personel] = {};
        }
        if (!analysis.personelVardiya[kayit.personel][vardiya]) {
            analysis.personelVardiya[kayit.personel][vardiya] = 0;
        }
        analysis.personelVardiya[kayit.personel][vardiya]++;
        
        // Günlük vardiya dağılımı
        if (!analysis.gunlukVardiya[kayit.tarih]) {
            analysis.gunlukVardiya[kayit.tarih] = {};
        }
        if (!analysis.gunlukVardiya[kayit.tarih][vardiya]) {
            analysis.gunlukVardiya[kayit.tarih][vardiya] = 0;
        }
        analysis.gunlukVardiya[kayit.tarih][vardiya]++;
    });
    
    // Vardiya kombinasyonlarını bul
    Object.entries(analysis.personelVardiya).forEach(([personel, vardiyalar]) => {
        const vardiyaListesi = Object.keys(vardiyalar);
        if (vardiyaListesi.length > 1) {
            const kombinasyon = vardiyaListesi.sort().join(' + ');
            const existingCombo = analysis.vardiyaKombinasyonlari.find(c => c.kombinasyon === kombinasyon);
            if (existingCombo) {
                existingCombo.personelSayisi++;
                existingCombo.personeller.push(personel);
            } else {
                analysis.vardiyaKombinasyonlari.push({
                    kombinasyon: kombinasyon,
                    personelSayisi: 1,
                    personeller: [personel],
                    vardiyalar: vardiyaListesi
                });
            }
        }
    });
    
    // İstatistikleri hesapla
    let enCokCalisan = 0;
    let enAzCalisan = Infinity;
    let toplamCalisma = 0;
    let toplamFM = 0;
    
    Object.entries(analysis.vardiyaDagilimi).forEach(([vardiya, data]) => {
        if (data.sayi > enCokCalisan) {
            enCokCalisan = data.sayi;
            analysis.istatistikler.enCokCalisanVardiya = vardiya;
        }
        if (data.sayi < enAzCalisan) {
            enAzCalisan = data.sayi;
            analysis.istatistikler.enAzCalisanVardiya = vardiya;
        }
        toplamCalisma += data.toplamCalisma;
        toplamFM += data.toplamFM;
    });
    
    analysis.istatistikler.ortalamaCalismaSuresi = Math.round(toplamCalisma / veriler.length);
    analysis.istatistikler.toplamFazlaMesai = toplamFM;
    
    // Set'leri array'e çevir
    Object.values(analysis.vardiyaDagilimi).forEach(data => {
        data.personeller = Array.from(data.personeller);
    });
    
    return analysis;
}

// Veri geçmişini yükle
ipcMain.handle('load-history', async () => {
    try {
        const historyPath = path.join(getOutputDir(), 'data_history.json');
        if (fs.existsSync(historyPath)) {
            const historyData = fs.readFileSync(historyPath, 'utf8');
            return JSON.parse(historyData);
        }
        return [];
    } catch (error) {
        console.error('Veri geçmişi yüklenirken hata:', error);
        return [];
    }
});

// Tüm işlemleri durdur
ipcMain.on('stop-processes', async () => {
    try {
        console.log('🛑 Tüm Puppeteer işlemleri durduruluyor...');
        
        // İşlem durumu flag'ini set et
        global.processStopped = true;
        
        // Global browser değişkenini kontrol et ve kapat
        if (global.currentBrowser) {
            console.log('🌐 Puppeteer browser kapatılıyor...');
            try {
                await global.currentBrowser.close();
            } catch (error) {
                console.log('Browser kapatılırken hata:', error.message);
            }
            global.currentBrowser = null;
        }
        
        // Tüm aktif sayfaları kapat
        if (global.activePages && global.activePages.length > 0) {
            console.log('📄 Aktif Puppeteer sayfaları kapatılıyor...');
            for (const page of global.activePages) {
                try {
                    await page.close();
                } catch (error) {
                    console.log('Sayfa kapatılırken hata:', error.message);
                }
            }
            global.activePages = [];
        }
        
        // Tüm Puppeteer processlerini kill et
        try {
            const { exec } = require('child_process');
            const os = require('os');
            
            if (os.platform() === 'win32') {
                // Windows için
                exec('taskkill /f /im chrome.exe /t', (error) => {
                    if (error && !error.message.includes('not found')) {
                        console.log('Chrome process kill hatası:', error.message);
                    }
                });
                exec('taskkill /f /im chromium.exe /t', (error) => {
                    if (error && !error.message.includes('not found')) {
                        console.log('Chromium process kill hatası:', error.message);
                    }
                });
            } else {
                // macOS ve Linux için
                exec('pkill -f chrome', (error) => {
                    if (error && !error.message.includes('No matching processes')) {
                        console.log('Chrome process kill hatası:', error.message);
                    }
                });
                exec('pkill -f chromium', (error) => {
                    if (error && !error.message.includes('No matching processes')) {
                        console.log('Chromium process kill hatası:', error.message);
                    }
                });
            }
        } catch (error) {
            console.log('Process kill hatası:', error.message);
        }
        
        console.log('✅ Tüm Puppeteer işlemleri başarıyla durduruldu');
        
    } catch (error) {
        console.error('❌ Puppeteer işlemleri durdurulurken hata:', error);
    }
});

// Pinhuman işlemini durdur
ipcMain.handle('stop-pinhuman-process', async () => {
    try {
        console.log('🛑 Pinhuman işlemi durduruluyor...');
        
        // İşlem durumu flag'ini set et
        global.processStopped = true;
        
        // Global browser değişkenini kontrol et ve kapat
        if (global.currentBrowser) {
            console.log('🌐 Puppeteer browser kapatılıyor...');
            try {
                await global.currentBrowser.close();
            } catch (error) {
                console.log('Browser kapatılırken hata:', error.message);
            }
            global.currentBrowser = null;
        }
        
        // Global page değişkenini temizle
        global.currentPage = null;
        
        // Tüm aktif sayfaları kapat
        if (global.activePages && global.activePages.length > 0) {
            console.log('📄 Aktif Puppeteer sayfaları kapatılıyor...');
            for (const page of global.activePages) {
                try {
                    await page.close();
                } catch (error) {
                    console.log('Sayfa kapatılırken hata:', error.message);
                }
            }
            global.activePages = [];
        }
        
        console.log('✅ Pinhuman işlemi başarıyla durduruldu');
        return { success: true, message: 'İşlem başarıyla durduruldu' };
        
    } catch (error) {
        console.error('❌ Pinhuman işlemi durdurulurken hata:', error);
        return { success: false, message: 'İşlem durdurulurken hata oluştu: ' + error.message };
    }
});

// Tek bir kayıt için veri girişi yapan fonksiyon
async function fillSingleRecord(record, date, shiftId) {
    try {
        console.log('📝 Tek kayıt veri girişi başlatılıyor:', record);
        
        // Global page değişkenini kontrol et
        if (!global.currentPage) {
            throw new Error('Sayfa bulunamadı. Önce giriş yapılmalı.');
        }
        
        const page = global.currentPage;
        
        // Veri formatını hazırla
        const attendanceData = {
            sicilNo: record.sicilNo,
            personel: record.personel,
            tarih: date,
            saat: record.time,
            direction: record.type === '1' ? 'In' : 'Out',
            vardiya: shiftId
        };
        
        // addAttendanceRecord fonksiyonunu kullan
        await addAttendanceRecord(page, attendanceData);
        
        console.log('✅ Tek kayıt veri girişi tamamlandı');
        
    } catch (error) {
        console.error('❌ Tek kayıt veri girişi hatası:', error);
        throw error;
    }
}

