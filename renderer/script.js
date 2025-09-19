const { ipcRenderer, shell } = require('electron');
const path = require('path');

// DOM elementleri
const fileInputArea = document.getElementById('fileInputArea');
const selectedFile = document.getElementById('selectedFile');
const fileName = document.getElementById('fileName');
const filePath = document.getElementById('filePath');
const selectFileBtn = document.getElementById('selectFileBtn');
const changeFileBtn = document.getElementById('changeFileBtn');
const processSection = document.getElementById('processSection');
const processBtn = document.getElementById('processBtn');
const progressSection = document.getElementById('progressSection');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultsSection = document.getElementById('resultsSection');
const errorSection = document.getElementById('errorSection');
const errorMessage = document.getElementById('errorMessage');
const totalRecords = document.getElementById('totalRecords');
const personelCount = document.getElementById('personelCount');
const errorCount = document.getElementById('errorCount');
const openFolderBtn = document.getElementById('openFolderBtn');
const retryBtn = document.getElementById('retryBtn');
const pinhumanBtn = document.getElementById('pinhumanBtn');
const logSection = document.getElementById('logSection');
const logContainer = document.getElementById('logContainer');
const closeLogBtn = document.getElementById('closeLogBtn');
const logoutBtn = document.getElementById('logoutBtn');

// Sidebar elementleri
const totalRecordsStat = document.getElementById('totalRecordsStat');
const personelCountStat = document.getElementById('personelCountStat');
const errorCountStat = document.getElementById('errorCountStat');
const personelFilter = document.getElementById('personelFilter');
const dataTableBody = document.getElementById('dataTableBody');

let selectedFilePath = null;
let outputPaths = {};
let currentData = [];
let filteredData = [];
let dataHistory = [];

// Event listeners
selectFileBtn.addEventListener('click', selectFile);
changeFileBtn.addEventListener('click', selectFile);
processBtn.addEventListener('click', processFile);
openFolderBtn.addEventListener('click', openOutputFolder);
retryBtn.addEventListener('click', () => {
    hideError();
    if (selectedFilePath) {
        processFile();
    } else {
        selectFile();
    }
});

// Pinhuman buton event listener
if (pinhumanBtn) {
    pinhumanBtn.addEventListener('click', () => {
        showLogSection();
        sendDataToPinhuman();
    });
}

// Log kapatma buton event listener
if (closeLogBtn) {
    closeLogBtn.addEventListener('click', () => {
        hideLogSection();
    });
}

// Çıkış butonu
if (logoutBtn) {
    logoutBtn.addEventListener('click', () => {
        if (confirm('Çıkış yapmak istediğinizden emin misiniz?')) {
            ipcRenderer.invoke('logout');
        }
    });
}

// Log görüntüleme fonksiyonları
function showLogSection() {
    // Mevcut alanları gizle
    hideAllSections();
    
    // Log alanını göster
    logSection.style.display = 'block';
    logSection.classList.add('slideInFromTop');
    
    // Log alanını temizle ve başlangıç mesajı ekle
    clearLogs();
    addLogEntry('info', 'Pinhuman\'a veri gönderimi başlatılıyor...');
}

function hideLogSection() {
    logSection.style.display = 'none';
    logSection.classList.remove('slideInFromTop');
    
    // Ana sayfaya geri dön
    resultsSection.style.display = 'block';
    resultsSection.classList.add('fade-in');
}

function clearLogs() {
    if (logContainer) {
        logContainer.innerHTML = '';
    }
}

function addLogEntry(type, message) {
    if (!logContainer) return;
    
    const timestamp = new Date().toLocaleTimeString('tr-TR');
    const logEntry = document.createElement('div');
    logEntry.className = `log-entry ${type}`;
    
    logEntry.innerHTML = `
        <span class="log-timestamp">[${timestamp}]</span>
        <span class="log-message">${message}</span>
    `;
    
    logContainer.appendChild(logEntry);
    
    // Scroll to bottom - log-content'e scroll yap (card içinde)
    const logContent = logContainer.parentElement;
    logContent.scrollTop = logContent.scrollHeight;
}

// Pinhuman'a veri gönderme fonksiyonu
async function sendDataToPinhuman() {
    try {
        // Buton durumunu değiştir
        pinhumanBtn.disabled = true;
        pinhumanBtn.innerHTML = '<i data-lucide="loader-2" class="btn-icon loading-spin"></i> Gönderiliyor...';
        lucide.createIcons();
        
        // Durdur butonunu oluştur ve yanına ekle
        const stopBtn = document.createElement('button');
        stopBtn.className = 'stop-btn';
        stopBtn.innerHTML = '<i data-lucide="square"></i> Durdur';
        stopBtn.onclick = stopPinhumanProcess;
        
        // Pinhuman butonunun yanına ekle
        pinhumanBtn.parentNode.insertBefore(stopBtn, pinhumanBtn.nextSibling);
        lucide.createIcons();
        
        addLogEntry('info', 'Bağlantı bilgileri kontrol ediliyor...');
        
        // Ayarları config.json'dan al
        const config = await ipcRenderer.invoke('get-config');
        const finalSettings = config?.pinhuman?.credentials;
        
        if (!finalSettings || !finalSettings.userName || !finalSettings.companyCode || !finalSettings.password || !finalSettings.totpSecret) {
            addLogEntry('error', '❌ Ayarlar eksik! Lütfen config.json dosyasını kontrol edin.');
            return;
        }
        
        addLogEntry('info', 'Pinhuman sistemine bağlanılıyor...');
        const result = await ipcRenderer.invoke('enter-data-pinhuman', finalSettings);
        
        if (result.success) {
            addLogEntry('success', '✅ ' + result.message);
        } else {
            addLogEntry('error', '❌ ' + result.message);
        }
        
    } catch (error) {
        console.error('Pinhuman gönderim hatası:', error);
        addLogEntry('error', '❌ Pinhuman\'a veri gönderilirken hata oluştu: ' + error.message);
    } finally {
        // Durdur butonunu kaldır
        const stopBtn = document.querySelector('.stop-btn');
        if (stopBtn) {
            stopBtn.remove();
        }
        
        // Buton durumunu geri getir
        pinhumanBtn.disabled = false;
        pinhumanBtn.innerHTML = '<i data-lucide="upload" class="btn-icon"></i> Pinhuman\'a Gönder';
        lucide.createIcons();
        
        addLogEntry('info', 'İşlem tamamlandı. Log alanını kapatmak için "Kapat" butonuna tıklayın.');
    }
}

// Pinhuman işlemini durdurma fonksiyonu
async function stopPinhumanProcess() {
    try {
        addLogEntry('warning', '⏹️ İşlem durduruluyor...');
        
        // Main process'e durdurma isteği gönder
        await ipcRenderer.invoke('stop-pinhuman-process');
        
        // Durdur butonunu kaldır
        const stopBtn = document.querySelector('.stop-btn');
        if (stopBtn) {
            stopBtn.remove();
        }
        
        // Buton durumunu geri getir
        pinhumanBtn.disabled = false;
        pinhumanBtn.innerHTML = '<i data-lucide="upload" class="btn-icon"></i> Pinhuman\'a Gönder';
        lucide.createIcons();
        
        addLogEntry('info', '✅ İşlem başarıyla durduruldu.');
    } catch (error) {
        console.error('Durdurma hatası:', error);
        addLogEntry('error', '❌ İşlem durdurulurken hata oluştu: ' + error.message);
    }
}

// Yardım maili açma fonksiyonu
function openHelpMail() {
    const subject = 'PDKS Sistemi - Hata Alıyorum';
    const body = `Merhaba,

PDKS sistemi ile ilgili bir sorun yaşıyorum. Lütfen yardımcı olabilir misiniz?

Detaylar:
- Hata mesajı: 
- Tarih: ${new Date().toLocaleDateString('tr-TR')}
- Saat: ${new Date().toLocaleTimeString('tr-TR')}

Teşekkürler.`;
    
    const mailtoLink = `mailto:furkan.ozmen@guleryuzgroup.com?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    
    // Outlook Classic'i açmak için özel protokol kullan
    const outlookLink = `outlook:${mailtoLink}`;
    
    try {
        // Önce outlook protokolünü dene
        window.open(outlookLink, '_blank');
    } catch (error) {
        // Outlook protokolü çalışmazsa normal mailto kullan
        window.open(mailtoLink, '_blank');
    }
}

// Sayfa geçiş sistemi
document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', () => {
        const pageId = item.dataset.page;
        switchPage(pageId);
    });
});

// Filtre event listeners (sadece personel sayfasında)
function setupFilterListeners() {
    const sicilFilterEl = document.getElementById('sicilFilter');
    if (sicilFilterEl) {
        sicilFilterEl.addEventListener('change', applyFilters);
    }
    
    const personelFilterEl = document.getElementById('personelFilter');
    if (personelFilterEl) {
        personelFilterEl.addEventListener('change', applyFilters);
    }
    
    const vardiyaFilterEl = document.getElementById('vardiyaFilter');
    if (vardiyaFilterEl) {
        vardiyaFilterEl.addEventListener('change', applyFilters);
    }
    
    const durumFilterEl = document.getElementById('durumFilter');
    if (durumFilterEl) {
        durumFilterEl.addEventListener('change', applyFilters);
    }
    
    const tarihFilterEl = document.getElementById('tarihFilter');
    if (tarihFilterEl) {
        tarihFilterEl.addEventListener('change', applyFilters);
    }
}

// Sayfa geçiş fonksiyonu
function switchPage(pageId) {
    // Tüm sayfaları gizle
    document.querySelectorAll('.page').forEach(page => {
        page.classList.remove('active');
    });
    
    // Tüm nav itemları pasif yap
    document.querySelectorAll('.nav-item').forEach(item => {
        item.classList.remove('active');
    });
    
    // Seçili sayfayı göster
    const targetPage = document.getElementById(pageId + '-page');
    if (targetPage) {
        targetPage.classList.add('active');
    }
    
    // Seçili nav item'ı aktif yap
    const activeNavItem = document.querySelector(`[data-page="${pageId}"]`);
    if (activeNavItem) {
        activeNavItem.classList.add('active');
    }
    
    // Personel sayfasına geçildiğinde filtreleri kur
    if (pageId === 'personel') {
        setupFilterListeners();
    }
    
    // Vardiya sayfasına geçildiğinde analizi güncelle
    if (pageId === 'vardiya') {
        updateVardiyaAnalysis();
    }
    
    // Vardiya analizi sayfasına geçildiğinde analizi yükle
    if (pageId === 'vardiya-analiz') {
        loadVardiyaAnalizi();
    }
    
    // Geçmiş sayfasına geçildiğinde geçmişi güncelle
    if (pageId === 'gecmis') {
        updateHistoryList();
    }
    
    // Ayarlar sayfasına geçildiğinde ayarları yükle
    if (pageId === 'ayarlar') {
        loadSettings();
    }
}

// Dosya seçme fonksiyonu
async function selectFile() {
    try {
        const filePath = await ipcRenderer.invoke('select-file');
        if (filePath) {
            selectedFilePath = filePath;
            showSelectedFile(filePath);
            showProcessSection();
        }
    } catch (error) {
        showError('Dosya seçilirken hata oluştu: ' + error.message);
    }
}

// Seçilen dosyayı göster
function showSelectedFile(filePath) {
    const fileNameOnly = path.basename(filePath);
    const fileDir = path.dirname(filePath);
    
    fileName.textContent = fileNameOnly;
    filePath.textContent = fileDir;
    
    fileInputArea.style.display = 'none';
    selectedFile.style.display = 'block';
    selectedFile.classList.add('fade-in');
}

// İşlem bölümünü göster
function showProcessSection() {
    processSection.style.display = 'block';
    processSection.classList.add('slide-up');
}

// Dosya işleme fonksiyonu
async function processFile() {
    if (!selectedFilePath) {
        showError('Lütfen önce bir dosya seçin.');
        return;
    }

    // Loading overlay göster
    showLoadingOverlay();
    
    try {
        // İşlemi başlat
        const result = await ipcRenderer.invoke('process-pdks', selectedFilePath);
        
        // 1.5 saniye bekle (loading overlay sabit kalacak)
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        // Loading overlay'i gizle
        hideLoadingOverlay();
        
        if (result.success) {
            showResults(result.data);
            outputPaths = {
                json: result.data.json_path,
                csv: result.data.csv_path
            };
            
            // Veri geçmişine ekle
            const filename = selectedFilePath ? path.basename(selectedFilePath) : 'Bilinmeyen Dosya';
            addToHistory(result.data, filename);
        } else {
            showError(result.message);
        }
    } catch (error) {
        hideLoadingOverlay();
        showError('İşlem sırasında hata oluştu: ' + error.message);
    }
}

// Progress animasyonu
function animateProgress() {
    let progress = 0;
    const interval = setInterval(() => {
        progress += Math.random() * 15;
        if (progress > 90) progress = 90;
        
        progressFill.style.width = progress + '%';
        
        if (progress < 30) {
            progressText.textContent = 'Dosya okunuyor...';
        } else if (progress < 60) {
            progressText.textContent = 'Veriler işleniyor...';
        } else if (progress < 90) {
            progressText.textContent = 'Raporlar oluşturuluyor...';
        }
    }, 200);
    
    // 3 saniye sonra animasyonu durdur
    setTimeout(() => {
        clearInterval(interval);
        progressFill.style.width = '100%';
        progressText.textContent = 'Tamamlandı!';
    }, 3000);
}

// Loading overlay göster
function showLoadingOverlay() {
    // Overlay oluştur
    const overlay = document.createElement('div');
    overlay.id = 'loadingOverlay';
    overlay.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.7);
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        z-index: 9999;
        font-family: 'Montserrat', sans-serif;
    `;
    
    overlay.innerHTML = `
        <div class="loading-card">
            <div class="loading-spinner"></div>
            <h3 class="loading-title">İşlem Yapılıyor...</h3>
            <h3 class="loading-subtitle">Lütfen bekleyin</h3>
        </div>
        <style>
            .loading-card {
                background: var(--card-bg, white);
                border-radius: 15px;
                padding: 40px;
                text-align: center;
                box-shadow: 0 10px 30px var(--shadow-color, rgba(0,0,0,0.3));
                max-width: 350px;
                width: 90%;
                border: 1px solid var(--border-color, #e1e8ed);
            }
            
            
            .loading-spinner {
                width: 50px;
                height: 50px;
                border: 4px solid var(--bg-tertiary, #f3f3f3);
                border-top: 4px solid #3498db;
                border-radius: 50%;
                animation: spin 1s linear infinite;
                margin: 0 auto 20px;
            }
            
            .loading-title {
                margin: 0;
                font-size: 1.3rem;
                font-weight: 600;
                color: var(--text-primary, #2c3e50);
                margin-bottom: 10px;
            }
            
            .loading-subtitle {
                margin: 0;
                color: var(--text-secondary, #7f8c8d);
                font-size: 0.9rem;
            }
            
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
        </style>
    `;
    
    document.body.appendChild(overlay);
}

// Loading overlay gizle
function hideLoadingOverlay() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.remove();
    }
}

// Sonuçları göster
function showResults(data) {
    hideAllSections();
    
    // Ana sonuçları güncelle
    totalRecords.textContent = data.toplam_kayit.toLocaleString('tr-TR');
    personelCount.textContent = data.personel_sayisi.toLocaleString('tr-TR');
    errorCount.textContent = data.hata_sayisi.toLocaleString('tr-TR');
    
    // Sidebar istatistiklerini güncelle (artık sidebar'da istatistik yok)
    
    // Dosya linklerini ayarla
    
    // Veriyi yükle ve sidebar'ı güncelle
    loadDataToSidebar(data);
    
    resultsSection.style.display = 'block';
    resultsSection.classList.add('fade-in');
}

// Hata göster
function showError(message) {
    hideAllSections();
    errorMessage.textContent = message;
    errorSection.style.display = 'block';
    errorSection.classList.add('fade-in');
}

// Hata gizle
function hideError() {
    errorSection.style.display = 'none';
}

// Progress göster
function showProgress() {
    progressSection.style.display = 'block';
    progressSection.classList.add('slide-up');
}

// Tüm bölümleri gizle
function hideAllSections() {
    processSection.style.display = 'none';
    progressSection.style.display = 'none';
    resultsSection.style.display = 'none';
    errorSection.style.display = 'none';
    logSection.style.display = 'none';
}

// Uygulamayı sıfırla
function resetApp() {
    selectedFilePath = null;
    outputPaths = {};
    currentData = [];
    filteredData = [];
    
    hideAllSections();
    fileInputArea.style.display = 'block';
    selectedFile.style.display = 'none';
    
    // Sidebar'ı sıfırla
    resetSidebar();
}

// Sidebar'ı sıfırla
function resetSidebar() {
    const sicilFilterEl = document.getElementById('sicilFilter');
    if (sicilFilterEl) {
        sicilFilterEl.innerHTML = '<option value="">Tüm Sicil Numaraları</option>';
    }
    
    const personelFilterEl = document.getElementById('personelFilter');
    if (personelFilterEl) {
        personelFilterEl.innerHTML = '<option value="">Tüm Personeller</option>';
    }
    
    if (dataTableBody) {
        dataTableBody.innerHTML = `
            <tr>
                <td colspan="9" style="text-align: center; padding: 40px; color: #7f8c8d;">
                    <i data-lucide="database" style="font-size: 2rem; margin-bottom: 10px; display: block;"></i>
                    Henüz veri yüklenmedi
                </td>
            </tr>
        `;
    }
}

// Veriyi sidebar'a yükle
function loadDataToSidebar(data) {
    // JSON dosyasını oku
    fetch('file://' + data.json_path)
        .then(response => response.json())
        .then(jsonData => {
            currentData = jsonData.veriler || [];
            filteredData = [...currentData];
            
            // Filtreleri güncelle
            updateSicilFilter();
            updatePersonelFilter();
            
            // Veri listesini güncelle
            updateDataList();
        })
        .catch(error => {
            console.error('Veri yüklenirken hata:', error);
        });
}

// Sicil filtresini güncelle
function updateSicilFilter() {
    const sicilNumaralari = [...new Set(currentData.map(item => item.sicilNo).filter(sicil => sicil && sicil.trim() !== ''))].sort();
    
    const sicilFilter = document.getElementById('sicilFilter');
    if (sicilFilter) {
        sicilFilter.innerHTML = '<option value="">Tüm Sicil Numaraları</option>';
        sicilNumaralari.forEach(sicil => {
            const option = document.createElement('option');
            option.value = sicil;
            option.textContent = sicil;
            sicilFilter.appendChild(option);
        });
    }
}

// Personel filtresini güncelle
function updatePersonelFilter() {
    // Personel isimlerini temizle ve benzersiz hale getir
    const personeller = [...new Set(currentData.map(item => {
        return item.personel ? item.personel.split(' ').filter(part => !part.match(/^\d{1,2}:\d{2}$/)).join(' ').trim() : 'Bilinmeyen';
    }))].sort();
    
    const personelFilter = document.getElementById('personelFilter');
    if (personelFilter) {
        personelFilter.innerHTML = '<option value="">Tüm Personeller</option>';
        personeller.forEach(personel => {
            const option = document.createElement('option');
            option.value = personel;
            option.textContent = personel;
            personelFilter.appendChild(option);
        });
    }
}

// Veri listesini güncelle
function updateDataList() {
    if (filteredData.length === 0) {
        dataTableBody.innerHTML = `
            <tr>
                <td colspan="9" style="text-align: center; padding: 20px; color: #7f8c8d;">
                    Filtre kriterlerine uygun veri bulunamadı
                </td>
            </tr>
        `;
        return;
    }
    
    dataTableBody.innerHTML = filteredData.map(item => {
        // Personel adını temizle (saat bilgisi varsa kaldır)
        const personelAdi = item.personel ? item.personel.split(' ').filter(part => !part.match(/^\d{1,2}:\d{2}$/)).join(' ').trim() : 'Bilinmeyen';
        
        
        return `
            <tr onclick="showItemDetails('${personelAdi}', '${item.tarih}')">
                <td>${item.sicilNo || '-'}</td>
                <td>${personelAdi}</td>
                <td>${formatDate(item.tarih)}</td>
                <td>${item.vardiya ? getVardiyaTimeRange(item.vardiya.kod) : '-'}</td>
                <td>${item.gercek ? item.gercek.gir : '-'}</td>
                <td>${item.gercek ? item.gercek.cik : '-'}</td>
                <td>${formatDuration(item.calisma_dk)}</td>
                <td>${item.fm_hhmm}</td>
                <td><span class="data-status ${item.durum}">${item.durum}</span></td>
            </tr>
        `;
    }).join('');
    
    // Iconları yeniden oluştur
    lucide.createIcons();
}

// Vardiya analizini güncelle
function updateVardiyaAnalysis() {
    if (!currentData || currentData.length === 0) {
        resetVardiyaAnalysis();
        return;
    }
    
    // Vardiya bazında verileri grupla
    const vardiyaData = {
        V1: { personel: new Set(), toplamCalisma: 0, toplamFM: 0, kayitlar: [] },
        V2: { personel: new Set(), toplamCalisma: 0, toplamFM: 0, kayitlar: [] },
        V3: { personel: new Set(), toplamCalisma: 0, toplamFM: 0, kayitlar: [] }
    };
    
    currentData.forEach(item => {
        if (item.vardiya && item.vardiya.kod) {
            const vardiya = item.vardiya.kod;
            if (vardiyaData[vardiya]) {
                vardiyaData[vardiya].personel.add(item.personel);
                vardiyaData[vardiya].toplamCalisma += item.calisma_dk;
                vardiyaData[vardiya].toplamFM += parseFMToMinutes(item.fm_hhmm);
                vardiyaData[vardiya].kayitlar.push(item);
            }
        }
    });
    
    // Vardiya kartlarını güncelle
    updateVardiyaCard('v1', vardiyaData.V1);
    updateVardiyaCard('v2', vardiyaData.V2);
    updateVardiyaCard('v3', vardiyaData.V3);
    
    // Vardiya tablosunu güncelle
    updateVardiyaTable(vardiyaData);
}

// Vardiya kartını güncelle
function updateVardiyaCard(vardiyaId, data) {
    const personelEl = document.getElementById(`${vardiyaId}-personel`);
    const calismaEl = document.getElementById(`${vardiyaId}-calisma`);
    const fmEl = document.getElementById(`${vardiyaId}-fm`);
    
    if (personelEl) personelEl.textContent = data.personel.size;
    if (calismaEl) calismaEl.textContent = formatDuration(data.toplamCalisma);
    if (fmEl) {
        const ortalamaFM = data.personel.size > 0 ? data.toplamFM / data.personel.size : 0;
        fmEl.textContent = formatDuration(Math.round(ortalamaFM));
    }
}

// Vardiya tablosunu güncelle
function updateVardiyaTable(vardiyaData) {
    const tbody = document.getElementById('vardiyaTableBody');
    if (!tbody) return;
    
    const allRecords = [];
    Object.keys(vardiyaData).forEach(vardiya => {
        vardiyaData[vardiya].kayitlar.forEach(kayit => {
            allRecords.push({ ...kayit, vardiyaKod: vardiya });
        });
    });
    
    if (allRecords.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" style="text-align: center; padding: 40px; color: #7f8c8d;">
                    <i data-lucide="clock" style="font-size: 2rem; margin-bottom: 10px; display: block;"></i>
                    Henüz veri yüklenmedi
                </td>
            </tr>
        `;
        return;
    }
    
    // Vardiya bazında sırala
    allRecords.sort((a, b) => {
        const vardiyaOrder = { V1: 1, V2: 2, V3: 3 };
        return vardiyaOrder[a.vardiyaKod] - vardiyaOrder[b.vardiyaKod];
    });
    
    tbody.innerHTML = allRecords.map(item => {
        // Personel adını temizle (saat bilgisi varsa kaldır)
        const personelAdi = item.personel ? item.personel.split(' ').filter(part => !part.match(/^\d{1,2}:\d{2}$/)).join(' ').trim() : 'Bilinmeyen';
        
        return `
            <tr onclick="showItemDetails('${personelAdi}', '${item.tarih}')">
                <td>${item.sicilNo || '-'}</td>
                <td><span class="vardiya-badge ${item.vardiyaKod.toLowerCase()}">${getVardiyaTimeRange(item.vardiyaKod)}</span></td>
                <td>${personelAdi}</td>
                <td>${item.gercek ? item.gercek.gir : '-'}</td>
                <td>${item.gercek ? item.gercek.cik : '-'}</td>
                <td>${formatDuration(item.calisma_dk)}</td>
                <td>${item.fm_hhmm}</td>
                <td><span class="data-status ${item.durum}">${item.durum}</span></td>
            </tr>
        `;
    }).join('');
    
    // Iconları yeniden oluştur
    lucide.createIcons();
}

// Vardiya analizini sıfırla
function resetVardiyaAnalysis() {
    ['v1', 'v2', 'v3'].forEach(vardiya => {
        const personelEl = document.getElementById(`${vardiya}-personel`);
        const calismaEl = document.getElementById(`${vardiya}-calisma`);
        const fmEl = document.getElementById(`${vardiya}-fm`);
        
        if (personelEl) personelEl.textContent = '0';
        if (calismaEl) calismaEl.textContent = '0:00';
        if (fmEl) fmEl.textContent = '0:00';
    });
    
    const tbody = document.getElementById('vardiyaTableBody');
    if (tbody) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" style="text-align: center; padding: 40px; color: #7f8c8d;">
                    <i data-lucide="clock" style="font-size: 2rem; margin-bottom: 10px; display: block;"></i>
                    Henüz veri yüklenmedi
                </td>
            </tr>
        `;
    }
}

// FM string'ini dakikaya çevir
function parseFMToMinutes(fmString) {
    if (!fmString || fmString === '00:00') return 0;
    const [hours, minutes] = fmString.split(':').map(Number);
    return hours * 60 + minutes;
}

// Vardiya saatlerini yuvarla
function roundToShiftTime(timeString, vardiyaKod) {
    if (!timeString || timeString === '-') return timeString;
    
    const [hours, minutes] = timeString.split(':').map(Number);
    const totalMinutes = hours * 60 + minutes;
    
    // Vardiya saatleri (dakika cinsinden)
    const shiftTimes = {
        'V1': { start: 8 * 60 + 30, end: 16 * 60 + 30 }, // 08:30 - 16:30
        'V2': { start: 16 * 60 + 30, end: 24 * 60 + 30 }, // 16:30 - 00:30
        'V3': { start: 0 * 60 + 30, end: 8 * 60 + 30 }    // 00:30 - 08:30
    };
    
    if (!vardiyaKod || !shiftTimes[vardiyaKod]) {
        return timeString; // Vardiya bilgisi yoksa orijinal saati döndür
    }
    
    const shift = shiftTimes[vardiyaKod];
    
    // Eğer çıkış saati vardiya bitiş saatinden önceyse, vardiya bitiş saatine yuvarla
    if (totalMinutes < shift.end) {
        const roundedHours = Math.floor(shift.end / 60);
        const roundedMinutes = shift.end % 60;
        return `${roundedHours.toString().padStart(2, '0')}:${roundedMinutes.toString().padStart(2, '0')}`;
    }
    
    // Eğer çıkış saati vardiya bitiş saatinden sonraysa, 30 dakikalık aralıklara yuvarla
    const roundedMinutes = Math.round(totalMinutes / 30) * 30;
    const roundedHours = Math.floor(roundedMinutes / 60);
    const finalMinutes = roundedMinutes % 60;
    
    return `${roundedHours.toString().padStart(2, '0')}:${finalMinutes.toString().padStart(2, '0')}`;
}

// FM hesapla (yuvarlanmış saatlere göre)
function calculateFMWithRoundedTime(giris, cikis, vardiyaKod) {
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

// Yuvarlanma kurallarını al
function getYuvarlanmaKurallari() {
    try {
        const kurallar = JSON.parse(localStorage.getItem('yuvarlanma_kurallari') || '{}');
        
        // Varsayılan değerler
        const defaultKurallar = {
            vardiyaV1Bitis: '16:30',
            vardiyaV2Bitis: '00:30',
            vardiyaV3Bitis: '08:30',
            yuvarlanmaAraligi: 30,
            yuvarlanmaToleransi: 5
        };
        
        return { ...defaultKurallar, ...kurallar };
    } catch (error) {
        console.error('Yuvarlanma kuralları yüklenirken hata:', error);
        return {
            vardiyaV1Bitis: '16:30',
            vardiyaV2Bitis: '00:30',
            vardiyaV3Bitis: '08:30',
            yuvarlanmaAraligi: 30,
            yuvarlanmaToleransi: 5
        };
    }
}

// Giriş ve çıkış saatlerinden vardiya tespit et
function detectShiftFromTimes(giris, cikis) {
    if (!giris || !cikis || giris === '-' || cikis === '-') {
        return 'Belirsiz';
    }
    
    const girisDakika = parseTimeToMinutes(giris);
    const cikisDakika = parseTimeToMinutes(cikis);
    
    // V1: 08:30 - 16:30 (510 - 990 dakika)
    // V2: 16:30 - 00:30 (990 - 30 dakika, gece vardiyası)
    // V3: 00:30 - 08:30 (30 - 510 dakika, gece vardiyası)
    
    // V1 kontrolü (gündüz vardiyası)
    if (girisDakika >= 480 && girisDakika <= 540 && cikisDakika >= 900 && cikisDakika <= 1020) {
        return 'V1';
    }
    
    // V2 kontrolü (akşam vardiyası)
    if ((girisDakika >= 960 || girisDakika <= 60) && (cikisDakika >= 0 && cikisDakika <= 120)) {
        return 'V2';
    }
    
    // V3 kontrolü (gece vardiyası)
    if (girisDakika >= 0 && girisDakika <= 120 && cikisDakika >= 480 && cikisDakika <= 540) {
        return 'V3';
    }
    
    // Alternatif V2 kontrolü (16:30-00:30 arası)
    if (girisDakika >= 990 && cikisDakika <= 30) {
        return 'V2';
    }
    
    // Alternatif V3 kontrolü (00:30-08:30 arası)
    if (girisDakika <= 30 && cikisDakika >= 480) {
        return 'V3';
    }
    
    return 'Belirsiz';
}


// Saati dakikaya çevir
function parseTimeToMinutes(timeString) {
    const [hours, minutes] = timeString.split(':').map(Number);
    return hours * 60 + minutes;
}

// Dakikayı saate çevir
function formatMinutesToTime(totalMinutes) {
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
}

// Veri geçmişini yükle
function loadDataHistory() {
    try {
        const historyData = localStorage.getItem('pdks_data_history');
        if (historyData) {
            dataHistory = JSON.parse(historyData);
        }
    } catch (error) {
        console.error('Veri geçmişi yüklenirken hata:', error);
        dataHistory = [];
    }
}

// Veri geçmişini kaydet
function saveDataHistory() {
    try {
        localStorage.setItem('pdks_data_history', JSON.stringify(dataHistory));
    } catch (error) {
        console.error('Veri geçmişi kaydedilirken hata:', error);
    }
}

// Veri geçmişine ekle
function addToHistory(data, filename) {
    const historyItem = {
        id: Date.now(),
        filename: filename,
        date: new Date().toISOString(),
        stats: {
            toplam_kayit: data.toplam_kayit,
            personel_sayisi: data.personel_sayisi,
            hata_sayisi: data.hata_sayisi
        },
        data: data
    };
    
    // Aynı dosya adı varsa güncelle, yoksa ekle
    const existingIndex = dataHistory.findIndex(item => item.filename === filename);
    if (existingIndex >= 0) {
        dataHistory[existingIndex] = historyItem;
    } else {
        dataHistory.unshift(historyItem);
    }
    
    // Son 50 kaydı tut
    if (dataHistory.length > 50) {
        dataHistory = dataHistory.slice(0, 50);
    }
    
    saveDataHistory();
}

// Geçmiş listesini güncelle
function updateHistoryList() {
    const historyTableBody = document.getElementById('historyTableBody');
    if (!historyTableBody) return;
    
    if (dataHistory.length === 0) {
        historyTableBody.innerHTML = `
            <tr>
                <td colspan="6" style="text-align: center; padding: 40px; color: #7f8c8d;">
                    <i data-lucide="history" style="font-size: 2rem; margin-bottom: 10px; display: block;"></i>
                    Henüz veri geçmişi yok
                </td>
            </tr>
        `;
        return;
    }
    
    historyTableBody.innerHTML = dataHistory.map(item => `
        <tr>
            <td class="history-filename">${item.filename}</td>
            <td><span class="history-date">${formatDate(item.date)}</span></td>
            <td class="history-stat">${item.stats.toplam_kayit}</td>
            <td class="history-stat">${item.stats.personel_sayisi}</td>
            <td class="history-stat">${item.stats.hata_sayisi}</td>
            <td class="history-actions">
                <button class="history-action-btn" onclick="loadHistoryData('${item.id}')" title="Veriyi Yükle">
                    <i data-lucide="upload"></i>
                    Yükle
                </button>
                <button class="history-action-btn delete" onclick="deleteHistoryItem('${item.id}')" title="Kaydı Sil">
                    <i data-lucide="trash-2"></i>
                    Sil
                </button>
            </td>
        </tr>
    `).join('');
    
    lucide.createIcons();
}

// Geçmiş verisini yükle
function loadHistoryData(historyId) {
    const historyItem = dataHistory.find(item => item.id == historyId);
    if (historyItem) {
        currentData = historyItem.data.kayitlar || [];
        filteredData = [...currentData];
        
        // Personel sayfasına geç ve verileri göster
        switchPage('personel');
        updateDataList();
        updateSicilFilter();
        updatePersonelFilter();
    }
}

// Geçmişi temizle
function clearHistory() {
    if (confirm('Tüm veri geçmişini silmek istediğinizden emin misiniz?')) {
        dataHistory = [];
        saveDataHistory();
        updateHistoryList();
    }
}

// Tek kayıt sil
function deleteHistoryItem(historyId) {
    const historyItem = dataHistory.find(item => item.id == historyId);
    if (historyItem && confirm(`"${historyItem.filename}" kaydını silmek istediğinizden emin misiniz?`)) {
        dataHistory = dataHistory.filter(item => item.id != historyId);
        saveDataHistory();
        updateHistoryList();
    }
}

// Rapor oluştur
async function generateReport(reportType) {
    if (!currentData || currentData.length === 0) {
        alert('Rapor oluşturmak için önce veri yükleyin!');
        return;
    }
    
    try {
        const reportData = await ipcRenderer.invoke('generate-report', {
            type: reportType,
            data: currentData
        });
        
        if (reportData.success) {
            // Dosyayı indir
            const link = document.createElement('a');
            link.href = 'file://' + reportData.filePath;
            link.download = reportData.filename;
            link.click();
        } else {
            alert('Rapor oluşturulurken hata: ' + reportData.message);
        }
    } catch (error) {
        console.error('Rapor oluşturma hatası:', error);
        alert('Rapor oluşturulurken hata oluştu!');
    }
}

// Filtreleri uygula
function applyFilters() {
    const sicilFilterEl = document.getElementById('sicilFilter');
    const personelFilterEl = document.getElementById('personelFilter');
    const vardiyaFilterEl = document.getElementById('vardiyaFilter');
    const durumFilterEl = document.getElementById('durumFilter');
    const tarihFilterEl = document.getElementById('tarihFilter');
    
    if (!sicilFilterEl || !personelFilterEl || !vardiyaFilterEl || !durumFilterEl || !tarihFilterEl) return;
    
    const selectedSicil = sicilFilterEl.value;
    const selectedPersonel = personelFilterEl.value;
    const selectedVardiya = vardiyaFilterEl.value;
    const selectedDurum = durumFilterEl.value;
    const selectedTarih = tarihFilterEl.value;
    
    // Filtrele
    filteredData = currentData.filter(item => {
        // Personel adını temizle
        const personelAdi = item.personel ? item.personel.split(' ').filter(part => !part.match(/^\d{1,2}:\d{2}$/)).join(' ').trim() : 'Bilinmeyen';
        
        const sicilMatch = !selectedSicil || item.sicilNo === selectedSicil;
        const personelMatch = !selectedPersonel || personelAdi === selectedPersonel;
        const vardiyaMatch = !selectedVardiya || (item.vardiya && item.vardiya.kod === selectedVardiya);
        const durumMatch = !selectedDurum || item.durum === selectedDurum;
        const tarihMatch = !selectedTarih || item.tarih === selectedTarih;
        
        return sicilMatch && personelMatch && vardiyaMatch && durumMatch && tarihMatch;
    });
    
    // Listeyi güncelle
    updateDataList();
}

// Yardımcı fonksiyonlar
function formatDate(dateStr) {
    const date = new Date(dateStr);
    return date.toLocaleDateString('tr-TR', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
    });
}

function formatDuration(minutes) {
    const hours = Math.floor(minutes / 60);
    const mins = minutes % 60;
    return `${hours}:${mins.toString().padStart(2, '0')}`;
}

function showItemDetails(personel, tarih) {
    const item = currentData.find(d => d.personel === personel && d.tarih === tarih);
    if (item) {
        alert(`Detaylar:\n\nPersonel: ${item.personel}\nTarih: ${formatDate(item.tarih)}\nVardiya: ${item.vardiya ? item.vardiya.kod : '-'}\nGiriş: ${item.gercek.gir}\nÇıkış: ${item.gercek.cik}\nÇalışma: ${formatDuration(item.calisma_dk)}\nFM: ${item.fm_hhmm}\nDurum: ${item.durum}`);
    }
}

// Çıktı klasörünü aç
function openOutputFolder() {
    if (outputPaths.json) {
        const outputDir = path.dirname(outputPaths.json);
        shell.openPath(outputDir);
    }
}

// Ayarlar sayfası fonksiyonları
async function loadSettings() {
    try {
        // Config.json'dan ayarları al
        const config = await ipcRenderer.invoke('get-config');
        const settings = config?.pinhuman?.credentials || {};
        
        document.getElementById('userName').value = settings.userName || '';
        document.getElementById('companyCode').value = settings.companyCode || '';
        
        // Şifre alanını kayıtlı şifre varsa gizli olarak göster, yoksa boş bırak
        if (settings.password && settings.password.trim() !== '') {
            document.getElementById('password').value = settings.password;
            document.getElementById('password').type = 'password'; // Şifreyi gizli göster
            document.getElementById('password').placeholder = 'Kayıtlı şifre mevcut. Yeni şifre girmek için yazın.';
        } else {
            document.getElementById('password').value = '';
            document.getElementById('password').type = 'password';
            document.getElementById('password').placeholder = 'Şifre kaydedilmemiş. Şifre girin.';
        }
        
        document.getElementById('totpSecret').value = settings.totpSecret || '';
        
        // UI ayarlarını yükle
        await loadUISettings();
        
        // Yuvarlanma kurallarını yükle
        loadYuvarlanmaKurallari();
        
        // Şifre alanı için event listener ekle
        const passwordField = document.getElementById('password');
        passwordField.addEventListener('focus', () => {
            // Alanı temizle ve text tipine çevir
            passwordField.value = '';
            passwordField.type = 'text';
            passwordField.placeholder = 'Yeni şifre girin...';
        });
        
        passwordField.addEventListener('blur', async () => {
            if (passwordField.value === '') {
                // Kayıtlı şifre varsa gerçek şifreyi gizli olarak göster, yoksa boş placeholder
                const config = await ipcRenderer.invoke('get-config');
                const settings = config?.pinhuman?.credentials || {};
                if (settings.password && settings.password.trim() !== '') {
                    passwordField.value = settings.password;
                    passwordField.type = 'password';
                    passwordField.placeholder = 'Kayıtlı şifre mevcut. Yeni şifre girmek için yazın.';
                } else {
                    passwordField.type = 'password';
                    passwordField.placeholder = 'Şifre kaydedilmemiş. Şifre girin.';
                }
            }
        });
        
    } catch (error) {
        console.error('Ayarlar yüklenirken hata:', error);
    }
}

// UI ayarlarını yükle
async function loadUISettings() {
    try {
        const uiSettings = await ipcRenderer.invoke('load-ui-settings');
        
        // Tema ayarını uygula
        if (uiSettings.theme) {
            document.documentElement.setAttribute('data-theme', uiSettings.theme);
            const darkModeToggle = document.getElementById('darkModeToggle');
            if (darkModeToggle) {
                darkModeToggle.checked = uiSettings.theme === 'dark';
            }
        } else {
            // Varsayılan tema light
            document.documentElement.setAttribute('data-theme', 'light');
            const darkModeToggle = document.getElementById('darkModeToggle');
            if (darkModeToggle) {
                darkModeToggle.checked = false;
            }
        }
        
        // Dil ayarını uygula
        if (uiSettings.language) {
            const languageSelect = document.getElementById('languageSelect');
            if (languageSelect) {
                languageSelect.value = uiSettings.language;
            }
        }
        
        // Otomatik kaydet ayarını uygula
        if (uiSettings.autoSave !== undefined) {
            const autoSaveToggle = document.getElementById('autoSaveToggle');
            if (autoSaveToggle) {
                autoSaveToggle.checked = uiSettings.autoSave;
            }
        }
        
    } catch (error) {
        console.error('UI ayarları yüklenirken hata:', error);
    }
}

// UI ayarlarını kaydet
async function saveUISettings() {
    try {
        const settings = {};
        
        // Tema ayarını al
        const theme = document.documentElement.getAttribute('data-theme') || 'light';
        settings.theme = theme;
        
        // Dil ayarını al
        const languageSelect = document.getElementById('languageSelect');
        if (languageSelect) {
            settings.language = languageSelect.value;
        }
        
        // Otomatik kaydet ayarını al
        const autoSaveToggle = document.getElementById('autoSaveToggle');
        if (autoSaveToggle) {
            settings.autoSave = autoSaveToggle.checked;
        }
        
        // Ayarları kaydet
        const result = await ipcRenderer.invoke('save-ui-settings', settings);
        
        if (result.success) {
            console.log('UI ayarları başarıyla kaydedildi');
        } else {
            console.error('UI ayarları kaydedilemedi:', result.message);
        }
        
    } catch (error) {
        console.error('UI ayarları kaydedilirken hata:', error);
    }
}

// Yuvarlanma kurallarını yükle
function loadYuvarlanmaKurallari() {
    try {
        const kurallar = getYuvarlanmaKurallari();
        
        document.getElementById('vardiyaV1Bitis').value = kurallar.vardiyaV1Bitis;
        document.getElementById('vardiyaV2Bitis').value = kurallar.vardiyaV2Bitis;
        document.getElementById('vardiyaV3Bitis').value = kurallar.vardiyaV3Bitis;
        document.getElementById('yuvarlanmaAraligi').value = kurallar.yuvarlanmaAraligi;
        document.getElementById('yuvarlanmaToleransi').value = kurallar.yuvarlanmaToleransi;
        
    } catch (error) {
        console.error('Yuvarlanma kuralları yüklenirken hata:', error);
    }
}

async function saveSettings() {
    const userName = document.getElementById('userName').value;
    const companyCode = document.getElementById('companyCode').value;
    const password = document.getElementById('password').value;
    const totpSecret = document.getElementById('totpSecret').value;
    
    const settings = {
        userName: userName,
        companyCode: companyCode,
        totpSecret: totpSecret
    };
    
    // Sadece şifre girilmişse şifreyi güncelle
    if (password && password.trim() !== '') {
        settings.password = password;
    }
    
    try {
        const result = await ipcRenderer.invoke('update-config', settings);
        if (result.success) {
            // UI ayarlarını da kaydet
            await saveUISettings();
            alert('Ayarlar başarıyla kaydedildi!');
            // Şifre alanını kaydedilen şifre ile doldur ve gizli göster
            if (password && password.trim() !== '') {
                document.getElementById('password').value = password;
                document.getElementById('password').type = 'password';
            } else {
                // Şifre değiştirilmemişse mevcut şifreyi göster
                const config = await ipcRenderer.invoke('get-config');
                const settings = config?.pinhuman?.credentials || {};
                if (settings.password) {
                    document.getElementById('password').value = settings.password;
                    document.getElementById('password').type = 'password';
                }
            }
            document.getElementById('password').placeholder = 'Şifre başarıyla kaydedildi! Yeni şifre girmek için yazın.';
        } else {
            alert('Ayarlar kaydedilirken hata oluştu: ' + result.message);
        }
    } catch (error) {
        console.error('Ayarlar kaydedilirken hata:', error);
        alert('Ayarlar kaydedilirken hata oluştu!');
    }
}

// Yuvarlanma kurallarını kaydet
function saveYuvarlanmaKurallari() {
    const kurallar = {
        vardiyaV1Bitis: document.getElementById('vardiyaV1Bitis').value,
        vardiyaV2Bitis: document.getElementById('vardiyaV2Bitis').value,
        vardiyaV3Bitis: document.getElementById('vardiyaV3Bitis').value,
        yuvarlanmaAraligi: parseInt(document.getElementById('yuvarlanmaAraligi').value),
        yuvarlanmaToleransi: parseInt(document.getElementById('yuvarlanmaToleransi').value)
    };
    
    try {
        localStorage.setItem('yuvarlanma_kurallari', JSON.stringify(kurallar));
        alert('Yuvarlanma kuralları başarıyla kaydedildi!');
        
        // Vardiya analizi sayfasındaysa tabloyu yenile
        if (currentPage === 'vardiya-analiz') {
            displayVardiyaGruplari();
        }
    } catch (error) {
        console.error('Yuvarlanma kuralları kaydedilirken hata:', error);
        alert('Yuvarlanma kuralları kaydedilirken hata oluştu!');
    }
}


function testConnection() {
    const userName = document.getElementById('userName').value;
    const companyCode = document.getElementById('companyCode').value;
    const password = document.getElementById('password').value;
    const totpSecret = document.getElementById('totpSecret').value;
    
    if (!userName || !companyCode || !password || !totpSecret) {
        alert('Lütfen tüm alanları doldurun!');
        return;
    }
    
    // Bağlantı testi simülasyonu
    alert('Bağlantı testi başlatılıyor...\n\nBu özellik geliştirilme aşamasındadır.');
}

function togglePasswordVisibility() {
    const passwordField = document.getElementById('password');
    const eyeIcon = document.getElementById('eyeIcon');
    
    if (passwordField.type === 'password') {
        passwordField.type = 'text';
        eyeIcon.setAttribute('data-lucide', 'eye-off');
    } else {
        passwordField.type = 'password';
        eyeIcon.setAttribute('data-lucide', 'eye');
    }
    
    // Lucide iconları yenile
    lucide.createIcons();
}

async function enterDataToPinhuman() {
    try {
        // Ayarları config.json'dan al
        const config = await ipcRenderer.invoke('get-config');
        const finalSettings = config?.pinhuman?.credentials;
        
        if (!finalSettings || !finalSettings.userName || !finalSettings.companyCode || !finalSettings.password || !finalSettings.totpSecret) {
            alert('Lütfen config.json dosyasını kontrol edin!');
            return;
        }
        
        const { userName, companyCode, password, totpSecret } = finalSettings;
        
        // Loading overlay göster
        showLoadingOverlay();
        
        // Playwright ile Pinhuman'a giriş yap
        const result = await ipcRenderer.invoke('enter-data-pinhuman', {
            userName: userName,
            companyCode: companyCode,
            password: password,
            totpSecret: totpSecret
        });
        
        // Loading overlay'i gizle
        hideLoadingOverlay();
        
        if (result.success) {
            console.log('Pinhuman işlemi başarıyla tamamlandı!');
        } else {
            alert('Hata: ' + result.message);
        }
    } catch (error) {
        hideLoadingOverlay();
        console.error('Pinhuman giriş hatası:', error);
        alert('Pinhuman\'a giriş yapılırken hata oluştu: ' + error.message);
    }
}

// Excel verilerini Pinhuman'a gir - dinamik personel eşleştirmesi ile
async function enterExcelDataToPinhuman() {
    try {
        // Ayarları config.json'dan al
        const config = await ipcRenderer.invoke('get-config');
        const finalSettings = config?.pinhuman?.credentials;
        
        if (!finalSettings || !finalSettings.userName || !finalSettings.companyCode || !finalSettings.password || !finalSettings.totpSecret) {
            alert('Lütfen config.json dosyasını kontrol edin!');
            return;
        }
        
        const { userName, companyCode, password, totpSecret } = finalSettings;
        
        // Loading overlay göster
        showLoadingOverlay();
        
        // Excel verilerini Pinhuman'a gir
        const result = await ipcRenderer.invoke('enter-excel-data-pinhuman', {
            userName: userName,
            companyCode: companyCode,
            password: password,
            totpSecret: totpSecret
        });
        
        // Loading overlay'i gizle
        hideLoadingOverlay();
        
        if (result.success) {
            console.log('Excel verileri Pinhuman\'a başarıyla girildi!');
            alert(result.message);
        } else {
            alert('Hata: ' + result.message);
        }
    } catch (error) {
        hideLoadingOverlay();
        console.error('Excel veri girişi hatası:', error);
        alert('Excel veri girişi sırasında hata oluştu: ' + error.message);
    }
}


// Vardiya analizi yükleme fonksiyonu
async function loadVardiyaAnalizi() {
    try {
        const result = await ipcRenderer.invoke('get-shift-analysis');
        
        if (result.success) {
            displayVardiyaAnalizi(result.data);
        } else {
            console.error('Vardiya analizi yüklenemedi:', result.message);
            showVardiyaAnaliziError(result.message);
        }
    } catch (error) {
        console.error('Vardiya analizi hatası:', error);
        showVardiyaAnaliziError('Vardiya analizi yüklenirken hata oluştu: ' + error.message);
    }
}

// Vardiya analizi verilerini ekranda göster
function displayVardiyaAnalizi(data) {
    // İstatistikleri güncelle
    document.getElementById('totalPersonel').textContent = Object.keys(data.personelVardiya).length;
    document.getElementById('totalKayit').textContent = data.toplamKayit;
    document.getElementById('enCokVardiya').textContent = data.istatistikler.enCokCalisanVardiya;
    document.getElementById('enAzVardiya').textContent = data.istatistikler.enAzCalisanVardiya;
    
    // Tarih bazlı vardiya gruplarını göster
    displayVardiyaGruplari();
    
    // Vardiya kombinasyonlarını göster
    displayVardiyaKombinasyonlari(data.vardiyaKombinasyonlari);
}

// Vardiya kodunu saat aralığına çevir
function getVardiyaTimeRange(vardiyaKod) {
    const vardiyaSaatleri = {
        'V1': '08:30 - 16:30',
        'V2': '16:30 - 00:30',
        'V3': '00:30 - 08:30'
    };
    return vardiyaSaatleri[vardiyaKod] || vardiyaKod;
}

// Vardiya gruplarını tablo formatında göster






// Verileri tarih ve vardiya bazında grupla
function groupDataByDateAndShift(data) {
    const gruplar = {};
    
    data.forEach(item => {
        const tarih = item.tarih;
        const vardiyaKod = item.vardiya ? item.vardiya.kod : 'Belirsiz';
        
        // Personel adını temizle
        const personelAdi = item.personel ? item.personel.split(' ').filter(part => !part.match(/^\d{1,2}:\d{2}$/)).join(' ').trim() : 'Bilinmeyen';
        
        // Yuvarlanmış saatleri hesapla
        const yuvarlanmisGiris = roundToShiftTime(item.gercek ? item.gercek.gir : '08:30', vardiyaKod);
        const yuvarlanmisCikis = roundToShiftTime(item.gercek ? item.gercek.cik : '16:30', vardiyaKod);
        
        if (!gruplar[tarih]) {
            gruplar[tarih] = {
                toplamPersonel: 0,
                vardiyalar: {}
            };
        }
        
        // FM durumunu hesapla
        const fmDakika = calculateFMWithRoundedTime(yuvarlanmisGiris, yuvarlanmisCikis, vardiyaKod);
        const fmSaat = Math.floor(fmDakika / 60);
        const fmDakikaKalan = fmDakika % 60;
        const fmStr = fmDakika > 0 ? `${fmSaat}:${fmDakikaKalan.toString().padStart(2, '0')}` : '0:00';
        
        // Grup anahtarını sadece çıkış saatine göre oluştur (FM durumuna göre ayrı gruplar)
        const vardiyaKey = `${vardiyaKod}-${yuvarlanmisCikis}-${fmStr}`;
        
        if (!gruplar[tarih].vardiyalar[vardiyaKey]) {
            gruplar[tarih].vardiyalar[vardiyaKey] = {
                vardiya: vardiyaKod,
                giris: yuvarlanmisGiris,
                cikis: yuvarlanmisCikis,
                fm: fmStr,
                fmDakika: fmDakika,
                personeller: []
            };
        }
        
        // Personel zaten eklenmiş mi kontrol et
        if (!gruplar[tarih].vardiyalar[vardiyaKey].personeller.includes(personelAdi)) {
            gruplar[tarih].vardiyalar[vardiyaKey].personeller.push(personelAdi);
            gruplar[tarih].toplamPersonel++;
        }
    });
    
    return gruplar;
}

// Vardiya kombinasyonlarını göster
function displayVardiyaKombinasyonlari(kombinasyonlar) {
    const kombinasyonList = document.getElementById('kombinasyonList');
    kombinasyonList.innerHTML = '';
    
    if (kombinasyonlar.length === 0) {
        kombinasyonList.innerHTML = `
            <div style="text-align: center; padding: 40px; color: #7f8c8d;">
                <i data-lucide="info" style="font-size: 2rem; margin-bottom: 10px; display: block;"></i>
                Hiç vardiya kombinasyonu bulunamadı
            </div>
        `;
        return;
    }
    
    kombinasyonlar.forEach(kombinasyon => {
        const itemElement = document.createElement('div');
        itemElement.className = 'kombinasyon-item';
        
        // Vardiya detaylarını analiz et
        const vardiyaDetaylari = analizVardiyaDetaylari(kombinasyon);
        
        itemElement.innerHTML = `
            <div class="kombinasyon-header">
                <div class="kombinasyon-name">${kombinasyon.kombinasyon}</div>
                <div class="kombinasyon-count">${kombinasyon.personelSayisi} kişi</div>
            </div>
            <div class="kombinasyon-detaylar">
                ${vardiyaDetaylari}
            </div>
            <div class="kombinasyon-personeller">
                <strong>Personeller:</strong>
                <ul class="personel-list">
                    ${kombinasyon.personeller.map(personel => `<li>${personel}</li>`).join('')}
                </ul>
            </div>
        `;
        
        kombinasyonList.appendChild(itemElement);
    });
}

// Vardiya detaylarını analiz et
function analizVardiyaDetaylari(kombinasyon) {
    if (!currentData || currentData.length === 0) {
        return '<div class="detay-bilgi">Detay bilgisi mevcut değil</div>';
    }
    
    // Bu kombinasyondaki personellerin verilerini al (temizlenmiş isimlerle)
    const personelVerileri = currentData.filter(item => {
        const personelAdi = item.personel ? item.personel.split(' ').filter(part => !part.match(/^\d{1,2}:\d{2}$/)).join(' ').trim() : 'Bilinmeyen';
        return kombinasyon.personeller.includes(personelAdi);
    });
    
    // Vardiya bazında grupla
    const vardiyaGruplari = {};
    personelVerileri.forEach(item => {
        if (item.vardiya && item.vardiya.kod) {
            const vardiya = item.vardiya.kod;
            if (!vardiyaGruplari[vardiya]) {
                vardiyaGruplari[vardiya] = {
                    personeller: new Set(),
                    cikisSaatleri: new Set(),
                    toplamCalisma: 0,
                    kayitSayisi: 0
                };
            }
            vardiyaGruplari[vardiya].personeller.add(item.personel);
            if (item.gercek && item.gercek.cik) {
                // Yuvarlanmış çıkış saatini ekle
                const yuvarlanmisCikis = roundToShiftTime(item.gercek.cik, vardiya);
                vardiyaGruplari[vardiya].cikisSaatleri.add(yuvarlanmisCikis);
            }
            vardiyaGruplari[vardiya].toplamCalisma += item.calisma_dk || 0;
            vardiyaGruplari[vardiya].kayitSayisi++;
        }
    });
    
    // Detay HTML'ini oluştur
    let detayHTML = '';
    Object.entries(vardiyaGruplari).forEach(([vardiya, data]) => {
        const ortalamaCalisma = data.kayitSayisi > 0 ? Math.round(data.toplamCalisma / data.kayitSayisi) : 0;
        const calismaSaat = Math.floor(ortalamaCalisma / 60);
        const calismaDakika = ortalamaCalisma % 60;
        const calismaStr = `${calismaSaat}:${calismaDakika.toString().padStart(2, '0')}`;
        
        const cikisSaatleri = Array.from(data.cikisSaatleri).sort();
        const cikisStr = cikisSaatleri.length > 0 ? cikisSaatleri.join(', ') : 'Belirsiz';
        
        detayHTML += `
            <div class="vardiya-detay">
                <div class="vardiya-detay-header">
                    <span class="vardiya-badge vardiya-${vardiya.toLowerCase()}">${vardiya}</span>
                    <span class="personel-sayisi">${data.personeller.size} kişi</span>
                </div>
                <div class="vardiya-detay-bilgiler">
                    <div class="detay-item">
                        <span class="detay-label">Çıkış Saatleri:</span>
                        <span class="detay-value">${cikisStr}</span>
                    </div>
                    <div class="detay-item">
                        <span class="detay-label">Ortalama Çalışma:</span>
                        <span class="detay-value">${calismaStr}</span>
                    </div>
                </div>
            </div>
        `;
    });
    
    return detayHTML || '<div class="detay-bilgi">Detay bilgisi mevcut değil</div>';
}


// Vardiya detayları tablosunu göster
function displayVardiyaDetaylari(vardiyaDetaylari) {
    const tableBody = document.getElementById('vardiyaTableBody');
    tableBody.innerHTML = '';
    
    if (vardiyaDetaylari.length === 0) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="8" style="text-align: center; padding: 40px; color: #7f8c8d;">
                    <i data-lucide="info" style="font-size: 2rem; margin-bottom: 10px; display: block;"></i>
                    Hiç vardiya detayı bulunamadı
                </td>
            </tr>
        `;
        return;
    }
    
    // İlk 100 kaydı göster (performans için)
    const displayData = vardiyaDetaylari.slice(0, 100);
    
    displayData.forEach(detay => {
        const row = document.createElement('tr');
        
        // Çalışma süresini saat:dakika formatına çevir
        const calismaSaat = Math.floor(detay.calisma_dk / 60);
        const calismaDakika = detay.calisma_dk % 60;
        const calismaStr = `${calismaSaat}:${calismaDakika.toString().padStart(2, '0')}`;
        
        // Fazla mesai süresini saat:dakika formatına çevir
        const fmSaat = Math.floor(detay.fm_dk / 60);
        const fmDakika = detay.fm_dk % 60;
        const fmStr = `${fmSaat}:${fmDakika.toString().padStart(2, '0')}`;
        
        row.innerHTML = `
            <td>${detay.tarih}</td>
            <td><span class="vardiya-badge vardiya-${detay.vardiya.toLowerCase()}">${detay.vardiya}</span></td>
            <td>${detay.personel}</td>
            <td>${detay.giris || '-'}</td>
            <td>${detay.cikis || '-'}</td>
            <td>${calismaStr}</td>
            <td>${fmStr}</td>
            <td><span class="durum-badge durum-${detay.durum.toLowerCase()}">${detay.durum}</span></td>
        `;
        
        tableBody.appendChild(row);
    });
    
    // Eğer daha fazla kayıt varsa bilgi göster
    if (vardiyaDetaylari.length > 100) {
        const infoRow = document.createElement('tr');
        infoRow.innerHTML = `
            <td colspan="8" style="text-align: center; padding: 20px; background: #f8f9fa; color: #6c757d; font-style: italic;">
                ${vardiyaDetaylari.length - 100} kayıt daha var. Performans için ilk 100 kayıt gösteriliyor.
            </td>
        `;
        tableBody.appendChild(infoRow);
    }
}

// Vardiya analizi hata mesajı
function showVardiyaAnaliziError(message) {
    const vardiyaBars = document.getElementById('vardiyaBars');
    const kombinasyonList = document.getElementById('kombinasyonList');
    
    vardiyaBars.innerHTML = `
        <div style="text-align: center; padding: 40px; color: #e74c3c;">
            <i data-lucide="alert-circle" style="font-size: 2rem; margin-bottom: 10px; display: block;"></i>
            ${message}
        </div>
    `;
    
    kombinasyonList.innerHTML = `
        <div style="text-align: center; padding: 40px; color: #e74c3c;">
            <i data-lucide="alert-circle" style="font-size: 2rem; margin-bottom: 10px; display: block;"></i>
            ${message}
        </div>
    `;
}

// IPC log dinleyicisi
ipcRenderer.on('log-message', (event, logData) => {
    if (logSection && logSection.style.display !== 'none') {
        addLogEntry(logData.type || 'info', logData.message);
    }
});

// Sayfa yüklendiğinde
document.addEventListener('DOMContentLoaded', async () => {
    // Lucide iconları başlat
    lucide.createIcons();
    
    // Önce UI ayarlarını yükle (tema için)
    await loadUISettings();
    
    // Veri geçmişini yükle
    loadDataHistory();
    
    // Ayarları yükle
    loadSettings();
    
    // Başlangıç durumunu ayarla
    hideAllSections();
    
    // Ana ayarlar formu
    const mainSettingsForm = document.getElementById('mainSettingsForm');
    if (mainSettingsForm) {
        mainSettingsForm.addEventListener('submit', (e) => {
            e.preventDefault();
            saveSettings();
        });
    }
    
    // Yuvarlanma kuralları formu
    const yuvarlanmaForm = document.getElementById('yuvarlanmaForm');
    if (yuvarlanmaForm) {
        yuvarlanmaForm.addEventListener('submit', (e) => {
            e.preventDefault();
            saveYuvarlanmaKurallari();
        });
    }
    
    
    const enterDataMainBtn = document.getElementById('enterDataMainBtn');
    if (enterDataMainBtn) {
        enterDataMainBtn.addEventListener('click', enterExcelDataToPinhuman);
    }
    
    
    // Şifre göster/gizle toggle
    const togglePasswordBtn = document.getElementById('togglePassword');
    if (togglePasswordBtn) {
        togglePasswordBtn.addEventListener('click', togglePasswordVisibility);
    }
    
    
    // Drag & drop desteği (gelecekte eklenebilir)
    fileInputArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        fileInputArea.style.borderColor = '#667eea';
        fileInputArea.style.background = '#f0f4ff';
    });
    
    fileInputArea.addEventListener('dragleave', (e) => {
        e.preventDefault();
        fileInputArea.style.borderColor = '#ddd';
        fileInputArea.style.background = '#fafafa';
    });
    
    fileInputArea.addEventListener('drop', (e) => {
        e.preventDefault();
        fileInputArea.style.borderColor = '#ddd';
        fileInputArea.style.background = '#fafafa';
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const file = files[0];
            if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
                selectedFilePath = file.path;
                showSelectedFile(file.path);
                showProcessSection();
            } else {
                showError('Lütfen Excel dosyası (.xlsx veya .xls) seçin.');
            }
        }
    });
});

// Klavye kısayolları
document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key === 'o') {
        e.preventDefault();
        selectFile();
    }
    
    if (e.ctrlKey && e.key === 'Enter' && processSection.style.display !== 'none') {
        e.preventDefault();
        processFile();
    }
    
    if (e.key === 'Escape') {
        resetApp();
    }
});

// Dark Mode Toggle
document.addEventListener('DOMContentLoaded', () => {
    const darkModeToggle = document.getElementById('darkModeToggle');
    
    // Toggle event listener
    darkModeToggle.addEventListener('change', (e) => {
        const theme = e.target.checked ? 'dark' : 'light';
        document.documentElement.setAttribute('data-theme', theme);
        
        // UI ayarlarını kaydet (localStorage yerine config dosyasına)
        saveUISettings();
    });
});
