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
const jsonFileLink = document.getElementById('jsonFileLink');
const csvFileLink = document.getElementById('csvFileLink');
const newProcessBtn = document.getElementById('newProcessBtn');
const openFolderBtn = document.getElementById('openFolderBtn');
const retryBtn = document.getElementById('retryBtn');

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
newProcessBtn.addEventListener('click', resetApp);
openFolderBtn.addEventListener('click', openOutputFolder);
retryBtn.addEventListener('click', () => {
    hideError();
    if (selectedFilePath) {
        processFile();
    } else {
        selectFile();
    }
});

// Sayfa geçiş sistemi
document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', () => {
        const pageId = item.dataset.page;
        switchPage(pageId);
    });
});

// Filtre event listeners (sadece personel sayfasında)
function setupFilterListeners() {
    const personelFilterEl = document.getElementById('personelFilter');
    if (personelFilterEl) {
        personelFilterEl.addEventListener('change', applyFilters);
    }
    
    ['vardiyaV1', 'vardiyaV2', 'vardiyaV3', 'durumOk', 'durumEksik', 'durumGecersiz'].forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.addEventListener('change', applyFilters);
        }
    });
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
    
    // Geçmiş sayfasına geçildiğinde geçmişi güncelle
    if (pageId === 'gecmis') {
        updateHistoryList();
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
        color: white;
        font-family: 'Montserrat', sans-serif;
    `;
    
    overlay.innerHTML = `
        <div style="text-align: center;">
            <div style="width: 60px; height: 60px; border: 4px solid #f3f3f3; border-top: 4px solid #3498db; border-radius: 50%; animation: spin 1s linear infinite; margin: 0 auto 20px;"></div>
            <h3 style="margin: 0; font-size: 1.5rem; font-weight: 600;">İşlem Yapılıyor...</h3>
            <p style="margin: 10px 0 0; opacity: 0.8;">Lütfen bekleyin</p>
        </div>
        <style>
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
    jsonFileLink.href = 'file://' + data.json_path;
    csvFileLink.href = 'file://' + data.csv_path;
    
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
    const personelFilterEl = document.getElementById('personelFilter');
    if (personelFilterEl) {
        personelFilterEl.innerHTML = '<option value="">Tüm Personeller</option>';
    }
    
    if (dataTableBody) {
        dataTableBody.innerHTML = `
            <tr>
                <td colspan="8" style="text-align: center; padding: 40px; color: #7f8c8d;">
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
            
            // Personel listesini güncelle
            updatePersonelFilter();
            
            // Veri listesini güncelle
            updateDataList();
        })
        .catch(error => {
            console.error('Veri yüklenirken hata:', error);
        });
}

// Personel filtresini güncelle
function updatePersonelFilter() {
    const personeller = [...new Set(currentData.map(item => item.personel))].sort();
    
    personelFilter.innerHTML = '<option value="">Tüm Personeller</option>';
    personeller.forEach(personel => {
        const option = document.createElement('option');
        option.value = personel;
        option.textContent = personel;
        personelFilter.appendChild(option);
    });
}

// Veri listesini güncelle
function updateDataList() {
    if (filteredData.length === 0) {
        dataTableBody.innerHTML = `
            <tr>
                <td colspan="8" style="text-align: center; padding: 20px; color: #7f8c8d;">
                    Filtre kriterlerine uygun veri bulunamadı
                </td>
            </tr>
        `;
        return;
    }
    
    dataTableBody.innerHTML = filteredData.map(item => `
        <tr onclick="showItemDetails('${item.personel}', '${item.tarih}')">
            <td>${item.personel}</td>
            <td>${formatDate(item.tarih)}</td>
            <td>${item.vardiya ? item.vardiya.kod : '-'}</td>
            <td>${item.gercek ? item.gercek.gir : '-'}</td>
            <td>${item.gercek ? item.gercek.cik : '-'}</td>
            <td>${formatDuration(item.calisma_dk)}</td>
            <td>${item.fm_hhmm}</td>
            <td><span class="data-status ${item.durum}">${item.durum}</span></td>
        </tr>
    `).join('');
    
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
    
    tbody.innerHTML = allRecords.map(item => `
        <tr onclick="showItemDetails('${item.personel}', '${item.tarih}')">
            <td><span class="vardiya-badge ${item.vardiyaKod.toLowerCase()}">${item.vardiyaKod}</span></td>
            <td>${item.personel}</td>
            <td>${item.gercek ? item.gercek.gir : '-'}</td>
            <td>${item.gercek ? item.gercek.cik : '-'}</td>
            <td>${formatDuration(item.calisma_dk)}</td>
            <td>${item.fm_hhmm}</td>
            <td><span class="data-status ${item.durum}">${item.durum}</span></td>
        </tr>
    `).join('');
    
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
    const personelFilterEl = document.getElementById('personelFilter');
    if (!personelFilterEl) return;
    
    const selectedPersonel = personelFilterEl.value;
    const selectedVardiyalar = [];
    const selectedDurumlar = [];
    
    // Seçili vardiyaları al
    const vardiyaV1 = document.getElementById('vardiyaV1');
    const vardiyaV2 = document.getElementById('vardiyaV2');
    const vardiyaV3 = document.getElementById('vardiyaV3');
    
    if (vardiyaV1 && vardiyaV1.checked) selectedVardiyalar.push('V1');
    if (vardiyaV2 && vardiyaV2.checked) selectedVardiyalar.push('V2');
    if (vardiyaV3 && vardiyaV3.checked) selectedVardiyalar.push('V3');
    
    // Seçili durumları al
    const durumOk = document.getElementById('durumOk');
    const durumEksik = document.getElementById('durumEksik');
    const durumGecersiz = document.getElementById('durumGecersiz');
    
    if (durumOk && durumOk.checked) selectedDurumlar.push('ok');
    if (durumEksik && durumEksik.checked) selectedDurumlar.push('eksik');
    if (durumGecersiz && durumGecersiz.checked) selectedDurumlar.push('geçersiz');
    
    // Filtrele
    filteredData = currentData.filter(item => {
        const personelMatch = !selectedPersonel || item.personel === selectedPersonel;
        const vardiyaMatch = !item.vardiya || selectedVardiyalar.includes(item.vardiya.kod);
        const durumMatch = selectedDurumlar.includes(item.durum);
        
        return personelMatch && vardiyaMatch && durumMatch;
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

// Sayfa yüklendiğinde
document.addEventListener('DOMContentLoaded', () => {
    // Lucide iconları başlat
    lucide.createIcons();
    
    // Veri geçmişini yükle
    loadDataHistory();
    
    // Başlangıç durumunu ayarla
    hideAllSections();
    
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
