// Firebase Login Sistemi - Basitleştirilmiş Versiyon
console.log('Login.js yükleniyor...');

// Electron IPC modülü
const { ipcRenderer } = require('electron');

// DOM elementleri - DOM yüklendikten sonra tanımlanacak
let emailInput, passwordInput, loginBtn, loadingOverlay, errorModal, errorMessage, errorModalClose, rememberMeCheckbox, themeToggle;

// Sayfa yüklendiğinde kullanıcı durumunu kontrol et
document.addEventListener('DOMContentLoaded', () => {
    console.log('DOM yüklendi, elementler alınıyor...');
    
    // DOM elementlerini al
    emailInput = document.getElementById('email');
    passwordInput = document.getElementById('password');
    loginBtn = document.getElementById('loginBtn');
    loadingOverlay = document.getElementById('loadingOverlay');
    errorModal = document.getElementById('errorModal');
    errorMessage = document.getElementById('errorMessage');
    errorModalClose = document.getElementById('errorModalClose');
    rememberMeCheckbox = document.getElementById('rememberMe');
    themeToggle = document.getElementById('themeToggle');

    // Elementlerin yüklenip yüklenmediğini kontrol et
    if (!emailInput || !passwordInput || !loginBtn) {
        console.error('Gerekli DOM elementleri bulunamadı!');
        console.error('emailInput:', emailInput);
        console.error('passwordInput:', passwordInput);
        console.error('loginBtn:', loginBtn);
        showError('Sayfa yüklenirken bir hata oluştu. Lütfen sayfayı yenileyin.');
        return;
    }

    console.log('Tüm elementler başarıyla yüklendi');

    // Event listener'ları ekle
    setupEventListeners();

    // Firebase auth state listener - DEVRE DIŞI (otomatik giriş yapmasın)
    // setTimeout(() => {
    //     if (window.firebaseAuth && window.onAuthStateChanged) {
    //         console.log('Firebase auth state listener başlatılıyor...');
    //         window.onAuthStateChanged(window.firebaseAuth, (user) => {
    //             if (user) {
    //                 // Kullanıcı zaten giriş yapmış, ana sayfaya yönlendir
    //                 console.log('Kullanıcı zaten giriş yapmış:', user.email);
    //                 
    //                 // Çıkış yapıldıktan sonra otomatik giriş yapmasın
    //                 const lastLogout = localStorage.getItem('lastLogout');
    //                 const now = Date.now();
    //                 
    //                 if (lastLogout && (now - parseInt(lastLogout)) < 5000) {
    //                     console.log('Yakın zamanda çıkış yapıldı, otomatik giriş yapılmıyor');
    //                     return;
    //                 }
    //                 
    //                 ipcRenderer.invoke('navigate-to-main');
    //             } else {
    //                 // Kullanıcı giriş yapmamış, login sayfasını göster
    //                 console.log('Kullanıcı giriş yapmamış');
    //             }
    //         });
    //     } else {
    //         console.error('Firebase auth bulunamadı!');
    //     }
    // }, 1000);
    
    console.log('Firebase auth state listener devre dışı bırakıldı - otomatik giriş yapılmayacak');

    // Lucide ikonlarını başlat
    if (typeof lucide !== 'undefined') {
        lucide.createIcons();
    }

    // Tema yönetimi
    initializeTheme();
    
    // Kayıtlı e-postayı yükle
    loadRememberedEmail();
});

// Event listener'ları kurma fonksiyonu
function setupEventListeners() {
    console.log('Event listener\'lar kuruluyor...');
    
    // ÖNCE tema butonunu tamamen izole et
    setupThemeButton();
    
    // SONRA diğer event listener'ları ekle
    setupOtherEventListeners();
}

// Tema butonunu tamamen yeniden oluşturma fonksiyonu
function setupThemeButton() {
    console.log('Tema butonu tamamen yeniden oluşturuluyor...');
    
    const container = document.getElementById('themeToggleContainer');
    if (!container) {
        console.error('Tema butonu container bulunamadı!');
        return;
    }
    
    // Container'ı temizle
    container.innerHTML = '';
    
    // Tema butonunu sıfırdan oluştur
    const themeToggleDiv = document.createElement('div');
    themeToggleDiv.className = 'theme-toggle';
    
    const themeButton = document.createElement('button');
    themeButton.type = 'button';
    themeButton.className = 'theme-btn';
    themeButton.title = 'Tema değiştir';
    themeButton.setAttribute('tabindex', '-1');
    themeButton.setAttribute('data-theme-button', 'true'); // Özel attribute
    
    // İkonları oluştur
    const sunIcon = document.createElement('i');
    sunIcon.setAttribute('data-lucide', 'sun');
    sunIcon.className = 'theme-icon light-icon';
    
    const moonIcon = document.createElement('i');
    moonIcon.setAttribute('data-lucide', 'moon');
    moonIcon.className = 'theme-icon dark-icon';
    
    themeButton.appendChild(sunIcon);
    themeButton.appendChild(moonIcon);
    themeToggleDiv.appendChild(themeButton);
    container.appendChild(themeToggleDiv);
    
    // Tema butonuna sadece click event'i ekle - hiçbir şey başka event ekleme
    themeButton.onclick = function(e) {
        console.log('=== TEMA BUTONU TIKLANDI (YENİ) ===');
        console.log('Event:', e);
        console.log('Target:', e.target);
        
        // Tüm event'leri engelle
        e.preventDefault();
        e.stopPropagation();
        e.stopImmediatePropagation();
        
        console.log('Tema değiştirme başlatılıyor...');
        toggleTheme();
        console.log('Tema değiştirme tamamlandı');
        
        return false;
    };
    
    // Lucide ikonlarını yeniden yükle
    if (typeof lucide !== 'undefined') {
        lucide.createIcons();
    }
    
    console.log('Tema butonu tamamen yeniden oluşturuldu');
}

// Diğer event listener'ları kurma fonksiyonu
function setupOtherEventListeners() {
    console.log('Diğer event listener\'lar kuruluyor...');
    
    // Login butonu event listener
    loginBtn.addEventListener('click', async (e) => {
        console.log('=== GİRİŞ BUTONU TIKLANDI ===');
        console.log('Event:', e);
        console.log('Target:', e.target);
        console.log('Current target:', e.currentTarget);
        console.log('Event type:', e.type);
        
        const email = emailInput.value.trim();
        const password = passwordInput.value;

        console.log('E-posta:', email);
        console.log('Şifre uzunluğu:', password.length);

        // Form validasyonu
        if (!email || !password) {
            console.log('Form validasyon hatası: Boş alan');
            showError('Lütfen e-posta ve şifre alanlarını doldurun.');
            return;
        }

        if (!isValidEmail(email)) {
            console.log('Form validasyon hatası: Geçersiz e-posta');
            showError('Lütfen geçerli bir e-posta adresi girin.');
            return;
        }

        console.log('Form validasyonu başarılı, giriş yapılıyor...');
        showLoading(true);

        try {
            // Firebase ile giriş yap
            console.log('Firebase giriş işlemi başlatılıyor...');
            
            if (!window.firebaseAuth || !window.signInWithEmailAndPassword) {
                throw new Error('Firebase henüz yüklenmedi. Lütfen tekrar deneyin.');
            }
            
            const userCredential = await window.signInWithEmailAndPassword(window.firebaseAuth, email, password);
            const user = userCredential.user;
            
            console.log('Firebase giriş başarılı:', user.email);
            
            // Beni hatırla seçiliyse e-postayı kaydet
            if (rememberMeCheckbox && rememberMeCheckbox.checked) {
                await ipcRenderer.invoke('save-remembered-email', user.email);
            }
            
            // Ana uygulamaya yönlendir
            ipcRenderer.invoke('navigate-to-main');
            
        } catch (error) {
            console.error('Firebase giriş hatası:', error);
            console.error('Hata kodu:', error.code);
            console.error('Hata mesajı:', error.message);
            showError(getErrorMessage(error.code));
        } finally {
            showLoading(false);
        }
    });

    // Enter tuşu ile giriş
    document.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            // Tema butonuna odaklanmışsa giriş yapma
            if (document.activeElement && document.activeElement.getAttribute('data-theme-button') === 'true') {
                console.log('Tema butonuna odaklanmış, giriş yapılmıyor');
                e.preventDefault();
                e.stopPropagation();
                return false;
            }
            
            // E-posta veya şifre input'una odaklanmışsa giriş yap
            if (document.activeElement === emailInput || document.activeElement === passwordInput) {
                console.log('Enter tuşu basıldı - giriş yapılıyor');
                e.preventDefault();
                loginBtn.click();
                return false;
            }
            
            // Diğer durumlarda giriş yapma
            console.log('Enter tuşu basıldı ama giriş yapılmıyor');
            e.preventDefault();
            return false;
        }
    });

    // Hata modalını kapat
    if (errorModalClose) {
        errorModalClose.addEventListener('click', () => {
            console.log('Hata modalı kapatılıyor');
            errorModal.style.display = 'none';
        });
    }

    // Modal dışına tıklayınca kapat
    if (errorModal) {
        errorModal.addEventListener('click', (e) => {
            if (e.target === errorModal) {
                console.log('Modal dışına tıklandı, kapatılıyor');
                errorModal.style.display = 'none';
            }
        });
    }

    // Tema butonu artık setupThemeButton() fonksiyonunda izole edildi

    // Remember me checkbox değişikliği
    if (rememberMeCheckbox) {
        rememberMeCheckbox.addEventListener('change', async (e) => {
            console.log('Remember me checkbox değişti:', e.target.checked);
            
            if (!e.target.checked) {
                // Checkbox kaldırıldıysa kayıtlı e-postayı sil
                try {
                    await ipcRenderer.invoke('save-remembered-email', '');
                    console.log('Kayıtlı e-posta silindi');
                } catch (error) {
                    console.error('E-posta silme hatası:', error);
                }
            }
        });
    }

    console.log('Tüm event listener\'lar başarıyla kuruldu');
}

// Yardımcı fonksiyonlar
function showLoading(show) {
    console.log('Loading gösteriliyor:', show);
    if (loadingOverlay) {
        loadingOverlay.style.display = show ? 'flex' : 'none';
    } else {
        console.error('Loading overlay bulunamadı!');
    }
}

function showError(message) {
    console.log('Hata gösteriliyor:', message);
    if (errorMessage && errorModal) {
        errorMessage.textContent = message;
        errorModal.style.display = 'flex';
        console.log('Hata modalı gösterildi');
    } else {
        console.error('Hata modalı veya mesaj elementi bulunamadı!');
        // Fallback: alert kullan
        alert('Hata: ' + message);
    }
}

function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

function getErrorMessage(errorCode) {
    const errorMessages = {
        'auth/user-not-found': 'Bu e-posta adresi ile kayıtlı kullanıcı bulunamadı.',
        'auth/wrong-password': 'Hatalı şifre girdiniz.',
        'auth/invalid-email': 'Geçersiz e-posta adresi.',
        'auth/user-disabled': 'Bu hesap devre dışı bırakılmış.',
        'auth/too-many-requests': 'Çok fazla başarısız giriş denemesi. Lütfen daha sonra tekrar deneyin.',
        'auth/network-request-failed': 'Ağ bağlantısı hatası. İnternet bağlantınızı kontrol edin.',
        'auth/invalid-credential': 'Geçersiz kimlik bilgileri.',
        'auth/requires-recent-login': 'Bu işlem için tekrar giriş yapmanız gerekiyor.',
        'auth/operation-not-allowed': 'E-posta/şifre girişi etkinleştirilmemiş.',
        'auth/weak-password': 'Şifre çok zayıf.',
        'auth/email-already-in-use': 'Bu e-posta adresi zaten kullanımda.'
    };

    return errorMessages[errorCode] || 'Bilinmeyen bir hata oluştu. Lütfen tekrar deneyin.';
}

// Tema yönetimi fonksiyonları
async function initializeTheme() {
    try {
        // Önce config'den tema tercihini al
        const configTheme = await ipcRenderer.invoke('get-theme-preference');
        console.log('Config\'den tema tercihi alındı:', configTheme);
        
        // Config'de tema varsa onu kullan, yoksa localStorage'dan al
        const savedTheme = configTheme || localStorage.getItem('theme') || 'light';
        setTheme(savedTheme);
        
        // Config'de tema yoksa localStorage'daki temayı config'e kaydet
        if (!configTheme && localStorage.getItem('theme')) {
            await ipcRenderer.invoke('save-theme-preference', localStorage.getItem('theme'));
        }
    } catch (error) {
        console.error('Tema yükleme hatası:', error);
        // Hata durumunda localStorage'dan yükle
        const savedTheme = localStorage.getItem('theme') || 'light';
        setTheme(savedTheme);
    }
}

async function toggleTheme() {
    const currentTheme = document.body.classList.contains('dark-theme') ? 'dark' : 'light';
    const newTheme = currentTheme === 'light' ? 'dark' : 'light';
    
    setTheme(newTheme);
    
    // Hem localStorage'a hem config'e kaydet
    localStorage.setItem('theme', newTheme);
    
    try {
        await ipcRenderer.invoke('save-theme-preference', newTheme);
        console.log('Tema tercihi config\'e kaydedildi:', newTheme);
    } catch (error) {
        console.error('Tema config\'e kaydetme hatası:', error);
    }
}

function setTheme(theme) {
    if (theme === 'dark') {
        document.body.classList.add('dark-theme');
        document.body.classList.remove('light-theme');
    } else {
        document.body.classList.add('light-theme');
        document.body.classList.remove('dark-theme');
    }
    
    // İkonları yeniden yükle
    if (typeof lucide !== 'undefined') {
        lucide.createIcons();
    }
    
    console.log('Tema uygulandı:', theme);
}

// Kayıtlı e-postayı yükleme fonksiyonu
async function loadRememberedEmail() {
    try {
        const rememberedEmail = await ipcRenderer.invoke('get-remembered-email');
        console.log('Kayıtlı e-posta alındı:', rememberedEmail);
        
        if (rememberedEmail && emailInput) {
            emailInput.value = rememberedEmail;
            // Remember me checkbox'ını işaretle
            if (rememberMeCheckbox) {
                rememberMeCheckbox.checked = true;
            }
            console.log('E-posta input\'a yüklendi:', rememberedEmail);
        }
    } catch (error) {
        console.error('Kayıtlı e-posta yükleme hatası:', error);
    }
}

console.log('Login.js yüklendi');