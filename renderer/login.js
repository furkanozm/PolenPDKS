// Electron IPC modülü
const { ipcRenderer } = require('electron');

// DOM elementleri
const emailInput = document.getElementById('email');
const passwordInput = document.getElementById('password');
const loginBtn = document.getElementById('loginBtn');
const togglePasswordBtn = document.getElementById('togglePassword');
const eyeIcon = document.getElementById('eyeIcon');
const loadingOverlay = document.getElementById('loadingOverlay');
const errorModal = document.getElementById('errorModal');
const errorMessage = document.getElementById('errorMessage');
const errorModalClose = document.getElementById('errorModalClose');
const rememberMeCheckbox = document.getElementById('rememberMe');
const requestPasswordBtn = document.getElementById('requestPasswordBtn');
const userInfoModal = document.getElementById('userInfoModal');
const closeUserInfoModal = document.getElementById('closeUserInfoModal');
const cancelUserInfo = document.getElementById('cancelUserInfo');
const sendMailBtn = document.getElementById('sendMailBtn');
const userNameInput = document.getElementById('userName');
const userPasswordInput = document.getElementById('userPassword');

// Şifre görünürlüğünü değiştir
togglePasswordBtn.addEventListener('click', () => {
    const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
    passwordInput.setAttribute('type', type);
    
    // İkonu değiştir
    if (type === 'text') {
        eyeIcon.setAttribute('data-lucide', 'eye-off');
    } else {
        eyeIcon.setAttribute('data-lucide', 'eye');
    }
    lucide.createIcons();
});

// E-posta ve şifre ile giriş
loginBtn.addEventListener('click', async () => {
    const email = emailInput.value.trim();
    const password = passwordInput.value;

    if (!email || !password) {
        showError('Lütfen e-posta ve şifre alanlarını doldurun.');
        return;
    }

    if (!isValidEmail(email)) {
        showError('Lütfen geçerli bir e-posta adresi girin.');
        return;
    }

    showLoading(true);

    try {
        // Config'deki Pinhuman şifresini doğrula
        const decryptedConfig = await ipcRenderer.invoke('verify-config-password', password);
        
        if (decryptedConfig) {
            console.log('Pinhuman şifre doğrulaması başarılı');
            
            // Beni hatırla seçiliyse e-postayı config'e kaydet
            if (rememberMeCheckbox && rememberMeCheckbox.checked) {
                await ipcRenderer.invoke('save-remembered-email', email);
            }
            
            // Ana uygulamaya yönlendir
            ipcRenderer.invoke('navigate-to-main');
        }
        
    } catch (error) {
        console.error('Giriş hatası:', error);
        showError('Geçersiz şifre. Lütfen şifrenizi girin.');
    } finally {
        showLoading(false);
    }
});


// Enter tuşu ile giriş
document.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
        loginBtn.click();
    }
});

// Şifre iste butonu
requestPasswordBtn.addEventListener('click', () => {
    userInfoModal.style.display = 'flex';
    userNameInput.focus();
});

// Modal kapatma
closeUserInfoModal.addEventListener('click', () => {
    userInfoModal.style.display = 'none';
});

cancelUserInfo.addEventListener('click', () => {
    userInfoModal.style.display = 'none';
});

// Modal dışına tıklayınca kapat
userInfoModal.addEventListener('click', (e) => {
    if (e.target === userInfoModal) {
        userInfoModal.style.display = 'none';
    }
});

// Mail gönderme
sendMailBtn.addEventListener('click', () => {
    const userName = userNameInput.value.trim();
    const userPassword = userPasswordInput.value.trim();
    
    if (!userName || !userPassword) {
        showError('Lütfen kullanıcı adı ve şifre alanlarını doldurun.');
        return;
    }
    
    // Outlook Classic ile mail gönder
    const subject = 'PDKS Sistemi - Hesap Oluşturulması İsteği';
    const body = `Merhaba,

PDKS İşleme Sistemi için hesap oluşturulmasını istiyorum.

Kullanıcı Bilgileri:
- Kullanıcı Adı: ${userName}
- Şifre: ${userPassword}

Lütfen bu bilgilerle hesabımı oluşturun.

Teşekkürler.`;

    const mailtoLink = `mailto:furkan.ozmen@guleryuzgroup.com?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    
    // Outlook Classic'i aç
    ipcRenderer.invoke('open-outlook-mail', mailtoLink);
    
    userInfoModal.style.display = 'none';
    userNameInput.value = '';
    userPasswordInput.value = '';
});

// Hata modalını kapat
errorModalClose.addEventListener('click', () => {
    errorModal.style.display = 'none';
});

// Modal dışına tıklayınca kapat
errorModal.addEventListener('click', (e) => {
    if (e.target === errorModal) {
        errorModal.style.display = 'none';
    }
});


// Yardımcı fonksiyonlar
function showLoading(show) {
    loadingOverlay.style.display = show ? 'flex' : 'none';
}

function showError(message) {
    errorMessage.textContent = message;
    errorModal.style.display = 'flex';
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
        'auth/popup-closed-by-user': 'Giriş penceresi kapatıldı.',
        'auth/cancelled-popup-request': 'Giriş işlemi iptal edildi.',
        'auth/popup-blocked': 'Popup engellendi. Lütfen popup engelleyicisini kapatın.',
        'auth/account-exists-with-different-credential': 'Bu e-posta adresi farklı bir giriş yöntemi ile kayıtlı.',
        'auth/email-already-in-use': 'Bu e-posta adresi zaten kullanımda.',
        'auth/weak-password': 'Şifre çok zayıf. Lütfen daha güçlü bir şifre seçin.',
        'auth/operation-not-allowed': 'Bu giriş yöntemi etkinleştirilmemiş.',
        'auth/invalid-credential': 'Geçersiz kimlik bilgileri.',
        'auth/requires-recent-login': 'Bu işlem için tekrar giriş yapmanız gerekiyor.'
    };

    return errorMessages[errorCode] || 'Bilinmeyen bir hata oluştu. Lütfen tekrar deneyin.';
}

// Şifre hashleme fonksiyonu
async function hashPassword(password) {
    const encoder = new TextEncoder();
    const data = encoder.encode(password);
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

// Kayıtlı e-postayı yükle
async function loadSavedEmail() {
    try {
        const savedEmail = await ipcRenderer.invoke('get-remembered-email');
        
        if (savedEmail) {
            emailInput.value = savedEmail;
            if (rememberMeCheckbox) {
                rememberMeCheckbox.checked = true;
            }
            // Şifre alanına odaklan
            passwordInput.focus();
        }
    } catch (error) {
        console.log('Kayıtlı e-posta yüklenemedi:', error);
    }
}

// Sayfa yüklendiğinde Lucide ikonlarını başlat ve kayıtlı e-postayı yükle
document.addEventListener('DOMContentLoaded', async () => {
    // İkonları başlat
    lucide.createIcons();
    
    // Kayıtlı e-postayı yükle
    await loadSavedEmail();
});
