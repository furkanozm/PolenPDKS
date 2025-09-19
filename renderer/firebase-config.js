// Firebase yapılandırması
// Not: Bu dosya sadece config bilgilerini içerir, Firebase'i başlatmaz
// Firebase başlatma işlemi login.js'de yapılır

// Firebase yapılandırma bilgileri
const firebaseConfig = {
  apiKey: "AIzaSyAYeEyGc3prE8XJQyfdtsJYBKtm-Skowco",
  authDomain: "polenpdks.firebaseapp.com",
  projectId: "polenpdks",
  storageBucket: "polenpdks.firebasestorage.app",
  messagingSenderId: "456234462898",
  appId: "1:456234462898:web:4646f358c88f127c0823ab",
  measurementId: "G-SFEWRJ8634"
};

// Config'i export et
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        firebaseConfig
    };
} else {
    // Browser ortamı için
    window.firebaseConfig = firebaseConfig;
}
