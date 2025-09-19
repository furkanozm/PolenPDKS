// Firebase yapılandırması
const { initializeApp } = require("firebase/app");
const { getAuth, GoogleAuthProvider } = require("firebase/auth");
const { getAnalytics } = require("firebase/analytics");

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

// Firebase'i başlat
const app = initializeApp(firebaseConfig);

// Authentication servisini al
const auth = getAuth(app);

// Google Auth Provider'ı oluştur
const googleProvider = new GoogleAuthProvider();

// Analytics'i başlat (opsiyonel)
const analytics = getAnalytics(app);

module.exports = {
    auth,
    googleProvider,
    analytics,
    app
};
