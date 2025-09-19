// Firebase yapılandırması
// Import the functions you need from the SDKs you need
import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getAnalytics } from "firebase/analytics";

console.log('Firebase config yükleniyor...');

// Your web app's Firebase configuration
const firebaseConfig = {
  apiKey: "AIzaSyAYeEyGc3prE8XJQyfdtsJYBKtm-Skowco",
  authDomain: "polenpdks.firebaseapp.com",
  projectId: "polenpdks",
  storageBucket: "polenpdks.firebasestorage.app",
  messagingSenderId: "456234462898",
  appId: "1:456234462898:web:4646f358c88f127c0823ab",
  measurementId: "G-SFEWRJ8634"
};

console.log('Firebase config:', firebaseConfig);

// Initialize Firebase
let app, auth, analytics;

try {
    app = initializeApp(firebaseConfig);
    console.log('Firebase app başlatıldı:', app);
    
    auth = getAuth(app);
    console.log('Firebase auth başlatıldı:', auth);
    
    analytics = getAnalytics(app);
    console.log('Firebase analytics başlatıldı:', analytics);
    
    console.log('Firebase başarıyla yapılandırıldı');
} catch (error) {
    console.error('Firebase yapılandırma hatası:', error);
    throw error;
}

// Export for use in other files
export { app, auth, analytics, firebaseConfig };