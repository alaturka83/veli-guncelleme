// ===============================================================
//  VELİ BİLGİ GÜNCELLEME — AYARLAR
//  Firebase konsolundan aldığın bilgileri aşağıya yapıştır.
//  (Firebase Console → Proje ayarları → Genel → Web uygulaması → SDK yapılandırması)
// ===============================================================

export const firebaseConfig = {
  apiKey: "AIzaSyBa-QT0FngwyCBLct7bOzfPV-WsLzYibQ4",
  authDomain: "veli-guncelleme.firebaseapp.com",
  projectId: "veli-guncelleme",
  storageBucket: "veli-guncelleme.firebasestorage.app",
  messagingSenderId: "340310024557",
  appId: "1:340310024557:web:8c75c6c4399bb18c6f3a64"
};

// Ekranlarda ve WhatsApp mesajlarında görünecek okul adı
export const OKUL_ADI = "Yunus Emre İlkokulu";

// Şube hesaplarının e-posta uzantısı (gerçek e-posta gönderilmez, değiştirmene gerek yok).
// DİKKAT: firestore.rules dosyasındaki uzantıyla aynı olmalı.
export const SUBE_ALANI = "sube.example.com";
