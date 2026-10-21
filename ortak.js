// Ortak yardımcılar — öğretmen ve yönetici ekranları birlikte kullanır.
import { firebaseConfig, OKUL_ADI, SUBE_ALANI } from './config.js?v=6';
import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js';
import { getAuth } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js';
import { getFirestore } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';

export { OKUL_ADI, SUBE_ALANI, firebaseConfig };
/** Her rol/şube kendi oturum alanını kullanır; böylece aynı tarayıcıda yönetici paneli
 *  ve farklı şubeler birbirinin oturumunu kapatmaz. */
export function baglan(ad) {
  const app = ad ? initializeApp(firebaseConfig, ad) : initializeApp(firebaseConfig);
  return { app, auth: getAuth(app), db: getFirestore(app) };
}

// ---------- Şubeler ----------
const HARFLER_1 = ['A','B','C','D','E','F','G','H','İ','J'];
const HARFLER = ['A','B','C','D','E','F','G','H','İ','J','K','L','M'];
const HARFLER_ANA = ['A','B','C','D','E','F'];
export const TUM_SUBELER = [
  ...HARFLER_ANA.map(h => anahtar(0, h)),
  ...HARFLER_1.map(h => anahtar(1, h)),
  ...[2,3,4].flatMap(g => HARFLER.map(h => anahtar(g, h)))
];
function anahtar(g, h) { return g + (h === 'İ' ? 'i' : h.toLowerCase()); }

/** "1A", "1/A", "1-a", "1 İ" → "1a" / "1i"; "ANA A", "Anasınıfı-B" → "0a" / "0b". Tanınmazsa null. */
export function subeAnahtari(v) {
  if (v == null) return null;
  const s = String(v).trim().replace(/[\s\/\-\.]/g, '').toLocaleUpperCase('tr-TR');
  let g, h, m = s.match(/^(ANASINIFI|ANASINIF|ANA)(.)$/u);
  if (m) { g = 0; h = m[2]; }
  else { m = s.match(/^([0-4])(.)$/u); if (!m) return null; g = Number(m[1]); h = m[2]; }
  if (h === 'I') h = 'İ';
  if (!HARFLER.includes(h)) return null;
  return anahtar(g, h);
}
/** "1i" → "1/İ", "0a" → "Ana/A" */
export function subeAdi(sk) { return sk ? (sk[0] === '0' ? 'Ana' : sk[0]) + '/' + harf(sk) : ''; }
/** "1i" → "1İ", "0a" → "ANA A" (Excel biçimi) */
export function subeExcel(sk) { return sk ? (sk[0] === '0' ? 'ANA ' : sk[0]) + harf(sk) : ''; }
function harf(sk) { return sk[1] === 'i' ? 'İ' : sk[1].toUpperCase(); }
export function subeSira(a, b) { return TUM_SUBELER.indexOf(a) - TUM_SUBELER.indexOf(b); }
export const subeEposta = sk => `${sk}@${SUBE_ALANI}`;

// ---------- Metin / telefon ----------
export function baslikYap(s) {
  return String(s || '').trim().replace(/\s+/g, ' ').split(' ')
    .map(k => k ? k[0].toLocaleUpperCase('tr-TR') + k.slice(1).toLocaleLowerCase('tr-TR') : k)
    .join(' ');
}
export function telTemizle(v) {
  let t = String(v ?? '').replace(/\D/g, '');
  if (t.length === 12 && t.startsWith('90')) t = t.slice(2);
  if (t.length === 11 && t.startsWith('0')) t = t.slice(1);
  return t;
}
export const telGecerli = t => /^5\d{9}$/.test(t);
export function telGoster(t) {
  if (!t) return '';
  if (t.length !== 10) return t;
  return `0 (${t.slice(0,3)}) ${t.slice(3,6)} ${t.slice(6,8)} ${t.slice(8)}`;
}
export const noTemizle = v => String(v ?? '').replace(/\D/g, '');
export function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
export function aramaMetni(s) {
  return String(s || '').toLocaleLowerCase('tr-TR')
    .replace(/[çğıöşü]/g, c => ({ç:'c',ğ:'g',ı:'i',ö:'o',ş:'s',ü:'u'}[c]));
}

// ---------- Durumlar ----------
export const DURUM = {
  bekliyor:   { ad: 'Kontrol bekliyor', kisa: 'Bekliyor',   renk: 'gri' },
  dogru:      { ad: 'Doğrulandı',       kisa: 'Doğru',      renk: 'yesil' },
  duzeltildi: { ad: 'Düzeltildi',       kisa: 'Düzeltildi', renk: 'mavi' },
  ayrildi:    { ad: 'Ayrıldı',          kisa: 'Ayrıldı',    renk: 'kirmizi' },
  yeni:       { ad: 'Yeni kayıt',       kisa: 'Yeni',       renk: 'mor' },
};
const ALANLAR = ['no', 'ad', 'veli', 'tel'];
/** Düzenleme sonrası yeni durumu hesaplar. */
export function durumHesapla(ogr, yeni) {
  if (ogr.durum === 'yeni' || !ogr.orj) return 'yeni';
  if (ogr.orj.sk && (yeni.sk || ogr.sk) !== ogr.orj.sk) return 'duzeltildi';
  return ALANLAR.some(a => (yeni[a] || '') !== (ogr.orj[a] || '')) ? 'duzeltildi' : 'dogru';
}
/** Hangi alanlar değişti? */
export function degisenAlanlar(o) {
  if (!o.orj) return [];
  const l = ALANLAR.filter(a => (o[a] || '') !== (o.orj[a] || ''));
  if (o.orj.sk && o.sk !== o.orj.sk) l.unshift('sk');
  return l;
}

export function sifreUret() {
  const a = new Uint32Array(1);
  let s;
  do { crypto.getRandomValues(a); s = String(a[0] % 1000000).padStart(6, '0'); }
  while (/^(\d)\1+$/.test(s) || '0123456789'.includes(s) || '9876543210'.includes(s));
  return s;
}

export function hataMesaji(e) {
  const c = e?.code || '';
  if (c.includes('invalid-credential') || c.includes('wrong-password') || c.includes('invalid-login') || c.includes('user-not-found')) return 'Şifre hatalı.';
  if (c.includes('too-many-requests')) return 'Çok fazla deneme yapıldı. Birkaç dakika sonra tekrar deneyin.';
  if (c.includes('network')) return 'İnternet bağlantısı yok gibi görünüyor.';
  if (c.includes('permission-denied')) return 'Bu işlem için yetkiniz yok (şube kilitli olabilir).';
  return 'Bir hata oluştu: ' + (e?.message || c || e);
}

let tostZaman;
export function tost(mesaj, tur = '') {
  let el = document.getElementById('tost');
  if (!el) { el = document.createElement('div'); el.id = 'tost'; document.body.appendChild(el); }
  el.textContent = mesaj;
  el.className = 'goster ' + tur;
  clearTimeout(tostZaman);
  tostZaman = setTimeout(() => el.className = '', 2600);
}

export function tarihGoster(ts) {
  const d = ts?.toDate ? ts.toDate() : (ts instanceof Date ? ts : null);
  if (!d) return '';
  return d.toLocaleString('tr-TR', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' });
}
