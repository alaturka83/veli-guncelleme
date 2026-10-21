// Biçimli Excel çıktıları — yönetici paneli ve öğretmen ekranı birlikte kullanır (aynı biçim).
import { OKUL_ADI, subeAdi, ozelMi, tarihGoster } from './ortak.js?v=19';

export const devamsizMi = o => o.durum !== 'ayrildi' && !!o.devamsiz;
export const devamsizMetni = o => devamsizMi(o) ? 'Sürekli devamsız' + (o.devamsizNot ? ' – ' + o.devamsizNot : '') : '';
export function bugun() { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; }
export const telSayi = t => t ? Number(t) : '';
// ---- Biçimli Excel (ExcelJS, ilk tıklamada yüklenir) ----
let excelJSYukleniyor = null;
export function excelJSYukle() {
  if (window.ExcelJS) return Promise.resolve();
  return excelJSYukleniyor ||= new Promise((ok, hata) => {
    const s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
    s.onload = ok; s.onerror = () => { excelJSYukleniyor = null; hata(new Error('Excel kütüphanesi yüklenemedi')); };
    document.head.appendChild(s);
  });
}
export const XR = { lacivert: 'FF1F3864', baslikAlt: 'FF2E4A7D', zebra: 'FFF4F6FA', cizgi: 'FFC9D2E0', soluk: 'FF5F6B7A', sari: 'FFB45309', beyaz: 'FFFFFFFF' };
export const XDURUM = {
  bekliyor: ['FFF1F3F5', 'FF5F6B7A', 'Bekliyor'], dogru: ['FFE6F4EA', 'FF1E7B34', 'Doğru'], duzeltildi: ['FFE8F0FE', 'FF1A56DB', 'Düzeltildi'],
  ayrildi: ['FFFDECEC', 'FFB42318', 'Ayrıldı'], yeni: ['FFF1EAFE', 'FF6D28D9', 'Yeni kayıt'],
  tamam: ['FFE6F4EA', 'FF1E7B34'], devam: ['FFE8F0FE', 'FF1A56DB'], baslamadi: ['FFF1F3F5', 'FF5F6B7A'],
};
export const XINCE = { style: 'thin', color: { argb: XR.cizgi } };
export const XKENAR = { top: XINCE, left: XINCE, bottom: XINCE, right: XINCE };
export const XTEL = '"0 ("000") "000" "00" "00';          // 5355733013 → 0 (535) 573 30 13
export const sayfaAdi = sk => subeAdi(sk).replace(/[\/\\?*\[\]:]/g, '-').slice(0, 31);
export const noSayi = n => /^\d+$/.test(n || '') ? Number(n) : (n || '');
export const dolgu = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });

/** Başlıklı, kenarlıklı, zebra satırlı standart tablo sayfası */
export function xSayfa(wb, ad, { baslik, alt, kolonlar, satirlar, durumlar, toplam, yatay, renk }) {
  const n = kolonlar.length;
  const ws = wb.addWorksheet(ad, {
    properties: { tabColor: { argb: renk || XR.lacivert } },
    views: [{ state: 'frozen', ySplit: 3, showGridLines: false }],
    pageSetup: { paperSize: 9, orientation: yatay ? 'landscape' : 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.55, header: 0.25, footer: 0.25 }, printTitlesRow: '3:3', horizontalCentered: true },
    headerFooter: { oddFooter: `&L&8${OKUL_ADI}&C&8${ad}&R&8Sayfa &P / &N` },
  });
  ws.columns = kolonlar.map(k => ({ width: k.w + 3 }));   // +3: başlıktaki filtre oku yazıyı kesmesin
  ws.mergeCells(1, 1, 1, n);
  Object.assign(ws.getCell(1, 1), { value: baslik, font: { name: 'Calibri', bold: true, size: 15, color: { argb: XR.lacivert } }, alignment: { vertical: 'middle' } });
  ws.getRow(1).height = 26;
  ws.mergeCells(2, 1, 2, n);
  Object.assign(ws.getCell(2, 1), { value: alt, font: { name: 'Calibri', size: 10, color: { argb: XR.soluk } }, alignment: { vertical: 'top' } });
  ws.getRow(2).height = 20;
  const hr = ws.getRow(3); hr.height = kolonlar.some(k => k.h.length > 11 && k.w < 12) ? 32 : 24;
  kolonlar.forEach((k, i) => Object.assign(hr.getCell(i + 1), {
    value: k.h, font: { name: 'Calibri', bold: true, size: 11, color: { argb: XR.beyaz } }, fill: dolgu(XR.lacivert), border: XKENAR,
    alignment: { vertical: 'middle', horizontal: k.orta ? 'center' : 'left', wrapText: true, indent: k.orta ? 0 : 1 },
  }));
  satirlar.forEach((s, ri) => {
    const r = ws.getRow(4 + ri); r.height = 19;
    s.forEach((v, i) => {
      const k = kolonlar[i], c = r.getCell(i + 1);
      const bos = v === '' || v == null;
      c.value = bos ? null : v;      // boş hücre boş kalır (yeniden içe aktarmada sorun çıkmasın), sarıyla işaretlenir
      c.font = { name: 'Calibri', size: 11, ...(k.kalin ? { bold: true } : {}), ...(k.vurgu && !bos ? { bold: true, color: { argb: 'FF1A56DB' } } : {}) };
      if (k.fmt && !bos) c.numFmt = k.fmt;
      if (bos && k.eksik) c.fill = dolgu('FFFFF1D6');
      else if (ri % 2) c.fill = dolgu(XR.zebra);
      c.border = XKENAR;
      c.alignment = { vertical: 'middle', horizontal: k.orta ? 'center' : 'left', indent: k.orta ? 0 : 1, wrapText: !!k.sar };
      if (k.turuncu && !bos) { c.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFC2410C' } }; c.fill = dolgu('FFFFEDD5'); }
    });
    const d = durumlar?.[ri];
    const di = kolonlar.findIndex(k => k.durum);
    if (d && di >= 0 && XDURUM[d]) {
      const c = r.getCell(di + 1);
      c.fill = dolgu(XDURUM[d][0]); c.font = { name: 'Calibri', bold: true, size: 10, color: { argb: XDURUM[d][1] } };
      c.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  });
  let son = 3 + satirlar.length;
  if (!satirlar.length) {
    son = 4; ws.mergeCells(4, 1, 4, n);
    Object.assign(ws.getCell(4, 1), { value: 'Kayıt yok', font: { name: 'Calibri', italic: true, color: { argb: XR.soluk } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: XKENAR });
    ws.getRow(4).height = 24;
  } else ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: son, column: n } };
  if (toplam) {
    const r = ws.getRow(son + 1); r.height = 21;
    toplam.forEach((v, i) => Object.assign(r.getCell(i + 1), {
      value: v === '' ? null : v, font: { name: 'Calibri', bold: true, size: 11, color: { argb: XR.beyaz } }, fill: dolgu(XR.baslikAlt), border: XKENAR,
      alignment: { vertical: 'middle', horizontal: kolonlar[i].orta ? 'center' : 'left', indent: kolonlar[i].orta ? 0 : 1 },
    }));
  }
  return ws;
}


export const subeRengi = sk => sk[0] === '0' ? 'FFF59E0B' : sk[0] === '5' ? 'FF64748B' : ['', 'FF0EA5E9', 'FF10B981', 'FF8B5CF6', 'FFEF4444'][sk[0]];

/** Tek şubenin sınıf listesi sayfası (+ varsa “Ayrılanlar” sayfası).
 *  hepsi: şubenin tüm kayıtları (sayılar buradan), liste: tabloda gösterilecek satırlar,
 *  filtre: başlıkta yazılacak filtre açıklaması (varsa ayrılanlar sayfası eklenmez). */
export function subeSayfalari(wb, { sk, sube = {}, hepsi, liste, filtre = '' }) {
  const tarih = new Date().toLocaleString('tr-TR', { day: 'numeric', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  const say = k => hepsi.filter(o => o.durum === k).length;
  const ayrSay = say('ayrildi');
  const ozet = `${filtre ? 'Listelenen' : 'Mevcut'}: ${liste.length}   ·   Düzeltilen: ${say('duzeltildi')}   ·   Yeni: ${say('yeni')}   ·   Ayrılan: ${ayrSay}${filtre || !ayrSay ? '' : ' (ayrı sayfada)'}`;
  xSayfa(wb, sayfaAdi(sk), {
    baslik: `${subeAdi(sk)}${ozelMi(sk) ? '' : ' Şubesi'}${sube.ogretmen ? ' · ' + sube.ogretmen : ''}`,
    alt: `${OKUL_ADI}   ·   ${ozet}${filtre ? '   ·   ' + filtre : ''}   ·   ${sube.tamamlandi ? 'Öğretmen listeyi tamamladı' : 'Öğretmen listeyi henüz tamamlamadı'}   ·   ${tarih}`,
    kolonlar: [{ h: 'Sıra', w: 6, orta: 1 }, { h: 'Okul No', w: 11, orta: 1 }, { h: 'Öğrenci Ad Soyad', w: 30 }, { h: 'Veli Ad Soyad', w: 28, eksik: 1 },
      { h: 'Telefon', w: 20, fmt: XTEL, eksik: 1 }, { h: 'Durum', w: 13, durum: 1 }, { h: 'Devamsızlık', w: 24, turuncu: 1, sar: 1 }],
    satirlar: liste.map((o, i) => [i + 1, noSayi(o.no), o.ad, o.veli, telSayi(o.tel), XDURUM[o.durum]?.[2] || '', devamsizMetni(o)]),
    durumlar: liste.map(o => o.durum),
    renk: subeRengi(sk),
  });
  const ayr = filtre ? [] : hepsi.filter(o => o.durum === 'ayrildi').sort((x, y) => (x.ad || '').localeCompare(y.ad || '', 'tr'));
  if (ayr.length) xSayfa(wb, 'Ayrılanlar', { baslik: `${subeAdi(sk)} · Ayrılan Öğrenciler`, alt: `${ayr.length} öğrenci   ·   Sınıf listesinde yer almazlar.`, renk: 'FFB42318',
    kolonlar: [{ h: 'Sıra', w: 6, orta: 1 }, { h: 'Okul No', w: 11, orta: 1 }, { h: 'Öğrenci Ad Soyad', w: 30 }, { h: 'Veli Ad Soyad', w: 28, eksik: 1 }, { h: 'Telefon', w: 20, fmt: XTEL, eksik: 1 }, { h: 'İşaretlenme', w: 15, orta: 1 }],
    satirlar: ayr.map((o, i) => [i + 1, noSayi(o.no), o.ad, o.veli, telSayi(o.tel), tarihGoster(o.guncel)]) });
}

/** Çalışma kitabını .xlsx olarak indirir */
export async function excelKaydet(wb, dosyaAdi) {
  const buf = await wb.xlsx.writeBuffer();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  // Dosya adında Türkçe karakter ve boşluk bazı tarayıcılarda sorun çıkarır
  a.download = dosyaAdi.replace(/[çğıöşüÇĞİÖŞÜ]/g, c => ({ ç: 'c', ğ: 'g', ı: 'i', ö: 'o', ş: 's', ü: 'u', Ç: 'C', Ğ: 'G', İ: 'I', Ö: 'O', Ş: 'S', Ü: 'U' }[c])).replace(/[^A-Za-z0-9_.\-]+/g, '_');
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}
export async function yeniKitap() {
  await excelJSYukle();
  const wb = new ExcelJS.Workbook(); wb.creator = OKUL_ADI; wb.created = new Date();
  return wb;
}
