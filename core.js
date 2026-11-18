/* Mali Tablo — çekirdek hesaplama (tarayıcı + Node)
 * - Fiş başlıklı yevmiye dökümünü VEDAS Excel yevmiye biçimine dönüştürür
 * - VEDAS Rehberi 7.16.4 mantığıyla bilanço / gelir tablosu üretir
 */
(function (root) {
  'use strict';

  // ------------------------------------------------------------ yardımcılar
  const TRMAP = { 'ı': 'I', 'İ': 'I', 'i': 'I', 'ş': 'S', 'Ş': 'S', 'ğ': 'G', 'Ğ': 'G', 'ü': 'U', 'Ü': 'U', 'ö': 'O', 'Ö': 'O', 'ç': 'C', 'Ç': 'C', 'â': 'A', 'Â': 'A' };
  const norm = s => String(s == null ? '' : s).replace(/[ıİişŞğĞüÜöÖçÇâÂ]/g, c => TRMAP[c]).toUpperCase();
  const basNorm = s => norm(s).replace(/[ _.\-:]/g, '');

  function hucre(row, c) {
    if (!row || c < 0 || c >= row.length) return '';
    const x = row[c];
    if (x == null) return '';
    if (x instanceof Date) return fmtGun(dateToGun(x));
    return String(x);
  }

  function sayi(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    if (typeof v === 'boolean' || v instanceof Date) return 0;
    let s = String(v).replace(/\s| /g, '').replace(/TL/gi, '');
    if (s === '' || s === '-') return 0;
    let neg = false;
    if (s[0] === '(' && s[s.length - 1] === ')') { neg = true; s = s.slice(1, -1); }
    const pd = s.lastIndexOf('.'), pc = s.lastIndexOf(',');
    if (pd >= 0 && pc >= 0) {
      s = pc > pd ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    } else if (pc >= 0) {
      s = (s.split(',').length - 1 > 1) ? s.replace(/,/g, '') : s.replace(',', '.');
    } else if (pd >= 0) {
      if (s.split('.').length - 1 > 1) s = s.replace(/\./g, '');
    }
    const x = parseFloat(s);
    if (!isFinite(x)) return 0;
    return neg ? -x : x;
  }

  // gün sayısı: 1970-01-01'den itibaren gün (UTC)
  const gunUTC = (y, m, d) => Math.round(Date.UTC(y, m - 1, d) / 86400000);
  function dateToGun(d) { return gunUTC(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
  function gunParca(g) { const d = new Date(g * 86400000); return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()]; }
  function fmtGun(g) { if (g == null) return ''; const [y, m, d] = gunParca(g); return String(d).padStart(2, '0') + '.' + String(m).padStart(2, '0') + '.' + y; }
  const ayEkle = (g, n) => { const [y, m] = gunParca(g); return gunUTC(y, m + n, 1); }; // ayın 1'i

  function tarih(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return dateToGun(v);
    if (typeof v === 'number') { if (v > 10000 && v < 100000) return Math.floor(v) - 25569; return null; }
    let s = String(v).trim();
    s = s.split(' ')[0].split('T')[0];
    const p = s.replace(/[.\-]/g, '/').split('/');
    if (p.length !== 3 || p.some(x => !/^\d+$/.test(x))) return null;
    if (p[0].length === 4) return gunUTC(+p[0], +p[1], +p[2]);
    let y = p[2]; if (y.length === 2) y = '20' + y;
    return gunUTC(+y, +p[1], +p[0]);
  }

  const isNumStr = s => /^\s*-?\d+(\.\d+)?\s*$/.test(String(s));
  function ykey(v) { const s = String(v == null ? '' : v).trim(); if (!s) return ''; return isNumStr(s) ? String(parseFloat(s)) : s; }

  function fisBasligi(a) {
    a = String(a || '').trim();
    if (a.indexOf('--') < 0) return null;
    while (a.indexOf('---') >= 0) a = a.split('---').join('--');
    const par = a.split('--').map(x => x.trim()).filter(x => x !== '');
    if (par.length < 2) return null;
    if (!isNumStr(par[0])) return null;
    const t = tarih(par[par.length - 1]);
    if (t == null) return null;
    let tur = '';
    if (par.length >= 4) tur = par[2]; else if (par.length === 3) tur = par[1];
    return { yev: par[0], gun: t, tur };
  }

  const anaKodMu = a => a.length > 0 && a[0] !== ' ' && /^\d{3}$/.test(a.trim());
  const anaKod = s => { s = String(s || '').trim(); return /^\d{3}/.test(s) ? s.slice(0, 3) : ''; };

  // ------------------------------------------------------------ çözümleme
  function yeniKayit(ctx, yev, gun, hk, ha, mk, ma, ac, b, al) {
    const yevS = String(yev == null ? '' : yev).trim();
    ctx.out.push({ yev: isNumStr(yevS) ? parseFloat(yevS) : yevS, key: ykey(yevS), gun, hk, ha, mk, ma, ac, b, a: al });
  }

  function hamCozumle(rows, ctx) {
    const i0 = rows.findIndex(r => fisBasligi(hucre(r, 0)));
    if (i0 < 0) return 0;
    const n0 = ctx.out.length;
    if (norm(hucre(rows[0], 0)).includes('YEVMIYE DEFTERI') && !ctx.unvan) ctx.unvan = hucre(rows[1], 0).trim();
    let fis = null, ana = null, bek = null;
    // 502 > 502.99 > 502.99.68 gibi kademeli kırılımlarda ara kırılım atılır, en alt kırılım alınır
    const bekleyenYaz = () => {
      if (bek && ana) yeniKayit(ctx, fis.yev, fis.gun, ana.kod, ana.ad, bek.kod, bek.ad, bek.ac, bek.b, bek.a);
      bek = null;
    };
    const anaKapat = () => {
      bekleyenYaz();
      if (ana && !ana.detay && ana.tutar !== 0) {
        yeniKayit(ctx, fis.yev, fis.gun, ana.kod, ana.ad, '', '', '', ana.taraf === 1 ? ana.tutar : 0, ana.taraf === 2 ? ana.tutar : 0);
      }
      ana = null;
    };
    const fisSonu = () => {
      anaKapat();
      if (fis.ac) for (let i = fis.bas; i < ctx.out.length; i++) if (!String(ctx.out[i].ac).trim()) ctx.out[i].ac = fis.ac;
      const k = ykey(fis.yev);
      if (k && !ctx.fis.has(k)) ctx.fis.set(k, fis.tur + ' | ' + fis.ac);
    };
    for (const r of rows) {
      const a = hucre(r, 0), at = a.trim(), na = norm(at);
      const fb = fisBasligi(at);
      if (fb) {
        if (fis) fisSonu();
        fis = { yev: fb.yev, gun: fb.gun, tur: fb.tur, ac: '', bas: ctx.out.length };
        ana = null;
      } else if (!fis || at === '') {
        // başlık bölgesi / boş satır
      } else if (na.startsWith('HESAP KODU')) {
        // sütun başlığı
      } else if (na.startsWith('FIS ACIKLAMA')) {
        anaKapat();
        let fa = hucre(r, 1).trim();
        if (!fa) { const p = at.indexOf(':'); if (p >= 0) fa = at.slice(p + 1).trim(); }
        fis.ac = fa;
      } else if (na.startsWith('TOPLAM')) {
        anaKapat();
        const e = sayi(r[4]), f = sayi(r[5]);
        let sb = 0, sa = 0;
        for (let i = fis.bas; i < ctx.out.length; i++) { sb += ctx.out[i].b; sa += ctx.out[i].a; }
        if (Math.abs(sb - e) > 0.01 || Math.abs(sa - f) > 0.01)
          ctx.uyari.push(`Yevmiye ${fis.yev}: fiş TOPLAM satırı (${tl(e)} / ${tl(f)}) ile aktarılan satırların toplamı (${tl(sb)} / ${tl(sa)}) farklı.`);
      } else if (na.includes('NAKLI') || na.includes('YEKUN')) {
        // nakli yekün
      } else if (anaKodMu(a)) {
        anaKapat();
        const e = sayi(r[4]), f = sayi(r[5]);
        ana = { kod: at, ad: hucre(r, 1).trim(), taraf: (f !== 0 && e === 0) ? 2 : 1, tutar: (f !== 0 && e === 0) ? f : e, detay: false };
      } else if (ana) {
        const t = sayi(r[3]);
        let e, f;
        if (t !== 0) { e = ana.taraf === 1 ? t : 0; f = ana.taraf === 2 ? t : 0; } else { e = sayi(r[4]); f = sayi(r[5]); }
        if (e !== 0 || f !== 0) {
          if (bek && !at.startsWith(bek.kod + '.')) yeniKayit(ctx, fis.yev, fis.gun, ana.kod, ana.ad, bek.kod, bek.ad, bek.ac, bek.b, bek.a);
          bek = { kod: at, ad: hucre(r, 1).trim(), ac: hucre(r, 2).trim(), b: e, a: f };
          ana.detay = true;
        }
      }
    }
    if (fis) fisSonu();
    return ctx.out.length - n0;
  }

  function tabloCozumle(rows, ctx) {
    let hr = -1, cH = -1, cB = -1, cAl = -1;
    for (let r = 0; r < Math.min(50, rows.length) && hr < 0; r++) {
      cH = cB = cAl = -1;
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c++) {
        const h = basNorm(hucre(row, c));
        if (h === 'HESAPKODU' || h === 'ANAHESAPKODU') cH = c;
        if (h === 'BORC' || h === 'BORCTUTARI' || h === 'BORCTUTAR') cB = c;
        if (h === 'ALACAK' || h === 'ALACAKTUTARI' || h === 'ALACAKTUTAR') cAl = c;
      }
      if (cH >= 0 && cB >= 0 && cAl >= 0) hr = r;
    }
    if (hr < 0) return 0;
    const H = (rows[hr] || []).map((x, c) => basNorm(hucre(rows[hr], c)));
    let cY = -1, cT = -1, cHA = -1, cM = -1, cMA = -1, cA = -1;
    H.forEach((h, c) => {
      if (h === 'HESAPADI' || h === 'ANAHESAPADI') cHA = c;
      if (['MUAVINKODU', 'MUAVINHESAPKODU', 'ALTHESAPKODU', 'DETAYHESAPKODU'].includes(h)) cM = c;
      if (['MUAVINADI', 'MUAVINHESAPADI', 'ALTHESAPADI', 'DETAYHESAPADI'].includes(h)) cMA = c;
      if (h === 'ACIKLAMA') cA = c;
      if (h.includes('TARIH') && (cT < 0 || h.includes('YEVMIYE'))) cT = c;
      if (h.includes('YEVMIYE') && !h.includes('TARIH') && (h.includes('NO') || h.includes('NUMARA') || h.includes('MADDE'))) cY = c;
    });
    if (cA < 0) cA = H.findIndex(h => h.includes('ACIKLAMA'));
    if (cY < 0) cY = H.findIndex(h => ['MADDENO', 'FISNO', 'YEVMIYE', 'YEVMIYEMADDENO'].includes(h));
    if (cY < 0 && cT < 0) return 0; // mizan vb. tablolar yevmiye sayılmaz
    const n0 = ctx.out.length;
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      let hk = hucre(row, cH).trim();
      const b = sayi(row[cB]), al = sayi(row[cAl]);
      if (!hk || (b === 0 && al === 0) || basNorm(hk).startsWith('TOPLAM')) continue;
      let mk = cM >= 0 ? hucre(row, cM).trim() : '';
      const p = hk.indexOf('.');
      if (p >= 0) { if (!mk) mk = hk; hk = hk.slice(0, p); }
      else if (hk.length > 3 && /^\d{3}/.test(hk)) { if (!mk) mk = hk; hk = hk.slice(0, 3); }
      yeniKayit(ctx, cY >= 0 ? hucre(row, cY).trim() : '', cT >= 0 ? tarih(row[cT]) : null, hk,
        cHA >= 0 ? hucre(row, cHA).trim() : '', mk, cMA >= 0 ? hucre(row, cMA).trim() : '', cA >= 0 ? hucre(row, cA).trim() : '', b, al);
    }
    if (ctx.out.length > n0 && cY < 0) ctx.uyari.push('Yevmiye numarası sütunu bulunamadı; açılış / kapanış / yansıtma kayıtları numara ile seçilemez.');
    return ctx.out.length - n0;
  }

  /** sheets: [{name, rows}] → {kayitlar, sayfalar:[{name,bas,son,adet}], unvan, uyari, fis} */
  function yevmiyeOku(sheets) {
    const ctx = { out: [], uyari: [], fis: new Map(), unvan: '' };
    const sayfalar = [];
    for (const sh of sheets) {
      const n0 = ctx.out.length;
      let k = hamCozumle(sh.rows, ctx);
      if (k === 0) k = tabloCozumle(sh.rows, ctx);
      if (k > 0) sayfalar.push({ name: sh.name, bas: n0, son: ctx.out.length, adet: k });
    }
    return { kayitlar: ctx.out, sayfalar, unvan: ctx.unvan, uyari: ctx.uyari, fis: ctx.fis };
  }

  function sayfaSec(y, sadeceEnBuyuk) {
    if (!sadeceEnBuyuk || y.sayfalar.length < 2) return y;
    const e = y.sayfalar.reduce((a, b) => (b.adet > a.adet ? b : a));
    return Object.assign({}, y, { kayitlar: y.kayitlar.slice(e.bas, e.son), sayfalar: [e] });
  }

  const YEV_BASLIK = ['Sıra No ', 'Yevmiye Numara', 'Yevmiye Tarihi', 'Hesap Kodu', 'Hesap Adı', 'Muavin Kodu', 'Muavin Adı', 'Açıklama', 'Borç', 'Alacak'];

  // ------------------------------------------------------------ özel fiş tespiti
  function ozelFisleriBul(y) {
    const K = y.kayitlar;
    let minT = null, maxT = null;
    K.forEach(k => { if (k.gun != null) { if (minT == null || k.gun < minT) minT = k.gun; if (maxT == null || k.gun > maxT) maxT = k.gun; } });
    const ac = [], kap = [], yan = [];
    const add = (arr, k) => { if (k && !arr.includes(k)) arr.push(k); };
    K.forEach(k => {
      if (!k.key) return;
      const t = norm(k.ac) + ' ' + norm(y.fis.get(k.key) || '');
      if (t.includes('ACILIS') && (k.gun === minT || k.gun == null)) add(ac, k.key);
      if (t.includes('KAPANIS') && (k.gun === maxT || k.gun == null)) add(kap, k.key);
    });
    K.forEach(k => { const h = anaKod(k.hk); if (k.key && (h === '690' || h === '692') && !kap.includes(k.key)) add(yan, k.key); });
    return { ac: ac.join(', '), kap: kap.join(', '), yan: yan.join(', ') };
  }

  function listeCoz(s) {
    const set = new Set();
    String(s || '').replace(/[; ]/g, ',').split(',').map(x => x.trim()).filter(Boolean).forEach(t => {
      const q = t.indexOf('-', 1);
      if (q > 0) {
        const a = parseFloat(t.slice(0, q)), b = parseFloat(t.slice(q + 1));
        if (isFinite(a) && isFinite(b) && b - a < 100000) for (let x = a; x <= b; x++) set.add(ykey(x));
      } else set.add(ykey(t));
    });
    return set;
  }

  // ------------------------------------------------------------ yapı tanımları
  const BIL_YAPI = [
    ['H', 'AKTİF (VARLIKLAR)'],
    ['R', 'I.', 'DÖNEN VARLIKLAR', 100, 199], ['G', 'A.', 'Hazır Değerler', 100, 109], ['G', 'B.', 'Menkul Kıymetler', 110, 119],
    ['G', 'C.', 'Ticari Alacaklar', 120, 129], ['G', 'D.', 'Diğer Alacaklar', 130, 149], ['G', 'E.', 'Stoklar', 150, 169],
    ['G', 'F.', 'Yıllara Yaygın İnşaat ve Onarım Maliyetleri', 170, 179], ['G', 'G.', 'Gelecek Aylara Ait Giderler ve Gelir Tahakkukları', 180, 189],
    ['G', 'H.', 'Diğer Dönen Varlıklar', 190, 199],
    ['R', 'II.', 'DURAN VARLIKLAR', 200, 299], ['G', 'A.', 'Ticari Alacaklar', 200, 229], ['G', 'B.', 'Diğer Alacaklar', 230, 239],
    ['G', 'C.', 'Mali Duran Varlıklar', 240, 249], ['G', 'D.', 'Maddi Duran Varlıklar', 250, 259], ['G', 'E.', 'Maddi Olmayan Duran Varlıklar', 260, 269],
    ['G', 'F.', 'Özel Tükenmeye Tabi Varlıklar', 270, 279], ['G', 'G.', 'Gelecek Yıllara Ait Giderler ve Gelir Tahakkukları', 280, 289],
    ['G', 'H.', 'Diğer Duran Varlıklar', 290, 299],
    ['T', 'AKTİF TOPLAMI (I + II)', 100, 299],
    ['H', 'PASİF (KAYNAKLAR)'],
    ['R', 'III.', 'KISA VADELİ YABANCI KAYNAKLAR', 300, 399], ['G', 'A.', 'Mali Borçlar', 300, 319], ['G', 'B.', 'Ticari Borçlar', 320, 329],
    ['G', 'C.', 'Diğer Borçlar', 330, 339], ['G', 'D.', 'Alınan Avanslar', 340, 349], ['G', 'E.', 'Yıllara Yaygın İnşaat ve Onarım Hakedişleri', 350, 359],
    ['G', 'F.', 'Ödenecek Vergi ve Diğer Yükümlülükler', 360, 369], ['G', 'G.', 'Borç ve Gider Karşılıkları', 370, 379],
    ['G', 'H.', 'Gelecek Aylara Ait Gelirler ve Gider Tahakkukları', 380, 389], ['G', 'I.', 'Diğer Kısa Vadeli Yabancı Kaynaklar', 390, 399],
    ['R', 'IV.', 'UZUN VADELİ YABANCI KAYNAKLAR', 400, 499], ['G', 'A.', 'Mali Borçlar', 400, 419], ['G', 'B.', 'Ticari Borçlar', 420, 429],
    ['G', 'C.', 'Diğer Borçlar', 430, 439], ['G', 'D.', 'Alınan Avanslar', 440, 449], ['G', 'E.', 'Borç ve Gider Karşılıkları', 450, 479],
    ['G', 'F.', 'Gelecek Yıllara Ait Gelirler ve Gider Tahakkukları', 480, 489], ['G', 'G.', 'Diğer Uzun Vadeli Yabancı Kaynaklar', 490, 499],
    ['R', 'V.', 'ÖZKAYNAKLAR', 500, 599], ['G', 'A.', 'Ödenmiş Sermaye', 500, 519], ['G', 'B.', 'Sermaye Yedekleri', 520, 539],
    ['G', 'C.', 'Kâr Yedekleri', 540, 569], ['G', 'D.', 'Geçmiş Yıllar Kârları', 570, 579], ['G', 'E.', 'Geçmiş Yıllar Zararları (-)', 580, 589],
    ['G', 'F.', 'Dönem Net Kârı (Zararı)', 590, 599],
    ['T', 'PASİF TOPLAMI (III + IV + V)', 300, 599],
  ];
  const GT_YAPI = [
    ['G', 'A.', 'BRÜT SATIŞLAR', 600, 609], ['G', 'B.', 'SATIŞ İNDİRİMLERİ (-)', 610, 619], ['S', 'C.', 'NET SATIŞLAR'],
    ['G', 'D.', 'SATIŞLARIN MALİYETİ (-)', 620, 629], ['S', '', 'BRÜT SATIŞ KÂRI VEYA ZARARI'],
    ['G', 'E.', 'FAALİYET GİDERLERİ (-)', 630, 639], ['S', '', 'FAALİYET KÂRI VEYA ZARARI'],
    ['G', 'F.', 'DİĞER FAALİYETLERDEN OLAĞAN GELİR VE KÂRLAR', 640, 649], ['G', 'G.', 'DİĞER FAALİYETLERDEN OLAĞAN GİDER VE ZARARLAR (-)', 650, 659],
    ['G', 'H.', 'FİNANSMAN GİDERLERİ (-)', 660, 669], ['S', '', 'OLAĞAN KÂR VEYA ZARAR'],
    ['G', 'I.', 'OLAĞANDIŞI GELİR VE KÂRLAR', 670, 679], ['G', 'J.', 'OLAĞANDIŞI GİDER VE ZARARLAR (-)', 680, 689],
    ['Y', '', '7 GRUBU MALİYET HESAPLARINDA YANSITILMAMIŞ BAKİYE (-)', 700, 799],
    ['S', '', 'DÖNEM KÂRI VEYA ZARARI'],
    ['G', 'K.', 'DÖNEM KÂRI VERGİ VE DİĞER YASAL YÜKÜMLÜLÜK KARŞILIKLARI (-)', 691, 691],
    ['S', '', 'DÖNEM NET KÂRI VEYA ZARARI'],
  ];
  const TERS = new Set('103,119,122,124,129,137,139,158,222,224,229,237,239,241,243,244,246,247,248,249,257,268,278,299,302,308,402,408,501,503,580,591'.split(','));
  const aralik = (kod, mn, mx) => /^\d{3}$/.test(kod) && +kod >= mn && +kod <= mx;

  // ------------------------------------------------------------ mali tablolar
  function maliTablolar(y, ayar) {
    const K = y.kayitlar;
    const cAc = listeCoz(ayar.ac), cKap = listeCoz(ayar.kap), cYan = listeCoz(ayar.yan);
    const hasAc = cAc.size > 0;
    let minT = null, maxT = null;
    K.forEach(k => { if (k.gun != null) { if (minT == null || k.gun < minT) minT = k.gun; if (maxT == null || k.gun > maxT) maxT = k.gun; } });
    if (minT == null) throw new Error('Kayıtlarda geçerli yevmiye tarihi bulunamadı.');
    const [by, bm] = gunParca(minT);
    const bas = gunUTC(by, bm, 1);
    const yilSonu = ayEkle(bas, 12) - 1;
    const donem = []; // {p, end, label}
    donem.push({ p: 0, end: null, label: 'Açılış Bilançosu\n' + fmtGun(minT) });
    if (ayar.gecici) for (let q = 1; q <= 3; q++) {
      const e = ayEkle(bas, 3 * q) - 1;
      if (e < maxT) donem.push({ p: donem.length, end: e, label: `${q}. Geçici Vergi Dönemi\n${fmtGun(bas)} - ${fmtGun(e)}` });
    }
    donem.push({ p: donem.length, end: maxT, label: (maxT >= yilSonu ? 'Yıllık' : 'Son Kayıt Tarihi İtibarıyla') + `\n${fmtGun(bas)} - ${fmtGun(maxT)}` });
    const NP = donem.length - 1;

    const hesap = new Map(); // kod -> {kod, ad, BB[], BA[], BC[], GB[], GA[], GC[]}
    const z = () => new Array(NP + 1).fill(0);
    const uyari = y.uyari.slice();
    K.forEach((k, i) => {
      const hk = anaKod(k.hk);
      if (!hk) { uyari.push(`Satır ${i + 1}: hesap kodu tanınamadı (${k.hk}), mali tablolara alınmadı.`); return; }
      let h = hesap.get(hk);
      if (!h) { h = { kod: hk, ad: '', BB: z(), BA: z(), BC: z(), GB: z(), GA: z(), GC: z() }; hesap.set(hk, h); }
      if (!h.ad) h.ad = String(k.ha || '').trim();
      const isK = cKap.has(k.key), isY = cYan.has(k.key), isA = cAc.has(k.key);
      if (isA && hasAc) { h.BB[0] += k.b; h.BA[0] += k.a; h.BC[0]++; }
      if (!isK) for (let p = 1; p <= NP; p++) {
        if (k.gun == null || k.gun <= donem[p].end) {
          h.BB[p] += k.b; h.BA[p] += k.a; h.BC[p]++;
          if (!isY && !isA) { h.GB[p] += k.b; h.GA[p] += k.a; h.GC[p]++; }
        }
      }
    });
    const hs = [...hesap.values()].sort((a, b) => a.kod < b.kod ? -1 : a.kod > b.kod ? 1 : 0);
    const res = z();
    hs.forEach(h => { if ('67'.includes(h.kod[0])) for (let p = 0; p <= NP; p++) res[p] += h.BA[p] - h.BB[p]; });
    const bilDeger = (h, p) => '12'.includes(h.kod[0]) ? h.BB[p] - h.BA[p] : '345'.includes(h.kod[0]) ? h.BA[p] - h.BB[p] : 0;

    // ---- bilanço
    const bilCols = (hasAc ? [0] : []).concat(donem.slice(1).map(d => d.p));
    const bil = [];
    let aktifR = null, pasifR = null;
    const resVar = res.some(v => Math.abs(v) > 0.005);
    BIL_YAPI.forEach(it => {
      if (it[0] === 'H') { bil.push({ tip: 4, kod: '', ad: it[1] }); return; }
      let row, mn, mx;
      if (it[0] === 'T') { row = { tip: 3, kod: '', ad: it[1], v: z() }; mn = it[2]; mx = it[3]; if (mx < 300) aktifR = row; else pasifR = row; }
      else { row = { tip: it[0] === 'R' ? 2 : 1, kod: it[1], ad: it[2], v: z() }; mn = it[3]; mx = it[4]; }
      bil.push(row);
      hs.forEach(h => { if (aralik(h.kod, mn, mx)) for (let p = 0; p <= NP; p++) row.v[p] += bilDeger(h, p); });
      if (mx >= 599 && mn <= 590) for (let p = 0; p <= NP; p++) row.v[p] += res[p];
      if (it[0] === 'G') {
        hs.forEach(h => {
          if (aralik(h.kod, mn, mx) && bilCols.some(p => h.BC[p] > 0)) {
            const v = z(); for (let p = 0; p <= NP; p++) v[p] = bilDeger(h, p);
            bil.push({ tip: 0, kod: h.kod, ad: h.ad, v });
          }
        });
        if (mn === 590 && resVar) bil.push({ tip: 0, kod: '6-7', ad: 'Dönem içi kâr / zarar (kapatılmamış 6 ve 7 grubu hesap bakiyeleri)', v: res.slice() });
      }
    });
    const fark = { tip: 3, kod: '', ad: 'FARK (AKTİF - PASİF)', v: z() };
    for (let p = 0; p <= NP; p++) fark.v[p] = aktifR.v[p] - pasifR.v[p];
    bil.push({ tip: 5, kod: '', ad: '' }, fark,
      { tip: 5, kod: '', ad: 'Not: Kapanış kayıtları hariç tutulmuştur. Gelir tablosu hesaplarının dönem içi bakiyesi Özkaynaklar / Dönem Net Kârı (Zararı) altında gösterilir.' },
      { tip: 5, kod: '', ad: 'Not: Aktif hesaplar borç - alacak, pasif hesaplar alacak - borç olarak gösterilir; (-) nitelikli hesaplar eksi değer alır.' });

    // ---- gelir tablosu
    function gtDoldur(p1, p2) {
      const rows = []; const cum = z(); let cB = 0, cA = 0, cC = 0;
      GT_YAPI.forEach(it => {
        if (it[0] === 'S') {
          rows.push({ tip: 6, kod: it[1], ad: it[2], v: cum.slice(), b: cB, a: cA, c: cC });
          return;
        }
        const [, kod, ad, mn, mx] = it;
        const ilgili = hs.filter(h => aralik(h.kod, mn, mx) && (() => { for (let p = p1; p <= p2; p++) if (h.GC[p] > 0) return true; return false; })());
        if (it[0] === 'Y' && ilgili.length === 0) return;
        const g = { tip: 1, kod, ad, v: z(), b: 0, a: 0, c: 0 };
        rows.push(g);
        ilgili.forEach(h => {
          const r = { tip: 0, kod: h.kod, ad: h.ad, v: z(), b: 0, a: 0, c: 0 };
          for (let p = p1; p <= p2; p++) { r.v[p] = h.GA[p] - h.GB[p]; g.v[p] += r.v[p]; }
          if (p1 === p2) { r.b = h.GB[p1]; r.a = h.GA[p1]; r.c = h.GC[p1]; g.b += r.b; g.a += r.a; g.c += r.c; }
          rows.push(r);
        });
        for (let p = p1; p <= p2; p++) cum[p] += g.v[p];
        cB += g.b; cA += g.a; cC += g.c;
      });
      return rows;
    }
    const gtCols = donem.slice(1).map(d => d.p);
    const gt = gtDoldur(1, NP).concat([
      { tip: 5, kod: '', ad: '' },
      { tip: 5, kod: '', ad: 'Not: Açılış, kapanış ve yansıtma kayıtları hariç tutulmuştur (VEDAS Rehberi 7.16.4).' },
      { tip: 5, kod: '', ad: 'Not: Gelir hesapları (+), gider ve indirim hesapları (-) işaretlidir; tutarlar dönem başından itibaren kümülatiftir.' }]);
    const gtd = [];
    for (let p = NP; p >= 1; p--) gtd.push({ p, label: donem[p].label.replace('\n', '  '), rows: gtDoldur(p, p) });

    // ---- mizan (son dönem)
    const mizan = hs.filter(h => h.BC[NP] > 0).map(h => {
      const bk = h.BB[NP] - h.BA[NP];
      return { kod: h.kod, ad: h.ad, b: h.BB[NP], a: h.BA[NP], bb: bk > 0 ? bk : 0, ab: bk > 0 ? 0 : -bk };
    });

    // ---- kontrol
    const fisMap = new Map(); let tb = 0, ta = 0;
    K.forEach(k => { let f = fisMap.get(k.key); if (!f) { f = { b: 0, a: 0 }; fisMap.set(k.key, f); } f.b += k.b; f.a += k.a; tb += k.b; ta += k.a; });
    const dengesiz = [...fisMap.entries()].filter(([, f]) => Math.abs(f.b - f.a) > 0.01).map(([k, f]) => ({ yev: k, b: f.b, a: f.a, fark: f.b - f.a }));
    const apFark = bilCols.map(p => {
      let v = 0; hs.forEach(h => { if (aralik(h.kod, 100, 299)) v += bilDeger(h, p); if (aralik(h.kod, 300, 599)) v -= bilDeger(h, p); });
      return { label: donem[p].label.replace('\n', '  '), v: v - res[p] };
    });
    const ters = [];
    bilCols.forEach(p => hs.forEach(h => {
      if (!aralik(h.kod, 100, 589) || h.BC[p] === 0) return;
      const v = bilDeger(h, p), kontra = TERS.has(h.kod);
      if ((!kontra && v < -0.005) || (kontra && v > 0.005)) ters.push({
        donem: donem[p].label.replace('\n', '  '), kod: h.kod, ad: h.ad, v,
        aciklama: '12'.includes(h.kod[0]) ? (kontra ? '(-) nitelikli aktif hesap borç bakiyesi veriyor' : 'Aktif hesap alacak bakiyesi veriyor')
          : (kontra ? '(-) nitelikli pasif hesap alacak bakiyesi veriyor' : 'Pasif hesap borç bakiyesi veriyor')
      });
    }));
    const kontrol = {
      unvan: y.unvan, satir: K.length, fisSayisi: fisMap.size, aralik: fmtGun(minT) + ' - ' + fmtGun(maxT),
      tb, ta, ayar: { ac: ayar.ac, kap: ayar.kap, yan: ayar.yan }, apFark, ters, dengesiz, uyari
    };
    return { donem, bilCols, gtCols, bil, gt, gtd, mizan, kontrol, NP };
  }

  function tl(x) { if (Math.abs(x || 0) < 0.005) x = 0; return (x || 0).toLocaleString('tr-TR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

  const api = { norm, sayi, tarih, fmtGun, yevmiyeOku, sayfaSec, ozelFisleriBul, maliTablolar, YEV_BASLIK, tl, ykey };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.MaliCore = api;
})(typeof self !== 'undefined' ? self : this);
