/* Mali Tablo — arayüz */
(function () {
  'use strict';
  const C = window.MaliCore;
  const $ = s => document.querySelector(s);
  const NF = '#,##0.00;[Red]-#,##0.00';
  let ham = null;      // yevmiyeOku sonucu (tüm sayfalar)
  let Y = null;        // seçilen kayıtlar
  let M = null;        // mali tablolar
  let dosyaAdi = '';

  // ------------------------------------------------------------ dosya seçimi
  const inp = $('#dosya');
  inp.addEventListener('change', e => { if (e.target.files[0]) dosyaYukle(e.target.files[0]); inp.value = ''; });
  const drop = $('#drop');
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) dosyaYukle(f); });

  function durum(el, msg, tip) { el.className = 'durum ' + (tip || ''); el.innerHTML = msg; }

  async function dosyaYukle(file) {
    dosyaAdi = file.name;
    durum($('#d1'), 'Okunuyor: <b>' + esc(file.name) + '</b> …');
    ['#adim2', '#adim3', '#adim4', '#sayfaSec'].forEach(s => $(s).hidden = true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: false });
      const sheets = wb.SheetNames.map(n => ({ name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null, blankrows: true }) }));
      ham = C.yevmiyeOku(sheets);
      if (!ham.kayitlar.length) {
        durum($('#d1'), 'Seçilen dosyada tanınabilir yevmiye kaydı bulunamadı. Desteklenen biçimler: fiş başlıklı döküm (ör. <code>00001-----00001-----MAHSUP-----01/01/2023</code>) veya başlık satırında <i>Hesap Kodu / Borç / Alacak</i> bulunan tablo.', 'hata');
        return;
      }
      if (ham.sayfalar.length > 1) {
        $('#sayfaSec').hidden = false;
        $('#sayfaListe').innerHTML = ham.sayfalar.map(s => `<li>${esc(s.name)} — ${s.adet.toLocaleString('tr-TR')} satır</li>`).join('');
        document.querySelectorAll('input[name=birlestir]').forEach(r => r.onchange = sec);
        $('#birlestirEvet').checked = true;
      }
      sec();
    } catch (err) {
      console.error(err);
      durum($('#d1'), 'Dosya okunamadı: ' + esc(err.message), 'hata');
    }
  }

  function sec() {
    const sadeceEnBuyuk = ham.sayfalar.length > 1 && $('#birlestirHayir').checked;
    Y = C.sayfaSec(ham, sadeceEnBuyuk);
    const K = Y.kayitlar;
    const fis = new Set(K.map(k => k.key)).size;
    let tb = 0, ta = 0, minT = null, maxT = null;
    K.forEach(k => { tb += k.b; ta += k.a; if (k.gun != null) { if (minT == null || k.gun < minT) minT = k.gun; if (maxT == null || k.gun > maxT) maxT = k.gun; } });
    durum($('#d1'), `<b>${esc(dosyaAdi)}</b> okundu${Y.unvan ? ' — ' + esc(Y.unvan) : ''}.`, 'ok');
    $('#ozet').innerHTML = [
      ['Satır', K.length.toLocaleString('tr-TR')], ['Yevmiye maddesi', fis.toLocaleString('tr-TR')],
      ['Tarih aralığı', minT != null ? C.fmtGun(minT) + ' – ' + C.fmtGun(maxT) : '—'],
      ['Toplam borç', C.tl(tb)], ['Toplam alacak', C.tl(ta)], ['Uyarı', Y.uyari.length]
    ].map(([a, b]) => `<div class="kutu"><span>${a}</span><b>${b}</b></div>`).join('');
    // önizleme
    const n = Math.min(200, K.length);
    let h = '<table class="tbl yev"><thead><tr>' + C.YEV_BASLIK.map(x => `<th>${esc(x.trim())}</th>`).join('') + '</tr></thead><tbody>';
    for (let i = 0; i < n; i++) {
      const k = K[i];
      h += `<tr><td>${i + 1}</td><td>${esc(yevStr(k.yev))}</td><td>${C.fmtGun(k.gun)}</td><td>${esc(k.hk)}</td><td>${esc(k.ha)}</td><td>${esc(k.mk)}</td><td>${esc(k.ma)}</td><td>${esc(k.ac)}</td><td class="n">${k.b ? C.tl(k.b) : ''}</td><td class="n">${k.a ? C.tl(k.a) : ''}</td></tr>`;
    }
    h += '</tbody></table>';
    if (K.length > n) h += `<p class="dip">İlk ${n} satır gösteriliyor; indirilen dosyada ${K.length.toLocaleString('tr-TR')} satırın tamamı vardır.</p>`;
    $('#onizleme').innerHTML = h;
    $('#adim2').hidden = false;
    // özel fişler
    const oz = C.ozelFisleriBul(Y);
    $('#acilis').value = oz.ac; $('#kapanis').value = oz.kap; $('#yansitma').value = oz.yan;
    document.querySelectorAll('input[name=gecici]').forEach(r => r.checked = false);
    $('#adim3').hidden = false;
    $('#adim4').hidden = true;
  }

  const yevStr = v => typeof v === 'number' ? String(v).padStart(5, '0') : String(v || '');

  // ------------------------------------------------------------ VEDAS yevmiye indir
  $('#yevIndir').addEventListener('click', () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, yevmiyeSheet(), 'Yevmiye');
    XLSX.writeFile(wb, 'VEDAS_Yevmiye.xlsx');
  });

  function yevmiyeSheet() {
    const aoa = [C.YEV_BASLIK];
    Y.kayitlar.forEach((k, i) => aoa.push([i + 1, k.yev, k.gun != null ? k.gun + 25569 : null, k.hk, k.ha, k.mk, k.ma, k.ac, k.b, k.a]));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    for (let r = 1; r < aoa.length; r++) {
      setZ(ws, r, 1, '00000'); setZ(ws, r, 2, 'dd.mm.yyyy'); setZ(ws, r, 8, NF); setZ(ws, r, 9, NF);
      [3, 4, 5, 6, 7].forEach(c => { const a = XLSX.utils.encode_cell({ r, c }); if (ws[a]) { ws[a].t = 's'; ws[a].v = String(ws[a].v); } });
    }
    ws['!cols'] = [8, 10, 12, 10, 30, 14, 30, 36, 16, 16].map(w => ({ wch: w }));
    ws['!autofilter'] = { ref: 'A1:J' + aoa.length };
    return ws;
  }
  function setZ(ws, r, c, z) { const a = XLSX.utils.encode_cell({ r, c }); if (ws[a] && ws[a].t === 'n') { ws[a].z = z; if (z === NF) ws[a].v = Math.round(ws[a].v * 100) / 100 || 0; } }

  // ------------------------------------------------------------ mali tablolar
  $('#olustur').addEventListener('click', () => {
    const g = document.querySelector('input[name=gecici]:checked');
    if (!g) { durum($('#d3'), 'Lütfen geçici vergi dönemleri sorusunu yanıtlayın.', 'hata'); $('#geciciSoru').classList.add('vurgu'); return; }
    $('#geciciSoru').classList.remove('vurgu');
    try {
      M = C.maliTablolar(Y, { ac: $('#acilis').value, kap: $('#kapanis').value, yan: $('#yansitma').value, gecici: g.value === 'evet' });
      durum($('#d3'), 'Tablolar oluşturuldu.', 'ok');
      $('#adim4').hidden = false;
      sekme('bil');
      $('#adim4').scrollIntoView({ behavior: 'smooth' });
    } catch (err) { console.error(err); durum($('#d3'), 'Hata: ' + esc(err.message), 'hata'); }
  });

  document.querySelectorAll('.sekmeler button').forEach(b => b.addEventListener('click', () => sekme(b.dataset.s)));
  function sekme(s) {
    document.querySelectorAll('.sekmeler button').forEach(b => b.classList.toggle('aktif', b.dataset.s === s));
    const el = $('#sonuc');
    if (s === 'bil') el.innerHTML = tabloHTML(M.bil, M.bilCols);
    if (s === 'gt') el.innerHTML = tabloHTML(M.gt, M.gtCols);
    if (s === 'gtd') el.innerHTML = M.gtd.map(b => `<h3>${esc(b.label)}</h3>` + gtdHTML(b)).join('');
    if (s === 'miz') el.innerHTML = mizanHTML();
    if (s === 'kon') el.innerHTML = kontrolHTML();
  }

  const sinif = t => ['hesap', 'grup', 'roma', 'toplam', 'bolum', 'not', 'ara'][t];
  const para = v => v == null ? '' : `<span class="${v < -0.005 ? 'neg' : ''}">${C.tl(v)}</span>`;

  function tabloHTML(rows, cols) {
    let h = '<div class="kaydir"><table class="tbl mali"><thead><tr><th>Kod</th><th>Hesap / Kalem</th>' +
      cols.map(p => `<th>${esc(M.donem[p].label).replace('\n', '<br>')}</th>`).join('') + '</tr></thead><tbody>';
    rows.forEach(r => {
      if (r.tip === 4 || r.tip === 5) { h += `<tr class="${sinif(r.tip)}"><td></td><td colspan="${cols.length + 1}">${esc(r.ad)}</td></tr>`; return; }
      h += `<tr class="${sinif(r.tip)}"><td>${esc(r.kod)}</td><td>${esc(r.ad)}</td>` + cols.map(p => `<td class="n">${para(r.v[p])}</td>`).join('') + '</tr>';
    });
    return h + '</tbody></table></div>';
  }
  function gtdHTML(b) {
    let h = '<div class="kaydir"><table class="tbl mali"><thead><tr><th>Hesap Kodu</th><th>Hesap Adı</th><th>Kayıt Sayısı</th><th>Borç Toplam</th><th>Alacak Toplam</th><th>Kalan</th></tr></thead><tbody>';
    b.rows.forEach(r => {
      const ad = r.tip === 0 ? r.ad : (r.kod ? r.kod.replace('.', '') + '-' : '') + r.ad;
      h += `<tr class="${sinif(r.tip)}"><td>${r.tip === 0 ? esc(r.kod) : ''}</td><td>${esc(ad)}</td><td class="n">${r.c.toLocaleString('tr-TR')}</td><td class="n">${para(r.b)}</td><td class="n">${para(r.a)}</td><td class="n">${para(r.v[b.p])}</td></tr>`;
    });
    return h + '</tbody></table></div>';
  }
  function mizanHTML() {
    let tb = 0, ta = 0, tbb = 0, tab = 0;
    let h = `<h3>Mizan (kapanış kayıtları hariç) — ${esc(M.donem[M.NP].label.replace('\n', '  '))}</h3><div class="kaydir"><table class="tbl mali"><thead><tr><th>Hesap Kodu</th><th>Hesap Adı</th><th>Borç</th><th>Alacak</th><th>Borç Bakiye</th><th>Alacak Bakiye</th></tr></thead><tbody>`;
    M.mizan.forEach(r => { tb += r.b; ta += r.a; tbb += r.bb; tab += r.ab; h += `<tr><td>${r.kod}</td><td>${esc(r.ad)}</td><td class="n">${para(r.b)}</td><td class="n">${para(r.a)}</td><td class="n">${para(r.bb)}</td><td class="n">${para(r.ab)}</td></tr>`; });
    h += `<tr class="toplam"><td></td><td>TOPLAM</td><td class="n">${para(tb)}</td><td class="n">${para(ta)}</td><td class="n">${para(tbb)}</td><td class="n">${para(tab)}</td></tr>`;
    return h + '</tbody></table></div>';
  }
  function kontrolHTML() {
    const k = M.kontrol;
    const sat = (a, b) => `<tr><td>${a}</td><td>${b}</td></tr>`;
    let h = '<div class="kaydir"><table class="tbl kon"><tbody>';
    h += '<tr class="bolum"><td colspan="2">GENEL</td></tr>' + sat('Kaynak dosya', esc(dosyaAdi)) + sat('Unvan', esc(k.unvan || '—')) + sat('Satır', k.satir.toLocaleString('tr-TR')) +
      sat('Yevmiye maddesi', k.fisSayisi.toLocaleString('tr-TR')) + sat('Tarih aralığı', k.aralik) + sat('Toplam borç', para(k.tb)) + sat('Toplam alacak', para(k.ta)) + sat('Fark', para(k.tb - k.ta));
    h += '<tr class="bolum"><td colspan="2">VEDAS ANALİZ AYARLARI (7.16.4)</td></tr>' + sat('Açılış kayıtları', esc(k.ayar.ac || '(yok)')) + sat('Kapanış kayıtları', esc(k.ayar.kap || '(yok)')) + sat('Yansıtma kayıtları', esc(k.ayar.yan || '(yok)'));
    h += '<tr class="bolum"><td colspan="2">AKTİF − PASİF KONTROLÜ</td></tr>' + k.apFark.map(x => sat(esc(x.label), `<span class="${Math.abs(x.v) > 0.01 ? 'neg' : ''}">${C.tl(x.v)}</span>`)).join('');
    h += '<tr class="bolum"><td colspan="2">TERS BAKİYE VEREN BİLANÇO HESAPLARI (VEDAS 7.16.3 benzeri)</td></tr>' +
      (k.ters.length ? k.ters.map(t => sat(esc(t.donem) + ' — <b>' + t.kod + '</b> ' + esc(t.ad), para(t.v) + ' · ' + esc(t.aciklama))).join('') : sat('Ters bakiye veren hesap yok.', ''));
    h += '<tr class="bolum"><td colspan="2">BORÇ − ALACAK EŞİTLİĞİ SAĞLANMAYAN YEVMİYE MADDELERİ</td></tr>' +
      (k.dengesiz.length ? k.dengesiz.map(d => sat('Yevmiye ' + esc(d.yev), `Borç ${C.tl(d.b)} · Alacak ${C.tl(d.a)} · Fark ${para(d.fark)}`)).join('') : sat('Tüm yevmiye maddelerinde borç = alacak.', ''));
    h += '<tr class="bolum"><td colspan="2">AKTARIM UYARILARI</td></tr>' + (k.uyari.length ? k.uyari.map(u => sat(esc(u), '')).join('') : sat('Uyarı yok.', ''));
    return h + '</tbody></table></div>';
  }

  // ------------------------------------------------------------ Excel'e aktar
  $('#excelIndir').addEventListener('click', () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, maliSheet('BİLANÇO', M.bil, M.bilCols), 'Bilanço');
    XLSX.utils.book_append_sheet(wb, maliSheet('GELİR TABLOSU', M.gt, M.gtCols), 'Gelir Tablosu');
    XLSX.utils.book_append_sheet(wb, gtdSheet(), 'GT Detay');
    XLSX.utils.book_append_sheet(wb, mizanSheet(), 'Mizan');
    XLSX.utils.book_append_sheet(wb, kontrolSheet(), 'Kontrol');
    XLSX.utils.book_append_sheet(wb, yevmiyeSheet(), 'Yevmiye');
    const ad = (Y.unvan ? Y.unvan.replace(/[\\/:*?"<>|]/g, '').slice(0, 40) + '_' : '') + 'Mali_Tablolar.xlsx';
    XLSX.writeFile(wb, ad);
  });

  function baslikSatirlari(t) { return [[t], [Y.unvan || ''], ['Kaynak: ' + dosyaAdi + '   |   Oluşturma: ' + new Date().toLocaleString('tr-TR')], []]; }
  function maliSheet(t, rows, cols) {
    const aoa = baslikSatirlari(t);
    aoa.push(['Kod', 'Hesap / Kalem'].concat(cols.map(p => M.donem[p].label.replace('\n', ' '))));
    const r0 = aoa.length;
    rows.forEach(r => aoa.push([r.kod, r.tip === 0 ? '      ' + r.ad : r.tip === 1 ? '   ' + r.ad : r.ad].concat(r.tip === 4 || r.tip === 5 ? [] : cols.map(p => r.v[p]))));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    for (let r = r0; r < aoa.length; r++) for (let c = 2; c < cols.length + 2; c++) setZ(ws, r, c, NF);
    ws['!cols'] = [{ wch: 8 }, { wch: 60 }].concat(cols.map(() => ({ wch: 22 })));
    return ws;
  }
  function gtdSheet() {
    const aoa = baslikSatirlari('GELİR TABLOSU DETAYI (VEDAS Gelir_Tablo görünümü)');
    const fmtRows = [];
    M.gtd.forEach(b => {
      aoa.push([b.label]); aoa.push(['HESAP KODU', 'HESAP ADI', 'KAYIT SAYISI', 'BORÇ TOPLAM', 'ALACAK TOPLAM', 'KALAN']);
      b.rows.forEach(r => { fmtRows.push(aoa.length); aoa.push([r.tip === 0 ? r.kod : '', r.tip === 0 ? r.ad : (r.kod ? r.kod.replace('.', '') + '-' : '') + r.ad, r.c, r.b, r.a, r.v[b.p]]); });
      aoa.push([], []);
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    fmtRows.forEach(r => [3, 4, 5].forEach(c => setZ(ws, r, c, NF)));
    ws['!cols'] = [12, 60, 13, 18, 18, 18].map(w => ({ wch: w }));
    return ws;
  }
  function mizanSheet() {
    const aoa = baslikSatirlari('MİZAN (Kapanış kayıtları hariç) — ' + M.donem[M.NP].label.replace('\n', ' '));
    aoa.push(['Hesap Kodu', 'Hesap Adı', 'Borç', 'Alacak', 'Borç Bakiye', 'Alacak Bakiye']);
    const r0 = aoa.length; let t = [0, 0, 0, 0];
    M.mizan.forEach(r => { aoa.push([r.kod, r.ad, r.b, r.a, r.bb, r.ab]); t = [t[0] + r.b, t[1] + r.a, t[2] + r.bb, t[3] + r.ab]; });
    aoa.push(['', 'TOPLAM'].concat(t));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    for (let r = r0; r < aoa.length; r++) for (let c = 2; c < 6; c++) setZ(ws, r, c, NF);
    ws['!cols'] = [11, 45, 18, 18, 18, 18].map(w => ({ wch: w }));
    return ws;
  }
  function kontrolSheet() {
    const k = M.kontrol, aoa = [['KONTROL VE UYARILAR'], []];
    aoa.push(['Kaynak dosya', dosyaAdi], ['Unvan', k.unvan], ['Satır', k.satir], ['Yevmiye maddesi', k.fisSayisi], ['Tarih aralığı', k.aralik], ['Toplam borç', k.tb], ['Toplam alacak', k.ta], ['Fark', k.tb - k.ta], []);
    aoa.push(['VEDAS ANALİZ AYARLARI (7.16.4)'], ['Açılış kayıtları', k.ayar.ac || '(yok)'], ['Kapanış kayıtları', k.ayar.kap || '(yok)'], ['Yansıtma kayıtları', k.ayar.yan || '(yok)'], []);
    aoa.push(['AKTİF - PASİF KONTROLÜ']); k.apFark.forEach(x => aoa.push([x.label, x.v])); aoa.push([]);
    aoa.push(['TERS BAKİYE VEREN BİLANÇO HESAPLARI'], ['Dönem', 'Hesap', 'Bakiye', 'Açıklama']);
    if (k.ters.length) k.ters.forEach(t => aoa.push([t.donem, t.kod + ' ' + t.ad, t.v, t.aciklama])); else aoa.push(['Ters bakiye veren hesap yok.']);
    aoa.push([], ['BORÇ - ALACAK EŞİTLİĞİ SAĞLANMAYAN YEVMİYE MADDELERİ']);
    if (k.dengesiz.length) k.dengesiz.forEach(d => aoa.push(['Yevmiye ' + d.yev, d.b, d.a, d.fark])); else aoa.push(['Tüm yevmiye maddelerinde borç = alacak.']);
    aoa.push([], ['AKTARIM UYARILARI']);
    if (k.uyari.length) k.uyari.forEach(u => aoa.push([u])); else aoa.push(['Uyarı yok.']);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    for (let r = 2; r < aoa.length; r++) for (let c = 1; c < 4; c++) setZ(ws, r, c, NF);
    ws['!cols'] = [48, 40, 22, 50].map(w => ({ wch: w }));
    return ws;
  }

  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

  // ------------------------------------------------------------ kurulum (PWA)
  let kurulumOlayi = null;
  window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); kurulumOlayi = e; $('#kur').hidden = false; });
  $('#kur').addEventListener('click', async () => { if (!kurulumOlayi) return; kurulumOlayi.prompt(); await kurulumOlayi.userChoice; kurulumOlayi = null; $('#kur').hidden = true; });
  if ('serviceWorker' in navigator) window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
})();
