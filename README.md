# Mali Tablo — Yevmiye Dönüştürücü, Bilanço ve Gelir Tablosu

Yevmiye defteri dökümünü **VEDAS Excel yevmiye biçimine** dönüştürür ve **VEDAS Kullanım Rehberi 7.16.4** mantığıyla bilanço ve gelir tablosunu oluşturur. İsterseniz tablolar **geçici vergi dönemleri itibarıyla** da oluşturulur. Tüm işlemler tarayıcıda yapılır. Dosyalar bilgisayarınızdan çıkmaz ve internet bağlantısı gerekmez.

## Özellikler
- **Yevmiye dönüştürme:** `00001-----00001-----MAHSUP-----01/01/2023` gibi fiş başlıklı dökümler, her muavin satırı tek satır olacak şekilde dönüştürülür. Sütunlar şunlardır: Sıra No, Yevmiye Numara, Yevmiye Tarihi, Hesap Kodu, Hesap Adı, Muavin Kodu, Muavin Adı, Açıklama, Borç, Alacak. Satırda açıklama yoksa fiş açıklaması kullanılır. `502 › 502.99 › 502.99.68` gibi kademeli kırılımlarda ara kırılım atlanır; ana hesap ve en alt kırılım alınır.
- **Desteklenen girdiler:** fiş başlıklı döküm, dönüştürülmüş yevmiye ve VEDAS / e-Defter Excel dökümü. Sonuncusunda YEVMIYE_NO, HESAPKODU, BORC, ALACAK gibi başlıklar bulunur.
- **Bilanço:** Tekdüzen hesap planı biçiminde hazırlanır (I–V grupları, harf grupları, ana hesaplar). Sütunlar açılış bilançosu, 1–2–3. geçici vergi dönemleri (kümülatif) ve yıllıktır.
- **Gelir tablosu:** A–K satırları ve ara toplamlardan oluşur. VEDAS'taki Gelir_Tablo görünümü de **GT Detay** sayfasındadır (kayıt sayısı, borç ve alacak toplamı, kalan).
- **Açılış, kapanış ve yansıtma kayıtları:** Otomatik bulunur, yevmiye numaraları üzerinde düzeltilebilir.
- **Kontrol:** Aktif–pasif farkını, ters bakiye veren hesapları, borç–alacak eşitliği sağlanmayan fişleri ve fiş toplam uyumsuzluklarını gösterir.
- Tüm sonuçlar tek bir Excel dosyası olarak indirilebilir.

## GitHub'a yükleme ve bilgisayara kurma
1. GitHub'da yeni bir depo oluşturun (ör. `mali-tablo`) ve **tüm dosyaları** deponun ana dizinine yükleyin. Klasör yoktur, hepsi aynı düzeydedir.
2. **Settings → Pages** bölümünde *Source: Deploy from a branch*, *Branch: main / (root)* seçeneklerini seçip kaydedin.
3. Birkaç dakika sonra uygulama `https://KULLANICIADI.github.io/mali-tablo/` adresinde açılır.
4. Chrome veya Edge'de sağ üstteki **"Bilgisayara yükle"** düğmesine ya da adres çubuğundaki yükleme simgesine tıklayın. Uygulama masaüstünde ayrı bir pencerede, çevrimdışı da çalışır.

Manifest kimliği (`mali-tablo-yevmiye-bilanco`) ve önbellek adı bu uygulamaya özeldir. Bu nedenle aynı GitHub Pages alanındaki diğer uygulamalarla (OCRİK, Kitaptan Sınav vb.) çakışmaz.

## Excel makro sürümü
`Yevmiye_MaliTablo.xlsm` aynı işlemleri Excel içinde yapar. Açılış ekranında dosya seçme düğmeleri vardır. Dosyayı açınca **İçeriği Etkinleştir** deyin. Dosya internetten indirildiyse önce *Özellikler → Engellemeyi kaldır* adımını uygulayın.

## Hesaplama kuralları
- Kapanış kayıtları hiçbir tabloya alınmaz.
- Yansıtma kayıtları (690/692 hesabı çalışan fişler) gelir tablosuna alınmaz, bilançoya alınır.
- Açılış kayıtları yalnızca açılış bilançosu ve bilanço sütunlarına girer.
- Bilançoda kapatılmamış 6–7 grubu bakiyeleri *Dönem Net Kârı (Zararı)* altında ayrı satırda gösterilir. Bu sayede ara dönemlerde de aktif = pasif olur.
- 4. dönem geçici vergi beyanı kaldırıldığından son sütun yıllıktır.

Kullanılan kütüphane: [SheetJS Community Edition](https://sheetjs.com) (Apache 2.0).
