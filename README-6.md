# veri.zip – Excel Veri Seti İnceleme Raporu

Bu belge, `veri.zip` içindeki **23 Excel dosyasının** (.xlsx) yapısını, içeriğini ve birbirleriyle ilişkisini ayrıntılı biçimde anlatır. Her dosya sayfa/sütun düzeyinde incelenmiş, sayısal özetler Python (pandas/openpyxl) ile hesaplanmıştır.

> **Gizlilik notu:** `kayit*`, `listeler` ve `isim1` dosyaları ad-soyad, ana/baba adı ve doğum tarihi gibi kişisel veri benzeri alanlar içerir. Bu raporda tekil kayıtlar yeniden üretilmemiş, örneklerde isimler maskelenmiştir; yalnızca toplu istatistikler verilmiştir.

---

## 1. Genel Bakış

| Bilgi | Değer |
|---|---|
| Arşiv adı | `veri.zip` |
| Excel dosyası sayısı | 23 (`veri/` klasöründe; 19'u benzersiz içerikli) |
| Ek içerik | `__MACOSX/` klasörü (24 adet 212 baytlık macOS meta dosyası – veri değil, silinebilir) |
| Dosya adlandırma düzeni | `<konu><sürüm no>.xlsx` (örn. `goc1`, `goc2`, `kayit3`) |
| Formül / grafik / görsel / birleştirilmiş hücre | **Hiçbirinde yok** – tüm dosyalar düz değer tablosu |
| Gizli sayfa | Yok |
| Tarihlerin saklanışı | Hepsi **metin** (`gg.aa.yyyy`; `isim1`'de sıfırsız `g.a.yyyy`) – gerçek Excel tarihi değil |
| Varsayılan sayfa adı | `tablo` (istisnalar aşağıda) |

### Klasör yapısı

```
veri.zip
├── __MACOSX/                  (macOS artığı, önemsiz)
└── veri/
    ├── analiz1.xlsx   analiz2.xlsx           → boy-kilo-yaş / cinsiyet / gözlük
    ├── goc1.xlsx  goc2.xlsx  goc3.xlsx  goc4.xlsx   → iller arası göç kayıtları
    ├── isim1.xlsx                            → isim-doğum tarihi + isim→cinsiyet sözlüğü
    ├── kayit1.xlsx … kayit5.xlsx             → öğrenci/kişi kayıt tablosu (5 sürüm)
    ├── lig1.xlsx  lig2.xlsx  lig3.xlsx       → 38 oyunculu lig maç sonuçları
    ├── listeler.xlsx                         → kayit tablosunun genişletilmiş hâli
    ├── pilav1.xlsx                           → pilav/içecek siparişleri + fiyat listesi
    ├── stadyum1.xlsx                         → 100×100 koltuk haritası
    ├── tablolar1.xlsx                        → sayı tablosu + kod tablosu
    ├── test1.xlsx  test2.xlsx                → sınav cevapları + cevap anahtarı
    ├── vefat1.xlsx                           → doğum/vefat kayıtları
    └── yolcu1.xlsx                           → günlük yolcu sayıları (2010–2019)
```

### Dosya özet tablosu

| Dosya | Boyut | Sayfa(lar) | Veri satırı × sütun | Ana konu |
|---|---:|---|---|---|
| analiz1.xlsx | 81 KB | tablo | 3200 × 3 | Fiziksel ölçüler |
| analiz2.xlsx | 81 KB | tablo | 3200 × 3 | analiz1 ile **birebir aynı** |
| goc1.xlsx | 145 KB | tablo | 4800 × 4 | Göç (Kayıt no'lu) |
| goc2.xlsx | 161 KB | tablo | 4800 × 4 | Göç (Aile Kodu'lu) |
| goc3.xlsx | 161 KB | tablo | 4800 × 4 | goc2 ile **birebir aynı** |
| goc4.xlsx | 161 KB | tablo | 4800 × 4 | goc2'den 3 hücre farklı |
| isim1.xlsx | 144 KB | tablo, cinsiyet | 4500 × 2 ve 453 × 2 | İsim/doğum tarihi, cinsiyet sözlüğü |
| kayit1.xlsx | 731 KB | tablo | 6100 × 13 | Kayıt tablosu v1 |
| kayit2.xlsx | 722 KB | tablo | 6100 × 13 | Kayıt tablosu v2 |
| kayit3.xlsx | 730 KB | tablo | 6100 × 13 | Kayıt tablosu v3 |
| kayit4.xlsx | 730 KB | tablo | 6100 × 13 | kayit3 ile **birebir aynı** |
| kayit5.xlsx | 731 KB | tablo | 6100 × 13 | kayit3'ün doğum tarihi sütunu kaydırılmış hâli |
| lig1.xlsx | 23 KB | tablo | 703 × 3 | Lig maçları |
| lig2.xlsx | 23 KB | tablo | 703 × 3 | lig1'den 1 hücre farklı |
| lig3.xlsx | 23 KB | tablo | 703 × 3 | lig1'den 74 satır farklı |
| listeler.xlsx | 660 KB | veri_tablosu | 6131 × 13 | Kayıt tablosunun genişletilmişi |
| pilav1.xlsx | 87 KB | tablo, fiyatlar | 3000 × 3 ve küçük fiyat tablosu | Sipariş verisi |
| stadyum1.xlsx | 42 KB | koltuklar | 100 × 100 ızgara (başlıksız) | Koltuk doluluk haritası |
| tablolar1.xlsx | 41 KB | tablolar | 1000 satır × 16 sütun (başlıksız) | Sayı ve kod tabloları |
| test1.xlsx | 56 KB | tablo, cevap_anahtarı | 2600 × 3 ve 30 × 2 | Sınav |
| test2.xlsx | 56 KB | tablo, cevap_anahtarı | 2600 × 3 ve 30 × 2 | test1 ile **birebir aynı** |
| vefat1.xlsx | 116 KB | tablo | 2900 × 5 | Doğum/vefat |
| yolcu1.xlsx | 151 KB | tablo | 3652 × 4 | Günlük yolcu |

### Birebir aynı dosyalar (MD5 doğrulamalı)

| Grup | Dosyalar |
|---|---|
| 1 | `analiz1.xlsx` = `analiz2.xlsx` |
| 2 | `goc2.xlsx` = `goc3.xlsx` |
| 3 | `kayit3.xlsx` = `kayit4.xlsx` |
| 4 | `test1.xlsx` = `test2.xlsx` |

Yani 23 dosyadan yalnızca **19'u benzersiz içeriğe** sahiptir. Diğer sürüm dosyaları (goc, kayit, lig) aynı temel verinin küçük/büyük değişiklik uygulanmış kopyalarıdır; bu yüzden aşağıda her aile için "sürümler arası fark" ayrıca verilmiştir.

---

## 2. Dosya Raporları

### 2.1 `analiz1.xlsx` / `analiz2.xlsx`

**Yapı:** tek sayfa `tablo`, A1:C3201, 3200 veri satırı. İki dosya birebir aynıdır.

| Sütun | Tür | İçerik |
|---|---|---|
| `Cinsiyet` | metin | `ERKEK` (1957) / `KADIN` (1243) |
| `Gözlük` | metin | `YOK` (2118) / `VAR` (1082) |
| `Boy - Kilo - Yaş` | metin | Üç değer tek hücrede, `" - "` ile ayrılmış (örn. `163 - 108 - 20`); 3172 benzersiz değer |

**Özellikler:** Üçüncü sütun analiz öncesi boy / kilo / yaş olarak **bölünmelidir** (tamamı sorunsuz ayrışıyor, 0 hatalı satır).

| Ölçü | Min | Maks | Ortalama |
|---|---:|---:|---:|
| Boy (cm) | 150 | 200 | 174,81 |
| Kilo (kg) | 65 | 140 | 103,18 |
| Yaş | 20 | 70 | 44,63 |

- Cinsiyete göre ortalamalar: Erkek boy 175,14 / kilo 102,86 / yaş 44,88; Kadın boy 174,30 / kilo 103,68 / yaş 44,23. Fark çok küçüktür (kadın-erkek boy farkı gerçek hayattan çok daha az → veri büyük olasılıkla **sentetik/rastgele üretilmiş**).
- Gözlük kullanım oranı: Erkek %32,8 (642/1957), Kadın %35,4 (440/1243).
- Yaş dağılımı: 20–29: 652, 30–44: 946, 45–59: 943, 60–70: 659.
- Hesaplanan VKİ (BMI) ortalaması 34,5: obez 2035, fazla kilolu 600, normal 519, zayıf 46.
- **5 adet tekrar eden satır çifti** var (örn. `ERKEK / YOK / 187 - 78 - 35` iki kez).

---

### 2.2 `goc1.xlsx`, `goc2.xlsx`, `goc3.xlsx`, `goc4.xlsx` (Göç kayıtları)

**Ortak yapı:** tek sayfa `tablo`, 4800 satır × 4 sütun. 81 ilin tamamı hem `Nereden` hem `Nereye` için kullanılıyor.

| Sütun | Açıklama |
|---|---|
| 1. sütun | `goc1`'de **`Kayıt`** (1–4800 sıra no, tamsayı); `goc2/3/4`'te **`Aile Kodu`** (`Y-4365` biçimi: 1 büyük harf + tire + 4 rakam) |
| `Nereden` | Çıkış ili (81 benzersiz) |
| `Nereye` | Varış ili (81 benzersiz) |
| `Tarih` | Metin, `gg.aa.yyyy`; **01.01.2015 – 28.12.2019** arası, 719 farklı gün |

**Sürümler arası fark**

| Karşılaştırma | Sonuç |
|---|---|
| goc1 → goc2 | `Nereden`, `Nereye`, `Tarih` **tamamen aynı**; yalnızca `Kayıt` sütunu `Aile Kodu` ile değiştirilmiş |
| goc2 → goc3 | Birebir aynı |
| goc2 → goc4 | Yalnızca **3 hücre** farklı (hepsi `Nereye`, hepsi `Nereden = İstanbul` satırları): Excel satır 540 (C-4368) Adana→Ağrı, 594 (M-4383) Adana→Ağrı, 731 (G-4826) Giresun→Adıyaman |

**İstatistikler (goc1 üzerinden; diğerleri aynı)**

- Aynı ile göç (`Nereden = Nereye`): **0**.
- Yıllara göre kayıt: 2015: 949 · 2016: 994 · 2017: 930 · 2018: 953 · 2019: 974.
- En çok çıkış: Hakkari 79, Ankara 77, Burdur 77, Ordu 75, Van 72. En az: Diyarbakır 41, Çankırı 42, Sakarya 45.
- En çok varış: Sinop 79, Sakarya 73, Çorum 71, Gümüşhane 71, Kocaeli 70. En az: Kahramanmaraş 44, Denizli 45, Osmaniye 46.
- Net göç (varış − çıkış): en yüksek Sinop +29, Sakarya +28, Nevşehir +20; en düşük Samsun −21, Muğla −18, Burdur −15.
- 4800 kayıtta **3407 farklı güzergâh**; en sık güzergâhlar 5–6 kez tekrarlanıyor (örn. Kırıkkale → Ardahan: 6).
- Aylık dağılım: en yoğun Mayıs (469), Şubat (460), Ocak (456); en az Aralık (346), Mart (347).
- `Aile Kodu` sütunu (goc2/3/4): 4740 benzersiz kod → **59 kod birden fazla kullanılmış** (119 satır etkileniyor; en çok tekrar eden `P-2721` 3 kez). Aile kodu bu yüzden tekil anahtar olarak kullanılamaz.

---

### 2.3 `isim1.xlsx`

**Yapı:** iki sayfa.

**Sayfa `tablo` (A1:B4501, 4500 satır)**

| Sütun | İçerik |
|---|---|
| `AD - SOYAD` | Hepsi BÜYÜK HARF, her satırda tam 2 sözcük (ad + soyad); 4500 satırın tamamı benzersiz |
| `DOĞUM TARİHİ` | Metin, **sıfırsız** `g.a.yyyy` (örn. `6.1.2004`); yıl aralığı **1975–2005**; 3693 benzersiz tarih |

**Sayfa `cinsiyet` (A1:B454, 453 satır):** `İSİM` → `CİNSİYET` sözlüğü (228 KADIN, 225 ERKEK), her isim bir kez geçer, çok sözcüklü isim yoktur.

**Bulgular**

- `tablo` sayfasındaki her kişinin **ilk adı sözlükte bulunuyor (4500/4500 eşleşme)**; eşleştirilince 2276 erkek, 2224 kadın çıkıyor. Yani iki sayfa birbirini tamamlayacak şekilde tasarlanmış (cinsiyet için `DÜŞEYARA/VLOOKUP` benzeri bir eşleştirme alıştırması).
- Doğum on yılları: 1970'ler 794 · 1980'ler 1418 · 1990'lar 1450 · 2000'ler 838.
- Ortalama yaş (31.10.2026 baz alınarak) ≈ 36,6.
- En sık soyadlar: Kaya 47, Yılmaz 43, Öztürk 39, Şahin 33, Yıldız 33. (`YILMAZ` 25 kişide ad olarak da geçiyor.)
- Doğum ayı dağılımı yaklaşık düzgün (344–402 arası) → rastgele üretim izlenimi.

---

### 2.4 `kayit1.xlsx` … `kayit5.xlsx` (Kayıt tablosu, 5 sürüm)

**Ortak yapı:** tek sayfa `tablo`, A1:M6101 → **6100 satır × 13 sütun**. Tüm sürümlerde sütun adları aynıdır:

| # | Sütun | Tür | Açıklama |
|---|---|---|---|
| 1 | `Sıra No` | tamsayı | 1–6133 arası, 6100 benzersiz (33 numara boşta) |
| 2 | `Kayıt No` | tamsayı | 13–20877, 6100 benzersiz |
| 3 | `Ad ve Soyad` | metin | BÜYÜK HARF; 6021 benzersiz (79 satır başka bir kayıtla aynı isimde) |
| 4 | `Soyad` | metin | `Ad ve Soyad` değerinin son kısmıyla her satırda tutarlı; 3139 benzersiz |
| 5 | `Ana Adı` | metin | Anne adı |
| 6 | `Baba Adı` | metin | Baba adı |
| 7 | `Doğum Tarihi` | metin | `gg.aa.yyyy`, 1956–1995; 159 kayıt `01.01.xxxx` (bilinmeyen tarih yerine kullanılan klasik değer olabilir) |
| 8 | `Doğum Yeri` | metin | İlçe/il veya yurt dışı yerleşim (sürüme göre değişir) |
| 9 | `Güvenlik Kodu` | tamsayı | Sürüme göre 3–6 haneli |
| 10 | `Üniversite` | metin | 143 farklı üniversite (hiç değişmez); en sık Marmara (1444), Anadolu (951), İstanbul (334) |
| 11 | `Bölüm` | metin | 118 farklı bölüm/fakülte |
| 12 | `Mezuniyet Durumu` | metin | `Mezun` / `Son Sınıf Öğrencisi` |
| 13 | `Başarı Notu` | ondalık | ~15 haneli ondalık sayı, 0,50–3,90 |

Tüm sürümlerde `Sıra No`, `Kayıt No`, `Ad ve Soyad`, `Soyad`, `Üniversite` sütunları **birebir aynıdır**; kayıtlar aynı kişilerdir, değişen alanlar aşağıdadır. Satırlar ad-soyada göre alfabetik dizilmiş görünür.

**Sürüm bazında farklar**

| Özellik | kayit1 | kayit2 | kayit3 (= kayit4) | kayit5 |
|---|---|---|---|---|
| Ana Adı | Nadir/garip adlar, en sık değer 12 kez (Arife, Feyfure…) | Gerçekçi dağılım: Fatma 311, Ayşe 253, Emine 194, Hatice 147 | kayit2 ile aynı (1 hücre farkı) | kayit3 ile aynı |
| Baba Adı | kayit3'ten 15 hücre farklı | kayit3 ile aynı | – | aynı |
| Doğum Yeri | 847 değer; küçük ilçeler + yurt dışı kasabaları (Siverek, Eşme 16'şar) | 855 değer; İstanbul ağırlıklı (Üsküdar 320, Bakırköy 320, Kadıköy 288, İstanbul 239) | kayit2 ile aynı | aynı |
| Bölüm | Yaklaşık düz dağılım (en sık 68 kayıt) | Ağırlıklı: İlahiyat Fakültesi 646, Türk Dili ve Edebiyatı 481, İşletme 419, Beden Eğitimi ve Spor YO 369 | kayit2 ile aynı | aynı |
| Doğum Tarihi | Özgün | kayit1 ile aynı | **942 satırda değişmiş**: ay değerleri 02↔05 (05→02: 490, 02→05: 452) | kayit3'ün tarih sütunu **bir satır aşağı kaydırılmış** (satır *n*'e satır *n−1*'in tarihi gelmiş; ilk satıra son satırın tarihi) → tarihler yanlış kişilerle eşleşmiş; 6097 satır farklı |
| Güvenlik Kodu | 401–9600 (3–4 hane, ort. 4984) | 400–9600 (3–4 hane, ort. 5013) | **1283–999659** (4–6 hane; 6 haneli 5454, 5 haneli 600) | kayit3 ile aynı |
| Mezun oranı | %49,1 (2996 / 6100) | %48,8 (2979) | %49,8 (3037) | %49,8 (kayit3 ile aynı) |
| Başarı Notu | ort. 2,185 | ort. 2,228 | ort. 2,190 | kayit3 ile aynı |

Notlar:
- `Mezuniyet Durumu` ve `Başarı Notu` her sürümde yeniden üretilmiş (kayit2→kayit3'te Mezuniyet Durumu 3034 satırda, Başarı Notu tüm satırlarda farklı). Yani sürümler **mantıksal olarak ilerleyen kopyalar değil, yeniden rastgele doldurulmuş alanlar** içeriyor.
- `kayit4`, `kayit3`'ün dosya düzeyinde kopyasıdır.
- `kayit5` yapısal olarak kayit3 ile aynıdır; tek fark tarih sütunundaki satır kayması olduğundan, **"yanlış hizalanmış veriyi fark etme/düzeltme" senaryosu** olarak okunabilir.

---

### 2.5 `lig1.xlsx`, `lig2.xlsx`, `lig3.xlsx` (Lig maçları)

**Ortak yapı:** tek sayfa `tablo`, 703 satır × 3 sütun.

| Sütun | İçerik |
|---|---|
| `Oyuncu 1` | Ev sahibi/ilk oyuncu |
| `Oyuncu 2` | Rakip |
| `Skor` | Metin, `"a - b"` biçimi (örn. `3 - 6`) |

**Yapı bulguları (üç dosyada da geçerli):** 38 oyuncu; 38×37/2 = **703 farklı eşleşme** → her çift tam bir kez oynamış (**tek devreli round-robin**); her oyuncu 37 maç yapmış; kendisiyle oynayan yok. Skorların her iki tarafı da 0–6 aralığında; hiçbirinde 6'yı aşan değer yok. 193 maçta bir taraf 6'ya ulaşmış, 510 maçta kimse ulaşmamış. Beraberlik: 93. En sık skorlar: `2 - 0` ve `3 - 5` (24'er), `1 - 2` (23), `3 - 0` (22).

**Puan durumu (galibiyet 3, beraberlik 1 varsayımıyla):**

| Dosya | Lider(ler) | Sonuncu |
|---|---|---|
| lig1 | Yunus Özoğul **66** (21G-3B-13M) ve Mert Can Gökçe **66** (20G-6B-11M) eşit; 3. Berkan Serdar Kavas 65 | Berranur Altıntaş 40 (12G-4B-21M) |
| lig2 | lig1 ile aynı | lig1 ile aynı |
| lig3 | Mert Can Gökçe **69** (21G-6B-10M), 2. Yunus Özoğul 66 | lig1 ile aynı |

**Sürümler arası fark**

| Karşılaştırma | Fark |
|---|---|
| lig1 → lig2 | **1 satır**: Excel satır 185 – Başak Cemre Karaca / Serra Ecem Demirel skoru `0 - 1` yerine `0 - 2` |
| lig1 → lig3 | **74 satır**: (a) **72 satırda** `Başak Cemre Karaca` ile `Haluk Bekleyen` isimleri karşılıklı yer değiştirmiş; (b) Excel satır 7: Mert Can Gökçe – Muhammed Enes Kaynar `0 - 5` yerine `5 - 0` (sonuç tersine çevrilmiş → liderlik değişiyor); (c) Excel satır 185'te hem isim değişimi hem `0 - 2` skoru (lig2'deki değişiklik) var |

---

### 2.6 `listeler.xlsx`

**Yapı:** tek sayfa `veri_tablosu` (diğerlerinden farklı sayfa adı), A1:M6132 → **6131 satır × 13 sütun**. Sütunlar `kayit*` dosyalarıyla **aynı 13 başlık**.

**kayit ailesiyle ilişkisi**
- `kayit*` dosyalarındaki 6100 kaydın tamamı (`Sıra No` + `Kayıt No` anahtarıyla) burada da var; ek olarak **31 yeni kayıt** içeriyor (6131 = 6100 + 31).
- `Ad ve Soyad`, `Soyad`, `Baba Adı`, `Doğum Yeri`, `Üniversite`, `Bölüm` değerleri kayit2/kayit3 ile aynı; `Doğum Tarihi` kayit1/kayit2'deki (özgün) tarihlerle **%100 aynı**.
- Yeniden üretilmiş alanlar: `Güvenlik Kodu` (62472–927076, çoğu 6 haneli), `Mezuniyet Durumu` (**Mezun 3826 = %62,4**, Son Sınıf 2305), `Başarı Notu` (**1,00–4,00**, en çok 2 ondalık, 256 farklı değer, ort. **2,855**; kayit dosyalarındaki ~15 haneli notlardan farklı).

**Veri kalitesi bulguları**
- `Sıra No` 3942 iki farklı kişide kullanılmış (diğer tüm `Sıra No`'lar benzersiz; `Kayıt No` benzersiz).
- `Bölüm` sütununda **543 hücrede sondaki boşluk** var (örn. `"Sosyoloji Bölümü "`) → gruplama/eşleştirmede ayrı kategori sayılır.
- Yazım hatası örneği: `ÜSKÜDÜR` (muhtemelen `ÜSKÜDAR`).
- 46 farklı yurt dışı doğum yeri değeri tutarsız biçimlerde yazılmış (`STUTTGART(D)`, `SCHAFFHAUSEN-(CH)`, `DORTMUND/B. ALMANYA`, `VÖLKLİNGEN/F.ALMANYA` …).
- `Hemşirelik / Bölümü`, `Ebelik / Bölümü` gibi eğik çizgili bölüm adları (322 + 95 satır) kayit dosyalarında da geçer; standart dışı yazım.
- Satırlar ad-soyada göre sıralı görünüyor; 143 üniversite, 118 bölüm, 859 farklı doğum yeri var.

---

### 2.7 `pilav1.xlsx`

**Sayfa `tablo` (A1:C3001, 3000 satır)**

| Sütun | İçerik |
|---|---|
| `Sipariş No` | `V - 89015` biçimi (harf + `" - "` + 5 rakam); 3000 satırın hepsi benzersiz, ancak yalnızca sayı kısmı 46 kez tekrar ediyor (harf farklı) |
| `Tarih` | Yalnızca **ay adı** (Ocak…Aralık); yıl/gün yok |
| `Sipariş İçeriği` | `Pilav - İçecek` biçiminde tek hücrede (24 farklı kombinasyon = 6 pilav × 4 içecek) |

**Sayfa `fiyatlar` (A1:E14):** iki küçük fiyat tablosu, C sütunu boş ayırıcı; 8–14. satırlar boş.

| Pilav Çeşidi | Fiyat | | İçecek Çeşidi | Fiyat |
|---|---:|---|---|---:|
| Nohutlu Pilav | 16 | | Ayran | 5 |
| Salatalı Pilav | 16 | | Gazoz | 7,5 |
| Fasülyeli Pilav | 17,5 | | Maden Suyu | 8,5 |
| Tavuklu Pilav | 19 | | Şalgam | 10 |
| Kavurmalı Pilav | 23 | | | |
| Karışık Pilav | 28 | | | |

(Para birimi dosyada belirtilmemiş.)

**Hesaplanan sonuçlar (fiyat tablosu ile eşleştirilerek; eşleşmeyen sipariş yok)**
- Toplam ciro **83.487**, sipariş başı ortalama **27,83** (pilav kısmı 60.260,5; içecek kısmı 23.226,5).
- Pilav adetleri: Kavurmalı 541 · Karışık 530 · Salatalı 507 · Nohutlu 496 · Tavuklu 483 · Fasülyeli 443.
- İçecek adetleri: Ayran 757 · Maden Suyu 754 · Şalgam 746 · Gazoz 743.
- Pilav bazında ciro: Karışık 18.941,5 · Kavurmalı 16.592 · Tavuklu 12.974,5 · Salatalı 11.981 · Nohutlu 11.793 · Fasülyeli 11.205.
- İçecek bazında ciro: Şalgam 22.379,5 · Maden Suyu 21.611,5 · Gazoz 20.511,5 · Ayran 18.984,5.
- En yoğun aylar: Temmuz (288 sipariş / 8.005,5), Ağustos (272), Ekim (265); en düşük Kasım (220 / 6.103). En sık kombinasyonlar: Salatalı Pilav – Ayran (148), Kavurmalı Pilav – Gazoz (144).
- Analiz için `Sipariş İçeriği` sütunu pilav ve içecek olarak **bölünmelidir**.

---

### 2.8 `stadyum1.xlsx`

**Yapı:** tek sayfa `koltuklar`, A1:CV100 → **100 × 100 hücrelik ızgara**; **başlık satırı yok** (ilk satır da veridir). Değerler yalnızca üç karakter: `E`, `K`, `-`.

- Dış çerçeve (1. ve 100. satır, A ve CV sütunları) tamamen `-`; dolayısıyla içerik 98 × 98 = 9604 hücrelik iç alandır.
- Tüm sayfa: `E` 3151 · `K` 2120 · `-` 4729. İç alanda `-` 4333 → **5271 dolu hücre (%54,9)**.
- `E` : `K` ≈ %59,8 : %40,2. `E`/`K` büyük olasılıkla **Erkek/Kadın** (veya benzeri iki kategori), `-` boş koltuk/koridor anlamındadır (dosyada açıklama yok; yorum varsayımdır).
- Satır başına E: 22–44, K: 13–34. En çok E olan satır Excel satır 7 (44), en çok K olan satır 44 (34).
- Dört çeyrek dağılımı benzer: Üst-sol E 744/K 557, Üst-sağ E 811/K 571, Alt-sol E 801/K 482, Alt-sağ E 795/K 510 (kuruluş homojen).
- Komşuluk: yatay bitişik çiftlerde EE 986, KK 482, E-K karışık 1394; dikeyde EE 1024, KK 446, karışık 1369 (yani karma oturma yaygın).
- Tamamen boş iç satır/sütun yok.
- Araç notu: dosyayı `pandas.read_excel` ile okurken `header=None` kullanılmalı, aksi hâlde ilk satır sütun adı olur.

---

### 2.9 `tablolar1.xlsx`

**Yapı:** tek sayfa `tablolar`, A1:P1000, **başlıksız**; iki bağımsız blok + boş ayırıcı sütun.

| Blok | Alan | İçerik |
|---|---|---|
| Sayı bloğu | A–E sütunları, 1000 satır | 10–99 arası tamsayılar (boş yok). Sütun ortalamaları: 53,98 · 54,82 · 56,60 · 55,84 · 53,01; genel ort. 54,85; satır toplamı 94–453 (ort. 274,24) |
| Ayırıcı | F sütunu | Tamamen boş |
| Kod bloğu | G–P sütunları, yalnızca **ilk 50 satır** | 50 × 10 = **500 kod**, biçim `A200` (A–E harfi + 1–1000 arası sayı); 500'ü de benzersiz. Harf dağılımı: A 102 · B 89 · C 104 · D 115 · E 90 |

Kodlardaki sayı aralığı 1–1000 olduğundan, sayı bloğundaki satır sayısıyla (1000) eşleşecek şekilde tasarlanmış görünür (kod → satır numarası arama alıştırması gibi); bu eşleşmeye dair dosyada açık bir açıklama yoktur. Okurken `header=None` gerekir.

---

### 2.10 `test1.xlsx` / `test2.xlsx` (Sınav)

İki dosya birebir aynıdır.

**Sayfa `tablo` (2600 satır):**

| Sütun | İçerik |
|---|---|
| `ÖĞRENCİ NO` | 9 haneli (100211045 – 100220012), 100 öğrenci |
| `SORU NO` | 1–30 |
| `VERİLEN CEVAP` | A–E |

**Sayfa `cevap_anahtarı` (30 satır):** `SORU NO` → `DOĞRU CEVAP` (A 10 · D 7 · E 6 · C 5 · B 2).

**Bulgular**
- 100 öğrenci × 30 soru = 3000 olası cevaptan **2600 satır dolu** → 400 soru **boş bırakılmış** (boş cevaplar satır olarak yazılmamış; öğrenci başına 21–30 cevap, ort. 26).
- Satırlar öğrenci no, sonra soru no sırasıyla dizili; tekrar eden (öğrenci, soru) çifti yok.
- Cevap anahtarına göre: doğru ort. **18,70** (10–26), yanlış ort. 7,30 (2–14), boş ort. 4,0. Cevaplananlar içinde doğruluk %71,9; tüm olası cevaplara göre %62,3.
- En kolay sorular: 8 (%83,3), 14 (%81,9), 24 (%77,3). En zor: 5 (%63,6), 18 (%65,0), 16 (%66,3).
- Verilen cevap dağılımı: A 739 · D 559 · E 502 · C 479 · B 321 (anahtardaki dağılımla uyumlu şekilde A en sık, B en az).
- En iyi öğrenci: 100219040 (28 cevap, 26 doğru, 2 yanlış); en düşük: 100217003 (21 cevap, 10 doğru, 11 yanlış).

---

### 2.11 `vefat1.xlsx`

**Yapı:** tek sayfa `tablo`, 2900 satır × 5 sütun.

| Sütun | İçerik |
|---|---|
| `Cinsiyet` | KADIN 1459 / ERKEK 1441 |
| `Doğum Yeri` | 81 il |
| `Vefat Yeri` | 81 il |
| `Doğum Tarihi` | Metin `gg.aa.yyyy`, **1940–1960** (1940'lar 1378, 1950'ler 1393, 1960: 129) |
| `Vefat Tarihi` | Metin `gg.aa.yyyy`, **2005–2019** |

**Bulgular**
- Tarih ayrıştırma hatası yok; vefat tarihi doğumdan önce olan satır yok.
- Vefat yaşı (hesaplanan): min 44,1 · maks 79,6 · ort. **62,11** · medyan 62,25; Erkek ort. 62,04, Kadın ort. 62,17 (fark yok).
- Yaş bantları: 40–50: 173 · 50–60: 978 · 60–70: 1274 · 70–80: 475.
- Yıllık vefat sayısı 167 (2007) ile 221 (2018) arasında; en yoğun ay Ağustos (279), en az Şubat (212).
- Doğum yeri ile vefat yeri aynı olan yalnızca 37 kayıt (%1,3) → yer alanları birbirinden bağımsız rastgele atanmış.
- En sık doğum yerleri: Edirne 48, Aydın 47, Yozgat/Çankırı/Kırıkkale 46. En sık vefat yerleri: Bursa 48, Osmaniye 48, Eskişehir 46, Adana 46, Sinop 45.

---

### 2.12 `yolcu1.xlsx`

**Yapı:** tek sayfa `tablo`, 3652 satır × 4 sütun.

| Sütun | İçerik | Aralık |
|---|---|---|
| `Tarih` | Metin `gg.aa.yyyy` | 01.01.2010 – 31.12.2019, **kesintisiz günlük** (3652 gün; 29.02.2012 ve 29.02.2016 dahil) |
| `Yetişkin (Erkek)` | tamsayı | 750.068 – 1.648.971 |
| `Yetişkin (Kadın)` | tamsayı | 750.049 – 1.649.998 |
| `Çocuk` | tamsayı | 200.069 – 899.906 |

**Bulgular**
- Toplamlar: Erkek **4.374.442.116** · Kadın **4.386.076.941** · Çocuk **2.008.262.339** · Genel **10.768.781.396** (pay: %40,6 / %40,7 / %18,6).
- Günlük ortalama toplam yolcu: **2.948.735**. En düşük gün 03.03.2011 (1.773.931), en yüksek gün 27.11.2013 (4.173.259).
- Yıllık toplamlar 1,060–1,091 milyar arasında; en yüksek 2013 (1.090.650.577) ve 2017 (1.090.471.871), en düşük 2012 (1.060.174.644; artık yıl olmasına rağmen).
- Mevsimsellik veya hafta içi/sonu etkisi **yok**: aylık günlük ortalamalar 2,91–2,98 milyon, haftanın günlerine göre 2,92–2,98 milyon; üç sütun arasındaki korelasyon ≈ 0 (|r| < 0,02) → değerler büyük olasılıkla bağımsız düzgün dağılımla üretilmiş.

---

## 3. Dosyalar Arası İlişkiler

| İlişki | Açıklama |
|---|---|
| `isim1` → `cinsiyet` sayfası | Ad-soyad listesindeki ilk adlar cinsiyet sözlüğüyle %100 eşleşir |
| `kayit1…5` ↔ `listeler` | Aynı kişi kümesi (`Sıra No` + `Kayıt No`); listeler 31 ek kayıt içerir ve notları/kodları farklıdır |
| `goc1` ↔ `goc2/3/4` | Aynı 4800 göç olayı; yalnızca kimlik sütunu (Kayıt/Aile Kodu) ve goc4'te 3 hücre değişir |
| `test1` = `test2`, `analiz1` = `analiz2`, `kayit3` = `kayit4`, `goc2` = `goc3` | Birebir kopya |
| `lig1/2/3` | Aynı maç listesi; küçük skor/isim değişiklikleri |
| `pilav1` iki sayfası | `tablo` siparişleri, `fiyatlar` fiyat listesi (ciro hesabı için birleştirilir) |
| `test1` iki sayfası | `tablo` cevaplar, `cevap_anahtarı` doğru cevaplar (puanlama için birleştirilir) |

Dosyalar arasında ortak bir anahtarla birleştirilebilecek **bağımsız konu dosyaları** (analiz, yolcu, vefat, stadyum…) yoktur; her konu kendi içinde bir alıştırma/veri seti niteliğindedir.

---

## 4. Veri Kalitesi ve Dikkat Edilmesi Gerekenler

1. **Tarihler metindir**, Excel tarihi değildir. `isim1`'de sıfırsız (`6.1.2004`), diğer dosyalarda sıfırlı (`06.01.2004`) yazılıdır. Analizden önce dönüştürülmelidir.
2. **Tek hücrede birleşik değerler:** `analiz`: `Boy - Kilo - Yaş`; `pilav1`: `Sipariş İçeriği`; `lig`: `Skor`; `pilav1`: `Sipariş No` (`V - 89015`). Bölünmeden analiz edilemezler.
3. **Başlıksız dosyalar:** `stadyum1` ve `tablolar1` (ilk satır veridir).
4. **Tekrarlar:** analiz1'de 5 çift tekrar eden satır; goc2/3/4'te 59 tekrar eden `Aile Kodu`; kayit'lerde 79 aynı `Ad ve Soyad` (farklı kayıt numaralarıyla); listeler'de tekrar eden `Sıra No` 3942.
5. **Hizalama hatası:** `kayit5.xlsx` doğum tarihi sütunu bir satır kaymış durumda.
6. **Sürümler arası sessiz değişiklikler:** goc4 (3 hücre), lig2 (1 hücre), lig3 (74 satır), kayit1/2/3 (çeşitli alanlar) – hangi sürümün "doğru" olduğu dosyalardan anlaşılamaz.
7. **Metin temizliği:** `listeler.xlsx` `Bölüm` sütununda 543 hücrede sondaki boşluk; `ÜSKÜDÜR` yazım hatası; yurt dışı doğum yerleri tutarsız yazılmış.
8. **Büyük olasılıkla sentetik veri:** düzgün (uniform) dağılımlar, cinsiyetler arası neredeyse sıfır fark, bağımsız yer alanları ve korelasyonsuz yolcu sayıları nedeniyle veriler gerçek gözleme değil, rastgele üretilmiş görünüyor. Gerçek dünya çıkarımı yapılmamalıdır.
9. **Para birimi/ölçü birimi** (`pilav1` fiyatları, `analiz` boy/kilo) dosyalarda belirtilmemiş; yorum varsayımdır.

---

## 5. Yöntem

- Dosyalar `unzip` ile açıldı; MD5 ile birebir kopyalar tespit edildi.
- Her sayfa `openpyxl` ile yapısal (boyut, birleşik hücre, formül, grafik, gizlilik) ve `pandas` ile içerik açısından (veri tipi, boş değer, benzersizlik, aralıklar, tekrarlar) incelendi.
- Sürümler arası farklar hücre düzeyinde karşılaştırıldı (`goc`, `kayit`, `lig`, `listeler`).
- Puan durumu (lig) için galibiyet=3 / beraberlik=1 / mağlubiyet=0; sınav sonuçları için yalnızca doğru/yanlış/boş sayıları verilmiştir (negatif puanlama kuralı dosyada bulunmadığı için net hesabı yapılmamıştır).
- Tüm rakamlar kodla hesaplanmıştır; "büyük olasılıkla" / "görünüyor" ifadeleri dosyalarda açık bir açıklama bulunmayan yorumları belirtir.
