# ALU İşlem Günlüğü: Veri Seti ve Soru Planı Raporu

Oct 31, 2026 · @Kerem Ali AVCIOGLU

## 1. Ödev isterleri ve seçilen fikir

Seçilen fikir bir ALU (aritmetik mantık birimi) işlem günlüğüdür: 8 bitlik bir işlemcinin 5.000 işlemlik kayıt defteri. Hesaplar Python'da yalnızca NOT ve AND kapılarıyla yapılır, 10 soru ise Excel'de cevaplanır.

- **Teslim:** 9 Kasım 2026 Pazartesi, saat 21:00'e kadar; YouTube linki e-posta ya da WhatsApp ile gönderilir.
- **Video:** sesli, ekran kaydı, en fazla 30 dakika, YouTube'da "liste dışı".
- **Veri seti:** kapsamlı bir Excel dosyası, yapay zekayla oluşturulacak; hocanın eski verilerine benzemeyecek.
- **Sorular:** 10 adet, derste çözülen tipte olabilir; tekrar etmeyen sorular ve doğru bir ilerleyiş (kolaydan zora) beklenir.
- **Etki yüzdesi:** %5.

Veri setinin hikâyesi: 8 çekirdekli bir işlemcinin 3 aylık işlem günlüğünde, aşırı ısınan çekirdeklerde bazı sonuçlar tek bir bit hatasıyla kaydedilmiştir. 10 soru, bu hatayı Excel'de adım adım ortaya çıkarır. Sorular hocanın eski sınav sorularına (sayma, en çok tekrar eden, hafta sonu, yüzde gibi) benzemez; bit düzeyinde doğrulama, zaman penceresi eşleme ve hata tespiti gibi ALU konusuna özgü tekniklere dayanır. Her sorunun tek net cevabı vardır.

## 2. Sadece NOT ve AND mi?

Hesap motoru yalnızca NOT ve AND kapılarıyla çalışır, Excel'deki 10 soru ise normal Excel işlevleriyle cevaplanır.

- **Python tarafı:** Her işlem (TOPLA, ÇIKAR, OR, XOR) sadece NOT ve AND kapılarından kurulur. Bu mümkündür, çünkü NOT ve AND birlikte "tam" bir kapı kümesidir. OR = NOT(AND(NOT a, NOT b)); XOR bunun üzerine kurulur; toplayıcı da XOR ve AND'lerden oluşur.
- **Excel tarafı:** Excel kapı mantığını yapmaz; sonuçları analiz eder ve doğrular.
- **Kapı Sayısı gerçektir:** Kod her NOT/AND çağrısını sayar. 8 bitlik işlemler için sabit değerler `islem_turleri` sayfasında da yer alır:

| İşlem | NOT sayısı | AND sayısı | Toplam kapı | Nasıl çıkıyor |
| --- | --: | --: | --: | --- |
| NOT | 8 | 0 | 8 | bit başına 1 NOT |
| AND | 0 | 8 | 8 | bit başına 1 AND |
| OR | 24 | 8 | 32 | bit başına 3 NOT + 1 AND |
| XOR | 32 | 24 | 56 | bit başına 4 NOT + 3 AND |
| TOPLA | 88 | 72 | 160 | bit başına 20 kapı (tam toplayıcı) |
| ÇIKAR | 96 | 72 | 168 | A + NOT(B) + 1: 8 NOT + TOPLA'nın 160'ı |

Hesap motoru her zaman doğru çalışır ve kendi sonucunu doğrular. Günlükteki hatalı kayıtlar motorun hatası değildir; ısınan çekirdeklerde sonradan oluşan bit hatalarıdır (bölüm 3).

Kesin kararlar: yalnızca **8 bit** kullanılır (Excel'in ikiliden onluğa çeviren işlevi en fazla 10 bitle çalıştığı için 16 bit sorun çıkarırdı) ve veri **5.000 satır** olur.

## 3. Veri setinin yapısı

Dosya `veri_alu.xlsx` olur ve 7 sayfadan oluşur; asıl veri 5.000 satırlık `islemler` sayfasındadır.

| Sayfa | İçerik | Satır |
| --- | --- | --: |
| `islemler` | Ana tablo, işlem günlüğü | 5.000 |
| `cekirdekler` | 8 çekirdeğin bilgisi: Çekirdek No, model, üretici, saat hızı (GHz) | 8 |
| `islem_turleri` | Her işlemin açıklaması, kullandığı NOT ve AND sayısı, beklenen sonucun Excel'deki karşılığı | 6 |
| `tarife` | Kapı başına enerji (nJ); başlangıç tarihine göre aralıklı tablo | 4 |
| `termal_olaylar` | Aşırı ısınma pencereleri: Çekirdek No, Başlangıç, Bitiş, Süre (dk) | 12 |
| `dogruluk_tablosu` | NOT, AND ve bunlardan kurulan OR, XOR, tam toplayıcı tabloları | \~30 |
| `aciklama` | Sütun sözlüğü ve veri setinin hikâyesi | – |

**`islemler` sütunları**

| Sütun | Sütun adı | Açıklama |
| --- | --- | --- |
| A | İşlem No | `ISL-00001` biçiminde |
| B | Tarih-Saat | Gerçek Excel tarihi, 1 Ağustos – 31 Ekim 2026 |
| C | Çekirdek No | 1–8 |
| D | İşlem Türü | TOPLA, ÇIKAR, AND, OR, XOR, NOT |
| E, F | A, B (ondalık) | 0–255 (NOT'ta B boş) |
| G, H | A, B (ikilik) | 8 karakterlik metin, örn. `01101001` |
| I | Sonuç (ikilik) | Kapılarla hesaplanmış; hatalı kayıtlarda tek bit ters dönmüş |
| J | Taşma | EVET/HAYIR (TOPLA'da işaretsiz taşıma, ÇIKAR'da A\<B) |
| K | Taşıma Zinciri | TOPLA'da en uzun taşıma zinciri (bit), diğer işlemlerde 0 |
| L | Kapı Sayısı | Kodun sayaçla ölçtüğü gerçek değer |
| M | Süre (ns) | Kapı sayısına, taşıma zincirine ve çekirdek hızına bağlı, üzerine biraz gürültü |

Süre şu modelle üretilir: (kapı sayısı × 0,04 + taşıma zinciri × 0,9) ÷ saat hızı (GHz) + gürültü; termal pencere içinde ×1,35. Böylece zincir uzadıkça süre doğrusal artar, hızlı çekirdek daha az ns harcar, ısınan çekirdek yavaşlar.

**Veriye kasıtlı konan yapılar (soruları taşıyan şey)**

| Yapı | Nasıl kurulur | Hangi soru |
| --- | --- | --: |
| Hatalı kayıt | Termal pencereye düşen işlemlerin yaklaşık %40'ında Sonuç'ta tek bit ters çevrilir; pencere dışında hata yoktur. Bozulan bit konumu eşit dağılmaz, bir konum ağırlıklıdır | 6, 7, 8 |
| Termal pencere | 12 olay, 30–90 dk; ilgili çekirdekte yoğun yük (pencere başına yaklaşık 12 işlem, toplam satırların yaklaşık %3'ü) | 4, 5, 7, 10 |
| Enerji tarifesi | 4 dönem; 4. dönem 12 Ekim'de başlar, yani Ekim ayı tek tarifeyle hesaplanamaz | 2 |
| Süre–zincir ilişkisi | Doğrusal, gürültülü | 10 |
| Sessiz çekirdek | Bir çekirdek yaklaşık 30 saat boyunca hiç işlem yapmaz | 9 |
| İşaretli sayılar | A ve B 0–255 aralığında düzgün dağılır; 128 ve üzeri değerler işaretli yorumda negatiftir | 3 |

**Üretim notları**

- Satırlar Tarih-Saat sırasına göre dizilir; İşlem No bu sıraya göre verilir.
- Arka plan işlemlerinin dağılımı yaklaşık TOPLA %30, ÇIKAR %20, AND/OR/XOR %15'er, NOT %5 olur.
- Süre değerleri 0,01 ns hassasiyetinde üretilir; en uzun süre gibi sorularda eşitlik bırakılmaz.
- İkilik değerler metin olarak saklanır, böylece baştaki sıfırlar kaybolmaz.
- Hatalı satırların listesi günlükte işaretlenmez; yalnızca cevap anahtarı dosyasında tutulur.
- En sık bozulan bit konumu ve en uzun bekleme gibi tek cevaplı sonuçlarda eşitlik çıkarsa kod veriyi yeniden üretir.

## 4. Nasıl hazırlayacağız?

Veri setini bir Python betiği üretir; motor önce doğrulanır, bit hataları ancak ondan sonra eklenir.

1. `alu_uret.py` yazılır: NOT/AND fonksiyonları, toplayıcı, kapı sayacı, sabit rastgele tohum.
2. Kod önce temiz günlüğü üretir ve her sonucu Python'un normal toplama, çıkarma ve bit işlemleriyle kıyaslar; tek bir fark bile varsa durur.
3. Sonra termal pencereler, yük işlemleri ve sessiz çekirdek eklenir; pencereye düşen işlemlere bit hataları işlenir ve bozulan satırların listesi ayrı tutulur.
4. Kod `veri_alu.xlsx` dosyasını (7 sayfa) ve `cevap_anahtari.xlsx` dosyasını yazar. Her cevap Python ile ayrıca hesaplanır; videoda Excel'de bulunan sonuç bununla birebir çıkmalıdır.
5. Kod, tek cevaplı soruların (en sık bozulan bit, en uzun bekleme, en uzun süre) eşitlik çıkarmadığını denetler; çıkarsa tohumu değiştirip yeniden üretir.

Hocanın "yapay zekayla Excel oluştur" şartı için en güvenlisi, videoda aynı isteği Claude'a verip dosyanın üretilmesini göstermektir. Hazırlığı şimdi yapılır, kayıtta temiz bir prompt'la tekrarlanır.

## 5. On soru (kolaydan zora, her biri farklı teknik)

Sorular tek bir hikâyeyi izler: önce veriyi tanı, sonra ısınan çekirdekleri bul, sonra hataları yakala. 6, 7 ve 8. sorular birbirine yaslanır; her soru ayrı bir tekniğe dayanır ve tek net cevabı vardır.

| # | Zorluk | Soru | Teknik | Cevap biçimi |
| --: | --- | --- | --- | --- |
| 1 | Kolay | Sonuç (ikilik) sütunundaki tüm kayıtlar birlikte düşünüldüğünde toplam kaç adet 1 biti vardır? | `TOPLA.ÇARPIM`, `UZUNLUK`, `YERİNEKOY` | Tam sayı |
| 2 | Kolay-orta | `tarife` sayfasındaki kapı başına enerjiye göre (Enerji = Kapı Sayısı × işlemin yapıldığı tarihteki tarife) Ekim 2026'daki tüm işlemlerin toplam enerjisi kaç nJ'dir? | Yaklaşık eşleşmeli `DÜŞEYARA`, `ÇOKETOPLA` | XXXXX,XX |
| 3 | Orta | A ve B değerleri 8 bitlik işaretli (2'ye tümleyen) sayı olarak yorumlanırsa kaç TOPLA işleminde işaretli taşma oluşur? (Taşma sütunu işaretsiz taşmayı gösterir, A ve B'ye bakınız.) | İşaretli dönüşüm, `EĞER`, `VEYA` | Tam sayı |
| 4 | Orta | Kaç işlem, kendi çekirdeğinin termal olay penceresi içinde yapılmıştır? | İki sayfadan aralık eşleme, `ÇOKEĞERSAY` | Tam sayı |
| 5 | Orta-zor | Tüm çekirdekler birlikte düşünüldüğünde, herhangi bir 60 dakikalık kesintisiz zaman diliminde yapılan en yüksek işlem sayısı kaçtır? (Pencere bir işlemin zamanında başlar.) | Kayan pencere, `ÇOKEĞERSAY`, `MAK` | Tam sayı |
| 6 | Orta-zor | Günlükteki Sonuç değeri, A ve B'den `islem_turleri` sayfasındaki kurallara göre hesaplanan beklenen sonuçtan farklı olan kaç kayıt vardır? | `BITAND`/`BITOR`/`BITXOR`, `MOD`, ikiliden onluğa çevirme, iç içe `EĞER` | Tam sayı |
| 7 | Zor | Termal olay penceresine düşen işlemlerin yüzde kaçı hatalı kayıttır? | 4. ve 6. sorunun sütunlarını birleştirip koşullu oran | % XX,XX |
| 8 | Zor | Hatalı kayıtlarda ters dönen bit incelendiğinde en sık bozulan bit konumu kaçtır? (0 = en sağdaki bit) | `BITXOR`, `LOG`, özet tablo | 0–7 arası tam sayı |
| 9 | Zor | Aynı çekirdeğin ardışık iki işlemi arasında geçen en uzun süre kaç saattir? | Grup içinde bir önceki zaman, `MAKSEĞER` | X,XX |
| 10 | En zor | 5 numaralı çekirdekte, termal pencere dışında yapılan TOPLA işlemlerinde taşıma zinciri 1 bit uzadığında süre ortalama kaç ns artar? | Regresyon eğimi `EĞİM`, koşullu dizi ya da filtre | X,XX |

**Excel formül iskeletleri** (`islemler` sayfasındaki düzene göre; yardımcı sütunlar N'den başlayarak eklenir; Türkçe Excel'de ayraç `;`)

```text
Soru 1   =TOPLA.ÇARPIM(UZUNLUK(I2:I5001)-UZUNLUK(YERİNEKOY(I2:I5001;"1";"")))
Soru 2   Q2: =L2*DÜŞEYARA(B2;tarife!$A$2:$B$5;2;DOĞRU)
         =ÇOKETOPLA(Q2:Q5001;B2:B5001;">="&TARİH(2026;10;1);B2:B5001;"<"&TARİH(2026;11;1))
Soru 3   N2: =EĞER(E2>127;E2-256;E2)    O2: =EĞER(F2>127;F2-256;F2)
         P2: =EĞER(D2="TOPLA";EĞER(VEYA(N2+O2<-128;N2+O2>127);1;0);0)    sonra =TOPLA(P2:P5001)
Soru 4   R2: =EĞER(ÇOKEĞERSAY(termal_olaylar!$A$2:$A$13;C2;termal_olaylar!$B$2:$B$13;"<="&B2;termal_olaylar!$C$2:$C$13;">="&B2)>0;1;0)
         =TOPLA(R2:R5001)
Soru 5   S2: =ÇOKEĞERSAY($B$2:$B$5001;">="&B2;$B$2:$B$5001;"<"&(B2+1/24))    sonra =MAK(S2:S5001)
Soru 6   T2: =EĞER(D2="TOPLA";MOD(E2+F2;256);EĞER(D2="ÇIKAR";MOD(E2-F2;256);EĞER(D2="AND";BITAND(E2;F2);
              EĞER(D2="OR";BITOR(E2;F2);EĞER(D2="XOR";BITXOR(E2;F2);255-E2)))))
         U2: =EĞER(BIN2DEC(I2)<>T2;1;0)    sonra =TOPLA(U2:U5001)
Soru 7   =ÇOKEĞERSAY(R2:R5001;1;U2:U5001;1)/ÇOKEĞERSAY(R2:R5001;1)*100
Soru 8   V2: =EĞER(U2=1;LOG(BITXOR(T2;BIN2DEC(I2));2);"")    Özet tablo: Satır = V (boşlar hariç), Değer = Sayı
Soru 9   W2: =EĞER(MAKSEĞER($B$2:$B$5001;$C$2:$C$5001;C2;$B$2:$B$5001;"<"&B2)=0;"";
              (B2-MAKSEĞER($B$2:$B$5001;$C$2:$C$5001;C2;$B$2:$B$5001;"<"&B2))*24)    sonra =MAK(W2:W5001)
Soru 10  (dizi formülü) =EĞİM(EĞER((C2:C5001=5)*(D2:D5001="TOPLA")*(R2:R5001=0);M2:M5001);
                              EĞER((C2:C5001=5)*(D2:D5001="TOPLA")*(R2:R5001=0);K2:K5001))
```

BITAND, BITOR, BITXOR ve BIN2DEC işlevlerinin Türkçe adları farklıdır; fx penceresinde "BIT" ve "BIN" yazıp aratın. `MAKSEĞER` Excel 2019 ve sonrasında vardır; yoksa veriyi önce Çekirdek No'ya, sonra Tarih-Saat'e göre sıralayıp bir önceki satırla çıkarma yapın. Soru 10 için dizi formülü yerine, yalnızca TOPLA işlemlerini, 5. çekirdeği ve pencere dışı satırları süzüp `=EĞİM(M;K)` de yazılabilir. Cevap anahtarındaki sayılar veri üretildikten sonra Python ile hesaplanır, bu yüzden bu tabloda kasıtlı olarak sayısal cevap yoktur.

## 6. Soruların hocanın eski sorularından farkı

Hocanın eski soruları ağırlıklı olarak koşullu sayma, en çok tekrar eden, ilk N, hafta sonu, yüzde ve bir önceki satırla karşılaştırma gibi tekniklere dayanır. Bu 10 soru ise bit düzeyi, zaman aralığı ve hata tespiti soruları içerir ve aynı tekniği ikinci kez kullanmaz.

| Teknik | Hocanın soru metinlerinde | Bizim sorularda |
| --- | --- | --- |
| Düz koşullu sayma | Çok (ör. "1990 yılında doğan kaç kişi?") | Yalnızca başka bir tekniğin çıktısı olarak (4, 6) |
| En çok tekrar eden / en uzun / ilk N | Var | Yalnızca hesaplanan sütunlar üzerinde (5, 8) |
| Hafta sonu, ay, yıl | Var | Yok |
| Yüzde | Var | 7. soru: iki hesaplanan sütunun birleşimi |
| Bir önceki satırla karşılaştırma | Var (yolcu dosyası) | 9. soru: çekirdek bazında, aralık süresi |
| Bit düzeyi işlemler | Yok | 1, 3, 6, 8 |
| Aralık tablosuyla eşleme, kayan pencere | Yok | 4, 5 |
| Yaklaşık eşleşmeli arama | Yok | 2 |
| Regresyon eğimi | Yok | 10 |
| Veriyi doğrulama, hata bulma | Yok | 6, 7, 8 |

**Kalite ölçütleri**

- Her sorunun tek bir net cevabı vardır; eşitlik riski (en sık bit, en uzun bekleme, en uzun süre) veri üretiminde kod tarafından denetlenir.
- Zorluk 1'den 10'a artar; 6, 7 ve 8. sorular aynı hata hikâyesini adım adım ilerletir.
- Her soru farklı bir işleve dayanır, yani tekrar yoktur.
- Çoğu soru derste görülen araçlarla (`ÇOKEĞERSAY`, `DÜŞEYARA`, özet tablo) çözülür; bit işlevleri ve `EĞİM` yenidir ama fx penceresinde bulunur.
- Risk: 5, 9 ve 10. sorular ağırdır (5.000 satırlık kayan formüller). Video süresi buna göre ayrıldı, hesaplama yavaşsa dosya önceden hazırlanıp videoda sonuç gösterilir.

## 7. Video akışı (hedef 28 dk, sınır 30)

Video yaklaşık 28 dakika sürer; ağır formüllü sorulara (5, 6, 9, 10) daha fazla süre ayrılır.

| Süre | Bölüm |
| --- | --- |
| 0:00–1:00 | Kendini tanıt, ödevi ve konuyu söyle: ısınan bir işlemcinin işlem günlüğünde bit hatası avı |
| 1:00–3:30 | NOT ve AND doğruluk tabloları; OR, XOR ve toplayıcının bunlardan nasıl kurulduğu (`dogruluk_tablosu` sayfası) |
| 3:30–7:00 | Claude'a verilen prompt; üretilen Excel'in 7 sayfası (özellikle `tarife` ve `termal_olaylar`) ve sütunları |
| 7:00–8:00 | Hızlı kontrol: satır sayısı, tarihlerin gerçek tarih olması, ikilik değerlerin metin olması |
| 8:00–10:00 | Soru 1, 2 (kolay, 1 dk'şar) |
| 10:00–11:30 | Soru 3 (işaretli taşma) |
| 11:30–13:00 | Soru 4 (termal pencere eşleme) |
| 13:00–15:30 | Soru 5 (60 dakikalık kayan pencere) |
| 15:30–18:00 | Soru 6 (beklenen sonuç ve hatalı kayıtlar) |
| 18:00–19:30 | Soru 7 (pencere içi hata oranı; 4. ve 6. soruyu birleştir) |
| 19:30–21:30 | Soru 8 (bozulan bit konumu) |
| 21:30–24:00 | Soru 9 (çekirdek içi en uzun bekleme) |
| 24:00–27:00 | Soru 10 (süre–zincir eğimi) |
| 27:00–28:00 | Kapanış: hatalı kayıtların tamamı termal pencerede çıktı; ne öğrendin, veri setinin özgün yanı |

**Her soruda izlenecek yöntem:** formülü ya da özet tabloyu kur, sonucu göster, neden o işlevi seçtiğini tek cümleyle söyle. Cevap anahtarıyla karşılaştırma kayıttan önce, hazırlıkta yapılır. Yardımcı sütunları (N–W) önceden hazırlayıp yalnızca fikri anlatmak süreyi rahatlatır, ama formülü bir kez ekranda yazmayı unutma.

Video sesli olmalı, ekranı kaydederken mikrofon açık kalmalı. YouTube'a yüklerken görünürlük **"Liste dışı"** seçilir; link gizli sekmede açılıp test edilir. 10. sorunun sonucunu bir dağılım grafiğiyle (zincir – süre) göstermek videoyu görsel olarak zenginleştirir.

## 8. Takvim ve hocaya sorulacaklar

Teslim 9 Kasım Pazartesi saat 21:00; bir günlük tampon için video en geç 8 Kasım'da YouTube'a yüklenmiş olmalı.

&#91;embedded content: ödev takvimi · 5 aşama, teslim 9 Kasım 21:00\]

Veri seti ve cevap anahtarı 1 Kasım'a kadar hazır olur; video 6–7 Kasım'da kaydedilir, 8 Kasım'da yüklenir.

**Hocaya sorulacaklar** (kısa bir mesajla netleşir):

- Dosyadaki bazı sonuçların kasıtlı olarak bozuk olduğu (bit hatalı) bir veri seti kabul edilir mi?
- Yardımcı sütun eklemek ve `BITAND`, `EĞİM`, `MAKSEĞER` gibi derste görülmemiş işlevleri kullanmak serbest mi?
- Cevap anahtarı yalnızca videoda mı gösterilsin, ayrıca dosya olarak da gönderilsin mi?
