# ALU İşlem Günlüğü: Veri Seti ve Soru Planı Raporu

Oct 31, 2026 · @Kerem Ali AVCIOGLU

## 1. Ödev isterleri ve seçilen fikir

Seçilen fikir bir ALU (aritmetik mantık birimi) işlem günlüğüdür: 8 bitlik bir işlemcinin 5.000 işlemlik kayıt defteri. Hesaplar Python'da yalnızca NOT ve AND kapılarıyla yapılır, 10 soru ise Excel'de cevaplanır.

- **Teslim:** 9 Kasım 2026 Pazartesi, saat 21:00'e kadar; YouTube linki e-posta ya da WhatsApp ile gönderilir.
- **Video:** sesli, ekran kaydı, en fazla 30 dakika, YouTube'da "liste dışı".
- **Veri seti:** kapsamlı bir Excel dosyası, yapay zekayla oluşturulacak; hocanın eski verilerine benzemeyecek.
- **Sorular:** 10 adet, derste çözülen tipte olabilir; tekrar etmeyen sorular ve doğru bir ilerleyiş (kolaydan zora) beklenir.
- **Etki yüzdesi:** %5.

Özgünlük veri setinin konusunda ve yapısında, sorular ise hocanın eski sınav tarzında olacak: tek net cevap, gerektiğinde cevap formatı (X,XX gibi) belirtilmiş.

## 2. Sadece NOT ve AND mi?

Hesap motoru yalnızca NOT ve AND kapılarıyla çalışır, Excel'deki 10 soru ise normal Excel işlevleriyle cevaplanır.

- **Python tarafı:** Her işlem (TOPLA, ÇIKAR, OR, XOR) sadece NOT ve AND kapılarından kurulur. Bu mümkündür, çünkü NOT ve AND birlikte "tam" bir kapı kümesidir. OR = NOT(AND(NOT a, NOT b)); XOR bunun üzerine kurulur; toplayıcı da XOR ve AND'lerden oluşur.
- **Excel tarafı:** Excel kapı mantığını yapmaz, sadece analiz eder (`ÇOKEĞERSAY`, `DÜŞEYARA`, özet tablo gibi).
- **Kapı Sayısı gerçektir:** Kod her NOT/AND çağrısını sayar. 8 bitlik işlemler için sabit değerler:

| İşlem | Kapı sayısı | Nasıl çıkıyor |
| --- | --: | --- |
| NOT | 8 | bit başına 1 NOT |
| AND | 8 | bit başına 1 AND |
| OR | 32 | bit başına 3 NOT + 1 AND |
| XOR | 56 | bit başına 7 kapı |
| TOPLA | 160 | bit başına 20 kapı (tam toplayıcı) |
| ÇIKAR | 168 | A + NOT(B) + 1: 8 NOT + TOPLA'nın 160'ı |

Kesin kararlar: yalnızca **8 bit** kullanılır (Excel'in ikiliden onluğa çeviren işlevi en fazla 10 bitle çalıştığı için 16 bit sorun çıkarırdı) ve veri **5.000 satır** olur.

## 3. Veri setinin yapısı

Dosya `veri_alu.xlsx` olur ve 5 sayfadan oluşur; asıl veri 5.000 satırlık `islemler` sayfasındadır.

| Sayfa | İçerik | Satır |
| --- | --- | --: |
| `islemler` | Ana tablo, işlem günlüğü | 5.000 |
| `cekirdekler` | 8 çekirdeğin bilgisi: model, üretici, saat hızı (GHz) | 8 |
| `islem_turleri` | Her işlemin açıklaması ve hangi kapılardan kurulduğu | 6 |
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
| I | Sonuç (ikilik) | Kapılarla hesaplanmış |
| J | Taşma | EVET/HAYIR (TOPLA'da taşıma, ÇIKAR'da A\<B) |
| K | Taşıma Zinciri | Toplamada en uzun taşıma zinciri (bit) |
| L | Kapı Sayısı | Kodun sayaçla ölçtüğü gerçek değer |
| M | Süre (ns) | Kapı sayısına, taşıma zincirine ve çekirdek hızına bağlı, üzerine biraz gürültü |

Süre sütunu gerçekçi kurulur: zincir uzadıkça süre artar, hızlı çekirdek daha az ns harcar. Böylece 5., 9. ve 10. sorularda anlamlı bir ilişki çıkar.

**Soruların çalışması için üretim notları**

- Satırlar Tarih-Saat sırasına göre dizilir (8. soru için şart).
- Dağılım yaklaşık TOPLA %30, ÇIKAR %20, AND/OR/XOR %15'er, NOT %5 olur; TOPLA satırları yeterince çok olsun diye.
- Yaklaşık %1 satırda A ve B kasıtlı olarak eşit yapılır (1. soru için).
- Süre değerleri 0,1 ns hassasiyetinde üretilir; en uzun süre ve ilk 100 sınırı için eşitlik bırakılmaz.
- İkilik değerler metin olarak saklanır, böylece baştaki sıfırlar kaybolmaz.
- Cevaplar tek sonuç vermeli: 4. sorudaki en çok tekrar eden değer için eşitlik çıkarsa kod veriyi yeniden üretir.

## 4. Nasıl hazırlayacağız?

Veri setini bir Python betiği üretir, doğruluğunu da yine kod denetler.

1. `alu_uret.py` yazılır: NOT/AND fonksiyonları, toplayıcı, kapı sayacı, sabit rastgele tohum.
2. Kod `veri_alu.xlsx` dosyasını üretir (5 sayfa).
3. Kod kendi sonuçlarını Python'un normal toplama/çıkarmasıyla kıyaslar; tek bir fark bile varsa durur.
4. Cevap anahtarı Python ile ayrıca hesaplanır. Videoda Excel'de bulunan sonuç bununla birebir çıkmalıdır.

Hocanın "yapay zekayla Excel oluştur" şartı için en güvenlisi, videoda aynı isteği Claude'a verip dosyanın üretilmesini göstermektir. Hazırlığı şimdi yapılır, kayıtta temiz bir prompt'la tekrarlanır.

## 5. On soru (kolaydan zora, her biri farklı teknik)

Sorular hocanın eski sınav tarzında yazıldı: tek net cevap, gerektiğinde cevap biçimi belirtilmiş, koşullar birleşik. Her sorunun tekniği farklıdır, böylece soru tekrarı olmaz.

| # | Zorluk | Soru | Teknik | Cevap biçimi |
| --: | --- | --- | --- | --- |
| 1 | Kolay | A ve B ondalık değerleri birbirine eşit olan kaç işlem vardır? (NOT işlemleri hariç) | Koşullu sayma, iki sütunu karşılaştırma | Tam sayı |
| 2 | Kolay | Cumartesi ya da Pazar günü yapılan TOPLA işlemlerinden kaçında taşma olmuştur? | `HAFTANINGÜNÜ`, `ÇOKEĞERSAY` | Tam sayı |
| 3 | Kolay-orta | Süresi en uzun olan işlemi gerçekleştiren çekirdeğin modeli nedir? | `MAK`, `KAÇINCI`, `İNDİS`, `DÜŞEYARA` (iki sayfa) | Model adı |
| 4 | Orta | Sonuç (ikilik) sütununda en çok tekrar eden değer hangisidir? | Özet tablo ya da `EĞERSAY` ile mod bulma (metin üzerinde) | 8 karakterli ikilik metin |
| 5 | Orta | Süresi en uzun ilk 100 işlemin kaçı 5 numaralı çekirdekte yapılmıştır? | `BÜYÜK` ile sıralama sınırı | Tam sayı |
| 6 | Orta-zor | Hem A hem B ikilik değerinde tam dört adet 1 bulunan kaç işlem vardır? | `UZUNLUK` − `YERİNEKOY` (metin sayma) | Tam sayı |
| 7 | Orta-zor | Gece (00:00–05:59) yapılan TOPLA işlemlerinin yüzde kaçı taşma ile sonuçlanmıştır? | `SAAT`, `ÇOKEĞERSAY`, yüzde | % XX,XX |
| 8 | Zor | İşlemler Tarih-Saat sırasıyla okunduğunda, hemen önceki işlemle hem aynı çekirdekte hem aynı türde olan kaç işlem vardır? | Kaydırmalı satır karşılaştırma, `TOPLA.ÇARPIM` | Tam sayı |
| 9 | Zor | İspat: Sonuç (ikilik) değeri onluğa çevrildiğinde Excel'in kendi hesabından (TOPLA için MOD(A+B;256), ÇIKAR için MOD(A−B;256)) farklı çıkan kaç işlem vardır? | İkiliden onluğa çevirme, `EĞER`, `MOD` | Tam sayı (0 beklenir) |
| 10 | En zor | Süre ÷ Kapı Sayısı (ns/kapı) ortalaması en yüksek olan çekirdek–işlem türü ikilisi hangisidir ve değeri kaçtır? | Hesaplanan sütun + özet tablo (çekirdek × tür) | Çekirdek n – TÜR, X,XX |

**Excel formül iskeletleri** (sütun harfleri `islemler` sayfasındaki düzene göre; yardımcı sütunlar N'den başlayarak eklenir, Türkçe Excel'de ayraç `;`)

```text
S1  =TOPLA.ÇARPIM((D2:D5001<>"NOT")*(E2:E5001=F2:F5001))
S2  N2: =HAFTANINGÜNÜ(B2;2)>=6
    =ÇOKEĞERSAY(D2:D5001;"TOPLA";J2:J5001;"EVET";N2:N5001;DOĞRU)
S3  =DÜŞEYARA(İNDİS(C2:C5001;KAÇINCI(MAK(M2:M5001);M2:M5001;0));cekirdekler!A2:D9;2;YANLIŞ)
S4  Özet tablo: Satır = Sonuç (ikilik), Değer = Sayı, büyükten küçüğe sırala
S5  =ÇOKEĞERSAY(M2:M5001;">="&BÜYÜK(M2:M5001;100);C2:C5001;5)
S6  O2: =UZUNLUK(G2)-UZUNLUK(YERİNEKOY(G2;"1";""))   P2: aynısı H2 için
    =ÇOKEĞERSAY(O2:O5001;4;P2:P5001;4)
S7  Q2: =SAAT(B2)
    =ÇOKEĞERSAY(D2:D5001;"TOPLA";J2:J5001;"EVET";Q2:Q5001;"<6")/ÇOKEĞERSAY(D2:D5001;"TOPLA";Q2:Q5001;"<6")*100
S8  =TOPLA.ÇARPIM((C3:C5001=C2:C5000)*(D3:D5001=D2:D5000))
S9  R2: =EĞER(D2="TOPLA";MOD(E2+F2;256);EĞER(D2="ÇIKAR";MOD(E2-F2;256);""))
    S2: =EĞER(R2="";0;EĞER(BIN2DEC(I2)<>R2;1;0))      sonra =TOPLA(S2:S5001)
S10 T2: =M2/L2   Özet tablo: Satır = Çekirdek No, Sütun = İşlem Türü, Değer = T ortalaması
```

İkiliden onluğa çeviren işlevin Türkçe adı farklıdır; fx penceresinde "BIN" ya da "ikili" yazıp aratın. Cevap anahtarındaki sayılar veri üretildikten sonra Python ile hesaplanır, bu yüzden bu tabloda kasıtlı olarak sayısal cevap yoktur.

## 6. Soruların hocanın eski sorularıyla ilişkisi

Her yeni soru, hocanın eski sınavlarındaki bir beceriyi ölçer ama konusu ALU'ya özgüdür; yani soru kopyası değil, aynı tarzda yeni soru.

| Bizim soru | Hocanın eski sorusu (örnek) | Ölçülen beceri |
| --: | --- | --- |
| 1 | "Sıra No'su ile Kayıt No'su aynı olan kaç kişi?" | İki sütunu karşılaştırıp sayma |
| 2 | "Kaç kadın hafta sonu dünyaya gelmiştir?" | Tarihten haftanın günü + koşullu sayma |
| 3 | "En uzun isme sahip kişinin doğum yeri neresidir?" | En büyük değerden başka bir alana ulaşma |
| 4 | "En çok tekrar eden soyadı / güvenlik kodu hangisidir?" | Mod bulma |
| 5 | "Başarı notuna göre ilk 100'e girenlerden kaç Şişli doğumludur?" | Sıralama sınırı + koşul |
| 6 | "Güvenlik kodunun ilk rakamı çift olan kaç kişi?" | Metin parçalama / rakam sayma |
| 7 | "Marmara Üniversitesi olanların yüzde kaçı mezun olmuştur?" | Yüzde ve biçimli cevap |
| 8 | "Bir önceki güne göre yolcu sayısı artan kaç gün var?" | Bir önceki satırla karşılaştırma |
| 9 | Test dosyasında cevap anahtarıyla karşılaştırıp doğru/yanlış sayma | İki kaynağı çapraz doğrulama |
| 10 | "En başarılı ikinci bölüm hangisidir?", "mezun sayısı fazla olan kaç üniversite?" | Gruplama, özet tablo, karşılaştırma |

**Kalite ölçütleri**

- Her sorunun tek bir net cevabı vardır; eşitlik riski (aynı en büyük değer gibi) veri üretiminde kod tarafından engellenir.
- Yardımcı sütun gerektiren sorular (2, 6, 7, 9, 10) hocanın "önce sütunu böl" alışkanlığına benzer; 1, 3, 5 ve 8 doğrudan formülle çözülür.
- Zorluk 1–10 arasında artar; her soru farklı bir işleve dayanır.
- İsteğe bağlı: hocanın 0/1 katılım listesine benzer bir `aktiflik` sayfası (8 çekirdek × 5 hafta, çalıştı = 1) eklenebilir ve yedek bir 11. soru çıkarılabilir. Gerek görürsen söyle.

## 7. Video akışı (hedef 28 dk, sınır 30)

Video yaklaşık 28 dakika sürer; soruların zorluğu arttıkça ayrılan süre de artar.

| Süre | Bölüm |
| --- | --- |
| 0:00–1:00 | Kendini tanıt, ödevi ve konuyu (ALU işlem günlüğü) söyle |
| 1:00–4:00 | NOT ve AND doğruluk tabloları; OR, XOR ve toplayıcının bunlardan nasıl kurulduğu (`dogruluk_tablosu` sayfası) |
| 4:00–8:00 | Claude'a verilen prompt, üretilen Excel'in sayfaları ve sütunları |
| 8:00–9:00 | Hızlı kontrol: satır sayısı, veri tipleri, tarihlerin gerçek tarih olması |
| 9:00–12:00 | Soru 1, 2, 3 (kolay, yaklaşık 1 dk'şar) |
| 12:00–15:00 | Soru 4, 5 (orta, 1,5 dk'şar) |
| 15:00–19:00 | Soru 6, 7 (orta-zor, 2 dk'şar) |
| 19:00–24:00 | Soru 8, 9 (zor, 2,5 dk'şar) |
| 24:00–27:00 | Soru 10 (en zor, özet tablo ve yorum) |
| 27:00–28:00 | Kapanış: ne öğrendin, veri setinin özgün yanı |

**Her soruda izlenecek yöntem:** formülü ya da özet tabloyu kur, sonucu göster, neden o işlevi seçtiğini tek cümleyle söyle. Cevap anahtarıyla karşılaştırma kayıttan önce, hazırlıkta yapılır.

Video sesli olmalı, ekranı kaydederken mikrofon açık kalmalı. YouTube'a yüklerken görünürlük **"Liste dışı"** seçilir; link gizli sekmede açılıp test edilir. Dilersen 10. sorunun sonucunu bir sütun grafiğiyle de göstererek videoyu görsel olarak zenginleştirebilirsin.
