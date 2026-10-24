# Cami Elektrik Dağıtım Projesi — çizim paketi üreteci

400/230 V, 50 Hz, TN-S sistemli, 250 kVA trafo ile beslenen cami binasının elektrik dağıtım projesi.
Tüm çizim, tablo ve rapor tek bir Python dosyasından (`proje_uret.py`) üretilir; tüm girdiler dosyanın
başındaki **`PARAMETRELER`** sözlüğündedir. Bir değer değişince çizimler, tablolar ve rapor birlikte güncellenir
(poz numaraları dört çizimde ve tablolarda aynıdır).

## Üretilen dosyalar (`cikti/`)

| Dosya | İçerik |
|---|---|
| `01_tek_hat_semasi.svg/.pdf` | Tek hat şeması (A3 yatay, IEC 60617 sembolleri, faz renkleri, çıkış tabloları, antet; dış aydınlatma kumandası, ses amfisi AMF1 ve SK ses hatları dahil) |
| `02_pano_on_gorunus.svg/.pdf` | Pano ön görünüş + iç yerleşim (1:10), DETAY-A (1:5), DIN ray modül hesabı, ısıl not, poz listesi |
| `03_topraklama_ve_yildirimdan_korunma.svg/.pdf` | Kat planı (1:150), bina kesiti (1:200), yuvarlanan küre (1:500), detaylar, EPB şeması |
| `04_hesap_tablolari.xlsx` | Yük Dengesi (gelecek payı, açgözlü faz ataması), Kablo Çizelgesi (ΔU formüllü), Pano Malzeme Listesi, Aydınlatma ve Priz Envanteri, Ses Sistemi Özeti, Topraklama ve Yıldırım Malzeme, Kesici Selektivite (formüllü) |
| `05_proje_raporu.pdf` | Kapak, varsayımlar, hesap özeti, çizimlerin küçük görüntüsü, notlar, poz tutarlılık matrisi |
| `06_aydinlatma_priz_yerlesim.svg/.pdf` | Aydınlatma, priz ve ses sistemi yerleşim planı (A3 yatay, 1:100): priz P, iç LED L, avize AVZ, dış projektör PR, hoparlör HP, amfi AMF1, ses kolları SS ve SK kabloları, kablo tavası, hat/faz renk tablosu, ses özet tablosu |

## Nasıl yeniden üretilir

```bash
pip install svgwrite cairosvg openpyxl reportlab pillow numpy pypdf
python3 proje_uret.py
```

* Betik `cikti/` klasörünü (betiğin yanında) oluşturur/üzerine yazar.
* Yazı tipi: Liberation Sans (Arial uyumlu) kullanılır; PDF'lere gömülür. SVG'ler Arial/Liberation Sans ile doğru görüntülenir.
* LibreOffice kuruluysa (`soffice`), xlsx formül sonuçları önbelleğe alınır ve Python hesabıyla karşılaştırılır; yoksa Excel dosyayı açınca hesaplar.
* Çalışma sonunda otomatik doğrulama yapılır: poz tutarlılığı (SVG ↔ tablolar), xlsx değerleri ↔ Python modeli,
  PDF'lerin A3 yatay/tek sayfa olması, DIN ray modül sığma kontrolü. Hata varsa konsola yazılır ve çıkış kodu 1 olur.

## Parametreleri değiştirme

`proje_uret.py` içindeki `PARAMETRELER` sözlüğü:

* `sistem`: gerilim, trafo gücü, **trafo-pano mesafesi** (varsayılan 30 m), trafo uk/ur.
* `hatlar`: klima/kamera/ses/acil hatlarının gücü, fazı (`"auto"` = açgözlü atama), cosφ, eşzamanlılık katsayısı, kablo kesiti/tipi, **uzunluğu**.
* `aydinlatma_priz`: priz (adet, birim W, dağılım, **toplam_kablo_m**), avize, iç LED, dış projektör (`PROJEKTOR_GUCU_W`, adet, toplam kablo, `spd2_goster`),
  dekoratif aydınlatma (kW). Hat başına kablo boyu = toplam kablo / hat sayısı; ΔU şartı sağlanmazsa kesit otomatik büyütülür ve uyarı verilir.
* `cihazlar`: kesici/RCD/RCBO/SPD/analizör değerleri (Q0 Ir, Icu, RCR0 ayarı, K1-K6 MPCB ayarı...).
* `pano`: **boyutlar (G x Y x D)**, modül genişlikleri, DIN ray satırları. Modüller sığmazsa uyarı verilir ve pano
  standart ölçülerden bir üst boyuta büyütülür (çizimler yeni boyutla yeniden üretilir).
* `topraklama`: bina ölçüleri, **minare/kubbe konumu**, halka ofseti/derinliği, çubuk boyu/aralığı, **yıldırım koruma seviyesi** (I-IV), yüksekliler.

Ondalık ayracı çizim ve raporlarda virgüldür. Excel hücrelerinde sayılar sayısal tutulur; virgül/nokta gösterimi
bilgisayarın bölgesel ayarına bağlıdır (Türkçe ayarda virgül görünür).

## Ses sistemi

Ses sistemi bir amfi (AMF1) ile üç hoparlör kolundan oluşur; tüm değerler `PARAMETRELER["aydinlatma_priz"]["ses"]` altındadır.

| Poz | Anlamı |
|---|---|
| `AM1` / `AMF1` | Amfi besleme hattı (panoda ayrı RCBO) ve amfi (elektrik odasında PN1 yanı, yer `amfi.konum_m`) |
| `SS1`–`SS3` | Hoparlör kolu güç hatları: sağ, sol ve minare kolu (varsayılan 100 m, NYY 3x2,5). Sağ + sol kollar ortak `RCD3` (2P 25 A, 30 mA) altında MCB C10, minare kolu ayrı RCBO ile korunur |
| `HP1`–`HP12` | Hoparlörler (varsayılan kol başına 4 adet; kol gücü hoparlörlere eşit bölünür) |
| `SK1`–`SK3` | Amfi çıkışlarını kollara bağlayan ekranlı ses kabloları (varsayılan 2x1,5 mm², kol boyuyla aynı) |

* **Çizim 1:** AM1 + AMF1 ve SK1–SK3 amfi çıkış hattı, RCD3 ve kol korumaları, HP pozları.
* **Çizim 2:** SES satırı (RCD3, AM1, SS1–SS3) DIN ray modülleri ve poz listesi.
* **Çizim 3:** ses kollarının topraklaması: kablo içi PE damarları (SS1–SS3 → PN1 BPE), amfi şasisi ve SK ekran barası `EB7`, minare ses grubu `EB8` (EPB1'e 16 mm² Cu), SK ekranlarının yalnız amfi ucunda topraklanması notu.
* **Çizim 6:** SS kol güzergâhları, HP konumları, AMF1 yeri, SK kabloları ve "Hoparlör ve kablo özeti" tablosu.
* **Excel:** `Ses Sistemi Özeti` sayfası (hoparlör, güç/ses kablosu, ΔU ve fire dahil sipariş metrajı formülleri); malzeme listesinde AM1, AMF1, SK1–SK3, HP1–HP12.
* **Rapor:** 3.6 Ses sistemi özeti tablosu ve "Varsayımlar" bölümü.
* **Parametreler:** `toplam_kw`, `kollar` (her kol için `ad`, `L`, `koruma`: `"ortak_rcd"` veya `"rcbo"`, `hoparlor_adet`), `amfi` (`konum_m`, `boyut_m`, `guc_kw`, `rcbo_In`, `kesit`, `L`, `sinyal_kablo`). Ortak RCD için en az iki kol gerekir; ortak RCD altındaki kollar aynı faza bağlanır.
* **Varsayımlar:** amfi gücü (0,3 kW), kol başına güç eşit bölme, hoparlör başına güç, SK kablo tipi/boyları ve amfi yeri. Amfi çıkış tipi (hat gerilimli/düşük empedans) ve hoparlör uyumu üretici verisiyle belirlenmelidir.

## Türkmen dili kopyası (`tm/`)

Betik aynı verilerden tüm çıktıların (01–06) Türkmen dilindeki kopyasını `tm/` klasörüne de üretir (dosya adları aynıdır).
Çeviri `tm_sozluk.json` sözlüğünden gelir: metindeki sayılar `{0}`, `{1}` yer tutucularına çevrilerek sözlükte aranır; çevirisi olmayan
kalıplar `tm_eksik.json` dosyasına yazılır (Türkçe kalır). Yalnız Türkçe çıktı için: `TM_URET=0 python3 proje_uret.py`.
Teknik terimlerin çevirisi, Türkmence ve elektrik terminolojisini bilen bir uzman tarafından gözden geçirilmelidir.

## Önemli not

Bu paket verilen veri ve raporun “Varsayımlar” bölümündeki kabullerle üretilmiştir. Uygulamadan önce **yetkili elektrik
mühendisi ve yerel şebeke (dağıtım şirketi) şartnamesi ile onaylanmalıdır.**
