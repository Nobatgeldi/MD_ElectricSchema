# Cami Elektrik Dağıtım Projesi — çizim paketi üreteci

400/230 V, 50 Hz, TN-S sistemli, 250 kVA trafo ile beslenen cami binasının elektrik dağıtım projesi.
Tüm çizim, tablo ve rapor tek bir Python dosyasından (`proje_uret.py`) üretilir; tüm girdiler dosyanın
başındaki **`PARAMETRELER`** sözlüğündedir. Bir değer değişince çizimler, tablolar ve rapor birlikte güncellenir
(poz numaraları üç çizimde ve tablolarda aynıdır).

## Üretilen dosyalar (`cikti/`)

| Dosya | İçerik |
|---|---|
| `01_tek_hat_semasi.svg/.pdf` | Tek hat şeması (A3 yatay, IEC 60617 sembolleri, faz renkleri, çıkış tabloları, antet) |
| `02_pano_on_gorunus.svg/.pdf` | Pano ön görünüş + iç yerleşim (1:10), DETAY-A (1:5), DIN ray modül hesabı, ısıl not, poz listesi |
| `03_topraklama_ve_yildirimdan_korunma.svg/.pdf` | Kat planı (1:150), bina kesiti (1:200), yuvarlanan küre (1:500), detaylar, EPB şeması |
| `04_hesap_tablolari.xlsx` | Yük Dengesi, Kablo Çizelgesi, Pano Malzeme Listesi, Topraklama ve Yıldırım Malzeme, Kesici Selektivite (formüllü) |
| `05_proje_raporu.pdf` | Kapak, varsayımlar, hesap özeti, 3 çizimin küçük görüntüsü, notlar, poz tutarlılık matrisi |

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
* `hatlar`: her hattın gücü, fazı, cosφ, eşzamanlılık katsayısı, kablo kesiti/tipi, **uzunluğu** (klima için varsayılan 25 m).
* `cihazlar`: kesici/RCD/RCBO/SPD/analizör değerleri (Q0 Ir, Icu, RCR0 ayarı, K1-K6 MPCB ayarı...).
* `pano`: **boyutlar (G x Y x D)**, modül genişlikleri, DIN ray satırları. Modüller sığmazsa uyarı verilir ve pano
  standart ölçülerden bir üst boyuta büyütülür (çizimler yeni boyutla yeniden üretilir).
* `topraklama`: bina ölçüleri, **minare/kubbe konumu**, halka ofseti/derinliği, çubuk boyu/aralığı, **yıldırım koruma seviyesi** (I-IV), yüksekliler.

Ondalık ayracı çizim ve raporlarda virgüldür. Excel hücrelerinde sayılar sayısal tutulur; virgül/nokta gösterimi
bilgisayarın bölgesel ayarına bağlıdır (Türkçe ayarda virgül görünür).

## Önemli not

Bu paket verilen veri ve raporun “Varsayımlar” bölümündeki kabullerle üretilmiştir. Uygulamadan önce **yetkili elektrik
mühendisi ve yerel şebeke (dağıtım şirketi) şartnamesi ile onaylanmalıdır.**
