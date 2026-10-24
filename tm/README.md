# Metjidiň Elektrik Paýlaýyş Taslamasy — çyzgy paketini öndüriji

400/230 V, 50 Hz, TN-S ulgamly, 250 kVA transformator bilen iýmitlenýän metjit binasynyň elektrik paýlaýyş taslamasy.
Ähli çyzgy, tablisa we hasabat ýekeje Python faýlyndan (`proje_uret.py`) öndürilýär; ähli giriş maglumatlary faýlyň
başyndaky **`PARAMETRELER`** sözlüginde. Bir baha üýtgände çyzgylar, tablisalar we hasabat bilelikde täzelenýär
(poz belgileri dört çyzgyda we tablisalarda birmeňzeşdir).

Bu papka (`tm/`) paketiň **Türkmen dilindäki nusgasydyr**; türk dilindäki esasy nusga `cikti/` papkasynda. Iki nusga hem şol bir hasaplardan alynýar.

## Öndürilýän faýllar (`tm/`)

| Faýl | Mazmuny |
|---|---|
| `01_tek_hat_semasi.svg/.pdf` | Bir liniýaly shema (A3 keseligine, IEC 60617 nyşanlary, faza reňkleri, çykyş tablisalary, çarçuwa; daşky ýagtylandyryş dolandyryşy, ses güýçlendirijisi AMF1 we SK ses liniýalary bilen) |
| `02_pano_on_gorunus.svg/.pdf` | Panel öň görnüşi + içki ýerleşiş (1:10), DETAL-A (1:5), DIN reýka modul hasaby, ýylylyk bellik, poz listi |
| `03_topraklama_ve_yildirimdan_korunma.svg/.pdf` | Gat meýilnamasy (1:150), bina kesimi (1:200), ýumlanýan sfera (1:500), detallar, EPB shemasy |
| `04_hesap_tablolari.xlsx` | Ýük deňagramlylygy (geljekki goýum, açgöz faza paýlanyşy), Kabel tablisasy (ΔU formulaly), Panel material listi, Ýagtylandyryş we rozetka enwentary, Ses ulgamynyň ýygnamasy, Zazemleniýe we ýyldyrym materialy, Awtomat saýlamalylygy (formulaly) |
| `05_proje_raporu.pdf` | Gapak, çaklamalar, hasap ýygnamasy, çyzgylaryň kiçi şekilleri, bellikler, poz yzygiderlilik matrisasy |
| `06_aydinlatma_priz_yerlesim.svg/.pdf` | Ýagtylandyryş, rozetka we ses ulgamynyň ýerleşiş meýilnamasy (A3 keseligine, 1:100): rozetka P, içki LED L, lýustra AVZ, daşky projektor PR, dinamik HP, güýçlendiriji AMF1, ses şahalary SS we SK kabelleri, kabel lotogy, liniýa/faza/reňk tablisasy, ses ýygnama tablisasy |

Faýl atlary türk nusgasy bilen birmeňzeş saklanýar (poz belgileri we faýl atlary üýtgemeýär).

## Nähili täzeden öndürilýär

```bash
pip install svgwrite cairosvg openpyxl reportlab pillow numpy pypdf
python3 proje_uret.py
```

* Skript `cikti/` (türkçe) we `tm/` (türkmençe) papkalaryny öndürýär/üstüne ýazýar. Diňe türkçe nusga üçin: `TM_URET=0 python3 proje_uret.py`.
* Terjime `tm_sozluk.json` sözlüginden alynýar: tekstdäki sanlar `{0}`, `{1}` ýaly ýer tutujylara öwrülip sözlükde gözlenýär. Terjimesi ýok tekst tapylsa, ol `tm_eksik.json` faýla ýazylýar we türkçe galýar.
* Ýazuw şrifti: Liberation Sans (Arial bilen sazlaşykly) ulanylýar; PDF-lere gömülýär.
* Işiň ahyrynda awtomatiki barlag edilýär: poz yzygiderliligi (SVG ↔ tablisalar; tm çyzgylarynyň poz toplumy türkçe bilen deň), xlsx bahalary ↔ Python modeli (tm xlsx-iň sanly öýjükleri türkçe bilen deň),
  PDF-lerde A3 keseligine/ýeke sahypa, DIN reýka modul sygymy. Ýalňyşlyk bolsa konsola ýazylýar.

## Parametrleri üýtgetmek

`proje_uret.py` içindäki `PARAMETRELER` sözlügi: `sistem`, `hatlar`, `aydinlatma_priz` (rozetka, lýustra, içki LED, daşky projektor, ses), `cihazlar`, `pano`, `topraklama`. Jikme-jiklik üçin türkçe `README.md`-e seret.

Ondalyk aýryjy çyzgylarda we hasabatlarda otur (,) bolup durýar.

## Ses ulgamy

Ses ulgamy bir güýçlendiriji (AMF1) we üç dinamik şahasyndan ybarat; ähli bahalar `PARAMETRELER["aydinlatma_priz"]["ses"]` astynda.

| Poz | Manysy |
|---|---|
| `AM1` / `AMF1` | Güýçlendirijiniň iýmitlendiriş liniýasy (panelde aýratyn RCBO) we güýçlendiriji (elektrik otagynda PN1-iň ýanynda, ýer `amfi.konum_m`) |
| `SS1`–`SS3` | Dinamik şahalarynyň güýç liniýalary: sag, sol we minara şahasy (adaty 100 m, NYY 3x2,5). Sag + sol şahalar umumy `RCD3` (2P 25 A, 30 mA) aşagynda MCB C10, minara şahasy aýratyn RCBO bilen goralýar |
| `HP1`–`HP12` | Dinamikler (adaty şaha başyna 4 sany; şaha kuwwaty dinamiklere deň bölünýär) |
| `SK1`–`SK3` | Güýçlendiriji çykyşlaryny şahalara birikdirýän ekranly ses kabelleri (adaty 2x1,5 mm², şaha uzynlygy bilen deň) |

* **Çyzgy 1:** AM1 + AMF1 we SK1–SK3 güýçlendiriji çykyş liniýasy, RCD3 we şaha goraglary, HP pozlary.
* **Çyzgy 2:** SES hatary (RCD3, AM1, SS1–SS3) DIN reýka modullary we poz listi.
* **Çyzgy 3:** ses şahalarynyň zazemleniýesi: kabel içindäki PE damarlary (SS1–SS3 → PN1 BPE), güýçlendirijiniň şassisi we SK ekran şinasy `EB7`, minara ses topary `EB8` (EPB1-e 16 mm² Cu), SK ekranlarynyň diňe güýçlendiriji ujunda zazemlenmegi bellik.
* **Çyzgy 6:** SS şaha ugurlary, HP ýerleri, AMF1 ýeri, SK kabelleri we "Dinamik we kabel ýygnamasy" tablisasy.
* **Excel:** `Ses Ulgamynyň Ýygnamasy` sahypasy (dinamik, güýç/ses kabeli, ΔU we galyndy goşulan sargyt metraj formulalary); material listinde AM1, AMF1, SK1–SK3, HP1–HP12.
* **Hasabat:** 3.6 Ses ulgamynyň ýygnamasy tablisasy we "Çaklamalar" bölümi.
* **Çaklamalar:** güýçlendiriji kuwwaty (0,3 kW), şaha başyna kuwwatyň deň bölünmegi, dinamik başyna kuwwat, SK kabel görnüşi/uzynlyklary we güýçlendiriji ýeri. Güýçlendiriji çykyş görnüşi (liniýa naprýaženiýeli/pes empedans) we dinamik sazlaşygy öndürijiniň maglumaty bilen kesgitlenmelidir.

## Möhüm bellik

Bu paket berlen maglumat we hasabatyň "Çaklamalar" bölümündäki kabul edilenler bilen öndürilen. Amala aşyrmazdan öň **ygtyýarly elektrik inženeri we ýerli şebeke (paýlaýyş kompaniýasy) şertnamasy bilen tassyklanmalydyr.**

Terjimäni türkmen dilini we elektrik terminologiýasyny bilýän hünärmen gözden geçirmelidir (tehniki söz düzümleri awtomatiki sözlük arkaly ulanylýar).
