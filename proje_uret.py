#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CAMİ ELEKTRİK DAĞITIM PROJESİ - ÇİZİM PAKETİ ÜRETECİ
=====================================================
Tek dosyadan şunları üretir (cikti/ klasörüne):
  01_tek_hat_semasi.svg/.pdf             (A3 yatay)
  02_pano_on_gorunus.svg/.pdf            (A3 yatay)
  03_topraklama_ve_yildirimdan_korunma.svg/.pdf (A3 yatay)
  04_hesap_tablolari.xlsx
  05_proje_raporu.pdf
Tüm değerler aşağıdaki PARAMETRELER sözlüğünden gelir. Bir değer değişince
çizimler, tablolar ve rapor birlikte güncellenir (aynı poz numaraları).

Çalıştırma:  python3 proje_uret.py
Gereksinim:  svgwrite, cairosvg, openpyxl, reportlab, pillow, numpy
             (xlsx formül değerlerini önbelleğe almak/doğrulamak için LibreOffice - isteğe bağlı)
"""
import os
import re
import sys
import math
import shutil
import copy
import subprocess
import tempfile
import datetime
import xml.etree.ElementTree as ET

import numpy as np
import svgwrite
import cairosvg
from PIL import ImageFont

# =====================================================================
#  PARAMETRELER  (tek doğruluk kaynağı)
# =====================================================================
PARAMETRELER = {
    # ------------------------------------------------------------------ proje / antet
    "proje": {
        "ad": "Cami Elektrik Dağıtım Projesi",
        "sistem_yazisi": "400 V TN-S",
        "tarih": "24.10.2026",
        "revizyon": "R0",
        "cizen": "proje_uret.py",
        "kontrol": "Yetkili Elk. Müh. (onay bekliyor)",
        "cizim_no": {1: "CEP-E-01", 2: "CEP-E-02", 3: "CEP-E-03", 6: "CEP-E-06"},
        "cizim_sayisi": 4,        # antette sayfa x/4 (çizim 01, 02, 03, 06)
        "cikti_klasoru": "cikti",
    },
    # ------------------------------------------------------------------ sistem / kaynak
    "sistem": {
        "U_LL": 400.0,            # V
        "U_LN": 230.0,            # V
        "frekans": 50,            # Hz
        "sebeke": "TN-S",
        "trafo_kva": 250.0,
        "trafo_grubu": "Dyn11",
        "trafo_uk_yuzde": 4.0,    # kısa devre gerilimi (varsayım)
        "trafo_ur_yuzde": 1.3,    # omik bileşen (varsayım)
        "trafo_pano_mesafe_m": 30.0,
        "ortam_sicakligi_C": 35.0,
        "pano_ic_sicaklik_max_C": 45.0,
        "cos_varsayilan": 0.90,   # hat bazında cosφ verilmemişse kullanılır
        "gelecek_payi_yuzde": 25.0,   # ana kesici için istenen gelecek payı
    },
    # ------------------------------------------------------------------ kablo hesap sabitleri
    "kablo": {
        "kappa": 44.0,            # m/(ohm*mm2)  Cu, 70 C çalışma sıcaklığı
        "X_mohm_m": 0.08,         # mohm/m  reaktans (tüm kesitler için sabit varsayım)
        "azaltma_k": 0.80,        # pano-alıcı arası kablolar için gruplama/sıcaklık azaltma katsayısı
        "azaltma_k_gomulu": 0.85, # gömülü besleme kablosu için (toprak ısıl direnci/gruplama)
        "limit_guc_yuzde": 3.0,
        "limit_aydinlatma_yuzde": 2.0,
        # Akım taşıma kapasiteleri (A) - IEC 60364-5-52 Ek B, PVC, Cu (yaklaşık; üretici kataloğu ile doğrulanmalı)
        "iz_C_2yuklu": {1.5: 17.5, 2.5: 24, 4: 32, 6: 41, 10: 57, 16: 76, 25: 101, 35: 125, 50: 151, 70: 192, 95: 232},
        "iz_C_3yuklu": {1.5: 15.5, 2.5: 21, 4: 28, 6: 36, 10: 50, 16: 68, 25: 89, 35: 110, 50: 134, 70: 171, 95: 207},
        "iz_D2_3yuklu": {1.5: 22, 2.5: 29, 4: 38, 6: 47, 10: 63, 16: 81, 25: 104, 35: 125, 50: 148, 70: 183, 95: 216},
    },
    # ------------------------------------------------------------------ besleme (trafo -> Q0)
    "besleme": {
        "kablo_no": "W0", "tip": "NYY", "damar": "4x95", "kesit": 95.0, "yontem": "D2",
        "pe_iletkeni": {"poz": "PE3", "kesit": 50.0, "tip": "NYY (sarı-yeşil) 1x50 mm² Cu"},
    },
    # ------------------------------------------------------------------ koruma / ölçüm cihazları
    "cihazlar": {
        "Q0": {"ad": "Ana kesici, MCCB 4P", "govde_A": 250, "Ir_A": 160, "Icu_kA": 36.0,
               "Isd_x": 6, "tsd_ms": 100, "Ii": "kapalı (OFF)", "Icw_kA": 10.0, "kutup": 4},
        "MX0": {"ad": "Q0 açtırma bobini (şönt tripli), 230 V a.a."},
        "RCR0": {"ad": "Gecikmeli kaçak akım rölesi", "I_dn_A": 1.0, "gecikme_ms": 300,
                 "aralik_I": "0,3-1 A", "aralik_t": "100-300 ms"},
        "TOR0": {"ad": "Toroid (kapalı tip), Ø140 mm", "cap_mm": 140},
        "SPD1": {"ad": "Parafudr Tip 1+2, 3+1 bağlantı", "I_imp_kA": 12.5, "In_kA": 25.0, "Up_kV": 1.5, "Uc_V": 275},
        "F1": {"ad": "SPD yedek sigortası, NH00 gG 3P ayırıcılı", "In_A": 125},
        "FSR1": {"ad": "Faz sırası ve gerilim izleme rölesi", "ayar": "±%10 Un, faz sırası/kaybı"},
        "F2": {"ad": "Ölçüm/kumanda devresi şalteri 3P MCB", "In_A": 6, "egri": "C", "Icu_kA": 10.0},
        "CT": {"ad": "Akım trafosu 200/5 A, 5 VA, Sınıf 0,5", "oran": "200/5 A", "adet": 3},
        "PM1": {"ad": "Enerji analizörü (3F4T, RS-485), 96x96 mm"},
        "A1": {"ad": "Ampermetre 0-200 A (200/5 A trafo), 96x96 mm"},
        "V1": {"ad": "Voltmetre 0-500 V, 96x96 mm"},
        "SA1": {"ad": "Ampermetre seçici anahtar (L1-L2-L3)"},
        "SA2": {"ad": "Voltmetre seçici anahtar (L1-L2-L3, L-N/L-L)"},
        "H": {"ad": "Sinyal lambası Ø22 mm, 230 V a.a., LED"},
        "SB1": {"ad": "Acil durdurma butonu Ø40 mm (Q0 açtırma bobinine bağlı)"},
        "Q_klima_grup": {"ad": "Grup besleme şalteri MCB 4P C63", "In_A": 63, "egri": "C", "Icu_kA": 10.0, "kutup": 4},
        "RCD_klima": {"ad": "Kaçak akım rölesi 4P 63 A, 300 mA, Tip B", "In_A": 63, "I_dn_mA": 300,
                      "tip": "B (alt: F)", "t_acma_ms": 40, "Inc_kA": 10.0, "kutup": 4},
        "K_klima": {"ad": "Motor koruma şalteri 3P (alt: C20 MCB)", "ayar_min_A": 11, "ayar_max_A": 16,
                    "ayar_A": 14, "Icu_kA": 50.0, "kutup": 3},
        "RCBO": {"tip": "A", "I_dn_mA": 30, "Icn_kA": 10.0, "t_acma_ms": 40, "kutup": 2},
        "RCD3": {"ad": "Ses sistemi ortak kaçak akım rölesi RCD 2P 25 A, 30 mA, Tip A", "In_A": 25, "I_dn_mA": 30, "tip": "A", "t_acma_ms": 40,
                 "Inc_kA": 10.0, "kutup": 2, "modul": 2},
        "MCB1": {"ad": "MCB 1P+N (ses kolları, RCD3 altında)", "egri": "C", "Icu_kA": 10.0, "kutup": 2},
        "UPS1": {"ad": "UPS, 2 kVA, online çift dönüşümlü (kamera/NVR)", "kva": 2.0},
        "secicilik": {"I_dn_oran_min": 3.0, "zaman_payi_ms": 50, "alt_kisa_devre_acma_ms": 10},
        # --- dış aydınlatma kumanda grubu ve avize
        "Q3": {"ad": "Dış aydınlatma besleme şalteri MCB 4P C25", "In_A": 25, "egri": "C", "Icu_kA": 10.0, "kutup": 4},
        "KM1": {"ad": "Dış aydınlatma kontaktörü 3P 25 A (AC-1/AC-5a), bobin 230 V", "In_A": 25, "modul": 3, "kutup": 3},
        "FT1": {"ad": "Fotosel (alacakaranlık anahtarı), DIN ray + harici sensör, IP65 sensör", "modul": 2, "lux_ayar": "5-50 lux"},
        "ZR1": {"ad": "Astronomik zaman rölesi (GPS/enlem-boylam programlı)", "modul": 2},
        "SA3": {"ad": "MAN-0-OTO seçici anahtar (kapıda, 3 konumlu)"},
        "SPD2": {"ad": "Parafudr Tip 2, 3+1 bağlantı (dış hatlar)", "I_n_kA": 20.0, "Imax_kA": 40.0, "Up_kV": 1.4, "Uc_V": 275, "modul": 4},
        "H4": {"ad": "Dış aydınlatma durum lambası Ø22 mm, 230 V a.a., LED"},
        "DM1": {"ad": "Avize dimmer (LED uyumlu, faz kesmeli, 2,5 kW) / kontaktör", "tip": "dimmer"},
    },
    # ------------------------------------------------------------------ hatlar (sabit kalemler)
    # faz: "3F" veya "L1"/"L2"/"L3" veya "auto" (açgözlü algoritma atar); kd: eşzamanlılık katsayısı; L: pano-alıcı kablo uzunluğu (m)
    # cosφ verilmemişse sistem.cos_varsayilan kullanılır; kablo numaraları (W1...) sırayla otomatik verilir.
    "hatlar": [
        {"id": "K1", "grup": "KLIMA", "ad": "Klima-1", "p_kw": 7.0, "faz": "3F", "kd": 0.95, "tip": "NYY", "damar": "5x4", "kesit": 4.0, "L": 25.0, "yuk": "M1", "limit": "guc"},
        {"id": "K2", "grup": "KLIMA", "ad": "Klima-2", "p_kw": 7.0, "faz": "3F", "kd": 0.95, "tip": "NYY", "damar": "5x4", "kesit": 4.0, "L": 25.0, "yuk": "M2", "limit": "guc"},
        {"id": "K3", "grup": "KLIMA", "ad": "Klima-3", "p_kw": 7.0, "faz": "3F", "kd": 0.95, "tip": "NYY", "damar": "5x4", "kesit": 4.0, "L": 25.0, "yuk": "M3", "limit": "guc"},
        {"id": "K4", "grup": "KLIMA", "ad": "Klima-4", "p_kw": 7.0, "faz": "3F", "kd": 0.95, "tip": "NYY", "damar": "5x4", "kesit": 4.0, "L": 25.0, "yuk": "M4", "limit": "guc"},
        {"id": "K5", "grup": "KLIMA", "ad": "Klima-5", "p_kw": 7.0, "faz": "3F", "kd": 0.95, "tip": "NYY", "damar": "5x4", "kesit": 4.0, "L": 25.0, "yuk": "M5", "limit": "guc"},
        {"id": "K6", "grup": "KLIMA", "ad": "Klima-6", "p_kw": 7.0, "faz": "3F", "kd": 0.95, "tip": "NYY", "damar": "5x4", "kesit": 4.0, "L": 25.0, "yuk": "M6", "limit": "guc"},
        # kamera (UPS üzerinden), ses, öncelikli hat (acil aydınlatma + yangın alarmı, FE180 halojensiz)
        {"id": "KR1", "grup": "KAMERA", "ad": "Kamera / NVR (UPS1)", "p_kw": 1.4, "faz": "auto", "kd": 1.00, "tip": "NYY", "damar": "3x2,5", "kesit": 2.5, "L": 28.0, "yuk": "UPS1", "limit": "guc", "rcbo_In": 10},
        {"id": "AC1", "grup": "ACIL", "ad": "Acil aydınlatma", "p_kw": 0.5, "faz": "auto", "cos": 0.95, "kd": 1.00, "tip": "NHXH FE180/E90", "damar": "3x2,5", "kesit": 2.5, "L": 35.0, "yuk": None, "limit": "aydinlatma", "rcbo_In": 10},
        {"id": "YA1", "grup": "ACIL", "ad": "Yangın alarm santrali", "p_kw": 0.3, "faz": "auto", "cos": 0.95, "kd": 1.00, "tip": "NHXH FE180/E90", "damar": "3x2,5", "kesit": 2.5, "L": 35.0, "yuk": None, "limit": "guc", "rcbo_In": 10},
    ],
    "yedek_rcbo": ["YD1", "YD2", "YD3"],
    "rcbo_defaults": {"egri": "C"},
    # ------------------------------------------------------------------ aydınlatma / priz / dış aydınlatma (yeni yük verileri)
    "aydinlatma_priz": {
        # TOPLAM_KABLO_M: grubun toplam kablo metrajı (kullanıcı verisi); hat başına uzunluk = toplam / hat sayısı.
        # du_yayilim: ΔU hesabında eşdeğer boy katsayısı (yük hat boyunca yayılı ise 0,5; uçta toplu ise 1,0).
        "priz": {"adet": 30, "dagilim": [10, 10, 10], "birim_w": 300.0, "kd": 0.5, "cos": 0.90, "rcbo_In": 16, "kesit": 2.5, "toplam_kablo_m": 800.0,
                 "du_yayilim": 0.5, "aciklama": "16 A, 230 V, topraklı priz girişi"},
        "avize": {"adet": 1, "w": 2000.0, "faz": "L1", "cos": 0.95, "kd": 1.0, "rcbo_In": 16, "kesit": 2.5, "toplam_kablo_m": 100.0,
                  "du_yayilim": 1.0,
                  "dimmer_goster": True},                      # DM1 dimmer/kontaktör bölümünü şemada göster (açık/kapalı)
        "led_ic": {"adet": 50, "birim_w": 15.0, "dagilim": [17, 17, 16], "cos": 0.90, "kd": 1.0, "rcbo_In": 10, "kesit": 1.5, "toplam_kablo_m": 800.0,
                   "du_yayilim": 0.5},
        "projektor": {"adet": 10, "PROJEKTOR_GUCU_W": 100.0, "dagilim": [4, 3, 3], "toplam_kablo_m": 200.0, "du_yayilim": 0.5, "cos": 0.90, "kd": 1.0,
                      "rcbo_In": 10, "kesit": 2.5, "yon": "dışa",
                      "spd2_goster": True},                    # dış hatlar için Tip 2 parafudr (SPD2) önerisini göster (açık/kapalı)
        # Ses sistemi: toplam güç kollara eşit bölünür. L: kol kablo boyu (m). koruma: "ortak_rcd" = ortak RCD3 (30 mA) altında MCB
        # (bu kollar aynı faza bağlanır, en az 2 kol gerekir), "rcbo" = kolun kendi RCBO'su (örn. minare kolu).
        "ses": {"toplam_kw": 1.8, "kollar": [{"ad": "sağ", "L": 100.0, "koruma": "ortak_rcd"}, {"ad": "sol", "L": 100.0, "koruma": "ortak_rcd"},
                                              {"ad": "minare", "L": 100.0, "koruma": "rcbo"}],
                "cos": 0.90, "kd": 0.90, "rcbo_In": 10, "kesit": 2.5},
        "dekoratif": {"kw": 0.0, "faz": "auto", "L": 35.0, "rcbo_In": 10, "kesit": 2.5},   # minare/kubbe dekoratif aydınlatma (0 = yok)
        "kesit_sirasi": [1.5, 2.5, 4.0, 6.0, 10.0, 16.0, 25.0],
    },
    # ------------------------------------------------------------------ pano (ön görünüş / iç yerleşim)
    "pano": {
        "G": 800, "Y": 2000, "D": 400,      # mm  genişlik x yükseklik x derinlik
        "IP": "IP54",
        "modul_mm": 17.5,
        "modul_genislik": {"1P+N": 2, "3P": 3, "4P": 4},
        "plaka_kenar_mm": 50,
        "yan_kanal_mm": 60,
        "cokul_mm": 100,
        "ust_bosluk_mm": 100,
        "ust_bolme_mm": 450,
        "orta_bolme_mm": 300,
        "ray_adim_mm": 150,
        "alt_bolge_min_mm": 300,
        "grup_boslugu_modul": 1,
        # DIN ray satırları (grup adları -> "devre_gruplari"); hazirla() yeni hat/gruplara göre satırları kurar.
        "satirlar": [["KLG1"], ["KLG2"], ["AVZ", "PRZ", "LED", "KAM", "SES"], ["PRJ", "KUM"], ["ACL", "DEK", "YED"]],
        "ust_ray_grup": ["RCR0", "F1", "SPD1", "FSR1", "F2"],
        "ust_ray_modul": {"RCR0": 4, "F1": 4, "SPD1": 4, "FSR1": 2, "F2": 3},
        "standart_genislik": [600, 800, 1000, 1200],
        "standart_yukseklik": [1800, 2000, 2200, 2400],
        "yuzey_isi_katsayisi_W_m2K": 5.5,   # boyalı sac, serbest duran
        # Cihaz kayıp gücü: (açıklama, adet, nominal kayıp W (tüm kutuplar), nominal akım A, işletme akımı referansı)
        # adet=None ise hat sayısından hesaplanır (RCBO).
        "kayiplar": [
            ("Q0 ana kesici", 1, 36.0, 250.0, "ana"),
            ("Ana bara sistemi (L1-L2-L3-N)", 1, 30.0, 250.0, "ana"),
            ("Q1-Q2 grup MCB", 2, 12.0, 63.0, "klima_grup"),
            ("RCD1-RCD2", 2, 6.0, 63.0, "klima_grup"),
            ("K1-K6 motor koruma şalteri", 6, 8.0, 14.0, "klima"),
            ("RCBO'lar (30 mA)", None, 2.5, 16.0, "rcbo"),
            ("Q3 MCB + KM1 kontaktör (bobin dahil)", 1, 14.0, 25.0, "sabit"),
            ("FT1, ZR1, SPD2 ve kumanda", 1, 6.0, 1.0, None),
            ("RCR0, FSR1, PM1, kumanda", 1, 25.0, 1.0, None),
            ("Akım trafoları ve sinyal lambaları", 1, 12.0, 1.0, None),
        ],
    },
    "devre_gruplari": {
        "KLG1": {"ad": "KLİMA GRUP 1", "cihazlar": ["Q1", "RCD1", "K1", "K2", "K3"]},
        "KLG2": {"ad": "KLİMA GRUP 2", "cihazlar": ["Q2", "RCD2", "K4", "K5", "K6"]},
        "KAM": {"ad": "KAMERA", "cihazlar": ["KR1"]},
        "ACL": {"ad": "ACİL/YANGIN", "cihazlar": ["AC1", "YA1"]},
        "YED": {"ad": "YEDEK", "cihazlar": ["YD1", "YD2", "YD3"]},
        # AVZ, PRZ, LED, PRJ, KUM, DEK grupları hazirla() içinde hatlardan oluşturulur
    },
    # ------------------------------------------------------------------ topraklama / yıldırımdan korunma
    "topraklama": {
        "bina_a_m": 20.0, "bina_b_m": 20.0,          # x (doğu-batı), y (kuzey-güney)
        "govde_yuksekligi_m": 10.0,
        "bina_yuksekligi_m": 16.0,                    # kubbe tepesi dahil en yüksek gövde noktası
        "kubbe_merkez_m": (10.0, 10.0), "kubbe_cap_m": 12.0,
        "minare_merkez_m": (18.5, 1.5), "minare_taban_m": 3.0, "minare_yuksekligi_m": 25.0,
        "halka_ofset_m": 1.0, "halka_derinlik_m": 0.8,
        "halka_iletkeni": "50 mm² çıplak Cu", "halka_kesit": 50.0,
        "cubuk_boy_m": 3.0, "cubuk_cap_mm": 18.0, "cubuk_aralik_max_m": 10.0, "cubuk_aralik_min_m": 6.0,
        "kontrol_cukuru_adet": 4,
        "hedef_direnc_ohm": 2.0,
        "pano_konum_m": (1.2, 19.4),                  # PN1 merkezi (elektrik odası, güney duvar)
        "epb_konum_m": (3.2, 19.4),
        "epb_halka_hatti_m": 3.0,                     # EPB1 - halka bağlantı uzunluğu
        "pano_epb_hatti_m": 2.5,                      # PN1 PE barası - EPB1 uzunluğu
        "trafo_elektrodu": {"cubuk_adet": 3, "cubuk_boy_m": 3.0, "aralik_m": 6.0},
        "ypk_seviye": "III",
        "ypk_tablo": {"I": {"R": 20, "aralik": 10}, "II": {"R": 30, "aralik": 10},
                      "III": {"R": 45, "aralik": 15}, "IV": {"R": 60, "aralik": 20}},
        "yk_boylari_m": {"YK1": 2.0, "YK2": 1.5, "YK3": 1.0},   # Franklin çubukları (minare, kubbe, çatı köşesi)
        "inis_kesit": 50.0, "olcum_ayirma_yuksekligi_m": 1.7,
        "cati_iletkeni": "30x3,5 mm galvaniz şerit",
        "yedek_pay_yuzde": 10.0,
        "eb_ortalama_uzunluk_m": 12.0,                 # EB1-EB6 bağlantı iletkenlerinin ortalama boyu (varsayım)
    },
}

# =====================================================================
#  Yardımcılar
# =====================================================================
FONT_AILE = "Liberation Sans, Arial, Helvetica, sans-serif"
_FONT_REG = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
_FONT_BOLD = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
_font_cache = {}
UYARILAR = []          # çalışma sırasında biriken uyarılar


def uyar(msg):
    UYARILAR.append(msg)
    print("  [UYARI]", msg)


def klima_mi(poz):
    """K1..K6 (klima motor koruma şalteri) mi?"""
    return re.fullmatch(r"K\d+", poz) is not None


def tr(x, nd=1):
    """Türkçe sayı biçimi: ondalık ayracı virgül."""
    if x is None:
        return "-"
    s = f"{x:.{nd}f}"
    return s.replace(".", ",")


def trg(x, nd=0):
    """Binlik ayraçlı Türkçe sayı (nokta binlik, virgül ondalık)."""
    s = f"{x:,.{nd}f}"
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def tw(text, size, bold=False):
    """Metin genişliği (mm), size = yazı yüksekliği (mm)."""
    key = (bold,)
    if key not in _font_cache:
        _font_cache[key] = ImageFont.truetype(_FONT_BOLD if bold else _FONT_REG, 200)
    return _font_cache[key].getlength(text) / 200.0 * size


def wrap(text, maxw, size, bold=False):
    """Kelime sarma; (satırlar) döner."""
    satirlar = []
    for par in str(text).split("\n"):
        kelimeler = par.split(" ")
        cur = ""
        for k in kelimeler:
            deneme = (cur + " " + k).strip()
            if tw(deneme, size, bold) <= maxw or not cur:
                cur = deneme
            else:
                satirlar.append(cur)
                cur = k
        satirlar.append(cur)
    return satirlar


def poz_aralik(pozlar):
    """['K1','K2','K3'] -> 'K1–K3' ; ardışık olmayanlar virgülle."""
    if not pozlar:
        return ""
    parcalar = []
    grp = [pozlar[0]]

    def ayir(p):
        m = re.match(r"^(.*?)(\d+)$", p)
        return (m.group(1), int(m.group(2))) if m else (p, None)

    for p in pozlar[1:]:
        a, n = ayir(grp[-1])
        b, m = ayir(p)
        if a == b and n is not None and m == n + 1:
            grp.append(p)
        else:
            parcalar.append(grp)
            grp = [p]
    parcalar.append(grp)
    out = []
    for g in parcalar:
        out.append(g[0] if len(g) == 1 else (g[0] + "–" + g[-1] if len(g) > 2 else g[0] + ", " + g[-1]))
    return ", ".join(out)


def poz_genislet(metin):
    """'K1–K6, W0' -> ['K1',...,'K6','W0']"""
    sonuc = []
    for parca in re.split(r"\s*,\s*", metin.strip()):
        if not parca:
            continue
        m = re.match(r"^(\D+?)(\d+)\s*[–-]\s*(\D*?)(\d+)$", parca)
        if m and (m.group(3) in ("", m.group(1))):
            for n in range(int(m.group(2)), int(m.group(4)) + 1):
                sonuc.append(f"{m.group(1)}{n}")
        else:
            sonuc.append(parca)
    return sonuc


# =====================================================================
#  HESAP MODELİ
# =====================================================================
def du_yuzde(I, L, S, cos, faz3, U_ll, U_ln, kappa, X_mohm_m):
    sin = math.sqrt(max(0.0, 1 - cos * cos))
    r = 1.0 / (kappa * S)               # ohm/m
    x = X_mohm_m / 1000.0
    if faz3:
        du = math.sqrt(3) * I * L * (r * cos + x * sin)
        return du / U_ll * 100.0, du
    du = 2 * I * L * (r * cos + x * sin)
    return du / U_ln * 100.0, du


def kes_txt(k):
    """Kesit metni: 1,5 / 2,5 / 4 / 10"""
    return tr(k, 1) if abs(k - round(k)) > 1e-9 else str(int(round(k)))


FAZLAR = ("L1", "L2", "L3")
_MCB_POZ = set()       # ortak RCD altındaki MCB'li ses kolları (hazirla doldurur; çizim türü seçimi için)
FAZ_FARKLI_GRUPLAR = ("PRIZ", "LED", "PROJ")      # aynı gruptaki hatlar farklı fazlara bağlanır


def hazirla(PARAM):
    """PARAMETRELER'den türetilen yük listesi: yeni hatlar (priz, avize, LED, projektör, dekoratif), kablo numaraları,
    açgözlü (greedy) faz ataması, devre grupları ve DIN ray satırları. Orijinal sözlüğü değiştirmez."""
    P = copy.deepcopy(PARAM)
    S, A, C = P["sistem"], P["aydinlatma_priz"], P["cihazlar"]
    cos0 = S["cos_varsayilan"]
    statik = P["hatlar"]
    klima = [dict(h) for h in statik if klima_mi(h["id"])]
    diger = [dict(h) for h in statik if not klima_mi(h["id"])]
    uret = []
    elemanlar = []                         # çizim 06 ve envanter için tekil cihaz listesi

    def bol(prefix, adet, dagilim, baslangic=1):
        if sum(dagilim) != adet:
            uyar(f"{prefix}: dağılım toplamı {sum(dagilim)} ≠ adet {adet}")
        sonuc, n0 = [], baslangic
        for n in dagilim:
            sonuc.append([f"{prefix}{k}" for k in range(n0, n0 + n)])
            n0 += n
        return sonuc

    # ---- avize
    av = A["avize"]
    if av["adet"] > 0:
        ids = [f"AVZ{k}" for k in range(1, av["adet"] + 1)]
        uret.append(dict(id="AV1", grup="AVIZE", ad="Büyük avize\n(" + poz_aralik(ids) + ")", p_kw=av["adet"] * av["w"] / 1000.0, faz=av["faz"], cos=av["cos"],
                         kd=av["kd"], tip="NYY", damar="3x" + kes_txt(av["kesit"]), kesit=av["kesit"], L=av["toplam_kablo_m"], yuk=None, limit="aydinlatma",
                         rcbo_In=av["rcbo_In"], du_yayilim=av["du_yayilim"], kesit_otomatik=True, elemanlar=ids))
        for e in ids:
            elemanlar.append({"poz": e, "tip": "avize", "hat": "AV1", "w": av["w"], "ad": "Büyük avize"})
    # ---- priz
    pz = A["priz"]
    for i, ids in enumerate(bol("P", pz["adet"], pz["dagilim"])):
        uret.append(dict(id=f"PZ{i + 1}", grup="PRIZ", ad=f"Priz hattı-{i + 1}\n({poz_aralik(ids)})", p_kw=len(ids) * pz["birim_w"] / 1000.0, faz="auto", cos=pz["cos"],
                         kd=pz["kd"], tip="NYY", damar="3x" + kes_txt(pz["kesit"]), kesit=pz["kesit"], L=pz["toplam_kablo_m"] / len(pz["dagilim"]), yuk=None, limit="guc", rcbo_In=pz["rcbo_In"],
                         du_yayilim=pz["du_yayilim"], kesit_otomatik=True, elemanlar=ids))
        for e in ids:
            elemanlar.append({"poz": e, "tip": "priz", "hat": f"PZ{i + 1}", "w": pz["birim_w"], "ad": "Priz girişi 16 A, 230 V"})
    # ---- iç LED
    ld = A["led_ic"]
    for i, ids in enumerate(bol("L", ld["adet"], ld["dagilim"])):
        uret.append(dict(id=f"LD{i + 1}", grup="LED", ad=f"İç LED hat-{i + 1}\n({poz_aralik(ids)})", p_kw=len(ids) * ld["birim_w"] / 1000.0, faz="auto", cos=ld["cos"],
                         kd=ld["kd"], tip="NYY", damar="3x" + kes_txt(ld["kesit"]), kesit=ld["kesit"], L=ld["toplam_kablo_m"] / len(ld["dagilim"]), yuk=None, limit="aydinlatma", rcbo_In=ld["rcbo_In"],
                         du_yayilim=ld["du_yayilim"], kesit_otomatik=True, elemanlar=ids))
        for e in ids:
            elemanlar.append({"poz": e, "tip": "led", "hat": f"LD{i + 1}", "w": ld["birim_w"], "ad": "İç mekan LED armatür"})
    # ---- dış projektör
    pj = A["projektor"]
    wpj = pj["PROJEKTOR_GUCU_W"]
    pj["PROJEKTOR_MESAFE_M"] = pj["toplam_kablo_m"] / len(pj["dagilim"])      # hat başına kablo boyu (türetilen)
    for i, ids in enumerate(bol("PR", pj["adet"], pj["dagilim"])):
        uret.append(dict(id=f"PJ{i + 1}", grup="PROJ", ad=f"Projektör hattı-{i + 1}\n({poz_aralik(ids)})", p_kw=len(ids) * wpj / 1000.0, faz="auto", cos=pj["cos"],
                         kd=pj["kd"], tip="NYY", damar="3x" + kes_txt(pj["kesit"]), kesit=pj["kesit"], L=pj["PROJEKTOR_MESAFE_M"], yuk=None, limit="guc",
                         rcbo_In=pj["rcbo_In"], du_yayilim=pj["du_yayilim"], kesit_otomatik=True, elemanlar=ids))
        for e in ids:
            elemanlar.append({"poz": e, "tip": "projektor", "hat": f"PJ{i + 1}", "w": wpj, "ad": "LED projektör IP65"})
    # ---- dekoratif (isteğe bağlı)
    dk = A["dekoratif"]
    if dk["kw"] > 0:
        uret.append(dict(id="DK1", grup="DEKOR", ad="Minare/kubbe dekoratif aydınlatma", p_kw=dk["kw"], faz=dk["faz"], cos=cos0, kd=1.0, tip="NYY",
                         damar="3x" + kes_txt(dk["kesit"]), kesit=dk["kesit"], L=dk["L"], yuk=None, limit="aydinlatma", rcbo_In=dk["rcbo_In"], elemanlar=[]))
    # ---- ses sistemi kolları (sağ / sol / minare), kamera hattından sonra
    ss = A["ses"]
    ses_hatlari = []
    ortak = [k_ for k_ in ss["kollar"] if k_.get("koruma") == "ortak_rcd"]
    if len(ortak) == 1:
        uyar("Ses sistemi: ortak RCD için en az 2 kol gerekir; tek kol RCBO'ya çevrildi.")
        ortak[0]["koruma"] = "rcbo"
    elif ortak:
        pass
    for i, kol_ in enumerate(ss["kollar"], start=1):
        ses_hatlari.append(dict(koruma_tur="mcb" if kol_.get("koruma") == "ortak_rcd" else "rcbo", id=f"SS{i}", grup="SES", ad=f"Ses sistemi - {kol_['ad']} kol", p_kw=ss["toplam_kw"] / len(ss["kollar"]), faz="auto", cos=ss["cos"],
                                kd=ss["kd"], tip="NYY", damar="3x" + kes_txt(ss["kesit"]), kesit=ss["kesit"], L=kol_["L"], yuk=None, limit="guc",
                                rcbo_In=ss["rcbo_In"], kesit_otomatik=True, kol_ad=kol_["ad"]))
    ort_ids = [h["id"] for h in ses_hatlari if h["koruma_tur"] == "mcb"]
    for h in ses_hatlari:
        if h["id"] in ort_ids:
            h["rcd_id"] = "RCD3"
            if h["id"] != ort_ids[0]:
                h["faz_ayni"] = ort_ids[0]
    _MCB_POZ.clear()
    _MCB_POZ.update(ort_ids)
    k_idx = next((i for i, h in enumerate(diger) if h["id"] == "KR1"), -1) + 1
    diger = diger[:k_idx] + ses_hatlari + diger[k_idx:]
    hatlar = klima + [h for h in uret if h["id"] != "DK1"] + ([h for h in uret if h["id"] == "DK1"]) + diger
    for h in hatlar:
        h.setdefault("cos", cos0)
        h.setdefault("yuk", None)
    # ---- kablo numaraları
    for i, h in enumerate(hatlar, start=1):
        h["kablo_no"] = f"W{i}"
    # ---- açgözlü faz ataması
    yuk = {f: 0.0 for f in FAZLAR}
    kullanilan = {}
    adimlar = []
    tekfaz = [h for h in hatlar if h["faz"] != "3F"]
    sabitler = [h for h in tekfaz if h["faz"] in FAZLAR]
    otomatik = [h for h in tekfaz if h["faz"] not in FAZLAR and not h.get("faz_ayni")]
    bagli = {}
    for h in tekfaz:
        if h.get("faz_ayni"):
            bagli.setdefault(h["faz_ayni"], []).append(h)
    otomatik.sort(key=lambda h: -(h["p_kw"] + sum(b["p_kw"] for b in bagli.get(h["id"], []))) * h["kd"])
    sira = 0
    for h in sabitler + otomatik:
        d = (h["p_kw"] + sum(b["p_kw"] for b in bagli.get(h["id"], []))) * h["kd"]
        if h["faz"] in FAZLAR:
            ph, neden = h["faz"], "parametre ile sabit"
        else:
            kullan = kullanilan.setdefault(h["grup"], set()) if h["grup"] in FAZ_FARKLI_GRUPLAR else set()
            aday = [f for f in FAZLAR if f not in kullan] or list(FAZLAR)
            ph = min(aday, key=lambda f: (yuk[f], f))
            neden = "en az yüklü faz" + (" (grupta farklı faz)" if h["grup"] in FAZ_FARKLI_GRUPLAR else "")
        h["faz"] = ph
        yuk[ph] += d
        if h["grup"] in FAZ_FARKLI_GRUPLAR:
            kullanilan.setdefault(h["grup"], set()).add(ph)
        sira += 1
        adimlar.append({"sira": sira, "id": h["id"], "ad": h["ad"].replace("\n", " "), "talep_kw": d, "faz": ph, "neden": neden,
                        "L1": yuk["L1"], "L2": yuk["L2"], "L3": yuk["L3"]})
        for b in bagli.get(h["id"], []):               # aynı RCD altındaki hatlar aynı faza bağlanır
            b["faz"] = ph
            sira += 1
            adimlar[-1]["ad"] += f" + {b['id']}"
            adimlar[-1]["neden"] += f" ({b['id']} aynı ortak RCD altında, aynı faz)"
    P["faz_atama"] = adimlar
    P["hatlar"] = hatlar
    P["elemanlar"] = elemanlar
    # ---- devre grupları ve DIN ray satırları
    G = P["devre_gruplari"]
    spd2 = pj.get("spd2_goster", True)
    G["AVZ"] = {"ad": "AVİZE", "cihazlar": ["AV1"] if av["adet"] > 0 else []}
    G["PRZ"] = {"ad": "PRİZ", "cihazlar": [h["id"] for h in hatlar if h["grup"] == "PRIZ"]}
    G["SES"] = {"ad": "SES", "cihazlar": (["RCD3"] if ort_ids else []) + [h["id"] for h in hatlar if h["grup"] == "SES"]}
    G["LED"] = {"ad": "İÇ LED", "cihazlar": [h["id"] for h in hatlar if h["grup"] == "LED"]}
    G["PRJ"] = {"ad": "DIŞ AYDINLATMA", "cihazlar": ["Q3"] + (["SPD2"] if spd2 else []) + ["KM1"] + [h["id"] for h in hatlar if h["grup"] == "PROJ"]}
    G["KUM"] = {"ad": "DIŞ AYD. KUMANDA", "cihazlar": ["FT1", "ZR1"]}
    if dk["kw"] > 0:
        G["DEK"] = {"ad": "DEKORATİF", "cihazlar": ["DK1"]}
    satirlar = []
    for sat in P["pano"]["satirlar"]:
        yeni = [g for g in sat if g in G and G[g]["cihazlar"]]
        if yeni:
            satirlar.append(yeni)
    P["pano"]["satirlar"] = satirlar
    return P


def hat_hesapla(h, kesit, P, M_du_besleme, U_ll, U_ln):
    """Bir hattın akım/koruma/ΔU alanlarını verilen kesit için (yeniden) hesaplar."""
    K, CIH = P["kablo"], P["cihazlar"]
    f3 = h["faz"] == "3F"
    h["faz3"] = f3
    h["kesit"] = kesit
    h["p_talep"] = h["p_kw"] * h["kd"]
    h["I"] = (h["p_kw"] * 1000.0 / (math.sqrt(3) * U_ll * h["cos"])) if f3 else (h["p_kw"] * 1000.0 / (U_ln * h["cos"]))
    tablo = K["iz_C_3yuklu"] if f3 else K["iz_C_2yuklu"]
    h["iz_tablo"] = tablo[kesit]
    h["iz"] = h["iz_tablo"] * K["azaltma_k"]
    h["du_hat"], h["du_V"] = du_yuzde(h["I"], h["L"] * h.setdefault("du_yayilim", 1.0), kesit, h["cos"], f3, U_ll, U_ln, K["kappa"], K["X_mohm_m"])
    h["du_limit"] = K["limit_guc_yuzde"] if h["limit"] == "guc" else K["limit_aydinlatma_yuzde"]
    h["du_toplam"] = h["du_hat"] + M_du_besleme
    h["uygun"] = (h["du_toplam"] <= h["du_limit"]) and (h["I"] <= h["In_koruma"] <= h["iz"])
    return h


def hesapla(P):
    S, K, CIH = P["sistem"], P["kablo"], P["cihazlar"]
    M = {}
    U_ll, U_ln = S["U_LL"], S["U_LN"]
    hat = []
    for h in P["hatlar"]:
        h = dict(h)
        if klima_mi(h["id"]):
            h["In_koruma"] = CIH["K_klima"]["ayar_A"]
            h["koruma_ad"] = "MPCB 3P %d-%d A (ayar %d A)" % (CIH["K_klima"]["ayar_min_A"], CIH["K_klima"]["ayar_max_A"], CIH["K_klima"]["ayar_A"])
            h["koruma_tablo"] = "MPCB 3P %d-%d A\n(ayar %d A)" % (CIH["K_klima"]["ayar_min_A"], CIH["K_klima"]["ayar_max_A"], CIH["K_klima"]["ayar_A"])
            h["Icu_kA"] = CIH["K_klima"]["Icu_kA"]
        elif h.get("koruma_tur") == "mcb":
            h["In_koruma"] = h["rcbo_In"]
            h["koruma_ad"] = "MCB 1P+N %s%d (%s altında)" % (CIH["MCB1"]["egri"], h["rcbo_In"], h["rcd_id"])
            h["koruma_tablo"] = "MCB 1P+N %s%d A\n(%s altında)" % (CIH["MCB1"]["egri"], h["rcbo_In"], h["rcd_id"])
            h["Icu_kA"] = CIH["MCB1"]["Icu_kA"]
        else:
            h["In_koruma"] = h["rcbo_In"]
            h["koruma_ad"] = "RCBO 1P+N %s%d, %d mA" % (P["rcbo_defaults"]["egri"], h["rcbo_In"], CIH["RCBO"]["I_dn_mA"])
            h["koruma_tablo"] = "RCBO 1P+N\n%s%d A, %d mA" % (P["rcbo_defaults"]["egri"], h["rcbo_In"], CIH["RCBO"]["I_dn_mA"])
            h["Icu_kA"] = CIH["RCBO"]["Icn_kA"]
        hat_hesapla(h, h["kesit"], P, 0.0, U_ll, U_ln)
        hat.append(h)
    M["hatlar"] = hat

    # faz toplamları (talep gücü, kW) ve akımlar
    pf = {"L1": 0.0, "L2": 0.0, "L3": 0.0}
    If = {"L1": 0.0, "L2": 0.0, "L3": 0.0}
    for h in hat:
        for ph in pf:
            if h["faz3"]:
                pf[ph] += h["p_talep"] / 3.0
                If[ph] += h["p_talep"] * 1000.0 / (3 * U_ln * h["cos"])
            elif h["faz"] == ph:
                pf[ph] += h["p_talep"]
                If[ph] += h["p_talep"] * 1000.0 / (U_ln * h["cos"])
    M["faz_kw"], M["faz_I"] = pf, If
    M["p_kurulu"] = sum(h["p_kw"] for h in hat)
    M["p_talep"] = sum(h["p_talep"] for h in hat)
    ort = sum(pf.values()) / 3.0
    M["dengesizlik"] = (max(pf.values()) - min(pf.values())) / ort * 100.0 if ort else 0.0
    M["Ib"] = max(If.values())
    M["Ib_ort"] = sum(If.values()) / 3.0
    M["cos_ort"] = M["p_talep"] * 1000.0 / (math.sqrt(3) * U_ll * M["Ib_ort"])
    M["S_talep_kva"] = math.sqrt(3) * U_ll * M["Ib_ort"] / 1000.0
    M["S_kurulu_kva"] = sum((h["p_kw"] / h["cos"]) for h in hat)
    Ir = CIH["Q0"]["Ir_A"]
    pay = S["gelecek_payi_yuzde"]
    M["Ir"] = Ir
    M["Ib_gelecek"] = M["Ib"] * (1 + pay / 100.0)
    M["gelecek_uygun"] = M["Ib_gelecek"] <= Ir
    M["yedek_A"] = Ir - M["Ib"]
    M["yedek_yuzde"] = (Ir - M["Ib"]) / Ir * 100.0
    M["faz_atama"] = P["faz_atama"]
    M["elemanlar"] = P["elemanlar"]

    # besleme kablosu
    B = P["besleme"]
    tablo = {"D2": K["iz_D2_3yuklu"], "C": K["iz_C_3yuklu"]}[B["yontem"]]
    kf = K["azaltma_k_gomulu"] if B["yontem"] == "D2" else K["azaltma_k"]
    M["besleme"] = dict(B)
    M["besleme"]["L"] = S["trafo_pano_mesafe_m"]
    M["besleme"]["iz_tablo"] = tablo[B["kesit"]]
    M["besleme"]["iz"] = tablo[B["kesit"]] * kf
    M["besleme"]["du"], M["besleme"]["du_V"] = du_yuzde(M["Ib"], S["trafo_pano_mesafe_m"], B["kesit"], M["cos_ort"], True,
                                                       U_ll, U_ln, K["kappa"], K["X_mohm_m"])
    # hat ΔU (besleme dahil) ve otomatik kesit büyütme (LED, projektör)
    sira = P["aydinlatma_priz"]["kesit_sirasi"]
    for h in hat:
        hat_hesapla(h, h["kesit"], P, M["besleme"]["du"], U_ll, U_ln)
        if h.get("kesit_otomatik"):
            k0 = h["kesit"]
            idx = sira.index(k0) if k0 in sira else 0
            while not h["uygun"] and idx + 1 < len(sira):
                idx += 1
                hat_hesapla(h, sira[idx], P, M["besleme"]["du"], U_ll, U_ln)
            if h["kesit"] != k0:
                h["damar"] = "3x" + kes_txt(h["kesit"])
                h["kesit_ilk"] = k0
                h["kesit_artirildi"] = True
                uyar(f"{h['id']} ({h['kablo_no']}): gerilim düşümü/akım şartı için kesit {kes_txt(k0)} → {kes_txt(h['kesit'])} mm² büyütüldü "
                     f"(ΔU toplam %{tr(h['du_toplam'], 2)} ≤ %{tr(h['du_limit'], 1)}).")
            if not h["uygun"]:
                uyar(f"{h['id']} ({h['kablo_no']}): en büyük kesitte de şart sağlanmadı (ΔU toplam %{tr(h['du_toplam'], 2)}).")
        elif not h["uygun"] and not klima_mi(h["id"]):
            h["sabit_kesit_uygunsuz"] = True
            uyar(f"{h['id']} ({h['kablo_no']}): kesit parametreyle sabitlendi ({kes_txt(h['kesit'])} mm²); ΔU toplam %{tr(h['du_toplam'], 2)} > limit %{tr(h['du_limit'], 1)} "
                 f"→ ŞART SAĞLANMIYOR (kullanıcı kararı; kesit/boy/yük konumu mühendislikçe yeniden değerlendirilmeli).")

    # kısa devre (trafo uçları ve pano barası)
    In_T = S["trafo_kva"] * 1000.0 / (math.sqrt(3) * U_ll)
    Z_T = S["trafo_uk_yuzde"] / 100.0 * U_ll ** 2 / (S["trafo_kva"] * 1000.0)
    R_T = S["trafo_ur_yuzde"] / 100.0 * U_ll ** 2 / (S["trafo_kva"] * 1000.0)
    X_T = math.sqrt(max(Z_T ** 2 - R_T ** 2, 0.0))
    Lb = S["trafo_pano_mesafe_m"]
    R_c = Lb / (56.0 * B["kesit"])          # 20 C direnci (maksimum Ik için)
    X_c = K["X_mohm_m"] / 1000.0 * Lb
    Z_tot = math.sqrt((R_T + R_c) ** 2 + (X_T + X_c) ** 2)
    M["In_trafo"] = In_T
    M["Ik_trafo_kA"] = In_T / (S["trafo_uk_yuzde"] / 100.0) / 1000.0
    M["Ik_pano_kA"] = U_ll / (math.sqrt(3) * Z_tot) / 1000.0
    return M


def poz_etiketleri(P):
    """Hat -> RCBO / MPCB poz sırası (çizimlerde kullanılacak)."""
    return [h["id"] for h in P["hatlar"]]


# ---------------------------------------------------------------------
#  Pano yerleşim hesabı (DIN ray modülleri)
# ---------------------------------------------------------------------
def cihaz_modul(P, poz):
    mg = P["pano"]["modul_genislik"]
    C = P["cihazlar"]
    if poz in C and "modul" in C[poz]:
        return C[poz]["modul"]           # KM1 (3P kontaktör), FT1, ZR1, SPD2 ...
    if poz in ("Q1", "Q2", "Q3", "RCD1", "RCD2"):
        return mg["4P"]
    if klima_mi(poz):
        return mg["3P"]
    return mg["1P+N"]       # RCBO / yedek yuva


def pano_hesapla(P, M):
    pn = dict(P["pano"])
    uyarilar = []
    G, Y = pn["G"], pn["Y"]
    for _ in range(10):
        ray_len = G - 2 * pn["plaka_kenar_mm"] - 2 * pn["yan_kanal_mm"]
        kap = int(ray_len // pn["modul_mm"])
        satirlar = []
        tasti = False
        for sat in pn["satirlar"]:
            seg, toplam = [], 0
            for gi, g in enumerate(sat):
                cih = P["devre_gruplari"][g]["cihazlar"]
                m = sum(cihaz_modul(P, c) for c in cih)
                toplam += m + (pn["grup_boslugu_modul"] if gi > 0 else 0)
                seg.append((g, m))
            satirlar.append({"gruplar": seg, "modul": toplam, "kapasite": kap})
            if toplam > kap:
                tasti = True
        # üst bölme ray (Q0 yanındaki)
        q0_gen = 205                                      # Q0 4P 250A gövde (~140 mm) + MX0 + boşluk
        ust_ray_len = G - 2 * pn["plaka_kenar_mm"] - 2 * pn["yan_kanal_mm"] - q0_gen
        ust_kap = int(ust_ray_len // pn["modul_mm"])
        ust_m = sum(pn["ust_ray_modul"][k] for k in pn["ust_ray_grup"]) + (len(pn["ust_ray_grup"]) - 1) * pn["grup_boslugu_modul"]
        if ust_m > ust_kap:
            tasti = True
        n_sat = len(satirlar)
        gerekli_Y = (pn["ust_bosluk_mm"] + pn["ust_bolme_mm"] + pn["orta_bolme_mm"] + n_sat * pn["ray_adim_mm"]
                     + pn["alt_bolge_min_mm"] + pn["cokul_mm"])
        yuk_tasti = gerekli_Y > Y
        if tasti:
            sonraki = [w for w in pn["standart_genislik"] if w > G]
            if not sonraki:
                pn["ek_bolme"] = True
                uyarilar.append(f"DIN ray modülleri en büyük standart genişliğe ({G} mm) de SIĞMIYOR -> ikinci kapı/ek bölme (yan pano) GEREKLİ.")
                break
            uyarilar.append(f"DIN ray modülleri {G} mm genişliğe SIĞMIYOR -> pano genişliği {sonraki[0]} mm'ye büyütüldü.")
            G = sonraki[0]
            continue
        if yuk_tasti:
            sonraki = [w for w in pn["standart_yukseklik"] if w >= gerekli_Y]
            if not sonraki:
                pn["ek_bolme"] = True
                uyarilar.append(f"Bölmeler en büyük standart yüksekliğe ({Y} mm) de SIĞMIYOR (gerekli {gerekli_Y} mm) -> ek bölme (yan pano) GEREKLİ.")
                break
            uyarilar.append(f"Bölmeler {Y} mm yüksekliğe SIĞMIYOR (gerekli {gerekli_Y} mm) -> pano yüksekliği {sonraki[0]} mm'ye büyütüldü.")
            Y = sonraki[0]
            continue
        break
    pn["G"], pn["Y"] = G, Y
    pn["ray_uzunlugu_mm"] = ray_len
    pn["modul_kapasite"] = kap
    pn["satir_hesap"] = satirlar
    pn["ust_ray_modul_toplam"] = ust_m
    pn["ust_ray_kapasite"] = ust_kap
    pn["gerekli_Y"] = gerekli_Y
    pn["uyarilar"] = uyarilar
    for u in uyarilar:
        uyar(u)

    # ısıl hesap
    Ib = M["Ib"]
    I_klima_grup = sum(h["I"] for h in M["hatlar"] if h["id"] in ("K1", "K2", "K3"))
    I_klima = M["hatlar"][0]["I"]
    tek = [h for h in M["hatlar"] if not h["faz3"]]
    I_rcbo = sum(h["I"] for h in tek) / max(len(tek), 1)
    ref = {"ana": Ib, "klima_grup": I_klima_grup, "klima": I_klima, "rcbo": I_rcbo, None: None, "sabit": None}
    kayip_satirlari = []
    toplam = 0.0
    n_rcbo = sum(1 for h in M["hatlar"] if not klima_mi(h["id"]))
    for ad, adet, w_nom, i_nom, r in pn["kayiplar"]:
        if adet is None:
            adet = n_rcbo
        if r is None or r == "sabit":
            w = w_nom * adet
        else:
            w = w_nom * adet * (ref[r] / i_nom) ** 2
        kayip_satirlari.append((ad, adet, w_nom, i_nom, ref[r], w))
        toplam += w
    Gm, Ym, Dm = G / 1000.0, Y / 1000.0, pn["D"] / 1000.0
    Ae = 1.8 * Ym * (Gm + Dm) + 1.4 * Gm * Dm                 # m2 (serbest duran, tüm yüzeyler açık)
    dT = P["sistem"]["pano_ic_sicaklik_max_C"] - P["sistem"]["ortam_sicakligi_C"]
    Pd = pn["yuzey_isi_katsayisi_W_m2K"] * Ae * dT
    fan_gerekli = toplam > 0.8 * Pd
    hava = max(0.0, 3.1 * (toplam - 0.8 * Pd) / dT) if fan_gerekli else 0.0
    pn["isil"] = {"satirlar": kayip_satirlari, "Pv": toplam, "Ae": Ae, "dT": dT, "Pd": Pd,
                  "fan_gerekli": fan_gerekli, "hava_m3h": hava}
    return pn


# ---------------------------------------------------------------------
#  Topraklama / yıldırımdan korunma geometrisi (m, kuzey-batı köşe orijin, y güneye)
# ---------------------------------------------------------------------
def topraklama_hesapla(P):
    T = dict(P["topraklama"])
    a, b, o = T["bina_a_m"], T["bina_b_m"], T["halka_ofset_m"]
    x0, y0, x1, y1 = -o, -o, a + o, b + o                       # halka dikdörtgeni
    Lx, Ly = x1 - x0, y1 - y0
    T["halka"] = (x0, y0, x1, y1)
    T["halka_cevre"] = 2 * (Lx + Ly)
    # çubuklar: saat yönünde, KB köşesinden başla
    nx = math.ceil(Lx / T["cubuk_aralik_max_m"])
    ny = math.ceil(Ly / T["cubuk_aralik_max_m"])
    cubuklar = []
    for i in range(nx):                                          # kuzey kenarı: batıdan doğuya
        cubuklar.append((x0 + Lx * i / nx, y0))
    for j in range(ny):                                          # doğu kenarı: kuzeyden güneye
        cubuklar.append((x1, y0 + Ly * j / ny))
    for i in range(nx):                                          # güney kenarı: doğudan batıya
        cubuklar.append((x1 - Lx * i / nx, y1))
    for j in range(ny):                                          # batı kenarı: güneyden kuzeye
        cubuklar.append((x0, y1 - Ly * j / ny))
    T["cubuklar"] = [{"poz": f"TÇ{i + 1}", "xy": c} for i, c in enumerate(cubuklar)]
    T["cubuk_aralik_gercek"] = max(Lx / nx, Ly / ny)
    T["cubuk_aralik_min_gercek"] = min(Lx / nx, Ly / ny)
    # kontrol çukurları: köşe çubuklarında (ilk 4 köşe)
    kose_idx = [0, nx, nx + ny, nx + ny + nx]
    n_cukur = int(T["kontrol_cukuru_adet"])
    cukur_idx = (kose_idx + [i for i in range(len(cubuklar)) if i not in kose_idx])[:n_cukur]
    T["cukurlar"] = [{"poz": f"KÇ{k + 1}", "cubuk": T["cubuklar"][i]["poz"], "xy": cubuklar[i]} for k, i in enumerate(sorted(cukur_idx))]
    # yıldırım koruma seviyesi
    sev = T["ypk_tablo"][T["ypk_seviye"]]
    T["kure_R"], T["inis_aralik"] = sev["R"], sev["aralik"]
    cevre_bina = 2 * (a + b)
    n_inis = max(2, math.ceil(cevre_bina / sev["aralik"]))
    # iniş iletkeni konumları: önce köşeler, sonra kenar ortaları (S: kubbe, N, E, W)
    kose = {"KD": (a, 0.0), "GD": (a, b), "GB": (0.0, b), "KB": (0.0, 0.0)}
    orta = {"G": (a / 2, b), "K": (a / 2, 0.0), "D": (a, b / 2), "B": (0.0, b / 2)}
    secim = list(kose.keys()) + ["G", "K", "D", "B"]
    allp = {**kose, **orta}
    n_inis = min(n_inis, 8)
    secili = secim[:n_inis] if n_inis >= 4 else list(kose.keys())[:n_inis]
    saat = ["KD", "D", "GD", "G", "GB", "B", "KB", "K"]            # saat yönünde sıralama (KD'den)
    secili = [k for k in saat if k in secili]
    T["inis"] = [{"poz": f"İL{i + 1}", "yer": k, "xy": allp[k]} for i, k in enumerate(secili)]
    for i, d in enumerate(T["inis"]):
        d["ok"] = f"ÖK{i + 1}"
    T["inis_ort_aralik"] = cevre_bina / len(T["inis"])
    # hava sonlandırma elemanları (Franklin çubukları)
    mx, my = T["minare_merkez_m"]
    kx, ky = T["kubbe_merkez_m"]
    T["yk"] = [
        {"poz": "YK1", "ad": "minare ucu", "xy": (mx, my), "z": T["minare_yuksekligi_m"], "boy": T["yk_boylari_m"]["YK1"]},
        {"poz": "YK2", "ad": "kubbe tepesi", "xy": (kx, ky), "z": T["bina_yuksekligi_m"], "boy": T["yk_boylari_m"]["YK2"]},
        {"poz": "YK3", "ad": "çatı kenarı (batı orta)", "xy": (0.0, b / 2.0), "z": T["govde_yuksekligi_m"], "boy": T["yk_boylari_m"]["YK3"]},
    ]
    # eş potansiyel bağlantı noktaları (EPB1'e)
    T["eb"] = [
        {"poz": "EB1", "ad": "Su borusu girişi", "xy": (0.0, 14.5)},
        {"poz": "EB2", "ad": "Gaz borusu (izolasyon flanşı İF1 sonrası)", "xy": (4.8, b)},
        {"poz": "EB3", "ad": "Kablo tavası / PE", "xy": (2.0, 17.2)},
        {"poz": "EB4", "ad": "HVAC kanalı / klima gövdeleri", "xy": (a, 12.0)},
        {"poz": "EB5", "ad": "Minare çelik konstrüksiyonu", "xy": (mx - 0.6, my + 0.6)},
        {"poz": "EB6", "ad": "Kubbe/çatı çelik konstrüksiyonu", "xy": (kx, ky - 4.5)},
    ]
    T["yuvarlanan_kure"] = kure_zarfi(P, T)
    return T


def kure_zarfi(P, T):
    """2B yuvarlanan küre: yapı noktalarına R yarıçaplı kürenin değebildiği alt zarfı hesaplar."""
    a = T["bina_a_m"]
    R = float(T["kure_R"])
    h_g = T["govde_yuksekligi_m"]
    mx = T["minare_merkez_m"][0]
    mt = T["minare_taban_m"]
    kx, kr, hk = T["kubbe_merkez_m"][0], T["kubbe_cap_m"] / 2.0, T["bina_yuksekligi_m"] - h_g
    yk = {d["poz"]: d for d in T["yk"]}
    # yapı noktaları (kesit düzlemi: x yatay, z düşey)
    pts = []
    for x in np.arange(0, a + 1e-9, 0.1):            # çatı çevre iletkeni (hava sonlandırma) çatı düzlemi
        pts.append((x, h_g))
    for z in np.arange(0, h_g + 1e-9, 0.1):          # duvarlar
        pts += [(0.0, z), (a, z)]
    for ang in np.linspace(0, math.pi, 120):         # kubbe
        pts.append((kx - kr * math.cos(ang), h_g + hk * math.sin(ang)))
    for z in np.arange(h_g, T["minare_yuksekligi_m"] + 1e-9, 0.1):   # minare gövdesi
        pts += [(mx - mt / 2, z), (mx + mt / 2, z)]
    for x in np.arange(mx - mt / 2, mx + mt / 2 + 1e-9, 0.1):
        pts.append((x, T["minare_yuksekligi_m"]))
    for d in T["yk"]:                                # Franklin çubukları
        x = d["xy"][0] if d["poz"] != "YK3" else 0.0
        for z in np.arange(d["z"], d["z"] + d["boy"] + 1e-9, 0.05):
            pts.append((x, z))
    S = np.array(pts)
    # izinli küre merkezleri
    xs = np.arange(-4 * R, a + 4 * R, 0.5)
    zs = np.arange(R, R + 60.0, 0.5)
    X, Z = np.meshgrid(xs, zs)
    C = np.stack([X.ravel(), Z.ravel()], axis=1)
    izinli = np.ones(len(C), dtype=bool)
    for i in range(0, len(S), 400):
        blok = S[i:i + 400]
        d2 = (C[:, None, 0] - blok[None, :, 0]) ** 2 + (C[:, None, 1] - blok[None, :, 1]) ** 2
        izinli &= (d2.min(axis=1) >= R * R - 1e-6)
    C = C[izinli]
    xg = np.arange(-1.6 * R, a + 1.6 * R, 0.5)
    zarf = np.full(len(xg), np.inf)
    merkez = np.zeros((len(xg), 2))
    for k, x in enumerate(xg):
        dx = x - C[:, 0]
        m = np.abs(dx) <= R
        if not m.any():
            continue
        zalt = C[m, 1] - np.sqrt(R * R - dx[m] ** 2)
        j = int(np.argmin(zalt))
        zarf[k] = max(zalt[j], 0.0)
        merkez[k] = C[m][j]
    # yapı korunuyor mu?  (çatı çevre iletkeni, Franklin uçları hava sonlandırma kabul edilir)
    ihlal = 0.0
    for (x, z) in S:
        k = int(round((x - xg[0]) / 0.5))
        if 0 <= k < len(xg) and np.isfinite(zarf[k]):
            ihlal = max(ihlal, z - zarf[k])
    return {"xg": xg, "zarf": zarf, "merkez": merkez, "R": R, "ihlal_m": ihlal, "korunuyor": ihlal <= 0.26}


# =====================================================================
#  SVG ÇİZİM KATMANI (mm cinsinden, A3 yatay 420x297)
# =====================================================================
RENK = {"L1": "#7b4a1a", "L2": "#151515", "L3": "#8e8e8e", "N": "#1d6fd8",
        "PE_G": "#1f9d4b", "PE_Y": "#f1c40f", "K": "#d0021b", "CU": "#b4590f", "GRI": "#666666",
        "ACIK": "#e6e6e6", "ORTA": "#bdbdbd", "MAVI": "#1d4f91"}
A3W, A3H = 420.0, 297.0


class Cizim:
    def __init__(self, yol, no):
        self.yol, self.no = yol, no
        self.d = svgwrite.Drawing(yol, size=(f"{A3W}mm", f"{A3H}mm"), viewBox=f"0 0 {A3W} {A3H}", debug=False)
        self.d.add(self.d.rect((0, 0), (A3W, A3H), fill="white"))
        self.g = self.d
        self.pozlar = set()
        self.kutular = []          # metin sınır kutuları (çakışma denetimi için)

    # ---- temel şekiller
    def line(self, x1, y1, x2, y2, w=0.3, c="#000", dash=None, cap="butt"):
        e = self.d.line((x1, y1), (x2, y2), stroke=c, stroke_width=w, stroke_linecap=cap)
        if dash:
            e["stroke-dasharray"] = dash
        self.g.add(e)

    def rect(self, x, y, w, h, sw=0.3, stroke="#000", fill="none", dash=None, rx=0):
        e = self.d.rect((x, y), (w, h), stroke=stroke if sw else "none", stroke_width=sw, fill=fill)
        if rx:
            e["rx"] = rx
            e["ry"] = rx
        if dash:
            e["stroke-dasharray"] = dash
        self.g.add(e)

    def circle(self, cx, cy, r, sw=0.3, stroke="#000", fill="none", dash=None):
        e = self.d.circle((cx, cy), r, stroke=stroke if sw else "none", stroke_width=sw, fill=fill)
        if dash:
            e["stroke-dasharray"] = dash
        self.g.add(e)

    def ellipse(self, cx, cy, rx, ry, sw=0.3, stroke="#000", fill="none", dash=None):
        e = self.d.ellipse((cx, cy), (rx, ry), stroke=stroke, stroke_width=sw, fill=fill)
        if dash:
            e["stroke-dasharray"] = dash
        self.g.add(e)

    def poly(self, pts, sw=0.3, stroke="#000", fill="none", closed=False, dash=None):
        if closed:
            e = self.d.polygon(pts, stroke=stroke, stroke_width=sw, fill=fill)
        else:
            e = self.d.polyline(pts, stroke=stroke, stroke_width=sw, fill=fill)
        e["stroke-linejoin"] = "round"
        if dash:
            e["stroke-dasharray"] = dash
        self.g.add(e)

    def path(self, dstr, sw=0.3, stroke="#000", fill="none", dash=None):
        e = self.d.path(d=dstr, stroke=stroke, stroke_width=sw, fill=fill)
        if dash:
            e["stroke-dasharray"] = dash
        self.g.add(e)

    def arc(self, cx, cy, r, a0, a1, sw=0.3, stroke="#000", dash=None):
        """a0->a1 derece (matematik yönü, y aşağı olduğu için işaret ayarlı), saat yönünde çizer."""
        x0, y0 = cx + r * math.cos(math.radians(a0)), cy - r * math.sin(math.radians(a0))
        x1, y1 = cx + r * math.cos(math.radians(a1)), cy - r * math.sin(math.radians(a1))
        large = 1 if abs(a1 - a0) > 180 else 0
        sweep = 0 if a1 > a0 else 1
        self.path(f"M {x0:.3f} {y0:.3f} A {r:.3f} {r:.3f} 0 {large} {sweep} {x1:.3f} {y1:.3f}", sw=sw, stroke=stroke, dash=dash)

    # ---- metin
    def t(self, x, y, s, size=2.4, anchor="start", bold=False, color="#000", rot=0, italic=False, poz=None,
          valign="base", maxw=None, opacity=None, halo=False, hat=None):
        s = str(s)
        if halo and not rot:
            w_ = tw(s, size, bold)
            x0_ = x if anchor == "start" else (x - w_ if anchor == "end" else x - w_ / 2)
            self.g.add(self.d.rect((x0_ - 0.4, y - size * 0.8), (w_ + 0.8, size * 1.0), fill="white", stroke="none"))
        if maxw is not None:
            w = tw(s, size, bold)
            if w > maxw + 1e-6:
                uyar(f"[{self.no}] Metin taşıyor ({w:.1f}>{maxw:.1f} mm): '{s}'")
        if valign == "mid":
            y = y + size * 0.35
        e = self.d.text(s, insert=(x, y), font_size=size, font_family=FONT_AILE, text_anchor=anchor, fill=color)
        if bold:
            e["font-weight"] = "bold"
        if italic:
            e["font-style"] = "italic"
        if rot:
            e["transform"] = f"rotate({rot},{x:.3f},{y:.3f})"
        else:
            w_ = tw(s, size, bold)
            x0_ = x if anchor == "start" else (x - w_ if anchor == "end" else x - w_ / 2)
            self.kutular.append((x0_, y - size * 0.74, x0_ + w_, y + size * 0.16, s))
        if opacity is not None:
            e["fill-opacity"] = opacity
        if poz:
            lst = [poz] if isinstance(poz, str) else list(poz)
            e["data-poz"] = " ".join(lst)
            self.pozlar.update(lst)
            if hat:
                e["data-hat"] = hat
        self.g.add(e)

    def tl(self, x, y, satirlar, size=2.4, lh=1.25, **kw):
        """Çok satırlı metin; her satır ayrı t()."""
        if isinstance(satirlar, str):
            satirlar = satirlar.split("\n")
        for i, s in enumerate(satirlar):
            self.t(x, y + i * size * lh, s, size=size, **kw)
        return len(satirlar) * size * lh

    def tw_(self, x, y, s, maxw, size=2.4, lh=1.25, **kw):
        """Sarmalı metin; bitiş y'sini döndürür."""
        lines = wrap(s, maxw, size, kw.get("bold", False))
        self.tl(x, y, lines, size=size, lh=lh, **kw)
        return y + len(lines) * size * lh

    # ---- iletkenler
    def iletken(self, x1, y1, x2, y2, renk, w=0.35):
        if renk == "PE":
            self.line(x1, y1, x2, y2, w=w, c=RENK["PE_G"])
            self.line(x1, y1, x2, y2, w=w, c=RENK["PE_Y"], dash="1.4,1.4")
        else:
            self.line(x1, y1, x2, y2, w=w, c=RENK.get(renk, renk))

    def nokta(self, x, y, r=0.6, c="#000"):
        self.circle(x, y, r, sw=0, fill=c)

    # ---- IEC 60617 sembolleri (düşey akışlı)
    def topraklama(self, x, y, pe=False, olcek=1.0, renk="#000"):
        """Genel toprak sembolü; (x,y) üst bağlantı noktası."""
        s = olcek
        self.line(x, y, x, y + 2.0 * s, c=renk, w=0.3)
        self.line(x - 2.0 * s, y + 2.0 * s, x + 2.0 * s, y + 2.0 * s, c=renk, w=0.3)
        self.line(x - 1.3 * s, y + 3.0 * s, x + 1.3 * s, y + 3.0 * s, c=renk, w=0.3)
        self.line(x - 0.6 * s, y + 4.0 * s, x + 0.6 * s, y + 4.0 * s, c=renk, w=0.3)
        return y + 4.0 * s

    def kutup(self, px, yc, h=2.4, renk="#000"):
        """Kesici kutbu: bıçak + çapraz (IEC 60617 devre kesici)."""
        self.line(px, yc + h, px + 1.7, yc - h + 0.7, c=renk, w=0.32)
        self.line(px - 0.9, yc - h - 0.9, px + 0.9, yc - h + 0.9, c=renk, w=0.3)
        self.line(px - 0.9, yc - h + 0.9, px + 0.9, yc - h - 0.9, c=renk, w=0.3)

    def kutup_kontaktor(self, px, yc, h=2.4, renk="#000"):
        """Kontaktör kutbu: bıçak + yarım daire (IEC 60617 kontaktör)."""
        self.line(px, yc + h, px + 1.7, yc - h + 0.7, c=renk, w=0.32)
        self.path(f"M {px - 1.0:.2f} {yc - h:.2f} A 1.0 1.0 0 0 1 {px + 1.0:.2f} {yc - h:.2f}", sw=0.3)

    def projektor(self, x, y, renk="#000", yon=0, sc=1.0):
        """LED projektör sembolü: gövde + ışın çizgileri; yon: derece (0 = sağa)."""
        a = math.radians(yon)
        ca, sa = math.cos(a), math.sin(a)
        def d(u, v):
            u, v = u * sc, v * sc
            return (x + u * ca - v * sa, y + u * sa + v * ca)
        self.poly([d(-1.4, -1.2), d(1.0, -1.2), d(1.0, 1.2), d(-1.4, 1.2)], closed=True, sw=0.3, stroke=renk, fill="#fff")
        for v in (-1.1, 0.0, 1.1):
            p0, p1 = d(1.4, v * 0.6), d(3.8, v * 1.8)
            self.line(p0[0], p0[1], p1[0], p1[1], w=0.25, c=renk)

    def avize(self, x, y, r=2.6, renk="#000"):
        """Avize sembolü: büyük lamba + kollar."""
        self.lamba(x, y, r=r, renk=renk)
        for k in range(8):
            ang = math.radians(45 * k + 22.5)
            x0, y0 = x + r * math.cos(ang), y + r * math.sin(ang)
            x1, y1 = x + (r + 1.6) * math.cos(ang), y + (r + 1.6) * math.sin(ang)
            self.line(x0, y0, x1, y1, w=0.25, c=renk)
            self.circle(x1, y1, 0.5, sw=0, fill=renk)

    def cihaz(self, xs, renkler, y1, y2, tur="mcb", pad=3.0, yc=None, dolgu="#fff", govde=True):
        """Çok kutuplu koruma cihazı. tur: mccb, mcb, mpcb, rcbo, rcd, sigorta, ayirici.
        Dönüş: (x_sol, x_sag, yc)"""
        xl, xr = min(xs) - pad, max(xs) + pad
        if govde:
            self.rect(xl, y1, xr - xl, y2 - y1, sw=0.3, fill=dolgu)
        if yc is None:
            yc = y1 + (y2 - y1) * (0.30 if tur in ("rcbo", "rcd") else 0.5)
        h = 2.0 if tur in ("rcbo", "rcd") else (2.4 if (y2 - y1) >= 9 else 2.0)
        for x, c in zip(xs, renkler):
            col = RENK.get(c, c)
            if tur == "sigorta":
                self.line(x, y1, x, y2, c=col, w=0.35)
                self.rect(x - 0.9, yc - 2.6, 1.8, 5.2, sw=0.3, fill="#fff")
                self.line(x, yc - 2.6, x, yc + 2.6, c="#000", w=0.25)
                continue
            self.line(x, y1, x, yc - h, c=col, w=0.35)
            self.line(x, yc + h, x, y2, c=col, w=0.35)
            if tur == "kontaktor":
                self.kutup_kontaktor(x, yc, h=h)
            else:
                self.kutup(x, yc, h=h)
        if tur != "sigorta" and len(xs) > 1:
            self.line(xs[0] + 0.85, yc, xs[-1] + 0.85, yc, w=0.2, dash="0.7,0.5")
        if tur in ("rcbo", "rcd"):
            yr = yc + h + 2.9
            self.ellipse((xs[0] + xs[-1]) / 2, yr, (xs[-1] - xs[0]) / 2 + 1.8, 1.3, sw=0.3)
            self.line((xs[0] + xs[-1]) / 2 + (xs[-1] - xs[0]) / 2 + 1.8, yr, xr - 0.6, yr, w=0.2, dash="0.5,0.4")
        if tur == "mpcb":
            yt = yc + h + 1.4
            for x in xs:
                self.rect(x - 0.9, yt - 0.9, 1.8, 1.8, sw=0.25, fill="#fff")
        if tur == "mccb":
            self.t(xl + 0.6, y2 - 0.9, "I>", size=1.9, bold=True)
        return xl, xr, yc

    def spd_kutu(self, x, y, genislik=2.6, yukseklik=5.0):
        """Parafudr (varistör) sembolü: dikdörtgen + diyagonal ok; (x,y) üst orta."""
        self.rect(x - genislik / 2, y, genislik, yukseklik, sw=0.3, fill="#fff")
        self.line(x - 1.8, y + yukseklik * 0.75, x + 1.8, y + yukseklik * 0.25, w=0.3)
        self.poly([(x + 1.8, y + yukseklik * 0.25), (x + 0.9, y + yukseklik * 0.22), (x + 1.45, y + yukseklik * 0.5)],
                  sw=0.2, fill="#000", closed=True)

    def akim_trafosu(self, x, y, r=2.0, renk="#000"):
        self.circle(x, y, r, sw=0.3, fill="#fff", stroke=renk)
        self.line(x, y - r - 1.2, x, y + r + 1.2, c=RENK.get(renk, renk) if renk != "#000" else "#000", w=0.0)

    def motor(self, x, y, r=5.0, faz="3~"):
        self.circle(x, y, r, sw=0.35, fill="#fff")
        self.t(x, y - 0.4, "M", size=r * 0.9, anchor="middle", bold=True)
        self.t(x, y + r * 0.62, faz, size=r * 0.55, anchor="middle")

    def olcu_aleti(self, x, y, harf, r=3.0):
        self.circle(x, y, r, sw=0.35, fill="#fff")
        self.t(x, y + r * 0.34, harf, size=r * 1.05, anchor="middle", bold=True)

    def lamba(self, x, y, r=2.0, renk="#000"):
        self.circle(x, y, r, sw=0.3, fill="#fff", stroke=renk)
        d = r * 0.707
        self.line(x - d, y - d, x + d, y + d, w=0.3, c=renk)
        self.line(x - d, y + d, x + d, y - d, w=0.3, c=renk)

    def priz(self, x, y, r=2.2):
        """Topraklı priz sembolü (yarım daire + toprak çizgileri), (x,y) bağlantı noktası üst."""
        self.line(x, y, x, y + 2.0, w=0.3)
        self.path(f"M {x - r:.2f} {y + 2.0 + r:.2f} A {r:.2f} {r:.2f} 0 0 1 {x + r:.2f} {y + 2.0 + r:.2f}", sw=0.35)
        self.line(x - r, y + 2.0 + r, x + r, y + 2.0 + r, w=0.35)
        self.line(x, y + 2.0 + r, x, y + 2.0 + r + 1.6, w=0.35)

    def hoparlor(self, x, y):
        """IEC hoparlör: küçük dikdörtgen + konik."""
        self.rect(x - 1.6, y - 1.6, 1.6, 3.2, sw=0.3, fill="#fff")
        self.poly([(x, y - 1.6), (x + 2.8, y - 3.4), (x + 2.8, y + 3.4), (x, y + 1.6)], sw=0.3, fill="#fff", closed=True)

    def zil(self, x, y):
        """Yangın alarm sireni/zil: yarım daire + vurgu."""
        self.path(f"M {x - 2.6:.2f} {y + 1.6:.2f} A 2.6 2.6 0 0 1 {x + 2.6:.2f} {y + 1.6:.2f} Z", sw=0.3, fill="#fff")
        self.circle(x, y + 2.9, 0.6, sw=0, fill="#000")

    def trafo(self, x, y, r=5.5, ust="D", alt="Yn"):
        """İki sargılı transformatör: üst (D=üçgen) alt (Y=yıldız + N). (x,y)= üst daire merkezi, alt daire merkezi y+0.65*2r."""
        dy = r * 1.3
        self.circle(x, y, r, sw=0.4, fill="#fff")
        self.circle(x, y + dy, r, sw=0.4, fill="#fff")
        # üçgen (D)
        s = r * 0.5
        self.poly([(x, y - s * 0.9), (x - s, y + s * 0.65), (x + s, y + s * 0.65)], sw=0.3, closed=True)
        # yıldız (Yn)
        yy = y + dy
        for ang in (90, 210, 330):
            self.line(x, yy, x + s * 0.95 * math.cos(math.radians(ang)), yy - s * 0.95 * math.sin(math.radians(ang)), w=0.3)
        return y + dy + r

    def ok(self, x1, y1, x2, y2, c="#000", w=0.3, boy=1.8):
        """Ok başlı çizgi."""
        self.line(x1, y1, x2, y2, c=c, w=w)
        a = math.atan2(y2 - y1, x2 - x1)
        for da in (2.6, -2.6):
            self.line(x2, y2, x2 - boy * math.cos(a + da * 0.3), y2 - boy * math.sin(a + da * 0.3), c=c, w=w)

    def kopma(self, x, y, yatay=True, c="#000"):
        """Kopuk çizgi (ölçekli değil) işareti: iki eğik çizgi."""
        if yatay:
            self.line(x - 1.2, y + 1.6, x - 0.2, y - 1.6, c=c, w=0.3)
            self.line(x + 0.2, y + 1.6, x + 1.2, y - 1.6, c=c, w=0.3)
        else:
            self.line(x - 1.6, y - 1.2, x + 1.6, y - 0.2, c=c, w=0.3)
            self.line(x - 1.6, y + 0.2, x + 1.6, y + 1.2, c=c, w=0.3)

    def olcu(self, x1, y1, x2, y2, metin, ofs=0.0, size=2.2, yatay=True, ters=False):
        """Ölçü çizgisi (ok uçlu), metin ortada. ofs: ölçü çizgisinin ötelenmesi (yatay için y, düşey için x)."""
        if yatay:
            ya = y1 + ofs
            self.line(x1, y1, x1, ya + (1.5 if ofs > 0 else -1.5), w=0.15, c="#444")
            self.line(x2, y2, x2, ya + (1.5 if ofs > 0 else -1.5), w=0.15, c="#444")
            self.line(x1, ya, x2, ya, w=0.2, c="#222")
            self.poly([(x1, ya), (x1 + 1.8, ya - 0.5), (x1 + 1.8, ya + 0.5)], sw=0, fill="#222", closed=True)
            self.poly([(x2, ya), (x2 - 1.8, ya - 0.5), (x2 - 1.8, ya + 0.5)], sw=0, fill="#222", closed=True)
            self.rect((x1 + x2) / 2 - tw(metin, size) / 2 - 0.6, ya - size * 0.5 - 0.6, tw(metin, size) + 1.2, size + 1.0, sw=0, fill="#fff")
            self.t((x1 + x2) / 2, ya + size * 0.33, metin, size=size, anchor="middle")
        else:
            xa = x1 + ofs
            self.line(x1, y1, xa + (1.5 if ofs > 0 else -1.5), y1, w=0.15, c="#444")
            self.line(x2, y2, xa + (1.5 if ofs > 0 else -1.5), y2, w=0.15, c="#444")
            self.line(xa, y1, xa, y2, w=0.2, c="#222")
            self.poly([(xa, y1), (xa - 0.5, y1 + 1.8), (xa + 0.5, y1 + 1.8)], sw=0, fill="#222", closed=True)
            self.poly([(xa, y2), (xa - 0.5, y2 - 1.8), (xa + 0.5, y2 - 1.8)], sw=0, fill="#222", closed=True)
            ym = (y1 + y2) / 2
            self.rect(xa - size * 0.5 - 0.6, ym - tw(metin, size) / 2 - 0.6, size + 1.0, tw(metin, size) + 1.2, sw=0, fill="#fff")
            self.t(xa + size * 0.33, ym, metin, size=size, anchor="middle", rot=-90)

    # ---- çerçeve + antet
    def cerceve(self, P, cizim_adi, sayfa_no, olcek_yazisi, aciklama=""):
        pr = P["proje"]
        self.rect(10, 10, 400, 277, sw=0.6)
        self.rect(5, 5, 410, 287, sw=0.18)
        x0, y0, w, h = 252, 250, 158, 37
        self.rect(x0, y0, w, h, sw=0.5, fill="#fff")
        # satırlar
        self.line(x0, y0 + 9, x0 + w, y0 + 9, w=0.3)
        self.line(x0, y0 + 21, x0 + w, y0 + 21, w=0.3)
        self.line(x0, y0 + 29, x0 + w, y0 + 29, w=0.3)
        self.line(x0 + 112, y0, x0 + 112, y0 + 9, w=0.3)
        self.t(x0 + 3, y0 + 6.3, pr["ad"], size=4.0, bold=True, maxw=106)
        self.t(x0 + 114.5, y0 + 6.3, pr["sistem_yazisi"], size=4.0, bold=True, maxw=42)
        self.t(x0 + 3, y0 + 13.5, "ÇİZİM ADI", size=1.9, color="#555")
        self.t(x0 + 3, y0 + 19.3, cizim_adi, size=4.2, bold=True, maxw=152)
        # alt hücreler
        kol = [("ÇİZİM NO", pr["cizim_no"][self.no], 0, 38), ("ÖLÇEK", olcek_yazisi, 38, 44), ("TARİH", pr["tarih"], 82, 28),
               ("REVİZYON", pr["revizyon"], 110, 22), ("SAYFA", f"{sayfa_no}/{pr['cizim_sayisi']}", 132, 26)]
        for ad, deger, ox, gw in kol:
            if ox:
                self.line(x0 + ox, y0 + 21, x0 + ox, y0 + 29, w=0.25)
            self.t(x0 + ox + 1.5, y0 + 24.2, ad, size=1.7, color="#555")
            self.t(x0 + ox + 1.5, y0 + 27.7, deger, size=2.5, bold=True, maxw=gw - 2.5)
        self.line(x0 + 79, y0 + 29, x0 + 79, y0 + h, w=0.25)
        self.t(x0 + 1.5, y0 + 32.4, "ÇİZEN", size=1.7, color="#555")
        self.t(x0 + 1.5, y0 + 35.6, pr["cizen"], size=2.4, bold=True, maxw=75)
        self.t(x0 + 80.5, y0 + 32.4, "KONTROL", size=1.7, color="#555")
        self.t(x0 + 80.5, y0 + 35.6, pr["kontrol"], size=2.4, bold=True, maxw=76)
        if aciklama:
            self.t(x0, y0 - 1.6, aciklama, size=1.9, color="#555", maxw=w)

    def cakismalar(self, tolerans=0.25):
        """Üst üste binen metin kutularını döndürür."""
        sonuc = []
        k = self.kutular
        for i in range(len(k)):
            for j in range(i + 1, len(k)):
                a, b = k[i], k[j]
                ox = min(a[2], b[2]) - max(a[0], b[0])
                oy = min(a[3], b[3]) - max(a[1], b[1])
                if ox > tolerans and oy > tolerans:
                    sonuc.append((a[4], b[4], round(a[0], 1), round(a[1], 1)))
        return sonuc

    def kaydet(self):
        self.d.save()
        pdf = os.path.splitext(self.yol)[0] + ".pdf"
        cairosvg.svg2pdf(url=self.yol, write_to=pdf)
        return pdf


# =====================================================================
#  POZ KAYDI (tek kaynak: çizimler ve tablolar bu listeden beslenir)
# =====================================================================
def poz_kayit(P, M, PN, T):
    """Her poz için: poz, ad, adet, deger, kat (pano|saha|topraklama), cizim (beklenen çizim no kümesi), grup."""
    C = P["cihazlar"]
    S = P["sistem"]
    kayit = []

    def ekle(poz, ad, adet, deger, kat, cizim, grup=None, **ek):
        d_ = {"poz": poz, "ad": ad, "adet": adet, "deger": deger, "kat": kat, "cizim": set(cizim), "grup": grup or poz}
        d_.update(ek)
        kayit.append(d_)

    G, Y, D = PN["G"], PN["Y"], P["pano"]["D"]
    ekle("PN1", "Ana dağıtım panosu, zemine monte, metal gövde, boyalı, çok bölmeli", 1,
         f"{G}x{Y}x{D} mm (GxYxD), {P['pano']['IP']}", "pano", {1, 2, 3, 6})
    ekle("TR1", "Güç transformatörü (OG tarafı blok gösterim)", 1,
         f"{tr(S['trafo_kva'], 0)} kVA, {S['trafo_grubu']}, uk=%{tr(S['trafo_uk_yuzde'])}, OG/0,4 kV", "saha", {1, 3})
    b = M["besleme"]
    ekle(b["kablo_no"], "Trafo - Q0 besleme kablosu", 1,
         f"{b['tip']} {b['damar'].replace('x', 'x')} mm² Cu, L={tr(b['L'])} m, gömülü", "saha", {1, 2, 3})
    pe3 = b["pe_iletkeni"]
    ekle(pe3["poz"], "Trafo yıldız noktası - PN1 PE barası koruma iletkeni", 1,
         f"{pe3['tip']}, L={tr(b['L'])} m (W0 ile aynı hendek)", "saha", {1, 3})
    te = T["trafo_elektrodu"]
    ekle("TE1", "Trafo yıldız noktası topraklama elektrodu (ayrı elektrot)", 1,
         f"{te['cubuk_adet']}x Ø{tr(T['cubuk_cap_mm'], 0)}x{tr(te['cubuk_boy_m'], 0)} m Cu kaplı çubuk, hedef ≤ {tr(T['hedef_direnc_ohm'], 0)} Ω",
         "topraklama", {1, 3})
    ekle("PE1", "PN1 PE barası (BPE) - EPB1 bağlantı iletkeni", 1,
         f"50 mm² Cu (sarı-yeşil), L={tr(T['pano_epb_hatti_m'])} m", "topraklama", {1, 3})
    ekle("PE2", "EPB1 - topraklama halkası (TH1) bağlantı iletkeni", 1,
         f"50 mm² Cu (sarı-yeşil), L={tr(T['epb_halka_hatti_m'])} m", "topraklama", {1, 3})
    ekle("EPB1", "Esas eş potansiyel barası (ana pano yanında)", 1, "Cu bara 40x5 mm, 12 klemens, ayırma bağlantılı", "topraklama", {1, 3})

    # ---- ana giriş ve koruma / ölçüm
    q0 = C["Q0"]
    ekle("Q0", q0["ad"], 1, f"{q0['govde_A']} A gövde, Ir={q0['Ir_A']} A, Icu={tr(q0['Icu_kA'], 0)} kA, Isd={q0['Isd_x']}xIr/{q0['tsd_ms']} ms, Ii {q0['Ii']}", "pano", {1, 2})
    ekle("MX0", C["MX0"]["ad"], 1, "230 V a.a., Q0 yardımcı kontak/shunt trip", "pano", {1, 2})
    r0 = C["RCR0"]
    ekle("RCR0", r0["ad"], 1, f"I∆n={tr(r0['I_dn_A'], 1)} A (ayar {r0['aralik_I']}), gecikme {r0['gecikme_ms']} ms (ayar {r0['aralik_t']})", "pano", {1, 2})
    ekle("TOR0", C["TOR0"]["ad"], 1, f"Ø{C['TOR0']['cap_mm']} mm iç çap, RCR0'a bağlı", "pano", {1, 2})
    s1 = C["SPD1"]
    ekle("SPD1", s1["ad"], 1, f"Iimp={tr(s1['I_imp_kA'])} kA/kutup, In={tr(s1['In_kA'], 0)} kA, Up≤{tr(s1['Up_kV'])} kV, Uc={s1['Uc_V']} V", "pano", {1, 2})
    ekle("F1", C["F1"]["ad"], 1, f"{C['F1']['In_A']} A gG", "pano", {1, 2})
    ekle("FSR1", C["FSR1"]["ad"], 1, C["FSR1"]["ayar"], "pano", {1, 2})
    ekle("F2", C["F2"]["ad"], 1, f"{C['F2']['egri']}{C['F2']['In_A']} A, {tr(C['F2']['Icu_kA'], 0)} kA", "pano", {1, 2})
    for i in (1, 2, 3):
        ekle(f"CT{i}", "Akım trafosu (L1, L2, L3)", 1, "200/5 A, 5 VA, Sınıf 0,5", "pano", {1, 2}, grup="CT")
    ekle("PM1", C["PM1"]["ad"], 1, "kapıda, 96x96 mm, CT sekonder 5 A", "pano", {1, 2})
    ekle("A1", C["A1"]["ad"], 1, "kapıda", "pano", {1, 2})
    ekle("V1", C["V1"]["ad"], 1, "kapıda", "pano", {1, 2})
    ekle("SA1", C["SA1"]["ad"], 1, "kapıda, 4 konumlu", "pano", {1, 2})
    ekle("SA2", C["SA2"]["ad"], 1, "kapıda, 7 konumlu", "pano", {1, 2})
    for i, rn in ((1, "L1 kahverengi"), (2, "L2 siyah"), (3, "L3 gri")):
        ekle(f"H{i}", "Faz sinyal lambası (L1, L2, L3)", 1, "Ø22 mm LED, 230 V a.a. (kapıda)", "pano", {1, 2}, grup="H")
    ekle("SB1", C["SB1"]["ad"], 1, "kapıda, kırmızı/sarı, kilitli, Q0 MX0'a bağlı", "pano", {1, 2})
    if any(h["grup"] == "PROJ" for h in M["hatlar"]):
        q3, km1, ft1, zr1 = C["Q3"], C["KM1"], C["FT1"], C["ZR1"]
        ekle("Q3", "Dış aydınlatma besleme şalteri (MCB 4P)", 1, f"4P {q3['egri']}{q3['In_A']} A, {tr(q3['Icu_kA'], 0)} kA", "pano", {1, 2})
        ekle("KM1", "Dış aydınlatma kontaktörü (3P)", 1, f"3P {km1['In_A']} A (AC-1/AC-5a), bobin 230 V, {km1['modul']} modül", "pano", {1, 2})
        ekle("FT1", "Fotosel (alacakaranlık anahtarı)", 1, f"DIN ray {ft1['modul']} modül + harici sensör IP65, ayar {ft1['lux_ayar']}", "pano", {1, 2})
        ekle("ZR1", "Astronomik zaman rölesi", 1, f"DIN ray {zr1['modul']} modül, enlem/boylam programlı, yedek pil", "pano", {1, 2})
        ekle("SA3", "MAN-0-OTO seçici anahtar (kapıda)", 1, "3 konumlu, kapı üzeri Ø22 mm", "pano", {1, 2})
        ekle("H4", "Dış aydınlatma durum lambası (kapıda)", 1, "Ø22 mm LED, 230 V a.a., yeşil", "pano", {1, 2})
        if P["aydinlatma_priz"]["projektor"].get("spd2_goster", True):
            sp2 = C["SPD2"]
            ekle("SPD2", "Parafudr Tip 2, 3+1 (dış hatlar)", 1,
                 f"In={tr(sp2['I_n_kA'], 0)} kA, Imax={tr(sp2['Imax_kA'], 0)} kA, Up≤{tr(sp2['Up_kV'])} kV, Uc={sp2['Uc_V']} V; yedek koruma Q3", "pano", {1, 2})
    ekle("XT1", "Ölçüm klemens sırası (CT sekonder, gerilim izleme)", 1, "12 klemens, CT kısa devre klemensli", "pano", {1, 2})
    for poz, ad in (("BL1", "L1"), ("BL2", "L2"), ("BL3", "L3")):
        ekle(poz, "Ana bara (L1, L2, L3)", 1, "Cu bara 30x5 mm, 250 A, Icw≥25 kA/1 s", "pano", {1, 2}, grup="BL")
    ekle("BN", "N barası (pano gövdesinden yalıtılmış)", 1, "Cu bara 30x5 mm, 250 A, N klemensli", "pano", {1, 2})
    ekle("BPE", "PE barası (pano gövdesine bağlı)", 1, "Cu bara 30x5 mm, PE klemensli", "pano", {1, 2})

    # ---- klima grupları
    qg, rc, kk = C["Q_klima_grup"], C["RCD_klima"], C["K_klima"]
    for gi in (1, 2):
        ekle(f"Q{gi}", "Klima grup besleme şalteri (MCB 4P)", 1, f"4P {qg['egri']}{qg['In_A']} A, {tr(qg['Icu_kA'], 0)} kA", "pano", {1, 2}, grup="Qg")
    for gi in (1, 2):
        ekle(f"RCD{gi}", "Klima grup kaçak akım rölesi (RCD 4P)", 1,
             f"4P {rc['In_A']} A, {rc['I_dn_mA']} mA, Tip {rc['tip']}, Inc={tr(rc['Inc_kA'], 0)} kA", "pano", {1, 2}, grup="RCD")
    for h in M["hatlar"]:
        if klima_mi(h["id"]):
            ekle(h["id"], "Klima motor koruma şalteri (MPCB 3P)", 1,
                 f"3P MPCB {kk['ayar_min_A']}-{kk['ayar_max_A']} A (ayar {kk['ayar_A']} A), Icu={tr(kk['Icu_kA'], 0)} kA; alternatif C20", "pano", {1, 2}, grup="K")
    rcbo_ad = {"AVIZE": "Avize hattı RCBO", "PRIZ": "Priz hattı RCBO", "LED": "İç LED hattı RCBO", "PROJ": "Dış projektör hattı RCBO",
               "KAMERA": "Kamera hattı RCBO", "SES": "Ses sistemi hattı RCBO", "DEKOR": "Dekoratif aydınlatma RCBO"}
    if any(h.get("koruma_tur") == "mcb" for h in M["hatlar"]):
        r3 = C["RCD3"]
        ekle("RCD3", "Ses sistemi ortak kaçak akım rölesi (RCD 2P)", 1,
             f"2P {r3['In_A']} A, {r3['I_dn_mA']} mA, Tip {r3['tip']}, Inc={tr(r3['Inc_kA'], 0)} kA", "pano", {1, 2}, grup="RCD3")
    for h in M["hatlar"]:
        if h.get("koruma_tur") == "mcb":
            ekle(h["id"], "Ses sistemi hattı MCB (RCD3 altında)", 1,
                 f"1P+N {C['MCB1']['egri']}{h['rcbo_In']} A, {tr(C['MCB1']['Icu_kA'], 0)} kA", "pano", {1, 2, 6}, grup="MCB_SES")
            continue
        if not klima_mi(h["id"]):
            ad_ = rcbo_ad.get(h["grup"]) or ("Acil aydınlatma RCBO" if h["id"] == "AC1" else "Yangın alarm hattı RCBO")
            ek_ = ", LED sürücü inrush için C eğrisi" if h["grup"] == "LED" else ""
            ekle(h["id"], ad_, 1,
                 f"1P+N {P['rcbo_defaults']['egri']}{h['rcbo_In']} A, {C['RCBO']['I_dn_mA']} mA, Tip {C['RCBO']['tip']}, {tr(C['RCBO']['Icn_kA'], 0)} kA{ek_}",
                 "pano", {1, 2, 6} if (h.get("elemanlar") or h["grup"] == "SES") else {1, 2}, grup="RCBO_" + h["grup"])
    for yd in P["yedek_rcbo"]:
        ekle(yd, "Yedek RCBO yuvası (boş, DIN ray + bara hazır)", 1, "1P+N RCBO için yer ayrıldı (2 modül)", "pano", {1, 2}, grup="YD")

    # ---- saha ekipmanı (pano dışı)
    for h in M["hatlar"]:
        if klima_mi(h["id"]):
            ekle(h["yuk"], "Dolap tipi klima (3 faz)", 1, f"{tr(h['p_kw'])} kW, 400 V, cosφ={tr(h['cos'], 2)}", "saha", {1}, grup="M")
    ekle("UPS1", C["UPS1"]["ad"], 1, f"{tr(C['UPS1']['kva'])} kVA, 230 V, kamera/NVR için", "saha", {1})
    if P["aydinlatma_priz"]["avize"].get("dimmer_goster") and P["aydinlatma_priz"]["avize"]["adet"] > 0:
        ekle("DM1", C["DM1"]["ad"], 1, "avize hattı AV1 üzerinde, 2,5 kW LED uyumlu", "saha", {1})
    ftip = {"priz": ("Priz girişi, 16 A, 230 V, topraklı (IP20)", "priz"), "led": ("İç mekan LED armatür/lamba", "led"),
            "projektor": ("LED projektör, IP65", "projektor"), "avize": ("Büyük avize + bağlantı ünitesi", "avize")}
    for e in M["elemanlar"]:
        ad_, tp_ = ftip[e["tip"]]
        ekle(e["poz"], ad_, 1, f"{tr(e['w'], 0)} W, hat {e['hat']}", "saha", {1, 6}, grup=e["hat"] + "_" + e["tip"], envanter=True, hat=e["hat"])
    pj_hatlar = [h for h in M["hatlar"] if h["grup"] == "PROJ"]
    for i, h in enumerate(pj_hatlar, start=1):
        ekle(f"JB{i}", "IP65 bağlantı kutusu (dış hat)", 1, f"hat {h['id']}, IP65, 3x2,5 mm² geçişli", "saha", {6}, grup="JB", envanter=True, hat=h["id"])
    ekle("KT1", "Kablo tavası (iç mekan, perfore, galvaniz)", 1, "100x50 mm, kapaklı; ana kablo yolları", "saha", {6}, envanter=True, hat="-")
    for h in M["hatlar"]:
        ekle(h["kablo_no"], "Alıcı besleme kablosu (pano - alıcı)", 1,
             f"{h['tip']} {h['damar']} mm² Cu, L={tr(h['L'])} m", "saha", {1, 6} if h["grup"] == "SES" else {1}, grup="W_SES" if h["grup"] == "SES" else "W")

    # ---- pano mekanik (yalnız çizim 2)
    ekle("MP1", "Montaj plakası (galvaniz, 2 mm)", 1, f"{G - 2 * P['pano']['plaka_kenar_mm']}x{Y - P['pano']['cokul_mm'] - 100} mm", "pano", {2})
    ekle("DR1", "DIN ray 35x7,5 mm (delikli)", len(PN["satir_hesap"]) + 1, f"{int(PN['ray_uzunlugu_mm'])} mm", "pano", {2})
    ekle("KN1", "Kablo kanalı (dikey/yatay), kapaklı, PVC halojensiz", 1, "60x80 mm dikey, 25x60 mm yatay", "pano", {2})
    ekle("GR1", "Havalandırma ızgarası + filtre (IP54)", 2, "Kapı alt/üst, 200x300 mm", "pano", {2})
    ekle("KL1", "Kapı kilidi (3 noktalı, kollu)", 1, "Çok noktalı kilit, anahtarlı", "pano", {2})
    ekle("RK0", "Kablo rakoru W0 giriş (alt plaka)", 1, "M63, IP68", "pano", {2})
    n_klima = sum(1 for h in M["hatlar"] if klima_mi(h["id"]))
    n_diger = sum(1 for h in M["hatlar"] if not klima_mi(h["id"]))
    ekle("RK1", "Kablo rakoru klima çıkışları (alt plaka)", n_klima, "M32, IP68", "pano", {2})
    ekle("RK2", "Kablo rakoru diğer çıkışlar (alt plaka)", n_diger, "M25, IP68", "pano", {2})
    if PN["isil"]["fan_gerekli"]:
        ekle("FN1", "Fanlı filtre ünitesi", 1, f"≥ {tr(PN['isil']['hava_m3h'], 0)} m³/h", "pano", {2})
        ekle("TS1", "Pano içi termostat", 1, "0-60 °C, ayar 35 °C", "pano", {2})

    # ---- topraklama / yıldırımdan korunma (çizim 3)
    ekle("TH1", "Topraklama halkası iletkeni (bina çevresi, derinlik %s m)" % tr(T["halka_derinlik_m"]), 1,
         f"{T['halka_iletkeni']}, ≈{tr(T['halka_cevre'], 1)} m", "topraklama", {3})
    for c in T["cubuklar"]:
        ekle(c["poz"], "Topraklama çubuğu", 1, f"Ø{tr(T['cubuk_cap_mm'], 0)} mm x {tr(T['cubuk_boy_m'], 0)} m bakır kaplı", "topraklama", {3}, grup="TC")
    for c in T["cukurlar"]:
        ekle(c["poz"], "Kontrol (ölçüm) çukuru (köşe çubuklarında)", 1, "300x300x300 mm, kapaklı, polimer beton", "topraklama", {3}, grup="KC")
    for d in T["inis"]:
        ekle(d["poz"], "Yıldırım iniş iletkeni", 1, f"{tr(T['inis_kesit'], 0)} mm² Cu veya 30x3,5 mm galvaniz şerit", "topraklama", {3}, grup="IL")
    for d in T["inis"]:
        ekle(d["ok"], "Ölçüm (test) ayırma noktası (iniş iletkeni başına)", 1, f"Test klemensi, zeminden {tr(T['olcum_ayirma_yuksekligi_m'])} m, korumalı kutu", "topraklama", {3}, grup="OK")
    for d in T["yk"]:
        ekle(d["poz"], f"Franklin çubuğu ({d['ad']})", 1, f"Ø16 mm Cu/paslanmaz, boy {tr(d['boy'])} m", "topraklama", {3}, grup="YK")
    ekle("ÇÇ1", "Çatı çevre iletkeni (hava sonlandırma)", 1, f"{T['cati_iletkeni']}", "topraklama", {3})
    for e in T["eb"]:
        ekle(e["poz"], f"Eş potansiyel bağlantı: {e['ad']}", 1, "16 mm² Cu (sarı-yeşil), EPB1'e", "topraklama", {3}, grup="EB")
    ekle("İF1", "İzolasyon flanşı (gaz borusu)", 1, "Gaz sayacı girişi, EB2 öncesi", "topraklama", {3})
    ekle("EK1", "Eksotermik kaynak bağlantısı (kalıp + toz)", eksotermik_say(T), "Cu-Cu / Cu-çelik, 50 mm² uyumlu", "topraklama", {3})
    return kayit


def eksotermik_say(T):
    """Eksotermik kaynak sayısı: çubuk-halka + iniş-halka + EPB/TE/ÇÇ bağlantıları."""
    return len(T["cubuklar"]) + len(T["inis"]) + 2 + 1 + len(T["eb"])


def satirlari_grupla(kayit):
    """Tablolar için ardışık aynı 'grup' kayıtlarını tek satırda toplar -> (poz_metni, ad, adet, deger, kat)."""
    out = []
    i = 0
    while i < len(kayit):
        j = i
        while j + 1 < len(kayit) and kayit[j + 1]["grup"] == kayit[i]["grup"] and kayit[j + 1]["kat"] == kayit[i]["kat"] \
                and kayit[j + 1]["deger"] == kayit[i]["deger"]:
            j += 1
        blok = kayit[i:j + 1]
        pozlar = [b["poz"] for b in blok]
        out.append({"poz": poz_aralik(pozlar), "pozlar": pozlar, "ad": blok[0]["ad"],
                    "adet": sum(b["adet"] for b in blok), "deger": blok[0]["deger"], "kat": blok[0]["kat"],
                    "cizim": blok[0]["cizim"]})
        i = j + 1
    return out


# =====================================================================
#  ÇİZİM 1 : TEK HAT ŞEMASI
# =====================================================================
def faz_adi(h):
    return "L1-L2-L3" if h["faz3"] else h["faz"]


def cizim1(P, M, PN, T, yol):
    c = Cizim(yol, 1)
    C = P["cihazlar"]
    S = P["sistem"]
    hat = {h["id"]: h for h in M["hatlar"]}

    # ---------------- sabitler (mm)
    XC0, XC1 = 15.2, 403.2
    Y = dict(panel_ust=58, bus={"L1": 107, "L2": 110, "L3": 113, "N": 116}, hdr=(120, 125.5), q1=(128, 137), rcd=(139, 151),
             sub={"L1": 155, "L2": 157.5, "L3": 160, "N": 162.5}, dev=(166, 177), lbl1=180.2, lbl2=183.2,
             pe=187.0, panel_alt=190.0, yuk_bar=203.5, tablo=216.0)
    INC = {"N": 46.75, "L1": 50.25, "L2": 53.75, "L3": 57.25}

    # kolonlar
    idler = []
    for g in ("KLG1", "KLG2", "AVZ", "PRZ", "LED", "PRJ", "KAM", "SES", "ACL", "DEK", "YED"):
        if g not in P["devre_gruplari"]:
            continue
        for cid in P["devre_gruplari"][g]["cihazlar"]:
            if cid in hat or g == "YED":
                idler.append((g, cid))
    YW = 8.0
    n_y = sum(1 for g, _ in idler if g == "YED")
    w_norm = ((XC1 - XC0) - n_y * YW) / (len(idler) - n_y)
    kol = {}
    x = XC0
    for g, cid in idler:
        w = YW if g == "YED" else w_norm
        kol[cid] = {"x0": x, "w": w, "cx": x + w / 2, "g": g}
        x += w

    # ---------------- başlık
    c.t(14, 17.2, "ÇİZİM 1 — TEK HAT ŞEMASI (ölçeksiz, prensip)", size=4.0, bold=True)
    c.t(14, 22.4, f"Sistem: {S['U_LL']:.0f}/{S['U_LN']:.0f} V, {S['frekans']} Hz, {S['sebeke']}  |  Kaynak: {tr(S['trafo_kva'], 0)} kVA trafo  |  Ana pano: PN1", size=2.4, color="#444")

    # ---------------- kaynak: OG blok, trafo, W0
    c.rect(34, 27, 36, 7, sw=0.4)
    c.t(52, 31.6, "OG ŞEBEKESİ (blok)", size=2.5, anchor="middle", bold=True)
    c.line(52, 34, 52, 36.5, w=0.4)
    yb = c.trafo(52, 42.0, r=5.5)                      # alt dairenin alt ucu
    c.t(60.5, 41.0, "TR1", size=3.0, bold=True, poz="TR1")
    c.t(60.5, 44.2, f"{tr(S['trafo_kva'], 0)} kVA, {S['trafo_grubu']}", size=2.1)
    c.t(60.5, 47.0, f"uk = %{tr(S['trafo_uk_yuzde'])}", size=2.1)
    # trafo AG terminal barı
    yt = 54.0
    c.line(44.5, yt, 59.5, yt, w=0.7)
    # (OG-trafo bağlantısı sembolik) -- AG iletkenleri W0
    for ph in ("N", "L1", "L2", "L3"):
        c.iletken(INC[ph], yt, INC[ph], 63.0, ph, w=0.45)
    c.rect(44.0, 55.0, 16.0, 6.0, sw=0.2, stroke="#777", dash="1,0.8")
    b = M["besleme"]
    c.t(62.5, 51.6, f"W0: {b['tip']} {b['damar']} mm² Cu", size=2.4, bold=True, poz=b["kablo_no"])
    c.t(62.5, 54.4, f"L = {tr(b['L'])} m, gömülü (TR1 → PN1)", size=2.0, maxw=48)
    # N-PE köprüsü (yalnız trafo yıldız noktası) + TE1 + PE3
    ypb = 50.0
    c.nokta(INC["N"], ypb, 0.7)
    c.iletken(INC["N"], ypb, 12.2, ypb, "PE", w=0.5)
    c.iletken(INC["N"], 42.0 + 7.15, INC["N"], yt, "N", w=0.45)
    c.t(14.6, 42.6, "N–PE köprüsü", size=2.1, bold=True)
    c.t(14.6, 45.0, "(yalnız trafoda)", size=2.0)
    c.t(14.6, 48.7, "PE3 1x50 mm² Cu", size=2.0, poz="PE3")
    c.nokta(37.5, ypb, 0.6)
    c.line(37.5, ypb, 37.5, 52.0, c=RENK["PE_G"], w=0.5)
    c.topraklama(37.5, 52.0, olcek=0.9)
    c.t(35.3, 53.0, "TE1", size=2.2, bold=True, anchor="end", poz="TE1")
    c.t(35.3, 55.4, "R ≤ 2 Ω", size=1.9, anchor="end")
    c.iletken(12.2, ypb, 12.2, Y["pe"], "PE", w=0.5)

    # ---------------- ana pano çerçevesi
    c.rect(14.2, Y["panel_ust"], 390.3, Y["panel_alt"] - Y["panel_ust"], sw=0.4, stroke="#444", dash="4,1.2,0.8,1.2")
    c.t(402.0, Y["panel_ust"] + 4.3, f"PN1 — ANA DAĞITIM PANOSU  {PN['G']}x{PN['Y']}x{P['pano']['D']} mm, {P['pano']['IP']}", size=2.7, bold=True, anchor="end", poz="PN1")

    # ---------------- Q0 ve kaçak akım koruması
    q0 = C["Q0"]
    xs_q0 = [INC["N"], INC["L1"], INC["L2"], INC["L3"]]
    xl, xr, yc = c.cihaz(xs_q0, ["N", "L1", "L2", "L3"], 63.0, 77.0, tur="mccb", pad=3.2)
    c.t(xr + 1.4, 66.0, "Q0", size=3.0, bold=True, poz="Q0")
    c.t(xr + 1.4, 69.0, f"MCCB 4P {q0['govde_A']} A gövde", size=2.0)
    c.t(xr + 1.4, 71.6, f"Ir = {q0['Ir_A']} A, Icu = {tr(q0['Icu_kA'], 0)} kA", size=2.0)
    c.t(xr + 1.4, 74.2, f"Isd = {q0['Isd_x']}xIr / {q0['tsd_ms']} ms", size=2.0)
    for ph in ("N", "L1", "L2", "L3"):
        c.iletken(INC[ph], 77.0, INC[ph], Y["bus"][ph], ph, w=0.45)
    # MX0 açtırma bobini + SB1
    c.rect(33.2, 66.8, 5.2, 4.2, sw=0.3, fill="#fff")
    c.line(33.2, 71.0, 38.4, 66.8, w=0.25)
    c.line(38.4, 68.9, xl, 68.9, w=0.25, dash="0.8,0.6")
    c.t(33.0, 66.0, "MX0", size=2.2, bold=True, anchor="end", poz="MX0")
    c.t(35.8, 74.0, "açtırma", size=1.8, anchor="middle")
    c.t(35.8, 76.0, "bobini", size=1.8, anchor="middle")
    # SB1 acil durdurma
    c.line(26.5, 68.9, 33.2, 68.9, w=0.25, dash="0.8,0.6")
    c.circle(21.7, 68.9, 2.4, sw=0.3, fill="#e74c3c")
    c.line(21.7, 71.3, 21.7, 74.0, w=0.3)
    c.line(19.3, 74.0, 24.1, 74.0, w=0.3)
    c.t(21.7, 78.3, "SB1", size=2.3, bold=True, anchor="middle", poz="SB1")
    c.t(21.7, 80.9, "acil durdurma", size=1.9, anchor="middle")
    c.line(24.1, 68.9, 26.5, 68.9, w=0.25, dash="0.8,0.6")
    # toroid TOR0 (halka) + RCR0
    c.ellipse(52.0, 82.5, 8.6, 2.8, sw=0.35)
    c.t(41.5, 85.2, "TOR0", size=2.2, bold=True, anchor="end", poz="TOR0")
    c.line(60.6, 82.5, 66.0, 82.5, w=0.25, dash="0.8,0.6")
    r0 = C["RCR0"]
    c.rect(66.0, 78.5, 36.0, 11.2, sw=0.35, fill="#fff")
    c.t(68.0, 82.3, "RCR0", size=2.8, bold=True, poz="RCR0")
    c.t(77.5, 82.3, "Gecikmeli kaçak akım rölesi", size=1.9)
    c.t(68.0, 85.4, f"I∆n = {tr(r0['I_dn_A'])} A  (ayar {r0['aralik_I']})", size=1.9)
    c.t(68.0, 88.3, f"t = {r0['gecikme_ms']} ms  (ayar {r0['aralik_t']})", size=1.9)
    # kumanda hattı (RCR0, FSR1, SB1 -> MX0)
    yk = 62.0
    c.line(35.8, yk, 35.8, 66.8, w=0.25, dash="0.8,0.6")
    c.line(35.8, yk, 312.0, yk, w=0.25, dash="0.8,0.6")
    c.line(95.0, yk, 95.0, 78.5, w=0.25, dash="0.8,0.6")
    c.nokta(95.0, yk, 0.45)
    c.t(104.0, yk - 0.8, "Açtırma kumanda hattı (RCR0, FSR1, SB1 → MX0)", size=1.8, italic=True)

    # akım trafoları CT1-CT3 (kademeli)
    ct_y = {"L1": 91.5, "L2": 96.0, "L3": 100.5}
    for i, ph in enumerate(("L1", "L2", "L3")):
        xx = INC[ph]
        c.circle(xx, ct_y[ph], 1.9, sw=0.3, fill="#fff")
        c.t(44.0 - (INC["N"] - 46.75), ct_y[ph] + 0.8, f"CT{i + 1}", size=2.2, bold=True, anchor="end", poz=f"CT{i + 1}")
    c.t(14.8, 93.2, "CT1–CT3: 200/5 A", size=2.0)
    c.t(14.8, 95.7, "5 VA, Sınıf 0,5", size=2.0)

    # ---------------- ana baralar
    for ph, poz in (("L1", "BL1"), ("L2", "BL2"), ("L3", "BL3"), ("N", "BN")):
        yy = Y["bus"][ph]
        c.iletken(40.5, yy, 402.5, yy, ph, w=0.75)
        c.t(38.8, yy + 0.7, poz, size=2.0, bold=True, anchor="end", poz=poz)
        c.nokta(INC[ph], yy, 0.7)
    c.t(404.0, Y["bus"]["L2"] + 0.4, "", size=1.0)

    # ---------------- SPD1 + F1
    sp_x = {"L1": 110.0, "L2": 114.0, "L3": 118.0, "N": 122.0}
    for ph, xx in sp_x.items():
        c.iletken(xx, Y["bus"][ph], xx, 76.5, ph, w=0.4)
        c.nokta(xx, Y["bus"][ph], 0.6)
    c.cihaz([sp_x["L1"], sp_x["L2"], sp_x["L3"]], ["L1", "L2", "L3"], 79.5, 88.5, tur="sigorta", pad=2.6)
    c.t(124.4, 82.0, "F1", size=2.5, bold=True, poz="F1")
    c.t(124.4, 84.6, f"NH00 gG {C['F1']['In_A']} A", size=1.9)
    for xx in sp_x.values():
        c.spd_kutu(xx, 70.5, genislik=2.8, yukseklik=5.0)
        c.line(xx, 75.5, xx, 76.5, w=0.3)
    c.iletken(107.5, 68.5, 124.5, 68.5, "PE", w=0.45)
    for xx in sp_x.values():
        c.line(xx, 68.5, xx, 70.5, w=0.3)
    c.iletken(124.5, 68.5, 128.0, 68.5, "PE", w=0.45)
    c.iletken(128.0, 68.5, 128.0, 70.0, "PE", w=0.45)
    c.topraklama(128.0, 70.0, olcek=0.8)
    s1 = C["SPD1"]
    c.t(131.0, 72.0, "SPD1", size=2.5, bold=True, poz="SPD1")
    c.t(131.0, 74.6, "Tip 1+2, 3+1", size=1.9)
    c.t(131.0, 77.0, f"Iimp {tr(s1['I_imp_kA'])} kA/k.", size=1.9)
    c.t(131.0, 79.4, f"Up ≤ {tr(s1['Up_kV'])} kV", size=1.9)
    c.t(105.0, 66.4, "PE: BPE'ye ≤ 0,5 m, 25 mm² Cu", size=1.8, italic=True)

    # ---------------- ölçüm ve izleme bloğu
    bx0, bx1, by0, by1 = 155.0, 352.0, 64.5, 105.5
    c.rect(bx0, by0, bx1 - bx0, by1 - by0, sw=0.3, stroke="#888", dash="2,1")
    c.t(bx0 + 2, by0 + 3.0, "ÖLÇÜM VE İZLEME (kapı üzeri cihazlar)", size=2.0, bold=True, color="#555")
    # gerilim alt barası ve F2
    f2x = {"L1": 331.0, "L2": 334.0, "L3": 337.0, "N": 341.0}
    sub = {"L1": 79.5, "L2": 82.0, "L3": 84.5, "N": 87.0}
    for ph, xx in f2x.items():
        c.iletken(xx, Y["bus"][ph], xx, sub[ph], ph, w=0.35)
        c.nokta(xx, Y["bus"][ph], 0.55)
    c.cihaz([f2x["L1"], f2x["L2"], f2x["L3"]], ["L1", "L2", "L3"], 92.0, 102.0, tur="mcb", pad=2.4)
    c.t(327.0, 96.0, "F2", size=2.5, bold=True, anchor="end", poz="F2")
    c.t(327.0, 98.6, f"3P {C['F2']['egri']}{C['F2']['In_A']} A", size=1.9, anchor="end")
    sx0 = 164.0
    for ph in ("L1", "L2", "L3", "N"):
        c.iletken(sx0, sub[ph], f2x[ph], sub[ph], ph, w=0.3)
        c.nokta(f2x[ph], sub[ph], 0.45)
    # PM1
    c.rect(162.0, 68.5, 38.0, 8.0, sw=0.35, fill="#fff")
    c.t(164.0, 72.4, "PM1", size=2.6, bold=True, poz="PM1")
    c.t(164.0, 75.2, "Enerji analizörü 3F4T", size=1.9)
    for k, ph in enumerate(("L1", "L2", "L3", "N")):
        xx = 176.0 + 3.0 * k
        c.iletken(xx, 76.5, xx, sub[ph], ph, w=0.25)
        c.nokta(xx, sub[ph], 0.4)
    c.t(188.8, 75.4, "U", size=1.8, anchor="start")
    # CT sekonderleri -> PM1 (I1,I2,I3)
    for k, ph in enumerate(("L1", "L2", "L3")):
        xct = 190.0 + 3.0 * k
        c.iletken(INC[ph] + 1.9, ct_y[ph], xct, ct_y[ph], ph, w=0.25)
        c.iletken(xct, ct_y[ph], xct, 76.5, ph, w=0.25)
    c.t(199.5, 75.4, "I", size=1.8, anchor="end")
    # XT1 klemens
    c.rect(150.0, 88.0, 5.0, 13.5, sw=0.3, fill="#fff")
    c.t(152.5, 86.8, "XT1", size=2.0, bold=True, anchor="middle", poz="XT1")
    for ph in ("L1", "L2", "L3"):
        c.circle(152.5, ct_y[ph], 0.5, sw=0.2, fill="#000")
    # A1 + SA1
    c.rect(207.0, 71.5, 7.0, 5.0, sw=0.3, fill="#fff")
    c.t(210.5, 75.0, "SA1", size=2.0, bold=True, anchor="middle", poz="SA1")
    c.olcu_aleti(224.0, 72.5, "A", r=3.4)
    c.t(224.0, 68.0, "A1", size=2.2, bold=True, anchor="middle", poz="A1")
    c.line(214.0, 74.0, 220.6, 74.0, w=0.25)
    c.line(210.5, 76.5, 210.5, 79.5, w=0.2, dash="0.7,0.5")
    c.t(210.5, 81.5, "I1-I2-I3", size=1.8, anchor="middle", color="#444")
    # V1 + SA2
    c.rect(234.0, 71.5, 7.0, 5.0, sw=0.3, fill="#fff")
    c.t(237.5, 75.0, "SA2", size=2.0, bold=True, anchor="middle", poz="SA2")
    c.olcu_aleti(251.0, 72.5, "V", r=3.4)
    c.t(251.0, 68.0, "V1", size=2.2, bold=True, anchor="middle", poz="V1")
    c.line(241.0, 74.0, 247.6, 74.0, w=0.25)
    for ph, xx in (("L1", 235.5), ("L2", 237.5), ("L3", 239.5)):
        c.iletken(xx, 76.5, xx, sub[ph], ph, w=0.25)
        c.nokta(xx, sub[ph], 0.4)
    # H1-H3
    for k, ph in enumerate(("L1", "L2", "L3")):
        xx = 264.0 + 11.0 * k
        c.lamba(xx, 72.0, r=2.1)
        c.t(xx, 68.0, f"H{k + 1}", size=2.2, bold=True, anchor="middle", poz=f"H{k + 1}")
        c.iletken(xx - 1.2, 74.0, xx - 1.2, sub[ph], ph, w=0.25)
        c.nokta(xx - 1.2, sub[ph], 0.4)
        c.iletken(xx + 1.2, 74.0, xx + 1.2, sub["N"], "N", w=0.25)
        c.nokta(xx + 1.2, sub["N"], 0.4)
    c.t(277.0, 90.6, "H1-H3: L1-L2-L3 sinyal lambaları", size=1.8, anchor="middle", color="#444")
    # FSR1
    c.rect(300.0, 68.5, 26.0, 8.0, sw=0.35, fill="#fff")
    c.t(302.0, 72.4, "FSR1", size=2.6, bold=True, poz="FSR1")
    c.t(302.0, 75.2, "Faz sırası/gerilim izl.", size=1.9, maxw=22.5)
    for ph, xx in (("L1", 304.0), ("L2", 307.0), ("L3", 310.0), ("N", 313.0)):
        c.iletken(xx, 76.5, xx, sub[ph], ph, w=0.25)
        c.nokta(xx, sub[ph], 0.4)
    c.line(320.0, 68.5, 320.0, yk, w=0.25, dash="0.8,0.6")
    c.nokta(320.0, yk, 0.45)
    c.line(312.0, yk, 320.0, yk, w=0.25, dash="0.8,0.6")

    # ---------------- EPB1, topraklama (sağ üst)
    c.t(405.5, 16.5, "TOPRAKLAMA / EŞ POTANSİYEL", size=2.4, bold=True, anchor="end")
    ex0, ex1, ey = 338.0, 405.0, 30.0
    c.rect(ex0, ey - 2.0, ex1 - ex0, 4.0, sw=0.35, fill="#e9d2b0")
    for k in range(7):
        c.rect(ex0 + 3.0 + k * 9.2, ey - 1.0, 2.0, 2.0, sw=0.2, fill="#8a5a1a")
    c.t(ex0 + 1, ey - 3.4, "EPB1 — Esas eş potansiyel barası", size=2.3, bold=True, poz="EPB1")
    c.iletken(ex1, ey, 407.2, ey, "PE", w=0.5)
    c.iletken(407.2, ey, 407.2, Y["pe"], "PE", w=0.5)
    c.iletken(407.2, Y["pe"], 403.0, Y["pe"], "PE", w=0.5)
    c.t(405.0, ey + 8.2, "PE1 50 mm² Cu (BPE ↔ EPB1)", size=2.0, anchor="end", poz="PE1")
    c.iletken(ex0 + 20.0, ey + 2.0, ex0 + 20.0, ey + 12.0, "PE", w=0.5)
    c.topraklama(ex0 + 20.0, ey + 12.0, olcek=1.0)
    c.t(ex0 + 24.0, ey + 14.2, "PE2 50 mm² Cu → TH1", size=2.0, poz="PE2")
    c.t(ex0 + 24.0, ey + 17.0, f"Topraklama halkası, hedef ≤ {tr(T['hedef_direnc_ohm'], 0)} Ω", size=1.9)
    c.t(ex0 + 24.0, ey + 19.8, "EB1–EB6, ÖK1–ÖK6 → Çizim 3", size=1.9, color="#444")
    for k in range(3):
        xx = 386.0 + 5.0 * k
        c.line(xx, ey - 2.0, xx, ey - 6.0, c=RENK["PE_G"], w=0.4)
    c.t(405.0, ey - 7.6, "EB1–EB6 (16 mm²)", size=1.9, color="#444", anchor="end")

    # ---------------- çıkışlar
    geciken = []

    def gec(*a, **k):
        geciken.append((a, k))

    # klima grup zinciri (Q1/RCD1 ... K1-K3)
    kq, rc, kk = C["Q_klima_grup"], C["RCD_klima"], C["K_klima"]
    KX = [-4.5, -1.5, 1.5, 4.5]
    KC = ["L1", "L2", "L3", "N"]
    for qpoz, rpoz, ids in (("Q1", "RCD1", ["K1", "K2", "K3"]), ("Q2", "RCD2", ["K4", "K5", "K6"])):
        cxg = kol[ids[1]]["cx"]
        xs = [cxg + d for d in KX]
        for x_, ph in zip(xs, KC):
            c.iletken(x_, Y["bus"][ph], x_, Y["q1"][0], ph, w=0.4)
            c.nokta(x_, Y["bus"][ph], 0.6)
        xl, xr, yc = c.cihaz(xs, KC, Y["q1"][0], Y["q1"][1], tur="mcb", pad=3.0)
        c.t(xr + 1.6, Y["q1"][0] + 3.6, qpoz, size=2.6, bold=True, poz=qpoz)
        c.t(xr + 1.6, Y["q1"][0] + 6.4, f"MCB 4P {kq['egri']}{kq['In_A']} A, {tr(kq['Icu_kA'], 0)} kA", size=1.9)
        for x_, ph in zip(xs, KC):
            c.iletken(x_, Y["q1"][1], x_, Y["rcd"][0], ph, w=0.4)
        xl2, xr2, yc2 = c.cihaz(xs, KC, Y["rcd"][0], Y["rcd"][1], tur="rcd", pad=3.0)
        c.t(xr2 + 1.6, Y["rcd"][0] + 3.4, rpoz, size=2.6, bold=True, poz=rpoz)
        c.t(xr2 + 1.6, Y["rcd"][0] + 6.0, f"4P {rc['In_A']} A, {rc['I_dn_mA']} mA, Tip B", size=1.9)
        c.t(xr2 + 1.6, Y["rcd"][0] + 8.4, "(alternatif Tip F)", size=1.9)
        # alt bara satırları
        xa = [kol[i]["cx"] for i in ids]
        for k, ph in enumerate(KC):
            yy = Y["sub"][ph]
            c.iletken(xa[0] + KX[k], yy, xa[2] + KX[k], yy, ph, w=0.5)
            c.iletken(xs[k], Y["rcd"][1], xs[k], Y["dev"][0], ph, w=0.4)
            for xc in xa:
                c.nokta(xc + KX[k], yy, 0.55)
        for i in ids:
            cx = kol[i]["cx"]
            for k, ph in enumerate(KC):
                c.iletken(cx + KX[k], Y["sub"][ph], cx + KX[k], Y["dev"][0], ph, w=0.4)
            h = hat[i]
            # N hattı cihaz yanından geçer
            c.iletken(cx + KX[3], Y["dev"][0], cx + KX[3], Y["yuk_bar"], "N", w=0.4)
            xl3, xr3, _ = c.cihaz([cx + KX[0], cx + KX[1], cx + KX[2]], ["L1", "L2", "L3"], Y["dev"][0], Y["dev"][1], tur="mpcb", pad=2.6)
            gec(cx, Y["lbl1"], i, size=2.6, bold=True, anchor="middle", poz=i, halo=True)
            gec(cx, Y["lbl2"], f"MPCB 3P {kk['ayar_A']} A", size=1.9, anchor="middle", maxw=kol[i]["w"] - 1, halo=True)
            for k, ph in enumerate(("L1", "L2", "L3")):
                c.iletken(cx + KX[k], Y["dev"][1], cx + KX[k], Y["yuk_bar"], ph, w=0.4)

    # ---- dış aydınlatma kumanda grubu: Q3 -> (SPD2) -> KM1 -> alt bara -> PJ RCBO'ları
    pj_ids = [h["id"] for h in M["hatlar"] if h["grup"] == "PROJ"]
    spd2_goster = P["aydinlatma_priz"]["projektor"].get("spd2_goster", True)
    sub_x = {}                                     # PJ grubu için alt bara ihtiyaçları
    if pj_ids:
        q3, km1, spd2 = C["Q3"], C["KM1"], C["SPD2"]
        cxg = kol[pj_ids[len(pj_ids) // 2]]["cx"]
        xs = [cxg + d for d in KX]
        for x_, ph in zip(xs, KC):
            c.iletken(x_, Y["bus"][ph], x_, Y["q1"][0], ph, w=0.4)
            c.nokta(x_, Y["bus"][ph], 0.6)
        xl, xr, yc = c.cihaz(xs, KC, Y["q1"][0], Y["q1"][1], tur="mcb", pad=3.0)
        c.t(xr + 1.6, Y["q1"][0] + 3.6, "Q3", size=2.6, bold=True, poz="Q3")
        c.t(xr + 1.6, Y["q1"][0] + 6.4, f"MCB 4P {q3['egri']}{q3['In_A']} A, {tr(q3['Icu_kA'], 0)} kA", size=1.9)
        for x_, ph in zip(xs[:3], KC[:3]):
            c.iletken(x_, Y["q1"][1], x_, Y["rcd"][0], ph, w=0.4)
        c.iletken(xs[3], Y["q1"][1], xs[3], Y["sub"]["N"], "N", w=0.4)
        xl2, xr2, yc2 = c.cihaz(xs[:3], KC[:3], Y["rcd"][0], Y["rcd"][1], tur="kontaktor", pad=2.6)
        c.rect(xr2 + 1.6, yc2 - 2.0, 5.0, 4.0, sw=0.3, fill="#fff")
        c.line(xr2, yc2, xr2 + 1.6, yc2, w=0.25, dash="0.8,0.6")
        c.t(xr2 + 9.0, Y["rcd"][0] + 3.4, "KM1", size=2.6, bold=True, poz="KM1")
        c.t(xr2 + 9.0, Y["rcd"][0] + 6.0, f"Kontaktör 3P {km1['In_A']} A", size=1.9)
        c.t(xr2 + 9.0, Y["rcd"][0] + 8.4, "bobin A1-A2: sağ üst", size=1.7, color="#444")
        c.t(xr2 + 9.0, Y["rcd"][0] + 10.8, "kumanda devresi", size=1.7, color="#444")
        # SPD2 (Tip 2) yan dal: Q3 çıkışından
        if spd2_goster:
            yt_ = Y["q1"][1] + 1.2
            xsp = xs[0] - 11.0
            c.nokta(xs[0], yt_, 0.55)
            c.line(xs[0], yt_, xsp, yt_, w=0.35)
            for k in range(4):
                c.line(xs[0] - 4.0 - k * 1.1, yt_ + 0.9, xs[0] - 3.4 - k * 1.1, yt_ - 0.9, w=0.2)
            c.line(xsp, yt_, xsp, yt_ + 1.6, w=0.35)
            c.spd_kutu(xsp, yt_ + 1.6, genislik=2.8, yukseklik=5.0)
            c.line(xsp, yt_ + 6.6, xsp, yt_ + 7.4, c=RENK["PE_G"], w=0.4)
            c.topraklama(xsp, yt_ + 7.4, olcek=0.8)
            c.t(xsp - 4.0, yt_ + 3.0, "SPD2", size=2.2, bold=True, anchor="end", poz="SPD2")
            c.t(xsp - 4.0, yt_ + 5.4, "Tip 2, 3+1", size=1.7, anchor="end")
            c.t(xsp - 4.0, yt_ + 7.6, f"In {tr(spd2['I_n_kA'], 0)} kA", size=1.7, anchor="end")
        # alt bara satırları (her faz için grup düşeyi + bu fazdaki hatlar)
        for k, ph in enumerate(KC):
            xsr = [xs[k]]
            for i in pj_ids:
                cxi = kol[i]["cx"]
                if ph == "N":
                    xsr.append(cxi + 1.6)
                elif hat[i]["faz"] == ph:
                    xsr.append(cxi - 1.6)
            yy = Y["sub"][ph]
            if ph != "N":
                c.iletken(xs[k], Y["rcd"][1], xs[k], yy, ph, w=0.4)
            if len(xsr) > 1:
                c.iletken(min(xsr), yy, max(xsr), yy, ph, w=0.5)
                for xx in xsr:
                    c.nokta(xx, yy, 0.55)

    # ---- ses sistemi: ortak RCD3 (2P) -> alt bara -> MCB'ler
    ses_mcb = [h["id"] for h in M["hatlar"] if h.get("koruma_tur") == "mcb"]
    if ses_mcb:
        rc3 = C["RCD3"]
        ph_s = hat[ses_mcb[0]]["faz"]
        cxm = sum(kol[i]["cx"] for i in ses_mcb) / len(ses_mcb)
        xp3, xn3 = cxm - 1.6, cxm + 1.6
        c.iletken(xp3, Y["bus"][ph_s], xp3, Y["rcd"][0], ph_s, w=0.4)
        c.iletken(xn3, Y["bus"]["N"], xn3, Y["rcd"][0], "N", w=0.4)
        c.nokta(xp3, Y["bus"][ph_s], 0.6)
        c.nokta(xn3, Y["bus"]["N"], 0.6)
        xl3_, xr3_, yc3_ = c.cihaz([xp3, xn3], [ph_s, "N"], Y["rcd"][0], Y["rcd"][1], tur="rcd", pad=3.0)
        c.t(xr3_ + 1.4, Y["rcd"][0] + 3.4, "RCD3", size=2.5, bold=True, poz="RCD3")
        c.t(xr3_ + 1.4, Y["rcd"][0] + 6.0, f"2P {rc3['In_A']} A", size=1.7)
        c.t(xr3_ + 1.4, Y["rcd"][0] + 8.2, f"{rc3['I_dn_mA']} mA, Tip {rc3['tip']}", size=1.7)
        c.t(xr3_ + 1.4, Y["rcd"][0] + 10.4, "ses ortak", size=1.6, color="#444")
        y_p3, y_n3 = Y["sub"][ph_s], Y["sub"]["N"]
        c.iletken(xp3, Y["rcd"][1], xp3, y_p3, ph_s, w=0.4)
        c.iletken(xn3, Y["rcd"][1], xn3, y_n3, "N", w=0.4)
        xs_p = [xp3] + [kol[i]["cx"] - 1.6 for i in ses_mcb]
        xs_n = [xn3] + [kol[i]["cx"] + 1.6 for i in ses_mcb]
        c.iletken(min(xs_p), y_p3, max(xs_p), y_p3, ph_s, w=0.5)
        c.iletken(min(xs_n), y_n3, max(xs_n), y_n3, "N", w=0.5)
        for xx in xs_p:
            c.nokta(xx, y_p3, 0.55)
        for xx in xs_n:
            c.nokta(xx, y_n3, 0.55)

    # ---- tek faz RCBO sütunları
    for cid, kd in kol.items():
        h = hat.get(cid)
        cx, x0, w = kd["cx"], kd["x0"], kd["w"]
        if klima_mi(cid):
            pe_x = cx + 7.6
            c.iletken(pe_x, Y["pe"], pe_x, Y["yuk_bar"], "PE", w=0.4)
            c.nokta(pe_x, Y["pe"], 0.6, c=RENK["PE_G"])
            c.t(cx - 7.2, 196.0, f"{h['kablo_no']} · {h['damar']}", size=1.9, rot=-90, anchor="middle", poz=h["kablo_no"])
            c.rect(cx - 6.3, Y["panel_alt"] + 1.0, 15.9, 10.0, sw=0.2, stroke="#888", dash="1,0.8")
            c.line(cx - 8.4, Y["yuk_bar"], cx + 8.8, Y["yuk_bar"], w=0.9)
            continue
        if kd["g"] == "YED":
            ph = {"YD1": "L1", "YD2": "L2", "YD3": "L3"}.get(cid, "L1")
            xp, xn = cx - 1.2, cx + 1.2
            c.iletken(xp, Y["bus"][ph], xp, Y["dev"][0], ph, w=0.3)
            c.iletken(xn, Y["bus"]["N"], xn, Y["dev"][0], "N", w=0.3)
            c.nokta(xp, Y["bus"][ph], 0.5)
            c.nokta(xn, Y["bus"]["N"], 0.5)
            c.rect(cx - 3.6, Y["dev"][0], 7.2, Y["dev"][1] - Y["dev"][0], sw=0.3, stroke="#555", dash="1.2,0.8")
            c.t(cx, Y["dev"][0] + 6.3, "boş", size=1.8, anchor="middle", color="#555")
            c.t(cx, Y["lbl1"], cid, size=2.1, bold=True, anchor="middle", poz=cid)
            c.t(cx, Y["lbl2"], "yedek", size=1.7, anchor="middle")
            continue
        ph = h["faz"]
        xp, xn = cx - 1.6, cx + 1.6
        alt_bara = kd["g"] == "PRJ" or h.get("koruma_tur") == "mcb"
        y_ph = Y["sub"][ph] if alt_bara else Y["bus"][ph]
        y_n = Y["sub"]["N"] if alt_bara else Y["bus"]["N"]
        c.iletken(xp, y_ph, xp, Y["dev"][0], ph, w=0.4)
        c.iletken(xn, y_n, xn, Y["dev"][0], "N", w=0.4)
        if not alt_bara:
            c.nokta(xp, y_ph, 0.6)
            c.nokta(xn, y_n, 0.6)
        mcb_ = h.get("koruma_tur") == "mcb"
        c.cihaz([xp, xn], [ph, "N"], Y["dev"][0], Y["dev"][1], tur="mcb" if mcb_ else "rcbo", pad=3.0)
        gec(cx, Y["lbl1"], cid, size=2.5, bold=True, anchor="middle", poz=cid, halo=True)
        gec(cx, Y["lbl2"], (f"MCB {P['rcbo_defaults']['egri']}{h['rcbo_In']}" if mcb_ else f"{P['rcbo_defaults']['egri']}{h['rcbo_In']}/{C['RCBO']['I_dn_mA']} mA"),
            size=1.8, anchor="middle", maxw=w - 0.8, halo=True)
        c.iletken(xp, Y["dev"][1], xp, Y["yuk_bar"], ph, w=0.4)
        c.iletken(xn, Y["dev"][1], xn, Y["yuk_bar"], "N", w=0.4)
        pe_x = cx + 4.8
        c.iletken(pe_x, Y["pe"], pe_x, Y["yuk_bar"], "PE", w=0.4)
        c.nokta(pe_x, Y["pe"], 0.6, c=RENK["PE_G"])
        c.rect(cx - 4.0, Y["panel_alt"] + 1.0, 10.6, 10.0, sw=0.2, stroke="#888", dash="1,0.8")
        c.t(cx - (7.6 if cid in ("KR1", "AV1") else 5.2), 196.0, f"{h['kablo_no']} · {h['damar']}", size=1.8, rot=-90, anchor="middle", poz=h["kablo_no"])

    def baslik_bandi(ilk, son, metin):
        x0 = kol[ilk]["x0"] + 0.4
        x1 = kol[son]["x0"] + kol[son]["w"] - 0.4
        c.rect(x0, Y["hdr"][0], x1 - x0, Y["hdr"][1] - Y["hdr"][0], sw=0.25, stroke="#7a8aa0", fill="#e8eef6")
        c.t((x0 + x1) / 2, Y["hdr"][0] + 3.9, metin, size=2.2, bold=True, anchor="middle", maxw=x1 - x0 - 0.6)

    for g in ("KLG1", "KLG2", "AVZ", "PRZ", "LED", "PRJ", "KAM", "SES", "ACL", "DEK", "YED"):
        if g not in P["devre_gruplari"]:
            continue
        ids = [i for gg, i in idler if gg == g]
        if not ids:
            continue
        if g in ("KLG1", "KLG2"):
            toplam = sum(hat[i]["p_kw"] for i in ids)
            metin = f"{P['devre_gruplari'][g]['ad']}  ({poz_aralik(ids)})  ΣP={tr(toplam)} kW"
        elif g == "ACL":
            metin = "ÖNCELİKLİ HAT"
        elif g == "YED":
            metin = "YEDEK"
        elif g in ("PRZ", "LED", "PRJ"):
            ad_ = {"PRZ": "PRİZ", "LED": "İÇ LED", "PRJ": "DIŞ AYDINLATMA"}[g]
            metin = f"{ad_} ({poz_aralik(ids)})"
        else:
            metin = P["devre_gruplari"][g]["ad"]
        baslik_bandi(ids[0], ids[-1], metin)

    # ---- BPE barası (PE) ve etiketi
    c.iletken(15.0, Y["pe"], 403.0, Y["pe"], "PE", w=0.9)
    for a_, k_ in geciken:
        c.t(*a_, **k_)
    c.t(402.5, Y["pe"] - 1.5, "BPE — PE barası (pano gövdesine bağlı)", size=2.0, bold=True, anchor="end", poz="BPE", color="#1c6a31", halo=True)

    # ---- yükler
    dimmer_goster = P["aydinlatma_priz"]["avize"].get("dimmer_goster", False)
    for cid, kd in kol.items():
        h = hat.get(cid)
        cx = kd["cx"]
        yb = Y["yuk_bar"]
        if kd["g"] == "YED":
            continue
        ara = None
        if cid == "KR1":
            ara = ("UPS1", f"{tr(C['UPS1']['kva'], 0)} kVA", "~ / = / ~")
        elif cid == "AV1" and dimmer_goster:
            ara = ("DM1", "dimmer", "LED uyumlu")
        if ara is not None:
            c.rect(cx - 6.6, 191.5, 13.2, 11.0, sw=0.35, fill="#fff")
            c.t(cx, 196.0, ara[0], size=2.4, bold=True, anchor="middle", poz=ara[0])
            c.t(cx, 198.8, ara[1], size=1.8, anchor="middle")
            c.t(cx, 201.3, ara[2], size=1.6, anchor="middle")
            c.line(cx - 8.4, yb + 3.0, cx + 8.8, yb + 3.0, w=0.9)
            for xx, col in ((cx - 1.6, h["faz"]), (cx + 1.6, "N"), (cx + 4.8, "PE")):
                c.iletken(xx, 202.5, xx, yb + 3.0, col, w=0.4)
            c.rect(cx - 8.0, yb + 4.0, 16.0, 7.6, sw=0.35, fill="#fff")
            if cid == "KR1":
                c.rect(cx - 6.0, yb + 5.2, 6.5, 3.0, sw=0.3)
                c.circle(cx + 3.8, yb + 6.7, 1.4, sw=0.3)
                c.t(cx, yb + 10.6, "kamera / NVR", size=1.8, anchor="middle")
            else:
                c.avize(cx - 4.2, yb + 7.6, r=1.7)
                c.t(cx + 3.4, yb + 8.0, "AVZ1", size=1.9, bold=True, anchor="middle", poz=h["elemanlar"])
            continue
        c.line(cx - 8.4, yb, cx + 8.8, yb, w=0.9)
        gr = h["grup"]
        if klima_mi(cid):
            yuk = h["yuk"]
            c.motor(cx - 1.0, yb + 6.2, r=4.6)
            c.t(cx + 4.4, yb + 6.4, yuk, size=2.4, bold=True, poz=yuk)
            c.t(cx + 4.4, yb + 9.2, f"{tr(h['p_kw'])} kW", size=1.9)
        elif gr == "AVIZE":
            c.avize(cx, yb + 6.6, r=2.4)
            c.t(cx, yb + 12.2, poz_aralik(h["elemanlar"]), size=1.9, bold=True, anchor="middle", poz=h["elemanlar"], maxw=kd["w"] - 0.6)
        elif gr == "PRIZ":
            for dx in (-5.0, 0.0, 5.0):
                c.priz(cx + dx, yb + 0.8, r=1.8)
            c.t(cx, yb + 11.4, poz_aralik(h["elemanlar"]), size=1.9, bold=True, anchor="middle", poz=h["elemanlar"], maxw=kd["w"] - 0.6)
        elif gr == "LED":
            for dx in (-5.0, 0.0, 5.0):
                c.lamba(cx + dx, yb + 5.4, r=1.7)
            c.t(cx, yb + 11.4, poz_aralik(h["elemanlar"]), size=1.9, bold=True, anchor="middle", poz=h["elemanlar"], maxw=kd["w"] - 0.6)
        elif gr == "PROJ":
            for dx in (-5.6, 0.0, 5.6):
                c.projektor(cx + dx - 1.4, yb + 5.4, yon=0, sc=0.7)
            c.t(cx, yb + 11.4, poz_aralik(h["elemanlar"]), size=1.9, bold=True, anchor="middle", poz=h["elemanlar"], maxw=kd["w"] - 0.4)
        elif gr == "SES":
            c.hoparlor(cx - 3.0, yb + 6.0)
            c.hoparlor(cx + 3.5, yb + 6.0)
            c.t(cx, yb + 11.2, f"hoparlör ({h.get('kol_ad', '')})", size=1.8, anchor="middle", maxw=kd["w"] - 0.4)
        elif cid == "AC1":
            c.lamba(cx - 2.0, yb + 6.0, r=2.2)
            c.rect(cx + 1.0, yb + 3.5, 5.5, 5.0, sw=0.3)
            c.t(cx + 3.75, yb + 7.4, "E", size=3.2, bold=True, anchor="middle")
            c.t(cx, yb + 11.2, "acil armatür", size=1.8, anchor="middle", maxw=kd["w"] - 0.4)
        elif cid == "YA1":
            c.zil(cx, yb + 3.2)
            c.t(cx, yb + 11.2, "yangın alarm", size=1.8, anchor="middle", maxw=kd["w"] - 0.4)
        elif gr == "DEKOR":
            c.lamba(cx, yb + 6.0, r=2.4)
            c.t(cx, yb + 11.2, "dekoratif", size=1.8, anchor="middle")

    # ---- dış aydınlatma kumanda devresi (üst sağ)
    if pj_ids:
        kx0, kx1, ky0, ky1 = 357.0, 403.5, 65.5, 106.0
        c.rect(kx0, ky0, kx1 - kx0, ky1 - ky0, sw=0.3, stroke="#888", dash="2,1")
        c.t(kx0 + 1.5, ky0 + 3.0, "DIŞ AYDINLATMA KUMANDASI", size=1.8, bold=True, color="#555")
        yl_, yn_ = 72.5, 101.0
        c.iletken(359.5, yl_, 401.5, yl_, "L1", w=0.4)
        c.iletken(359.5, yn_, 401.5, yn_, "N", w=0.4)
        c.t(401.5, yl_ - 1.0, "L (F2'den, 230 V)", size=1.6, anchor="end", color="#444")
        c.t(401.5, yn_ + 2.4, "N", size=1.6, anchor="end", color="#444")

        def kontak(x, y0, y1, ad, ad2, rn="#000", poz=None, yan="sag"):
            c.line(x, y0, x, y0 + 1.2, w=0.3)
            c.line(x, y0 + 3.6, x, y1, w=0.3)
            c.line(x, y0 + 3.6, x + 1.5, y0 + 1.2, w=0.3)
            c.circle(x, y0 + 1.2, 0.35, sw=0, fill="#000")
            c.circle(x, y0 + 3.6, 0.35, sw=0, fill="#000")
            c.t(x + 2.6, y0 + 2.4, ad, size=1.8, bold=True, poz=poz)
            if ad2:
                c.t(x + 2.6, y0 + 4.5, ad2, size=1.5, color="#444")
        xr1, xr2 = 364.0, 391.0
        c.line(xr1, yl_, xr1, 74.5, w=0.3)
        kontak(xr1, 74.5, 79.0, "SA3", "OTO", poz="SA3")
        c.line(xr1, 79.0, xr1, 81.0, w=0.3)
        kontak(xr1, 81.0, 85.5, "FT1", "fotosel", poz="FT1")
        c.line(xr1, 85.5, xr1, 87.5, w=0.3)
        kontak(xr1, 87.5, 92.0, "ZR1", "astronomik", poz="ZR1")
        c.line(xr1, 92.0, xr1, 94.0, w=0.3)
        c.nokta(xr1, 94.0, 0.5)
        c.rect(xr1 - 2.6, 95.2, 5.2, 4.0, sw=0.3, fill="#fff")
        c.t(xr1 + 3.6, 98.0, "KM1 bobin", size=1.8, bold=True, poz="KM1")
        c.line(xr1, 99.2, xr1, yn_, w=0.3)
        # MAN kolu (SA3 - MAN)
        c.line(xr2, yl_, xr2, 74.5, w=0.3)
        kontak(xr2, 74.5, 79.0, "SA3", "MAN", poz="SA3")
        c.line(xr2, 79.0, xr2, 94.0, w=0.3)
        c.line(xr2, 94.0, xr1, 94.0, w=0.3)
        c.nokta(xr1, 94.0, 0.5)
        if "H4" in P["cihazlar"]:
            xh = 378.0
            c.nokta(xh, 94.0, 0.5)
            c.line(xh, 94.0, xh, 96.2, w=0.3)
            c.lamba(xh, 98.2, r=2.0)
            c.line(xh, 100.2, xh, yn_, w=0.3)
            c.t(xh + 2.8, 97.6, "H4", size=1.8, bold=True, poz="H4")
            c.t(xh + 2.8, 99.8, "durum lambası", size=1.4, color="#444")
        c.t(kx1 - 1.0, 104.6, "SA3: MAN–0–OTO (kapıda)", size=1.6, anchor="end", color="#444")

    # ---- çıkış tabloları
    ty = Y["tablo"]
    rows_h = [7.4, 4.0, 4.0, 9.4, 7.4]
    for cid, kd in kol.items():
        h = hat.get(cid)
        x0, w = kd["x0"] + 0.4, kd["w"] - 0.8
        yy = ty
        if kd["g"] == "YED":
            c.rect(x0, yy, w, sum(rows_h), sw=0.25, stroke="#777")
            c.t(x0 + w / 2, yy + 5.0, cid, size=2.0, bold=True, anchor="middle")
            c.t(x0 + w / 2, yy + 8.6, "YEDEK", size=1.7, anchor="middle", maxw=w - 0.4)
            c.t(x0 + w / 2, yy + 11.8, "boş", size=1.7, anchor="middle")
            c.t(x0 + w / 2, yy + 14.4, "yuva", size=1.7, anchor="middle")
            c.t(x0 + w / 2, yy + 18.0, "P: -", size=1.7, anchor="middle")
            c.t(x0 + w / 2, yy + 21.0, "Faz: -", size=1.7, anchor="middle")
            c.t(x0 + w / 2, yy + 24.0, "Kablo: -", size=1.6, anchor="middle", maxw=w - 0.4)
            continue
        guc_txt = f"{tr(h['p_kw'], 2)} kW" if h["p_kw"] < 1.0 or h["grup"] in ("PRIZ", "LED", "PROJ") else f"{tr(h['p_kw'])} kW"
        kablo_txt = f"{h['kablo_no']}: {h['tip']} {h['damar']} mm²" + (" ↑" if h.get("kesit_artirildi") else "")
        veriler = [
            ("HAT ADI", wrap(h["ad"], w - 1.0, 1.9, True), True),
            ("GÜÇ", [guc_txt], True),
            ("FAZ", [faz_adi(h)], False),
            ("KABLO", wrap(kablo_txt, w - 1.0, 1.8), False),
            ("KORUMA", wrap(h["koruma_tablo"], w - 1.0, 1.8), False),
        ]
        for (etk, satir, bold), rh in zip(veriler, rows_h):
            c.rect(x0, yy, w, rh, sw=0.25, stroke="#555")
            c.t(x0 + 0.5, yy + 1.9, etk, size=1.3, color="#777")
            if etk in ("GÜÇ", "FAZ"):
                c.t(x0 + w - 0.6, yy + rh - 0.9, satir[0], size=1.9, anchor="end", bold=bold, maxw=w - 1.0)
            else:
                sz = 1.8 if etk in ("KABLO", "KORUMA") else 1.9
                lim = 3 if etk == "KABLO" else 2
                if len(satir) > lim:
                    uyar(f"[1] Tablo hücresi {lim} satırı aşıyor: {cid} {etk} -> {satir}")
                for li, ln in enumerate(satir[:3]):
                    c.t(x0 + 0.6, yy + 3.9 + li * 2.3, ln, size=sz, bold=bold, maxw=w - 0.8)
            yy += rh

    # ---------------- alt bölge: notlar, açıklamalar, güç bilançosu
    ny = 253.5
    c.t(14.0, ny, "NOTLAR", size=2.6, bold=True)
    rcr = C["RCR0"]
    A_ = P["aydinlatma_priz"]
    notlar = [
        f"1. Seçicilik: Q0 (Ir={q0['Ir_A']} A, Isd={q0['Isd_x']}xIr/{q0['tsd_ms']} ms) > RCR0 ({tr(rcr['I_dn_A'])} A/{rcr['gecikme_ms']} ms) > RCD1-2 (300 mA) / RCD3 ve RCBO (30 mA): I∆n oranı ≥ 3, zaman payı ≥ {C['secicilik']['zaman_payi_ms']} ms.",
        "2. N ve PE iletkenleri bina içinde birleştirilmez; N–PE köprüsü yalnızca trafo yıldız noktasındadır (TN-S). BN barası pano gövdesinden yalıtılmıştır.",
        f"3. Tek faz yüklerin faz ataması açgözlü (greedy) algoritma ile otomatik yapılmıştır (04_hesap_tablolari.xlsx → Yük Dengesi); faz dengesizliği %{tr(M['dengesizlik'])} (hedef ≤ %10).",
        "4. İç LED hatlarında LED sürücü kalkış akımı (inrush) nedeniyle C tipi eğri RCBO kullanılmıştır. Acil/yangın hatları: NHXH FE180/E90, öncelikli ayrı hat.",
        f"5. Dış aydınlatma: Q3 → KM1 (SA3 MAN-0-OTO, FT1 fotosel, ZR1 astronomik röle) → PJ1-PJ3; SPD2 Tip 2 önerilmiştir. Projektör gücü {tr(A_['projektor']['PROJEKTOR_GUCU_W'], 0)} W/adet varsayımdır.",
        "6. Boyler/ısıtıcı vb. yük eklenirse yük dengesi, kablo, kesici ve pano ısıl hesabı yeniden yapılır. Kumanda devresi prensip gösterimidir.",
        "7. Poz numaraları Çizim 2, 3, 6 ve tablolarla aynıdır. Uygulamadan önce yetkili elektrik mühendisi ve şebeke şartnamesi ile onaylanmalıdır.",
    ]
    yy = ny + 3.3
    for n_ in notlar:
        yy = c.tw_(14.0, yy, n_, 132.0, size=1.85, lh=1.17)
    # renk açıklaması
    lx = 150.0
    c.t(lx, ny, "İLETKEN RENKLERİ / SEMBOLLER", size=2.6, bold=True)
    for k, (ph, ad) in enumerate((("L1", "L1 kahverengi"), ("L2", "L2 siyah"), ("L3", "L3 gri"), ("N", "N mavi"), ("PE", "PE sarı-yeşil"))):
        yy2 = ny + 4.5 + k * 3.9
        c.iletken(lx, yy2 - 0.6, lx + 8.0, yy2 - 0.6, ph, w=0.9)
        c.t(lx + 10.0, yy2, ad, size=2.0)
    c.line(lx + 30.0, ny + 3.9, lx + 37.0, ny + 3.9, w=0.25, dash="0.8,0.6")
    c.t(lx + 39.0, ny + 4.5, "kumanda hattı", size=1.9)
    c.rect(lx + 30.0, ny + 7.6, 7.0, 3.2, sw=0.3, fill="#fff")
    c.t(lx + 39.0, ny + 10.4, "cihaz gövdesi", size=1.9)
    c.line(lx + 30.0, ny + 14.0, lx + 37.0, ny + 14.0, w=0.3)
    c.nokta(lx + 33.5, ny + 14.0, 0.6)
    c.t(lx + 39.0, ny + 14.7, "bağlantı noktası", size=1.9)
    c.t(lx + 30.0, ny + 20.6, "Semboller: IEC 60617", size=1.9, italic=True)
    # güç bilançosu
    gx = 209.0
    c.t(gx, ny, "GÜÇ BİLANÇOSU", size=2.6, bold=True)
    bil = [
        ("Kurulu güç", f"{tr(M['p_kurulu'], 2)} kW"),
        ("Talep gücü / kVA", f"{tr(M['p_talep'], 2)} kW / {tr(M['S_talep_kva'], 1)} kVA"),
        ("Hesap akımı Ib", f"{tr(M['Ib'])} A"),
        ("cosφ (ort.)", f"{tr(M['cos_ort'], 2)}"),
        ("Faz dengesizliği", f"%{tr(M['dengesizlik'])} (≤ %10)"),
        (f"Ib×{tr(1 + P['sistem']['gelecek_payi_yuzde'] / 100, 2)} ≤ Q0 Ir", f"{tr(M['Ib_gelecek'], 0)} ≤ {M['Ir']} A"),
        ("Ik (PN1 barası)", f"{tr(M['Ik_pano_kA'])} kA"),
        ("ΔU besleme (W0)", f"%{tr(M['besleme']['du'], 2)}"),
    ]
    for k, (a_, b_) in enumerate(bil):
        c.t(gx, ny + 4.6 + k * 3.8, a_, size=1.9)
        c.t(gx + 40.0, ny + 4.6 + k * 3.8, b_, size=1.9, bold=True, anchor="end", maxw=26)
        c.line(gx, ny + 5.6 + k * 3.8, gx + 40.0, ny + 5.6 + k * 3.8, w=0.1, c="#aaa")

    c.cerceve(P, "TEK HAT ŞEMASI", 1, "ölçeksiz")
    return c


# =====================================================================
#  ÇİZİM 2 : PANO ÖN GÖRÜNÜŞ VE YERLEŞİM
# =====================================================================
def satir_cihazlari(P, PN, satir):
    """Bir DIN ray satırı için (poz, x_mm_rayboyu, modul) listesi."""
    pn = P["pano"]
    out, cur = [], 0
    for gi, (g, m) in enumerate(satir["gruplar"]):
        if gi > 0:
            cur += pn["grup_boslugu_modul"]
        for poz in P["devre_gruplari"][g]["cihazlar"]:
            mw = cihaz_modul(P, poz)
            out.append((poz, cur * pn["modul_mm"], mw))
            cur += mw
    return out


def cihaz_turu(poz):
    if poz in ("Q1", "Q2", "Q3"):
        return "mcb4"
    if poz.startswith("RCD"):
        return "rcd"
    if poz == "KM1":
        return "kontaktor"
    if poz == "SPD2":
        return "spd"
    if poz in ("FT1", "ZR1"):
        return "role"
    if klima_mi(poz):
        return "mpcb"
    if poz.startswith("YD"):
        return "yedek"
    if poz in _MCB_POZ:
        return "mcb1"
    return "rcbo"


DOLGU = {"mcb1": "#e3e3e3", "mcb4": "#e3e3e3", "rcd": "#f1dca0", "mpcb": "#cfe5cf", "rcbo": "#d3e3f4", "yedek": "#ffffff",
         "kontaktor": "#dcd0f0", "spd": "#ebc9c9", "role": "#f6e4c4"}
_ETK = {"Q1": "KLM GRP1", "Q2": "KLM GRP2", "RCD1": "KLM GRP1", "RCD2": "KLM GRP2", "KR1": "KAM.", "RCD3": "SES RCD", "AC1": "ACİL", "YA1": "YANG.",
        "Q3": "DIŞ AYD", "KM1": "KONT.", "SPD2": "SPD T2", "FT1": "FOTO", "ZR1": "ZAMAN", "AV1": "AVİZE", "DK1": "DEKOR"}


def etiket_detay(poz):
    """Detay-A'da cihaz altına yazılan kısa hat adı."""
    if poz in _ETK:
        return _ETK[poz]
    if klima_mi(poz):
        return "KLM-" + poz[1:]
    m = re.fullmatch(r"([A-Z]+)(\d+)", poz)
    if m:
        return {"PZ": "PRZ-", "LD": "LED-", "PJ": "PRJ-", "SS": "SES-", "YD": "YEDEK"}.get(m.group(1), m.group(1) + "-") + (m.group(2) if m.group(1) != "YD" else "")
    return poz


def cizim2(P, M, PN, T, KAYIT, yol):
    c = Cizim(yol, 2)
    pn = P["pano"]
    C = P["cihazlar"]
    G, Y, D = PN["G"], PN["Y"], pn["D"]
    cokul = pn["cokul_mm"]
    Hb = Y - cokul
    S1 = 0.1
    mod = pn["modul_mm"]

    c.t(14, 17.2, f"ÇİZİM 2 — PANO ÖN GÖRÜNÜŞ VE İÇ YERLEŞİM (PN1)", size=4.0, bold=True)
    c.t(14, 22.4, f"{G} (G) x {Y} (Y) x {D} (D) mm, zemine monte metal pano, {pn['IP']}  |  Ön ve iç görünüş 1:10, DETAY-A 1:5  |  Ölçüler mm", size=2.4, color="#444")

    # ------------------------------------------------------------------ ÖN GÖRÜNÜŞ
    X0F, Y0F = 38.0, 40.0
    fx = lambda x: X0F + x * S1
    fy = lambda y: Y0F + y * S1
    c.t(X0F + G * S1 / 2, 27.5, "ÖN GÖRÜNÜŞ (kapı kapalı) — 1:10", size=2.8, bold=True, anchor="middle")
    # çokul ve gövde
    c.rect(fx(0), fy(Hb), G * S1, cokul * S1, sw=0.3, fill="#8d8d8d")
    c.t(fx(G / 2), fy(Hb) + cokul * S1 / 2 + 0.8, f"ÇOKUL {cokul}", size=2.0, anchor="middle", color="#fff")
    c.rect(fx(0), fy(0), G * S1, Hb * S1, sw=0.5, fill="#f4f6f8")
    c.rect(fx(20), fy(20), (G - 40) * S1, (Hb - 40) * S1, sw=0.2, stroke="#888")
    for yy in (200, 900, 1600):                                  # menteşeler
        c.rect(fx(-8), fy(yy), 8 * S1, 80 * S1, sw=0.25, fill="#555")
    # isim plakası
    c.rect(fx(230), fy(50), 340 * S1, 70 * S1, sw=0.3, fill="#fff")
    c.t(fx(400), fy(50) + 4.5, "PN1  ANA PANO", size=3.0, bold=True, anchor="middle", poz="PN1", maxw=33)
    c.poly([(fx(700), fy(130)), (fx(670), fy(50)), (fx(730), fy(50))], sw=0.3, fill="#ffd400", closed=True)
    c.t(fx(700), fy(110), "!", size=3.0, bold=True, anchor="middle")
    # ölçü cihazları
    def kare(x, y, w, poz, ad, ek=""):
        c.rect(fx(x), fy(y), w * S1, w * S1, sw=0.35, fill="#222")
        c.rect(fx(x + 8), fy(y + 8), (w - 16) * S1, (w - 16) * S1, sw=0.15, stroke="#777", fill="#e9f2ea")
        c.t(fx(x + w / 2), fy(y + w / 2) + 0.6, ad, size=min(3.2, w * S1 * 0.28), bold=True, anchor="middle")
        c.t(fx(x + w / 2), fy(y + w) + 2.6, poz, size=2.2, bold=True, anchor="middle", poz=poz)
    kare(90, 170, 96, "A1", "A")
    kare(210, 170, 96, "V1", "V")
    kare(330, 150, 144, "PM1", "kWh")
    c.t(fx(402), fy(150) + 4.0, "", size=1.0)
    for x, poz in ((138, "SA1"), (258, "SA2")):
        c.circle(fx(x), fy(350), 16 * S1, sw=0.3, fill="#333")
        c.line(fx(x), fy(350), fx(x) + 1.0, fy(350) - 1.2, c="#fff", w=0.3)
        c.t(fx(x), fy(350) + 4.2, poz, size=2.2, bold=True, anchor="middle", poz=poz)
    for k, (x, ph, col) in enumerate(((530, "L1", RENK["L1"]), (590, "L2", RENK["L2"]), (650, "L3", RENK["L3"]))):
        c.circle(fx(x), fy(190), 14 * S1, sw=0.3, fill=col)
        c.t(fx(x), fy(190) + 3.2, f"H{k + 1}", size=2.0, bold=True, anchor="middle", poz=f"H{k + 1}")
        c.t(fx(x), fy(190) - 2.0, ph, size=1.7, anchor="middle")
    # acil durdurma ve Q0 kolu
    c.circle(fx(700), fy(400), 38 * S1, sw=0.3, fill="#ffd400")
    c.circle(fx(700), fy(400), 24 * S1, sw=0.3, fill="#d0021b")
    c.t(fx(700), fy(400) + 6.4, "SB1", size=2.2, bold=True, anchor="middle", poz="SB1")
    c.rect(fx(150), fy(480), 150 * S1, 150 * S1, sw=0.3, fill="#d9d9d9")
    c.rect(fx(175), fy(520), 100 * S1, 60 * S1, sw=0.3, fill="#d0021b")
    c.t(fx(225), fy(480) - 1.0, "Q0", size=2.2, bold=True, anchor="middle", poz="Q0")
    c.t(fx(225), fy(630) + 3.0, "ana kesici kolu", size=1.7, anchor="middle")
    # dış aydınlatma kumanda bölmesi (kapıda): SA3 MAN-0-OTO + H4 durum lambası
    if "SA3" in P["cihazlar"] and any(h["grup"] == "PROJ" for h in M["hatlar"]):
        c.rect(fx(340), fy(500), 300 * S1, 190 * S1, sw=0.3, stroke="#444", fill="#fffbe6")
        c.t(fx(490), fy(500) + 3.0, "DIŞ AYDINLATMA KUMANDA", size=1.9, bold=True, anchor="middle")
        c.circle(fx(430), fy(600), 17 * S1, sw=0.3, fill="#333")
        c.line(fx(430), fy(600), fx(430) + 1.2, fy(600) - 0.8, c="#fff", w=0.3)
        c.t(fx(430) - 3.2, fy(600) - 2.2, "MAN", size=1.5, anchor="middle")
        c.t(fx(430), fy(600) - 2.6, "0", size=1.5, anchor="middle")
        c.t(fx(430) + 3.4, fy(600) - 2.2, "OTO", size=1.5, anchor="middle")
        c.t(fx(430), fy(600) + 4.6, "SA3", size=2.1, bold=True, anchor="middle", poz="SA3")
        c.circle(fx(550), fy(600), 11 * S1, sw=0.3, fill="#2e9e44")
        c.t(fx(550), fy(600) + 4.6, "H4", size=2.1, bold=True, anchor="middle", poz="H4")
        c.t(fx(550), fy(600) - 2.2, "dış aydınlatma açık", size=1.4, anchor="middle")

    # kilit
    c.circle(fx(745), fy(1000), 18 * S1, sw=0.3, fill="#bbb")
    c.line(fx(745) - 1.0, fy(1000), fx(745) + 1.0, fy(1000), w=0.3)
    c.t(fx(745) - 3.0, fy(1000) + 4.4, "KL1", size=2.0, bold=True, anchor="middle", poz="KL1")
    c.line(fx(775), fy(120), fx(775), fy(Hb - 120), w=0.2, dash="1,0.8", c="#888")
    # havalandırma ızgaraları
    for gx0 in (70, 530):
        c.rect(fx(gx0), fy(1450), 200 * S1, 300 * S1, sw=0.3, fill="#e9e9e9")
        for k in range(11):
            c.line(fx(gx0) + 0.8, fy(1450) + 1.5 + k * 2.6, fx(gx0 + 200) - 0.8, fy(1450) + 1.5 + k * 2.6, w=0.2, c="#666")
    c.t(fx(170), fy(1750) + 3.0, "GR1", size=2.2, bold=True, anchor="middle", poz="GR1")
    c.t(fx(630), fy(1750) + 3.0, "GR1", size=2.2, bold=True, anchor="middle", poz="GR1")
    c.t(fx(G / 2), fy(1150), "IP54 · Kapı contalı · Üç noktalı kilit", size=2.0, anchor="middle", color="#555")
    # ölçüler
    c.olcu(fx(0), fy(0), fx(G), fy(0), f"{G}", ofs=-6.0, size=2.4)
    c.olcu(fx(0), fy(0), fx(0), fy(Y), f"{Y}", ofs=-11.0, size=2.4, yatay=False)
    c.olcu(fx(0), fy(Hb), fx(0), fy(Y), f"{cokul}", ofs=-5.0, size=2.0, yatay=False)
    c.t(fx(G / 2), fy(Y) + 7.0, f"Derinlik D = {D} mm", size=2.4, anchor="middle", bold=True)

    # ------------------------------------------------------------------ İÇ GÖRÜNÜŞ
    X0I, Y0I = 145.0, 40.0
    ix = lambda x: X0I + x * S1
    iy = lambda y: Y0I + y * S1
    c.t(X0I + G * S1 / 2, 25.5, "İÇ GÖRÜNÜŞ (kapı açık, montaj plakası) — 1:10", size=2.8, bold=True, anchor="middle")
    c.rect(ix(0), iy(Hb), G * S1, cokul * S1, sw=0.3, fill="#8d8d8d")
    c.rect(ix(0), iy(0), G * S1, Hb * S1, sw=0.5, fill="#fbfbfb")
    pk, yk = pn["plaka_kenar_mm"], pn["yan_kanal_mm"]
    plate_y1 = Hb - 50
    c.rect(ix(pk), iy(100), (G - 2 * pk) * S1, (plate_y1 - 100) * S1, sw=0.25, stroke="#999", fill="#f1f1f1")
    c.t(ix(pk) + 1.0, iy(100) + 2.6, "MP1", size=2.0, bold=True, color="#555", poz="MP1")
    # dikey kanallar
    xd0 = pk
    xd1 = G - pk - yk
    for xk in (xd0, xd1):
        c.rect(ix(xk), iy(100), yk * S1, (plate_y1 - 100) * S1, sw=0.25, stroke="#777", fill="#dcdcdc")
        for yy in range(120, int(plate_y1) - 20, 40):
            c.line(ix(xk) + 0.5, iy(yy), ix(xk + yk) - 0.5, iy(yy), w=0.1, c="#999")
    # bölme sınırları
    a_top, a_bot = 100, 100 + pn["ust_bolme_mm"]
    b_bot = a_bot + pn["orta_bolme_mm"]
    n_sat = len(PN["satir_hesap"])
    rows_y0 = b_bot
    rows_y1 = b_bot + n_sat * pn["ray_adim_mm"]
    for yy in (a_bot, b_bot, rows_y1):
        c.line(ix(pk), iy(yy), ix(G - pk), iy(yy), w=0.2, dash="2,1", c="#555")
    ray_x0 = pk + yk
    ray_x1 = G - pk - yk
    c.t(ix(ray_x1 - 40), iy(rows_y0 + pn["ray_adim_mm"] - 8), "KN1", size=1.9, bold=True, anchor="middle", poz="KN1", halo=True)

    # ---- a) ÜST BÖLME : Q0, MX0, TOR0, CT'ler, üst ray cihazları
    q0x, q0y, q0w, q0h = ray_x0 + 10, 130, 140, 160
    c.rect(ix(q0x), iy(q0y), q0w * S1, q0h * S1, sw=0.35, fill="#4a4a4a")
    c.rect(ix(q0x + 20), iy(q0y + 20), (q0w - 40) * S1, 40 * S1, sw=0.2, fill="#2b2b2b")
    for k in range(4):
        c.rect(ix(q0x + 22 + k * 25), iy(q0y + 30), 14 * S1, 20 * S1, sw=0.15, fill="#ddd")
    c.t(ix(q0x + q0w / 2), iy(q0y + 110) + 0.6, "Q0", size=2.8, bold=True, anchor="middle", color="#fff", poz="Q0")
    c.t(ix(q0x + q0w / 2), iy(q0y + 140) + 0.5, f"{C['Q0']['govde_A']} A / {C['Q0']['Ir_A']} A", size=1.6, anchor="middle", color="#fff")
    c.rect(ix(q0x + q0w + 6), iy(q0y + 10), 40 * S1, 50 * S1, sw=0.3, fill="#fff")
    c.t(ix(q0x + q0w + 6 + 20), iy(q0y + 35) + 0.5, "MX0", size=1.7, bold=True, anchor="middle", poz="MX0")
    riser = {"N": q0x + 20, "L1": q0x + 55, "L2": q0x + 90, "L3": q0x + 125}
    bus_y = {"L1": b_bot - 270, "L2": b_bot - 230, "L3": b_bot - 190, "N": b_bot - 150}
    for ph, xr in riser.items():
        c.rect(ix(xr - 10), iy(q0y + q0h), 20 * S1, (bus_y[ph] + 15 - (q0y + q0h)) * S1, sw=0.15, fill=RENK[ph])
    # toroid
    c.circle(ix(q0x + 70), iy(380), 60 * S1, sw=0.5, fill="none", stroke="#222")
    c.circle(ix(q0x + 70), iy(380), 52 * S1, sw=0.15, fill="none", stroke="#555")
    c.t(ix(q0x + 70), iy(380) + 0.5, "TOR0", size=1.9, bold=True, anchor="middle", poz="TOR0", halo=True)
    # akım trafoları
    for k, ph in enumerate(("L1", "L2", "L3")):
        xr = riser[ph]
        c.rect(ix(xr - 16), iy(450), 32 * S1, 60 * S1, sw=0.3, fill="#e8e0d0")
        c.rect(ix(xr - 8), iy(462), 16 * S1, 36 * S1, sw=0.15, fill=RENK[ph])
        c.t(ix(xr), iy(450) + 8.4 + (k % 2) * 2.2, f"CT{k + 1}", size=1.7, bold=True, anchor="middle", poz=f"CT{k + 1}")
    # üst ray cihazları
    x_cur = q0x + q0w + 6 + 40 + 10
    ust_y = 140
    ust_h = 90
    c.rect(ix(x_cur), iy(ust_y + 35), (ray_x1 - x_cur) * S1, 20 * S1, sw=0.15, fill="#c8c8c8")
    labels_top = []
    gap = pn["grup_boslugu_modul"] * mod
    xx = x_cur
    for poz in pn["ust_ray_grup"]:
        mw = pn["ust_ray_modul"][poz]
        wmm = mw * mod
        fillc = {"RCR0": "#f1dca0", "F1": "#e3e3e3", "SPD1": "#e8c9c9", "FSR1": "#d3e3f4", "F2": "#e3e3e3"}[poz]
        c.rect(ix(xx), iy(ust_y), wmm * S1, ust_h * S1, sw=0.3, fill=fillc)
        for k in range(mw):
            c.rect(ix(xx + k * mod + 3), iy(ust_y + 12), (mod - 6) * S1, 22 * S1, sw=0.12, fill="#fff")
        labels_top.append((poz, xx + wmm / 2))
        xx += wmm + gap
    for k, (poz, xc) in enumerate(labels_top):
        c.t(ix(xc), iy(ust_y + ust_h) + 2.4 + (k % 2) * 2.4, poz, size=1.8, bold=True, anchor="middle", poz=poz)
    # XT1 ölçüm klemensi
    c.rect(ix(x_cur), iy(330), 130 * S1, 55 * S1, sw=0.3, fill="#d9d9d9")
    for k in range(12):
        c.rect(ix(x_cur + 5 + k * 10.3), iy(340), 8 * S1, 35 * S1, sw=0.12, fill="#fff")
    c.t(ix(x_cur + 65), iy(330) + 7.6, "XT1", size=1.8, bold=True, anchor="middle", poz="XT1")
    c.line(ix(riser["L3"] + 28), iy(480), ix(x_cur), iy(355), w=0.2, dash="1.2,0.8", c="#444")
    c.t(ix(x_cur + 140), iy(345), "CT sekonder → XT1 →", size=1.7, color="#444")
    c.t(ix(x_cur + 140), iy(375), "kapıdaki cihazlar:", size=1.7, color="#444")
    c.t(ix(x_cur + 140), iy(405), "PM1, A1, V1, SA1, SA2,", size=1.7, color="#444")
    c.t(ix(x_cur + 140), iy(435), "H1–H3, SB1", size=1.7, color="#444")
    # a) bölüm etiketi
    # ---- b) ORTA BÖLME : ana bara
    for ph, poz in (("L1", "BL1"), ("L2", "BL2"), ("L3", "BL3"), ("N", "BN")):
        yy = bus_y[ph]
        c.rect(ix(ray_x0 - 20), iy(yy), (ray_x1 - ray_x0 + 40) * S1, 30 * S1, sw=0.25, fill=RENK[ph])
        c.t(ix(ray_x1 + 5) - 1.0, iy(yy) + 2.5, poz, size=1.7, bold=True, anchor="end", color="#fff" if ph != "L3" else "#000", poz=poz)
    # bara besleme düşüşü (sağ kanal içinde)
    xb = G - pk - yk / 2
    for k, ph in enumerate(("L1", "L2", "L3", "N")):
        c.line(ix(xb - 16 + k * 10.7), iy(bus_y[ph] + 15), ix(xb - 16 + k * 10.7), iy(rows_y1 - 40), w=0.35, c=RENK[ph])
    c.t(ix(xb), iy(rows_y1 - 40) + 2.6, "besleme", size=1.5, anchor="middle", color="#444")

    # ---- c) ALT BÖLMELER : DIN ray satırları
    satir_yer = []
    for ri, sat in enumerate(PN["satir_hesap"]):
        ry = rows_y0 + ri * pn["ray_adim_mm"]
        dev_y = ry + 12
        dev_h = 80
        c.rect(ix(ray_x0), iy(dev_y + 22), (ray_x1 - ray_x0) * S1, 35 * S1, sw=0.15, fill="#c8c8c8")   # DIN ray
        c.rect(ix(ray_x0), iy(ry + pn["ray_adim_mm"] - 24), (ray_x1 - ray_x0) * S1, 22 * S1, sw=0.2, stroke="#888", fill="#e2e2e2")  # yatay kanal
        devs = satir_cihazlari(P, PN, sat)
        for k, (poz, xmm, mw) in enumerate(devs):
            xa = ray_x0 + xmm
            tur = cihaz_turu(poz)
            if tur == "yedek":
                c.rect(ix(xa), iy(dev_y), mw * mod * S1, dev_h * S1, sw=0.25, stroke="#666", fill="#fff", dash="0.8,0.6")
            else:
                c.rect(ix(xa), iy(dev_y), mw * mod * S1, dev_h * S1, sw=0.3, fill=DOLGU[tur])
                nl = mw if tur in ("mcb4", "rcd", "mpcb") else 1
                lw = (mw * mod) / nl
                for q in range(nl):
                    c.rect(ix(xa + q * lw + lw * 0.2), iy(dev_y + 10), lw * 0.6 * S1, 24 * S1, sw=0.12, fill="#fff")
            c.t(ix(xa + mw * mod / 2), iy(dev_y + dev_h) + 2.3 + (k % 2) * 2.1, poz, size=1.7, bold=True, anchor="middle", poz=poz)
        satir_yer.append((ry, devs))
    # ray etiketleri
    c.t(ix(ray_x0) + 0.5, iy(rows_y0) - 1.0, "DR1", size=1.7, bold=True, color="#555", poz="DR1")

    # ---- alt bölge: N ve PE klemens barası, rakorlar
    bz0 = rows_y1
    yN, yPE = bz0 + 90, bz0 + 150
    c.rect(ix(ray_x0), iy(yN), (ray_x1 - ray_x0) * S1, 28 * S1, sw=0.25, fill=RENK["N"])
    c.rect(ix(ray_x0), iy(yPE), (ray_x1 - ray_x0) * S1, 28 * S1, sw=0.25, fill=RENK["PE_G"])
    for k in range(int((ray_x1 - ray_x0) // 24)):
        c.rect(ix(ray_x0 + 6 + k * 24), iy(yN + 7), 10 * S1, 14 * S1, sw=0.1, fill="#fff")
        c.rect(ix(ray_x0 + 6 + k * 24), iy(yPE + 7), 10 * S1, 14 * S1, sw=0.1, fill=RENK["PE_Y"])
    c.t(ix(ray_x0) + 1.0, iy(yN) - 1.0, "BN — N klemens barası (yalıtımlı)", size=1.7, bold=True, color="#1d4f91", poz="BN")
    c.t(ix(ray_x0) + 1.0, iy(yPE) - 1.0, "BPE — PE klemens barası (gövdeye bağlı)", size=1.7, bold=True, color="#1c6a31", poz="BPE")
    # N-PE ayrımı notu
    c.t(ix(ray_x1) - 1.0, iy(yN) - 1.0, "BN–BPE köprüsü YOK", size=1.6, anchor="end", color="#a00", bold=True)
    # rakor plakası
    c.rect(ix(pk), iy(Hb - 50), (G - 2 * pk) * S1, 50 * S1, sw=0.3, fill="#b9b9b9")
    n_k = sum(1 for h in M["hatlar"] if klima_mi(h["id"]))
    n_d = len(M["hatlar"]) - n_k
    tot = 1 + n_k + n_d
    xs_g = [ray_x0 + 30 + k * (ray_x1 - ray_x0 - 60) / (tot - 1 + 1e-9) for k in range(tot)]
    kinds = ["RK0"] + ["RK1"] * n_k + ["RK2"] * n_d
    for xg, kd in zip(xs_g, kinds):
        r = {"RK0": 31, "RK1": 16, "RK2": 12.5}[kd] * S1
        c.circle(ix(xg), iy(Hb - 25), r, sw=0.3, fill="#444")
        c.circle(ix(xg), iy(Hb - 25), r * 0.55, sw=0.15, fill="#bbb")
    c.t(ix(xs_g[0]), iy(Hb - 50) - 4.6, "RK0", size=1.8, bold=True, anchor="middle", poz="RK0")
    c.t(ix(xs_g[1] + 18), iy(Hb - 50) - 1.4, "RK1", size=1.8, bold=True, anchor="middle", poz="RK1")
    c.t(ix(xs_g[1 + n_k] + 25), iy(Hb - 50) - 1.4, "RK2", size=1.8, bold=True, anchor="middle", poz="RK2")
    c.t(ix(G / 2), iy(Hb - 50) + 0.4, "", size=1.0)
    # W0 yolu: RK0 -> sol kanal -> Q0
    xw = xd0 + yk / 2
    c.poly([(ix(xs_g[0]), iy(Hb - 50)), (ix(xs_g[0]), iy(bz0 + 190)), (ix(xw), iy(bz0 + 190)), (ix(xw), iy(q0y + q0h * 0.5)), (ix(q0x), iy(q0y + q0h * 0.5))],
           sw=0.7, stroke="#7a1f1f")
    c.t(ix(xw) - 0.6, iy(1180), "W0 NYY 4x95 → Q0", size=1.7, bold=True, rot=-90, anchor="middle", color="#7a1f1f", poz="W0")

    # ---- bölme etiketleri ve ölçü zinciri
    def bolme(y0, y1, metin):
        c.line(ix(0) - 3.0, iy(y0) + 0.5, ix(0) - 3.0, iy(y1) - 0.5, w=0.4, c="#555")
        c.t(ix(0) - 4.4, (iy(y0) + iy(y1)) / 2, metin, size=2.0, bold=True, anchor="middle", rot=-90, color="#333")
    bolme(a_top, a_bot, "a) ÜST BÖLME")
    bolme(a_bot, b_bot, "b) ORTA BÖLME")
    bolme(b_bot, rows_y1, "c) ALT BÖLMELER (DIN ray)")
    bolme(rows_y1, Hb, "kablo/klemens bölgesi")
    zinc = [(0, 100, "100"), (100, a_bot, f"{pn['ust_bolme_mm']}"), (a_bot, b_bot, f"{pn['orta_bolme_mm']}")]
    for ri in range(n_sat):
        zinc.append((rows_y0 + ri * pn["ray_adim_mm"], rows_y0 + (ri + 1) * pn["ray_adim_mm"], f"{pn['ray_adim_mm']}"))
    zinc.append((rows_y1, Hb, f"{int(Hb - rows_y1)}"))
    zinc.append((Hb, Y, f"{cokul}"))
    for y0, y1, tx in zinc:
        c.olcu(ix(G), iy(y0), ix(G), iy(y1), tx, ofs=5.0, size=1.8, yatay=False)
    c.olcu(ix(G), iy(0), ix(G), iy(Y), f"{Y}", ofs=11.0, size=2.2, yatay=False)
    # genişlik zinciri
    c.olcu(ix(0), iy(0), ix(G), iy(0), f"{G}", ofs=-9.0, size=2.2)
    segs = [(0, pk, f"{pk}"), (pk, pk + yk, f"{yk}"), (pk + yk, G - pk - yk, f"{int(PN['ray_uzunlugu_mm'])}"), (G - pk - yk, G - pk, f"{yk}"), (G - pk, G, f"{pk}")]
    for x0_, x1_, tx in segs:
        c.olcu(ix(x0_), iy(0), ix(x1_), iy(0), tx, ofs=-3.0, size=1.7)

    # ------------------------------------------------------------------ DETAY-A (1:5)
    S2 = 0.2
    DX0, DY0 = 250.0, 37.0
    dxf = lambda x: DX0 + (x - ray_x0) * S2
    c.t(DX0, 27.5, "DETAY-A — DIN ray satırları (alt bölmeler) — 1:5", size=2.8, bold=True)
    for ri, (ry, devs) in enumerate(satir_yer):
        y0 = DY0 + ri * pn["ray_adim_mm"] * S2
        sat = PN["satir_hesap"][ri]
        c.rect(DX0, y0 + 22 * S2 + 1.6, (ray_x1 - ray_x0) * S2, 35 * S2, sw=0.12, fill="#cfcfcf")
        for k, (poz, xmm, mw) in enumerate(devs):
            xa = DX0 + xmm * S2
            tur = cihaz_turu(poz)
            wmm = mw * mod * S2
            if tur == "yedek":
                c.rect(xa, y0 + 12 * S2, wmm, 80 * S2, sw=0.25, stroke="#555", fill="#fff", dash="0.8,0.6")
            else:
                c.rect(xa, y0 + 12 * S2, wmm, 80 * S2, sw=0.3, fill=DOLGU[tur])
                nl = mw if tur in ("mcb4", "rcd", "mpcb") else 1
                lw = wmm / nl
                for q in range(nl):
                    c.rect(xa + q * lw + lw * 0.2, y0 + 12 * S2 + 2.2, lw * 0.6, 4.4, sw=0.1, fill="#fff")
            c.t(xa + wmm / 2, y0 + 12 * S2 + 16.0 + 2.4, poz, size=2.1, bold=True, anchor="middle", poz=poz)
            c.t(xa + wmm / 2, y0 + 12 * S2 + 16.0 + 4.8, etiket_detay(poz), size=1.55, anchor="middle", maxw=max(wmm, 6) + 0.8)
        # satır başlığı (kanal şeridi)
        grup_ad = " + ".join(P["devre_gruplari"][g]["ad"] for g, _ in sat["gruplar"])
        c.rect(DX0, y0 + (pn["ray_adim_mm"] - 24) * S2, (ray_x1 - ray_x0) * S2, 22 * S2, sw=0.15, stroke="#999", fill="#ececec")
        c.t(DX0 + 1.0, y0 + (pn["ray_adim_mm"] - 24) * S2 + 3.2, f"SATIR {ri + 1}: {grup_ad} — {sat['modul']}/{sat['kapasite']} modül (%{tr(100.0 * sat['modul'] / sat['kapasite'], 0)})",
            size=1.9, bold=True, color="#333", maxw=(ray_x1 - ray_x0) * S2 - 1.5)
    # renk açıklaması (başlık altında tek satır)
    for k, (tur, ad) in enumerate((("mcb4", "MCB 1P+N/4P"), ("rcd", "RCD 2P/4P"), ("mpcb", "MPCB 3P"), ("rcbo", "RCBO 1P+N"),
                                    ("kontaktor", "kontaktör 3P"), ("spd", "SPD"), ("role", "röle"), ("yedek", "boş yuva"))):
        lx = DX0 + k * 19.6
        c.rect(lx, 30.2, 3.0, 2.4, sw=0.25, fill=DOLGU[tur], dash="0.6,0.4" if tur == "yedek" else None)
        c.t(lx + 4.0, 32.2, ad, size=1.5)
    c.t(DX0 + (ray_x1 - ray_x0) * S2 + 1.5, 36.2, f"modül {tr(mod)} mm", size=1.4, color="#555")

    # ------------------------------------------------------------------ MALZEME / POZ LİSTESİ
    liste = [r for r in satirlari_grupla(KAYIT) if r["kat"] == "pano" and 2 in r["cizim"]]
    LY0 = 193.0
    BX0 = 241.0
    c.t(BX0, LY0 - 2.4, "MALZEME / POZ LİSTESİ", size=2.6, bold=True)
    ncol = 3
    colw = 54.0
    xs_ = [BX0 + k * (colw + 2.0) for k in range(ncol)]
    hdr_h = 3.4
    rh_ = 2.85

    def kisa(ad, mx=34):
        t_ = re.split(r" \(|, ", ad)[0]
        if tw(t_, 1.65) > mx:
            while tw(t_ + "…", 1.65) > mx and len(t_) > 5:
                t_ = t_[:-1]
            t_ += "…"
        return t_
    n_per = math.ceil(len(liste) / ncol)
    for k_ in range(ncol):
        xcol = xs_[k_]
        c.rect(xcol, LY0, colw, hdr_h, sw=0.25, fill="#e3e8ef")
        c.t(xcol + 0.8, LY0 + 2.5, "POZ", size=1.7, bold=True)
        c.t(xcol + 16.0, LY0 + 2.5, "AÇIKLAMA", size=1.7, bold=True)
        c.t(xcol + colw - 0.8, LY0 + 2.5, "ADET", size=1.7, bold=True, anchor="end")
    for i_, r in enumerate(liste):
        k_, j_ = divmod(i_, n_per)
        xcol = xs_[k_]
        yy = LY0 + hdr_h + j_ * rh_
        c.rect(xcol, yy, colw, rh_, sw=0.1, stroke="#999")
        c.t(xcol + 0.8, yy + 2.1, r["poz"], size=1.65, bold=True, maxw=15.0)
        c.t(xcol + 16.0, yy + 2.1, kisa(r["ad"]), size=1.65, maxw=colw - 22.5)
        c.t(xcol + colw - 0.8, yy + 2.1, f"{r['adet']}", size=1.65, anchor="end")
    liste_son = LY0 + hdr_h + n_per * rh_
    if liste_son > 248.5:
        uyar(f"[2] Malzeme listesi alt sınırı aşıyor ({liste_son:.1f} mm)")

    # ------------------------------------------------------------------ alt bölge: hesaplar ve notlar
    ny = 253.5
    c.t(14.0, ny, "DIN RAY YERLEŞİM HESABI", size=2.5, bold=True)
    c.t(14.0, ny + 3.6, f"Ray boyu {int(PN['ray_uzunlugu_mm'])} mm → {PN['modul_kapasite']} modül; 1P+N=2, 3P=3, 4P=4 modül", size=1.8, color="#444")
    for k, sat in enumerate(PN["satir_hesap"]):
        gtxt = "+".join(g for g, _ in sat["gruplar"])
        uyg = "UYGUN" if sat["modul"] <= sat["kapasite"] else "SIĞMIYOR"
        c.t(14.0, ny + 7.4 + k * 3.2, f"Satır {k + 1} ({gtxt})", size=1.9)
        c.t(66.0, ny + 7.4 + k * 3.2, f"{sat['modul']}/{sat['kapasite']} modül, %{tr(100.0 * sat['modul'] / sat['kapasite'], 0)}  {uyg}", size=1.9, bold=True, anchor="start")
    k = len(PN["satir_hesap"])
    c.t(14.0, ny + 7.4 + k * 3.2, "Üst ray (Q0 yanı)", size=1.9)
    c.t(66.0, ny + 7.4 + k * 3.2, f"{PN['ust_ray_modul_toplam']}/{PN['ust_ray_kapasite']} modül  " + ("UYGUN" if PN["ust_ray_modul_toplam"] <= PN["ust_ray_kapasite"] else "SIĞMIYOR"), size=1.9, bold=True)
    if PN["uyarilar"]:
        c.t(14.0, ny + 7.4 + (k + 1) * 3.2, "UYARI: " + PN["uyarilar"][0][:70], size=1.8, color="#a00")
    else:
        c.t(14.0, ny + 7.4 + (k + 1) * 3.2, f"Sığma kontrolü: tüm satırlar uygun (boş yer: Satır 1-2 %{tr(100.0 - 100.0 * PN['satir_hesap'][0]['modul'] / PN['modul_kapasite'], 0)})", size=1.8, color="#1c6a31")

    isl = PN["isil"]
    hx = 112.0
    c.t(hx, ny, "ISIL KONTROL (kayıp güç / havalandırma)", size=2.5, bold=True)
    yy = ny + 3.7
    yy = c.tw_(hx, yy, f"Toplam kayıp güç Pv ≈ {tr(isl['Pv'], 0)} W (çalışma akımlarında).", 62.0, size=1.9, bold=True)
    yy = c.tw_(hx, yy, f"Pano yüzeyi Ae = {tr(isl['Ae'], 2)} m²; ΔT = {tr(isl['dT'], 0)} K (iç {P['sistem']['pano_ic_sicaklik_max_C']:.0f} °C, ortam {P['sistem']['ortam_sicakligi_C']:.0f} °C); ısı atımı Pd = k·Ae·ΔT ≈ {tr(isl['Pd'], 0)} W (k = {tr(pn['yuzey_isi_katsayisi_W_m2K'])} W/m²K).", 62.0, size=1.8)
    if isl["fan_gerekli"]:
        yy = c.tw_(hx, yy + 0.6, f"Pv > 0,8·Pd → FAN GEREKLİ: ≥ {tr(isl['hava_m3h'], 0)} m³/h (FN1 + termostat TS1).", 62.0, size=1.9, bold=True, color="#a00")
    else:
        yy = c.tw_(hx, yy + 0.6, "Pv < 0,8·Pd → ayrıca fan gerekmez; doğal havalandırma + filtreli ızgara GR1 (IP54) yeterli.", 62.0, size=1.9, bold=True, color="#1c6a31")
    yy = c.tw_(hx, yy + 0.6, "Ortam > 35 °C veya ek yük (boyler vb.) halinde FN1 + TS1 eklenir ve hesap yenilenir.", 62.0, size=1.8)
    nx = 178.0
    notlar = [
        "NOTLAR",
        "• Etiketler: poz no + hat adı (DETAY-A'da okunaklı ölçekte).",
        "• BN barası gövdeden yalıtılmış; BN–BPE köprüsü yok.",
        "• Giriş/çıkış kabloları alt rakor plakasından (RK0-RK2).",
        "• KN1: güç ve zayıf akım kabloları ayrı bölmede.",
        "• Pano IP54, kapı contalı, 3 noktalı kilitli (KL1).",
        "• Poz no'lar Çizim 1 ve 3 ile tablolarla aynıdır.",
    ]
    for k, n_ in enumerate(notlar):
        c.t(nx, ny + k * 3.3, n_, size=2.5 if k == 0 else 1.85, bold=(k == 0), maxw=70)

    c.cerceve(P, "PANO ÖN GÖRÜNÜŞ VE İÇ YERLEŞİM", 2, "1:10 / 1:5")
    return c


# =====================================================================
#  ÇİZİM 3 : TOPRAKLAMA VE YILDIRIMDAN KORUNMA
# =====================================================================
def klip_baslat(c, x, y, w, h):
    c._klip_id = getattr(c, "_klip_id", 0) + 1
    cid = f"klip{c._klip_id}"
    cp = c.d.defs.add(c.d.clipPath(id=cid))
    cp.add(c.d.rect((x, y), (w, h)))
    grp = c.d.g(clip_path=f"url(#{cid})")
    c.d.add(grp)
    c._g_onceki = c.g
    c.g = grp


def klip_bitir(c):
    c.g = c._g_onceki


def tarama(c, x0, y0, x1, y1, aralik=2.2, renk="#8a7a60", w=0.15):
    """Toprak taraması (45° çizgiler), dikdörtgen içinde."""
    klip_baslat(c, x0, y0, x1 - x0, y1 - y0)
    k = x0 - (y1 - y0)
    while k < x1:
        c.line(k, y1, k + (y1 - y0), y0, w=w, c=renk)
        k += aralik
    klip_bitir(c)


def cizim3(P, M, PN, T, KAYIT, yol):
    c = Cizim(yol, 3)
    a, b = T["bina_a_m"], T["bina_b_m"]
    o = T["halka_ofset_m"]
    S = 1000.0 / 150.0                  # 1 m = 6,667 mm (1:150)
    X0, Y0 = 41.0, 47.0
    pm = lambda x, y: (X0 + x * S, Y0 + y * S)
    KOL = {"halka": "#a8520c", "kur": "#c0151f", "pe": RENK["PE_G"]}
    GRN = "#1c6a31"

    # ------------------------------------------------------------------ A) KAT PLANI
    c.t(14, 17.5, "ÇİZİM 3 — TOPRAKLAMA VE YILDIRIMDAN KORUNMA", size=4.0, bold=True)
    c.t(14, 22.8, "A) KAT PLANI (üstten görünüş) — ölçek 1:150", size=2.8, bold=True)

    # bina
    c.rect(*pm(0, 0), a * S, b * S, sw=0.7, fill="#f3f3f3")
    c.rect(*pm(0.4, 0.4), (a - 0.8) * S, (b - 0.8) * S, sw=0.25, stroke="#555")
    # kubbe
    kx, ky = T["kubbe_merkez_m"]
    kr = T["kubbe_cap_m"] / 2
    cx_, cy_ = pm(kx, ky)
    c.circle(cx_, cy_, kr * S, sw=0.4, fill="#e9eef5")
    c.circle(cx_, cy_, kr * S * 0.66, sw=0.15, stroke="#7a8aa0")
    c.circle(cx_, cy_, kr * S * 0.33, sw=0.15, stroke="#7a8aa0")
    for ang in range(0, 360, 45):
        c.line(cx_, cy_, cx_ + kr * S * math.cos(math.radians(ang)), cy_ + kr * S * math.sin(math.radians(ang)), w=0.12, c="#9aa7ba")
    c.t(cx_, cy_ + kr * S * 0.66 + 4.2, f"KUBBE Ø{tr(T['kubbe_cap_m'])} m", size=2.2, anchor="middle", color="#355", halo=True)
    # minare
    mx, my = T["minare_merkez_m"]
    mt = T["minare_taban_m"]
    mxp, myp = pm(mx, my)
    c.rect(mxp - mt * S / 2, myp - mt * S / 2, mt * S, mt * S, sw=0.5, fill="#e9eef5")
    c.line(mxp - mt * S / 2, myp - mt * S / 2, mxp + mt * S / 2, myp + mt * S / 2, w=0.15, c="#888")
    c.line(mxp - mt * S / 2, myp + mt * S / 2, mxp + mt * S / 2, myp - mt * S / 2, w=0.15, c="#888")
    c.circle(mxp, myp, mt * S * 0.33, sw=0.3, fill="#fff")
    c.t(mxp - mt * S / 2 - 1.4, myp + 6.6, "MİNARE", size=2.0, anchor="end", color="#355")
    c.t(mxp - mt * S / 2 - 1.4, myp + 9.0, f"{tr(mt)}x{tr(mt)} m, h={tr(T['minare_yuksekligi_m'], 0)} m", size=1.8, anchor="end", color="#355")
    # elektrik odası, PN1, EPB1
    px, py = T["pano_konum_m"]
    ex, ey = T["epb_konum_m"]
    c.rect(*pm(0.4, b - 3.7), 4.7 * S, 3.3 * S, sw=0.25, stroke="#666", dash="1.2,0.8")
    c.t(*pm(0.7, b - 3.95), "ELEKTRİK ODASI", size=1.7, color="#555")
    xa, ya = pm(px, py)
    xb, yb = pm(ex, ey)
    # eş potansiyel bağlantı noktaları (EB1-EB6) önce çizilir (altta kalır)
    for e in T["eb"]:
        x, y = e["xy"]
        xp_, yp_ = pm(x, y)
        c.line(xb, yb - 1.0, xp_, yp_, c=KOL["pe"], w=0.28, dash="2,1")
    c.rect(*pm(px - 0.4, py - 0.2), 0.8 * S, 0.4 * S, sw=0.4, fill="#d8d8d8")
    c.rect(*pm(ex - 0.2, ey - 0.07), 0.4 * S, 0.14 * S, sw=0.3, fill="#e9d2b0")
    c.line(xa + 2.7, ya, xb - 1.4, yb, c=KOL["pe"], w=0.5)

    # çatı çevre iletkeni ÇÇ1 ve çatı bağlantıları (kırmızı kesik çizgi)
    kd = 0.25
    c.poly([pm(kd, kd), pm(a - kd, kd), pm(a - kd, b - kd), pm(kd, b - kd)], closed=True, sw=0.5, stroke=KOL["kur"], dash="2.4,1.0")
    c.line(*pm(kx, ky), *pm(kx, b - kd), w=0.5, c=KOL["kur"], dash="2.4,1.0")
    c.line(*pm(mx, my), *pm(a - kd, kd), w=0.5, c=KOL["kur"], dash="2.4,1.0")
    c.t(*pm(a * 0.27, b * 0.082), "ÇÇ1 çatı çevre iletkeni", size=2.0, color=KOL["kur"], bold=True, poz="ÇÇ1", halo=True)

    # topraklama halkası
    x0r, y0r, x1r, y1r = T["halka"]
    c.poly([pm(x0r, y0r), pm(x1r, y0r), pm(x1r, y1r), pm(x0r, y1r)], closed=True, sw=1.0, stroke=KOL["halka"])

    # iniş iletkenleri (duvardan halkaya) + ölçüm ayırma noktaları
    def kenar_norm(yer):
        return {"KD": (1, -1), "GD": (1, 1), "GB": (-1, 1), "KB": (-1, -1), "G": (0, 1), "K": (0, -1), "D": (1, 0), "B": (-1, 0)}[yer]
    inis_lbl = {}
    for d in T["inis"]:
        x, y = d["xy"]
        nx, ny = kenar_norm(d["yer"])
        xw, yw = pm(x, y)
        xr_, yr_ = pm(x + nx * o, y + ny * o)
        c.line(xw, yw, xr_, yr_, c=KOL["kur"], w=0.9)
        xm, ym = (xw + xr_) / 2, (yw + yr_) / 2
        c.rect(xm - 1.0, ym - 1.0, 2.0, 2.0, sw=0.3, fill="#fff", stroke=KOL["kur"])
        c.circle(xw, yw, 1.35, sw=0, fill=KOL["kur"])
        inis_lbl[d["poz"]] = (xw, yw, xr_, yr_, nx, ny, d["ok"])

    # topraklama çubukları ve kontrol çukurları
    cukur = {k["cubuk"]: k["poz"] for k in T["cukurlar"]}
    for cb in T["cubuklar"]:
        x, y = cb["xy"]
        xp_, yp_ = pm(x, y)
        if cb["poz"] in cukur:
            c.rect(xp_ - 2.0, yp_ - 2.0, 4.0, 4.0, sw=0.5, fill="#fff", stroke="#222")
        c.circle(xp_, yp_, 1.05, sw=0.3, fill="#000")

    def dis_yon(x, y):
        sx = -1 if x <= x0r + 1e-6 else (1 if x >= x1r - 1e-6 else 0)
        sy = -1 if y <= y0r + 1e-6 else (1 if y >= y1r - 1e-6 else 0)
        return sx, sy
    for cb in T["cubuklar"]:
        x, y = cb["xy"]
        xp_, yp_ = pm(x, y)
        sx, sy = dis_yon(x, y)
        kose = (sx != 0 and sy != 0)
        if kose:
            satir = [cb["poz"]]
            if cb["poz"] in cukur:
                satir.append(cukur[cb["poz"]])
            for d in T["inis"]:
                if abs(d["xy"][0] - (x - sx * o)) < 1e-6 and abs(d["xy"][1] - (y - sy * o)) < 1e-6:
                    satir.append(d["ok"])
            ax = "end" if sx < 0 else "start"
            bx_ = xp_ + sx * 3.4
            for i, t_ in enumerate(satir):
                if sy > 0 or (sx < 0 and sy < 0):
                    yy_ = yp_ + (4.2 if sy > 0 else 5.0) + i * 2.5
                else:
                    yy_ = yp_ - 2.0 - (len(satir) - 1 - i) * 2.5
                col = "#222" if t_.startswith("TÇ") else ("#555" if t_.startswith("KÇ") else KOL["kur"])
                c.t(bx_, yy_, t_, size=2.0, bold=True, anchor=ax, poz=t_, color=col)
        else:
            c.t(xp_ + sx * 3.8, yp_ + sy * 3.2 + (1.0 if sy >= 0 else 0.5), cb["poz"], size=2.0, bold=True,
                anchor=("middle" if sx == 0 else ("end" if sx < 0 else "start")), poz=cb["poz"])
    # orta kenar iniş iletkenleri: ÖK etiketi dışarıda
    for pz, (xw, yw, xr_, yr_, nx, ny, ok) in inis_lbl.items():
        if nx != 0 and ny != 0:
            continue
        c.t(xr_ + nx * 5.2, yr_ + ny * 4.0 + 1.0, ok, size=2.0, bold=True, anchor="middle", color=KOL["kur"], poz=ok)
    # iniş iletkeni etiketleri (bina içi)
    for pz, (xw, yw, xr_, yr_, nx, ny, ok) in inis_lbl.items():
        if nx != 0 and ny != 0:
            ix_, iy_ = xw - nx * 3.6, yw - ny * 3.4
        else:
            ix_, iy_ = xw - nx * 5.0, yw - ny * 4.4
        if pz == "İL1":
            ix_, iy_ = xw - 15.5, yw + 3.0
        c.t(ix_, iy_ + 0.8, pz, size=2.0, bold=True, anchor="middle", color=KOL["kur"], poz=pz, halo=True)

    # Franklin çubukları YK1-YK3
    for d in T["yk"]:
        xp_, yp_ = pm(*d["xy"])
        r = 1.7
        c.circle(xp_, yp_, r, sw=0.5, stroke=KOL["kur"], fill="#fff")
        for ang in (0, 45, 90, 135):
            c.line(xp_ - r * 1.5 * math.cos(math.radians(ang)), yp_ - r * 1.5 * math.sin(math.radians(ang)),
                   xp_ + r * 1.5 * math.cos(math.radians(ang)), yp_ + r * 1.5 * math.sin(math.radians(ang)), w=0.3, c=KOL["kur"])
    xyk1, yyk1 = pm(*T["yk"][0]["xy"])
    c.t(xyk1 - 3.0, yyk1 - 3.8, "YK1", size=2.0, bold=True, color=KOL["kur"], anchor="end", poz="YK1", halo=True)
    xyk2, yyk2 = pm(*T["yk"][1]["xy"])
    c.t(xyk2 + 3.0, yyk2 - 2.4, "YK2", size=2.0, bold=True, color=KOL["kur"], poz="YK2", halo=True)
    xyk3, yyk3 = pm(*T["yk"][2]["xy"])
    c.t(xyk3 + 3.6, yyk3 - 2.6, "YK3", size=2.0, bold=True, color=KOL["kur"], poz="YK3", halo=True)

    # eş potansiyel noktaları + etiketler
    for e in T["eb"]:
        xp_, yp_ = pm(*e["xy"])
        c.poly([(xp_, yp_ - 1.7), (xp_ + 1.7, yp_), (xp_, yp_ + 1.7), (xp_ - 1.7, yp_)], closed=True, sw=0.4, fill="#fff", stroke=KOL["pe"])
    ebp = {e["poz"]: pm(*e["xy"]) for e in T["eb"]}
    c.t(ebp["EB1"][0] + 3.6, ebp["EB1"][1] + 0.7, "EB1 su borusu", size=1.9, bold=True, color=GRN, poz="EB1", halo=True)
    gx_, gy_ = ebp["EB2"]
    c.t(gx_ + 2.6, gy_ - 2.2, "EB2 gaz borusu", size=1.9, bold=True, color=GRN, poz="EB2", halo=True)
    c.line(gx_, gy_ + 1.8, gx_, gy_ + 3.4, w=0.3, c="#555")
    c.rect(gx_ - 1.8, gy_ + 3.4, 3.6, 0.9, sw=0.3, fill="#fff")
    c.rect(gx_ - 1.8, gy_ + 4.8, 3.6, 0.9, sw=0.3, fill="#fff")
    c.t(gx_ + 2.6, gy_ + 5.2, "İF1", size=1.9, bold=True, poz="İF1")
    c.t(ebp["EB3"][0] + 3.0, ebp["EB3"][1] - 1.4, "EB3 kablo tavası", size=1.9, bold=True, color=GRN, poz="EB3", halo=True)
    c.t(ebp["EB4"][0] - 3.0, ebp["EB4"][1] + 0.7, "EB4 HVAC kanalı", size=1.9, bold=True, color=GRN, anchor="end", poz="EB4", halo=True)
    c.t(ebp["EB5"][0] - 2.0, ebp["EB5"][1] + 4.6, "EB5", size=1.9, bold=True, color=GRN, anchor="end", poz="EB5", halo=True)
    c.t(ebp["EB6"][0] + 3.0, ebp["EB6"][1] + 0.5, "EB6 kubbe çeliği", size=1.9, bold=True, color=GRN, poz="EB6", halo=True)
    # PN1 / EPB1 etiketleri (çizim üstüne)
    c.t(xa - 0.5, ya - 3.2, "PN1", size=2.2, bold=True, anchor="middle", poz="PN1", halo=True)
    c.t(xb + 1.0, yb - 3.2, "EPB1", size=2.2, bold=True, anchor="middle", poz="EPB1", halo=True)
    c.t((xa + xb) / 2 + 1.0, ya + 3.6, "PE1", size=1.8, bold=True, color=GRN, anchor="middle", poz="PE1", halo=True)

    # trafo merkezi, TE1, W0/PE3 hattı
    ttx, tty = 17.5, 28.2
    c.circle(ttx, tty, 3.2, sw=0.4, fill="#fff")
    c.circle(ttx, tty + 4.2, 3.2, sw=0.4, fill="#fff")
    c.poly([(ttx, tty - 1.6), (ttx - 1.6, tty + 1.0), (ttx + 1.6, tty + 1.0)], sw=0.25, closed=True)
    for ang in (90, 210, 330):
        c.line(ttx, tty + 4.2, ttx + 1.7 * math.cos(math.radians(ang)), tty + 4.2 - 1.7 * math.sin(math.radians(ang)), w=0.25)
    c.t(ttx + 5.2, tty - 0.2, "TR1 trafo merkezi", size=2.2, bold=True, poz="TR1")
    c.t(ttx + 5.2, tty + 2.6, "(plan dışı)", size=1.8)
    # TE1 (yıldız noktası -> ayrı elektrot)
    yte = tty + 4.2
    c.iletken(ttx + 3.2, yte, 72.0, yte, "PE", w=0.5)
    c.nokta(72.0, yte, 0.6)
    c.line(72.0, yte, 72.0, yte + 1.6, c=RENK["PE_G"], w=0.5)
    c.topraklama(72.0, yte + 1.6, olcek=0.8)
    c.t(77.0, yte - 1.0, "TE1: trafo yıldız noktası topraklaması (ayrı elektrot, ≤ 2 Ω)", size=1.9, bold=True, poz="TE1")
    c.t(77.0, yte + 2.1, "N–PE köprüsü yalnızca trafo yıldız noktasında (TN-S)", size=1.8)
    # W0 + PE3 güzergâhı
    yl = Y0 + (b - 1.5) * S
    c.iletken(ttx, tty + 7.4 + 0.6, ttx, yl, "PE", w=0.5)
    c.line(ttx + 1.5, tty + 8.0, ttx + 1.5, yl + 1.5, c="#333", w=0.45)
    c.iletken(ttx, yl, X0, yl, "PE", w=0.5)
    c.line(ttx + 1.5, yl + 1.5, X0, yl + 1.5, c="#333", w=0.45)
    c.kopma(ttx + 0.75, 95.0, yatay=False)
    c.kopma(ttx + 0.75, 99.0, yatay=False)
    c.t(ttx + 5.0, 99.8, f"{tr(M['besleme']['L'])} m", size=1.9, bold=True)
    b0 = M["besleme"]
    c.t(ttx + 4.0, 160.0, f"W0: {b0['tip']} {b0['damar']} mm² Cu + PE3: 1x50 mm² Cu (aynı hendek), L = {tr(b0['L'])} m", size=1.9, bold=True, rot=-90, poz=["W0", "PE3"])
    c.t(ttx - 3.4, 160.0, "trafo–pano mesafesi ölçekli değildir", size=1.7, rot=-90, color="#555")
    # TH1 etiketi
    c.t(*pm(2.7, y0r - 0.45), "TH1", size=2.2, bold=True, anchor="middle", color=KOL["halka"], poz="TH1", halo=True)
    # ölçüler
    c.olcu(*pm(a, 0), *pm(a, b), f"{tr(b)} m", ofs=19.0, size=2.0, yatay=False)
    c.t(*pm(a / 2, 18.2), f"ANA İBADET HACMİ {tr(a)} x {tr(b)} m", size=2.0, anchor="middle", color="#555", halo=True)
    # kuzey oku ve ölçek çubuğu
    c.poly([(100.0, 20.5), (97.8, 28.5), (100.0, 26.5), (102.2, 28.5)], closed=True, sw=0.3, fill="#222")
    c.t(103.6, 24.4, "K", size=2.0, bold=True)
    for k in range(4):
        c.rect(112.0 + k * S * 2.5, 25.2, S * 2.5, 1.8, sw=0.25, fill="#fff" if k % 2 else "#222")
    c.t(112.0, 24.2, "0", size=1.7, anchor="middle")
    c.t(112.0 + S * 5, 24.2, "5", size=1.7, anchor="middle")
    c.t(112.0 + S * 10, 24.2, "10 m", size=1.7, anchor="middle")

    # ------------------------------------------------------------------ B2) YUVARLANAN KÜRE (1:500)
    c.t(198.0, 17.5, "B) DETAY VE KESİTLER", size=3.0, bold=True)
    yz = T["yuvarlanan_kure"]
    R = yz["R"]
    sc = 2.0                                                   # 1 m = 2 mm (1:500)
    GX0, GX1, GY0, GY1 = 198.0, 408.0, 29.0, 88.0
    gxp = lambda x: 303.0 + (x - 10.0) * sc
    gzp = lambda z: GY1 - 2.5 - z * sc
    hg = T["govde_yuksekligi_m"]
    hk = T["bina_yuksekligi_m"] - hg
    hm = T["minare_yuksekligi_m"]
    mtw = T["minare_taban_m"]
    c.t(198.0, 26.4, f"B2) YUVARLANAN KÜRE YÖNTEMİ — koruma seviyesi {T['ypk_seviye']} (R = {tr(R, 0)} m) — 1:500", size=2.6, bold=True)
    klip_baslat(c, GX0, GY0, GX1 - GX0, GY1 - GY0)
    c.rect(GX0, GY0, GX1 - GX0, GY1 - GY0, sw=0, fill="#fbfdff")
    xg, zarf = yz["xg"], yz["zarf"]
    pts = [(gxp(x), gzp(max(0.0, z))) for x, z in zip(xg, zarf) if np.isfinite(z)]
    poly_ = [(gxp(xg[0]), gzp(0))] + pts + [(gxp(xg[-1]), gzp(0))]
    c.poly(poly_, closed=True, sw=0, fill="#cdeccd")
    c.poly(pts, sw=0.55, stroke="#2e7d32")
    for xs_ in (-20.0, 2.0, 16.0, 40.0):
        k = int(round((xs_ - xg[0]) / 0.5))
        cxm, czm = yz["merkez"][k]
        c.circle(gxp(cxm), gzp(czm), R * sc, sw=0.35, stroke="#d46a00", dash="2.4,1.2")
        c.circle(gxp(cxm), gzp(czm), 0.9, sw=0, fill="#d46a00")
    k0 = int(round((-20.0 - xg[0]) / 0.5))
    ccx, ccz = yz["merkez"][k0]
    c.line(gxp(ccx), gzp(ccz), gxp(ccx) + R * sc * math.cos(math.radians(-60)), gzp(ccz) - R * sc * math.sin(math.radians(-60)), w=0.3, c="#d46a00")
    c.line(GX0, gzp(0), GX1, gzp(0), w=0.6)
    klip_bitir(c)
    c.rect(GX0, GY0, GX1 - GX0, GY1 - GY0, sw=0.3, stroke="#888")
    tarama(c, GX0, gzp(0), GX1, GY1, aralik=2.0)
    # bina silueti
    c.rect(gxp(0), gzp(hg), a * sc, hg * sc, sw=0.5, fill="#f3f3f3")
    c.path(f"M {gxp(kx - kr):.2f} {gzp(hg):.2f} A {kr * sc:.2f} {hk * sc:.2f} 0 0 1 {gxp(kx + kr):.2f} {gzp(hg):.2f}", sw=0.5, fill="#e9eef5")
    c.rect(gxp(mx - mtw / 2), gzp(hm - 4.5), mtw * sc, (hm - 4.5 - hg) * sc, sw=0.4, fill="#eef0f3")
    c.poly([(gxp(mx - mtw / 2), gzp(hm - 4.5)), (gxp(mx), gzp(hm)), (gxp(mx + mtw / 2), gzp(hm - 4.5))], closed=True, sw=0.4, fill="#cfd4dc")
    for d in T["yk"]:
        xx = gxp(d["xy"][0])
        c.line(xx, gzp(d["z"]), xx, gzp(d["z"] + d["boy"]), c=KOL["kur"], w=0.7)
        c.circle(xx, gzp(d["z"] + d["boy"]), 0.8, sw=0, fill=KOL["kur"])
    c.t(gxp(mx) + 2.4, gzp(T["yk"][0]["z"] + 1.0), "YK1", size=1.8, bold=True, color=KOL["kur"], poz="YK1", halo=True)
    c.t(gxp(kx) + 1.8, gzp(T["yk"][1]["z"] + 1.0) - 1.6, "YK2", size=1.8, bold=True, color=KOL["kur"], poz="YK2", halo=True)
    c.t(gxp(0.0) - 2.2, gzp(T["yk"][2]["z"]) - 1.2, "YK3", size=1.8, bold=True, color=KOL["kur"], anchor="end", poz="YK3", halo=True)
    c.line(gxp(0), gzp(hg) - 0.4, gxp(a), gzp(hg) - 0.4, c=KOL["kur"], w=0.5, dash="1.6,0.8")
    c.t(gxp(ccx) + 8.0, GY0 + 13.0, f"R = {tr(R, 0)} m", size=2.0, bold=True, color="#b45309")
    c.t(GX0 + 2.0, GY1 - 8.0, "Korunan hacim (yeşil bölge)", size=1.9, bold=True, color="#2e7d32", halo=True)
    sonuc = "Gövde, kubbe, minare ve çatı kenarları korunan hacim içinde" if yz["korunuyor"] else "UYARI: yapı korunan hacim dışına çıkıyor — ek yakalayıcı gerekli"
    c.t(GX1 - 2.0, GY0 + 4.6, sonuc, size=1.8, bold=True, anchor="end", color="#2e7d32" if yz["korunuyor"] else "#a00", maxw=100)
    c.t(GX1 - 2.0, GY0 + 7.6, f"(zarf kontrolü: en büyük sapma {tr(max(0.0, yz['ihlal_m']), 2)} m, ağ toleransı 0,25 m)", size=1.6, anchor="end", color="#444")
    c.t(GX1 - 2.0, GY0 + 10.6, "Hava sonlandırma: YK1-YK3 Franklin çubukları + ÇÇ1 çatı çevre iletkeni", size=1.6, anchor="end", color="#444")

    # ------------------------------------------------------------------ B1) BİNA KESİTİ (1:200)
    BX0, BYG = 202.0, 227.0
    s2 = 5.0
    bx = lambda x: BX0 + x * s2
    bz = lambda z: BYG - z * s2
    c.t(198.0, 92.4, "B1) BİNA KESİTİ (yan görünüş) — 1:200", size=2.6, bold=True)
    c.line(bx(-5), bz(0), bx(a + 5), bz(0), w=0.7)
    tarama(c, bx(-5), bz(0), bx(a + 5), BYG + 21.0, aralik=2.4)
    c.rect(bx(0), bz(hg), a * s2, hg * s2, sw=0.6, fill="#f3f3f3")
    c.path(f"M {bx(kx - kr):.2f} {bz(hg):.2f} A {kr * s2:.2f} {hk * s2:.2f} 0 0 1 {bx(kx + kr):.2f} {bz(hg):.2f}", sw=0.6, fill="#e9eef5")
    c.line(bx(kx - kr), bz(hg), bx(kx + kr), bz(hg), w=0.3)
    mxl, mxr = bx(mx - mtw / 2), bx(mx + mtw / 2)
    c.rect(mxl, bz(hm - 4.5), mxr - mxl, (hm - 4.5 - hg) * s2, sw=0.5, fill="#eef0f3")
    c.rect(mxl - 1.5, bz(hm - 8.0), mxr - mxl + 3.0, 1.2, sw=0.4, fill="#cfd4dc")
    c.poly([(mxl, bz(hm - 4.5)), ((mxl + mxr) / 2, bz(hm)), (mxr, bz(hm - 4.5))], closed=True, sw=0.5, fill="#cfd4dc")
    for d in T["yk"]:
        xx = bx(d["xy"][0])
        c.line(xx, bz(d["z"]), xx, bz(d["z"] + d["boy"]), c=KOL["kur"], w=0.9)
        c.circle(xx, bz(d["z"] + d["boy"]), 0.7, sw=0, fill=KOL["kur"])
    c.t(bx(mx) - 2.0, bz(T["yk"][0]["z"] + T["yk"][0]["boy"]) + 3.6, f"YK1 Franklin L={tr(T['yk'][0]['boy'])} m", size=1.8, bold=True, color=KOL["kur"], anchor="end", poz="YK1", halo=True)
    c.t(bx(kx), bz(T["yk"][1]["z"] + T["yk"][1]["boy"]) - 1.6, f"YK2 Franklin L={tr(T['yk'][1]['boy'])} m", size=1.8, bold=True, color=KOL["kur"], anchor="middle", poz="YK2", halo=True)
    c.t(bx(0) + 1.8, bz(T["yk"][2]["z"] + T["yk"][2]["boy"]) - 1.2, f"YK3 L={tr(T['yk'][2]['boy'])} m", size=1.8, bold=True, color=KOL["kur"], anchor="start", poz="YK3", halo=True)
    c.line(bx(0), bz(hg) - 0.5, bx(a), bz(hg) - 0.5, c=KOL["kur"], w=0.6, dash="2.2,0.9")
    c.t(bx(a * 0.30), bz(hg) - 1.8, "ÇÇ1 çatı çevre iletkeni", size=1.8, color=KOL["kur"], bold=True, poz="ÇÇ1", halo=True)
    c.line(bx(kx), bz(hg + hk), bx(kx), bz(hg), c=KOL["kur"], w=0.7, dash="2.2,0.9")
    bxlist = {"İL1": bx(mx + mtw / 2 - 0.25), "İL4": bx(0.25), "İL3": bx(kx)}
    for pz in ("İL1", "İL3", "İL4"):
        if not any(d["poz"] == pz for d in T["inis"]):
            continue
        xx = bxlist[pz]
        ztop = {"İL1": hm - 1.0, "İL3": hg, "İL4": hg}[pz]
        c.line(xx, bz(ztop), xx, bz(-0.8), c=KOL["kur"], w=1.0)
        c.rect(xx - 1.1, bz(T["olcum_ayirma_yuksekligi_m"]) - 1.1, 2.2, 2.2, sw=0.3, fill="#fff", stroke=KOL["kur"])
        okp = [d["ok"] for d in T["inis"] if d["poz"] == pz][0]
        ofx = 3.0 if pz != "İL1" else -3.0
        c.t(xx + ofx, bz(T["olcum_ayirma_yuksekligi_m"]) + 0.8, f"{pz}/{okp}", size=1.8, bold=True, color=KOL["kur"], anchor="start" if ofx > 0 else "end", poz=[pz, okp], halo=True)
    c.line(bx(-o), bz(-0.8), bx(a + o), bz(-0.8), c=KOL["halka"], w=1.3)
    c.t(bx(a * 0.5), bz(-0.8) + 4.6, f"TH1 halka, {T['halka_iletkeni']}, −{tr(T['halka_derinlik_m'])} m", size=1.8, bold=True, color=KOL["halka"], anchor="middle", poz="TH1")
    for xx_m in (-o, a + o):
        xx = bx(xx_m)
        c.line(xx, bz(-0.8), xx, bz(-0.8) + T["cubuk_boy_m"] * s2, c="#222", w=1.0)
        c.rect(xx - 2.2, bz(0) - 0.1, 4.4, 4.6, sw=0.35, fill="#fff")
    c.t(bx(-o) - 2.8, bz(-0.8) + 9.0, "TÇ1", size=1.8, bold=True, anchor="end", poz="TÇ1", halo=True)
    c.t(bx(a + o) + 2.8, bz(-0.8) + 9.0, "TÇ4", size=1.8, bold=True, anchor="start", poz="TÇ4", halo=True)
    c.t(bx(-o) - 2.8, bz(0) + 3.0, "KÇ1", size=1.8, bold=True, anchor="end", poz="KÇ1", halo=True)
    c.t(bx(a + o) + 2.8, bz(0) + 3.0, "KÇ2", size=1.8, bold=True, anchor="start", poz="KÇ2", halo=True)
    c.t(bx(a + o) + 2.8, bz(-0.8) + 12.0, f"çubuk {tr(T['cubuk_boy_m'], 0)} m", size=1.6, anchor="start")
    c.olcu(bx(a + 3.0), bz(0), bx(a + 3.0), bz(hg), f"{tr(hg, 0)} m", ofs=0.0, size=1.8, yatay=False)
    c.olcu(bx(a + 4.9), bz(0), bx(a + 4.9), bz(T["bina_yuksekligi_m"]), f"{tr(T['bina_yuksekligi_m'], 0)} m", ofs=0.0, size=1.8, yatay=False)
    c.olcu(bx(a + 6.8), bz(0), bx(a + 6.8), bz(hm), f"{tr(hm, 0)} m", ofs=0.0, size=1.8, yatay=False)

    # ------------------------------------------------------------------ B6) EPB ŞEMASI
    ex0 = 341.0
    c.t(ex0, 92.4, "B6) EPB ŞEMASI", size=2.6, bold=True)
    c.t(ex0, 95.6, "(hangi parçalar bağlanıyor)", size=1.9, color="#444")
    bar_x = 352.0
    y_top, y_bot = 112.0, 232.0
    c.rect(bar_x - 1.8, y_top, 3.6, y_bot - y_top, sw=0.4, fill="#e9d2b0")
    c.t(bar_x - 4.0, y_top + 6.0, "EPB1", size=2.4, bold=True, anchor="end", poz="EPB1", rot=0) if False else None
    c.t(bar_x, y_top - 8.6, "EPB1", size=2.6, bold=True, anchor="middle", poz="EPB1")
    c.iletken(bar_x, y_top - 7.0, bar_x, y_top, "PE", w=0.6)
    c.t(bar_x + 2.6, y_top - 4.2, "PE1 50 mm² Cu", size=1.8, bold=True, poz="PE1")
    c.t(bar_x + 2.6, y_top - 1.8, "← PN1 BPE barası", size=1.7)
    items = [
        ("EB1", "Su borusu", "16 mm² Cu", "EB1"),
        ("EB2", "Gaz borusu (İF1 sonrası)", "16 mm² Cu", "EB2"),
        ("EB3", "Kablo tavası / PE", "16 mm² Cu", "EB3"),
        ("EB4", "HVAC kanalı, klima gövde", "16 mm² Cu", "EB4"),
        ("EB5", "Minare çelik konstr.", "16 mm² Cu", "EB5"),
        ("EB6", "Kubbe/çatı çeliği", "16 mm² Cu", "EB6"),
        ("İL1–İL6", "İniş iletkenleri (ÖK)", "→ halka TH1", None),
        ("SPD1", "Parafudr PE ucu", "25 mm² Cu", "SPD1"),
        ("HB", "Haberleşme, ses, kamera", "6 mm² Cu", None),
        ("TE1", "Trafo (PE3 ile)", "50 mm² Cu", "TE1"),
    ]
    pitch = (y_bot - y_top - 8.0) / (len(items) - 1)
    for k, (kod, ad, kes, pz) in enumerate(items):
        yy = y_top + 4.0 + k * pitch
        c.rect(bar_x + 1.8, yy - 0.9, 2.0, 1.8, sw=0.2, fill="#8a5a1a")
        c.iletken(bar_x + 3.8, yy, bar_x + 7.0, yy, "PE", w=0.5)
        c.rect(bar_x + 7.0, yy - 3.4, 47.0, 6.8, sw=0.3, fill="#fff")
        c.t(bar_x + 8.0, yy - 0.5, kod, size=1.9, bold=True, poz=pz if pz else None, maxw=12)
        c.t(bar_x + 8.0, yy + 2.5, kes, size=1.7, color="#444", maxw=14)
        c.t(bar_x + 22.5, yy - 0.5, ad, size=1.7, maxw=31)
    c.iletken(bar_x, y_bot, bar_x, y_bot + 6.0, "PE", w=0.6)
    c.t(bar_x + 2.6, y_bot + 3.2, "PE2 50 mm² Cu → halka TH1", size=1.8, bold=True, poz=["PE2", "TH1"])
    c.t(ex0 - 0.5, y_bot + 11.0, "Tüm sistemler (güç, yıldırım, haberleşme) tek", size=1.7, color="#444")
    c.t(ex0 - 0.5, y_bot + 13.4, "ortak topraklamaya bağlanır.", size=1.7, color="#444")

    # ------------------------------------------------------------------ B3-B5 DETAYLAR (sol alt)
    dy0 = 199.0
    dx = 14.0
    c.t(dx, dy0, "B3) EKSOTERMİK KAYNAK BAĞLANTISI", size=2.1, bold=True)
    c.t(dx, dy0 + 2.9, "EK1 — şematik, ölçeksiz", size=1.7, color="#444", poz="EK1")
    cx0, cy0 = dx + 25.0, dy0 + 24.0
    c.rect(cx0 - 13, cy0 - 12, 26, 22, sw=0.5, fill="#d8d8d8")
    c.rect(cx0 - 5, cy0 - 4, 10, 8, sw=0.3, fill="#ff9d4a")
    c.circle(cx0, cy0, 2.4, sw=0.3, fill="#fff")
    c.rect(cx0 - 19, cy0 - 2.2, 12, 4.4, sw=0.3, fill="#e9b97d")
    c.rect(cx0 - 2.2, cy0 - 18, 4.4, 14, sw=0.3, fill="#e9b97d")
    c.rect(cx0 + 7, cy0 - 2.2, 12, 4.4, sw=0.3, fill="#e9b97d")
    c.rect(cx0 - 2.2, cy0 + 4, 4.4, 8, sw=0.3, fill="#e9b97d")
    c.t(cx0 + 15.0, cy0 - 7.5, "grafit kalıp", size=1.7)
    c.line(cx0 + 14.0, cy0 - 6.2, cx0 + 11.0, cy0 - 8.0, w=0.2)
    c.t(cx0 + 15.0, cy0 + 9.5, "kaynak metali", size=1.7)
    c.line(cx0 + 5.5, cy0 + 3.0, cx0 + 14.5, cy0 + 8.2, w=0.2)
    c.t(dx, dy0 + 7.6, "Cu iletken / çubuk", size=1.7)
    c.t(dx, dy0 + 43.0, "T bağlantı (şerit–çubuk, halka–iniş iletkeni)", size=1.6, color="#444")

    dx4 = 74.0
    c.t(dx4, dy0, "B4) ÇUBUK + KONTROL ÇUKURU KESİTİ", size=2.1, bold=True)
    c.t(dx4, dy0 + 2.9, "TÇ / KÇ — şematik, ölçeksiz", size=1.7, color="#444", poz=["TÇ1", "KÇ1"])
    gy = dy0 + 12.0
    c.line(dx4, gy, dx4 + 56.0, gy, w=0.6)
    tarama(c, dx4, gy, dx4 + 56.0, dy0 + 47.5, aralik=2.2)
    cxp = dx4 + 20.0
    c.rect(cxp - 8.0, gy - 0.5, 16.0, 11.0, sw=0.4, fill="#fff")
    c.rect(cxp - 9.5, gy - 1.6, 19.0, 1.6, sw=0.4, fill="#8a8a8a")
    c.line(cxp, gy - 0.5, cxp, dy0 + 35.0, c="#222", w=1.2)
    c.line(cxp - 8.0, gy + 7.0, cxp + 8.0, gy + 7.0, c=KOL["halka"], w=1.0)
    c.rect(cxp - 1.8, gy + 4.8, 3.6, 4.4, sw=0.3, fill="#fff")
    c.kopma(cxp, dy0 + 28.0, yatay=False)
    c.t(cxp + 11.5, gy - 0.4, "kapak (zemin seviyesi)", size=1.7, halo=True)
    c.t(cxp + 11.5, gy + 3.4, "kontrol çukuru 300x300 mm", size=1.7, halo=True)
    c.t(cxp + 11.5, gy + 7.2, "halka iletkeni, −0,8 m", size=1.7, color=KOL["halka"], halo=True)
    c.t(cxp + 11.5, gy + 11.0, "kelepçe + ölçüm klemensi", size=1.7, halo=True)
    c.t(cxp + 3.0, dy0 + 33.0, f"çubuk Ø{tr(T['cubuk_cap_mm'], 0)} x {tr(T['cubuk_boy_m'], 0)} m, bakır kaplı", size=1.7, bold=True, halo=True)

    dx5 = 136.0
    c.t(dx5, dy0, "B5) ŞERİT–ÇUBUK BAĞLANTISI", size=2.1, bold=True)
    c.t(dx5, dy0 + 2.9, "şematik, ölçeksiz", size=1.7, color="#444")
    sy_ = dy0 + 14.0
    c.rect(dx5 + 1, sy_ - 1.7, 38.0, 3.4, sw=0.3, fill="#cfd8de")
    c.t(dx5 + 1, sy_ - 6.4, "30x3,5 mm galvaniz şerit (veya 50 mm² Cu)", size=1.7)
    c.rect(dx5 + 24.0, sy_ - 5.0, 9.0, 10.0, sw=0.4, fill="#d9c28f")
    c.circle(dx5 + 26.8, sy_ - 2.4, 0.9, sw=0.25, fill="#fff")
    c.circle(dx5 + 30.2, sy_ + 2.4, 0.9, sw=0.25, fill="#fff")
    c.rect(dx5 + 26.2, sy_ + 5.0, 4.6, 26.0, sw=0.3, fill="#e9b97d")
    c.t(dx5 + 35.0, sy_ + 8.0, "kelepçe (M10 cıvata,", size=1.7)
    c.t(dx5 + 35.0, sy_ + 10.6, "paslanmaz) veya", size=1.7)
    c.t(dx5 + 35.0, sy_ + 13.2, "eksotermik kaynak", size=1.7)
    c.t(dx5 + 33.0, sy_ + 24.0, f"Ø{tr(T['cubuk_cap_mm'], 0)} çubuk", size=1.7)
    c.t(dx5 + 1.0, sy_ + 21.0, "Bağlantı bitümlü bantla", size=1.7)
    c.t(dx5 + 1.0, sy_ + 23.6, "korunur (korozyon).", size=1.7)

    # ------------------------------------------------------------------ alt bölge: notlar
    ny = 253.5
    c.t(14.0, ny, "NOTLAR", size=2.6, bold=True)
    notlar = [
        "1. Tüm sistemler (güç, yıldırım, haberleşme) tek ortak topraklamaya bağlanır; ölçüm hedefi ≤ 1–2 Ω, kuru mevsimde tekrar ölçülür.",
        f"2. Halka: bina çevresinde {tr(o)} m dışta, {tr(T['halka_derinlik_m'])} m derinlikte {T['halka_iletkeni']} (≈ {tr(T['halka_cevre'], 0)} m); {len(T['cubuklar'])} adet Ø{tr(T['cubuk_cap_mm'], 0)}x{tr(T['cubuk_boy_m'], 0)} m çubuk (aralık ≤ {tr(T['cubuk_aralik_gercek'], 1)} m), {len(T['cukurlar'])} kontrol çukuru.",
        f"3. Yıldırım koruma seviyesi {T['ypk_seviye']} (R = {tr(T['kure_R'], 0)} m): {len(T['inis'])} iniş iletkeni (ort. aralık {tr(T['inis_ort_aralik'], 1)} m ≤ {tr(T['inis_aralik'], 0)} m), her biri ölçüm ayırma noktalı (ÖK) ve halkaya eksotermik kaynaklıdır.",
        "4. Gaz borusunda izolasyon flanşı (İF1) kullanılır; su/gaz borusu, kablo tavası, HVAC kanalı ve çelik konstrüksiyon EPB1'e bağlanır.",
        "5. TE1: trafo yıldız noktası için ayrı elektrot (≤ 2 Ω); PE3 ile PN1 BPE barasına bağlanır. N–PE köprüsü yalnız trafodadır.",
        "6. Yeraltı bağlantıları eksotermik kaynaklıdır; ölçümler kontrol çukurlarından yapılır. Poz no'lar Çizim 1-2 ve tablolarla aynıdır.",
    ]
    yy = ny + 3.3
    for n_ in notlar:
        yy = c.tw_(14.0, yy, n_, 158.0, size=1.9, lh=1.2)
    lx = 176.0
    c.t(lx, ny, "GÖSTERİMLER", size=2.6, bold=True)
    leg = [("halka", f"TH1 halka ({T['halka_iletkeni']}, −{tr(T['halka_derinlik_m'])} m)"), ("kur", "yıldırım iletkenleri (iniş, ÇÇ1)"), ("pe", "PE / eş potansiyel (EB)")]
    for k, (kk, ad) in enumerate(leg):
        yy2 = ny + 4.5 + k * 3.8
        if kk == "pe":
            c.iletken(lx, yy2 - 0.6, lx + 8, yy2 - 0.6, "PE", w=0.7)
        else:
            c.line(lx, yy2 - 0.6, lx + 8, yy2 - 0.6, c=KOL[kk], w=0.9, dash="2.4,1.0" if kk == "kur" else None)
        c.t(lx + 10, yy2, ad, size=1.8)
    yy2 = ny + 4.5 + 3 * 3.8
    c.circle(lx + 1.5, yy2 - 0.8, 1.0, sw=0, fill="#000")
    c.t(lx + 4.0, yy2, "topraklama çubuğu", size=1.8)
    c.rect(lx + 32.0, yy2 - 2.4, 3.0, 3.0, sw=0.4)
    c.t(lx + 36.0, yy2, "kontrol çukuru", size=1.8)
    yy3 = yy2 + 3.8
    c.rect(lx + 0.5, yy3 - 2.0, 2.0, 2.0, sw=0.3, stroke=KOL["kur"], fill="#fff")
    c.t(lx + 4.0, yy3, "ölçüm ayırma noktası", size=1.8)
    c.poly([(lx + 33.0, yy3 - 1.0), (lx + 34.6, yy3 - 2.5), (lx + 36.2, yy3 - 1.0), (lx + 34.6, yy3 + 0.5)], closed=True, sw=0.3, stroke=KOL["pe"], fill="#fff")
    c.t(lx + 37.5, yy3, "eş pot. noktası", size=1.8)

    c.cerceve(P, "TOPRAKLAMA VE YILDIRIMDAN KORUNMA", 3, "1:150 / 1:200 / 1:500")
    return c


# =====================================================================
#  ÇİZİM 6 : AYDINLATMA VE PRİZ YERLEŞİM PLANI (1:100)
# =====================================================================
HAT_RENK_PALET = ["#c2185b", "#1565c0", "#ef6c00", "#2e7d32", "#6a1b9a", "#b8860b", "#00838f", "#c62828", "#37474f", "#827717", "#000000"]


SES_RENK_PALET = ["#795548", "#039be5", "#558b2f", "#ff8a80"]


def hat_renkleri(M):
    ids = [h["id"] for h in M["hatlar"] if h.get("elemanlar")]
    sonuc = {i: HAT_RENK_PALET[k % len(HAT_RENK_PALET)] for k, i in enumerate(ids)}
    for k, h in enumerate(x for x in M["hatlar"] if x["grup"] == "SES"):      # ses kolları: aydınlatma renklerinden ayrı renkler
        sonuc[h["id"]] = SES_RENK_PALET[k % len(SES_RENK_PALET)]
    return sonuc


def cizim6(P, M, PN, T, KAYIT, yol):
    c = Cizim(yol, 6)
    a, b = T["bina_a_m"], T["bina_b_m"]
    S = 10.0                                   # 1 m = 10 mm  (1:100)
    X0, Y0 = 26.0, 35.0
    pm = lambda x, y: (X0 + x * S, Y0 + y * S)
    renk = hat_renkleri(M)
    hat = {h["id"]: h for h in M["hatlar"]}
    el = {e["poz"]: e for e in M["elemanlar"]}
    PJ = P["aydinlatma_priz"]["projektor"]
    kx, ky = T["kubbe_merkez_m"]
    kr = T["kubbe_cap_m"] / 2
    mx, my = T["minare_merkez_m"]
    mt = T["minare_taban_m"]
    px, py = T["pano_konum_m"]
    oda = (0.4, b - 3.7, 5.1, b - 0.4)         # elektrik odası (x0,y0,x1,y1)

    c.t(14, 17.5, "ÇİZİM 6 — AYDINLATMA VE PRİZ YERLEŞİM PLANI", size=4.0, bold=True)
    c.t(14, 22.8, "Kat planı 1:100 — priz (P), iç LED lamba (L), avize (AVZ), dış projektör (PR); renk = bağlı olduğu RCBO hattı (PZ, AV, LD, PJ)", size=2.4, color="#444")

    # ---- bina
    c.rect(*pm(0, 0), a * S, b * S, sw=0.8, fill="#fafafa")
    c.rect(*pm(0.4, 0.4), (a - 0.8) * S, (b - 0.8) * S, sw=0.25, stroke="#666")
    cx_, cy_ = pm(kx, ky)
    c.circle(cx_, cy_, kr * S, sw=0.3, stroke="#7a8aa0", fill="#eef2f8", dash="3,1.5")
    c.t(cx_ - 13.0, cy_ + kr * S - 2.0, f"KUBBE Ø{tr(T['kubbe_cap_m'], 0)} m", size=1.9, anchor="middle", color="#557")
    mxp, myp = pm(mx, my)
    c.rect(mxp - mt * S / 2, myp - mt * S / 2, mt * S, mt * S, sw=0.5, fill="#e9eef5")
    c.circle(mxp, myp, mt * S * 0.33, sw=0.3, fill="#fff")
    c.t(mxp - 6.0, myp + mt * S / 2 + 3.2, "MİNARE", size=1.9, anchor="middle", color="#557")
    c.rect(*pm(oda[0], oda[1]), (oda[2] - oda[0]) * S, (oda[3] - oda[1]) * S, sw=0.3, stroke="#666", dash="2,1", fill="#f1f1f1")
    c.t(*pm(oda[0] + 1.6, oda[1] + 0.9), "ELEKTRİK ODASI", size=1.7, color="#555", halo=True)
    xa, ya = pm(px, py)
    c.rect(xa - 4.0, ya - 2.0, 8.0, 4.0, sw=0.5, fill="#d0d0d0")
    c.t(xa + 5.2, ya + 0.6, "PN1 ana pano", size=2.0, bold=True, poz="PN1", halo=True)

    # ---- kablo tavası KT1 (iç çevre + güney-merkez omurga)
    inset = 1.25
    tray = [pm(inset, inset), pm(a - inset, inset), pm(a - inset, b - inset), pm(inset, b - inset)]
    c.poly(tray, closed=True, sw=1.6, stroke="#c9ccd1")
    c.poly(tray, closed=True, sw=0.25, stroke="#7d8590", dash="2.2,1.2")
    c.line(*pm(a / 2, b - inset), *pm(kx, ky + 0.6), w=1.6, c="#c9ccd1")
    c.line(*pm(a / 2, b - inset), *pm(kx, ky + 0.6), w=0.25, c="#7d8590", dash="2.2,1.2")
    c.line(xa, ya - 2.0, *pm(px, b - inset), w=1.6, c="#c9ccd1")
    c.t(*pm(a / 2 + 0.5, b - inset + 0.9), "KT1 kablo tavası (ana kablo yolu)", size=1.8, bold=True, color="#555", poz="KT1", halo=True)

    # ---- SES KOLLARI güzergâhı (SS1-SSn): PN1'den kablo tavası boyunca; sağ kol doğu duvarı, sol kol batı + kuzey duvarı, minare kolu kuzey duvarından minareye
    ses_yasak = []                                     # etiket/sembol alanları (m): LED lambaları buraya konmaz
    ses_h = [h for h in M["hatlar"] if h["grup"] == "SES"]

    def ses_rota(h, sira, n):
        ad = h.get("kol_ad", "")
        ofs = (sira - (n - 1) / 2.0) * 0.34
        if ad.startswith("sa"):                            # sağ kol: güney duvar + doğu duvar
            return [(px, b - 0.55), (px, b - inset), (a - inset, b - inset), (a - inset, 9.0)], "son_k"
        if ad.startswith("min"):                           # minare kolu: batı + kuzey duvar, sonra minareye
            return [(px + 0.17, b - 0.55), (px + 0.17, inset + 0.17), (mx - mt / 2, inset + 0.17), (mx, my)], "son_m"
        return [(px - 0.17, b - 0.55), (px - 0.17, inset - 0.17), (7.0, inset - 0.17)], "son_b"   # sol kol: batı + kuzey duvar yarısı
    ses_cizim = []
    for i_, h in enumerate(ses_h):
        rota, tip_ = ses_rota(h, i_, len(ses_h))
        ses_cizim.append((h, [pm(*q) for q in rota], tip_))
        xe, ye = rota[-1]
        if tip_ == "son_k":
            ses_yasak.append((xe - 1.2, ye - 0.5, xe + 0.1, ye + 0.9))
        elif tip_ == "son_b":
            ses_yasak.append((xe - 0.7, ye, xe + 0.7, ye + 1.4))
        else:
            ses_yasak.append((xe - 3.5, ye + 0.1, xe - 1.2, ye + 0.9))
    for h, pts_, tip_ in ses_cizim:
        col_ = renk[h["id"]]
        c.poly(pts_, sw=0.95, stroke=col_)
        xe_, ye_ = pts_[-1]
        x0_, y0_ = pts_[-2]
        ang_ = math.atan2(ye_ - y0_, xe_ - x0_)
        c.poly([(xe_, ye_), (xe_ - 2.2 * math.cos(ang_) + 1.0 * math.sin(ang_), ye_ - 2.2 * math.sin(ang_) - 1.0 * math.cos(ang_)),
                (xe_ - 2.2 * math.cos(ang_) - 1.0 * math.sin(ang_), ye_ - 2.2 * math.sin(ang_) + 1.0 * math.cos(ang_))], closed=True, sw=0, fill=col_)
        etik = f"{h['id']} · {h['kablo_no']}"
        if tip_ == "son_k":
            c.hoparlor(xe_ - 7.5, ye_)
            c.t(xe_ - 2.0, ye_ + 6.2, etik, size=1.6, bold=True, anchor="end", color=col_, poz=[h["id"], h["kablo_no"]], halo=True)
        elif tip_ == "son_b":
            c.hoparlor(xe_ - 1.0, ye_ + 3.8)
            c.t(xe_ + 0.5, ye_ + 11.0, etik, size=1.6, bold=True, anchor="middle", color=col_, poz=[h["id"], h["kablo_no"]], halo=True)
        else:
            c.hoparlor(xe_ - 1.2, ye_ - 0.2)
            c.t(xe_ - 12.0, ye_ + 6.2, etik, size=1.6, bold=True, anchor="middle", color=col_, poz=[h["id"], h["kablo_no"]], halo=True)

    # ---- sembol yardımcıları
    def prizs(x, y, nx, ny, col):
        """Yarım daire priz, düz kenar duvara yaslı; (nx,ny) duvardan içe doğru birim vektör."""
        r = 1.3
        ang = math.degrees(math.atan2(ny, nx))
        a0, a1 = ang - 90, ang + 90
        x0_, y0_ = x + r * math.cos(math.radians(a0)), y + r * math.sin(math.radians(a0))
        x1_, y1_ = x + r * math.cos(math.radians(a1)), y + r * math.sin(math.radians(a1))
        c.path(f"M {x0_:.2f} {y0_:.2f} A {r} {r} 0 0 1 {x1_:.2f} {y1_:.2f} Z", sw=0.3, stroke=col, fill=col)
        c.line(x, y, x - nx * 1.0, y - ny * 1.0, w=0.35, c=col)

    def lamps(x, y, col):
        c.circle(x, y, 1.15, sw=0.35, stroke=col, fill="#fff")
        d = 0.8
        c.line(x - d, y - d, x + d, y + d, w=0.3, c=col)
        c.line(x - d, y + d, x + d, y - d, w=0.3, c=col)

    # ---- PRİZLER P1-Pn: iç duvar boyunca (0,9 m içeride), saat yönünde GB'den başlar
    soc = [e for e in M["elemanlar"] if e["tip"] == "priz"]
    ins = 0.9
    yol_ = [(5.6, b - ins), (a - ins, b - ins), (a - ins, ins), (ins, ins), (ins, oda[1] - 0.2)]
    seg = [(yol_[i], yol_[i + 1]) for i in range(len(yol_) - 1)]
    uz = [math.hypot(q[0] - p_[0], q[1] - p_[1]) for p_, q in seg]
    toplam = sum(uz)

    def yol_nokta(s_):
        for (p_, q), u in zip(seg, uz):
            if s_ <= u + 1e-9:
                t_ = s_ / u
                return (p_[0] + (q[0] - p_[0]) * t_, p_[1] + (q[1] - p_[1]) * t_, (q[0] - p_[0]) / u, (q[1] - p_[1]) / u)
            s_ -= u
        return (yol_[-1][0], yol_[-1][1], 0, -1)
    # minare alanı (genişletilmiş) dışındaki yol uzunluklarına göre eşit dağıt
    nokta_sayisi = int(toplam / 0.05)
    izinli = []
    for i_ in range(nokta_sayisi + 1):
        s_ = i_ * 0.05
        x_, y_, _, _ = yol_nokta(s_)
        if not (x_ > mx - mt / 2 - 0.9 and y_ < my + mt / 2 + 0.9):
            izinli.append(s_)
    adim = len(izinli) / max(len(soc), 1)
    for k, e in enumerate(soc):
        s_ = izinli[min(int((k + 0.5) * adim), len(izinli) - 1)]
        x_, y_, dx_, dy_ = yol_nokta(s_)
        nxi, nyi = (-dy_, dx_)                      # saat yönünde ilerlerken içeri = sağ normal
        xp_, yp_ = pm(x_ + (-nxi) * 0.55, y_ + (-nyi) * 0.55)
        prizs(xp_, yp_, nxi, nyi, renk[e["hat"]])
        lx_, ly_ = xp_ + nxi * 4.2, yp_ + nyi * 4.2
        c.t(lx_, ly_ + 0.6, e["poz"], size=1.6, bold=True, anchor="middle", color=renk[e["hat"]], poz=e["poz"], hat=e["hat"])

    # ---- AVİZE
    for e in [e for e in M["elemanlar"] if e["tip"] == "avize"]:
        xv, yv = pm(kx, ky)
        c.avize(xv, yv, r=4.2, renk=renk[e["hat"]])
        c.t(xv, yv + 10.6, e["poz"], size=2.0, bold=True, anchor="middle", color=renk[e["hat"]], poz=e["poz"], hat=e["hat"], halo=True)
        c.t(xv, yv + 13.2, f"{tr(e['w'], 0)} W", size=1.6, anchor="middle", color="#444", halo=True)

    # ---- İÇ LED LAMBALAR L1-L50: halkalar halinde, açıya göre numaralı
    leds = [e for e in M["elemanlar"] if e["tip"] == "led"]
    n_led = len(leds)
    halka_r = [2.6, 4.8, 6.4, 8.0]
    oranlar = [0.12, 0.22, 0.29, 0.37]
    sayilar = [max(1, round(n_led * o)) for o in oranlar]
    sayilar[-1] += n_led - sum(sayilar)
    pos = []
    for r_, n_, off_ in zip(halka_r, sayilar, (0.0, 10.0, 0.0, 6.0)):
        for i in range(n_):
            ang = math.radians(off_ + i * 360.0 / n_)
            pos.append((kx + r_ * math.sin(ang), ky - r_ * math.cos(ang), (math.degrees(ang) % 360)))
    def cakisma(x_, y_):
        for bx0, by0, bx1, by1 in ses_yasak:
            if bx0 - 0.45 < x_ < bx1 + 0.45 and by0 - 0.45 < y_ < by1 + 0.45:
                return True
        if x_ > mx - mt / 2 - 0.8 and y_ < my + mt / 2 + 0.8:
            return True
        if oda[0] - 0.8 < x_ < oda[2] + 0.8 and oda[1] - 0.8 < y_ < oda[3] + 0.8:
            return True
        return False
    # halka konumları sırayla kabul edilir; yasak bölgeye (minare, elektrik odası, duvar payı) düşen veya komşusuna 1,9 m'den
    # yakın olanlar, ince bir ızgaradan kendilerine en yakın uygun noktaya taşınır
    def gecerli(x_, y_, kabul):
        if cakisma(x_, y_) or not (2.2 <= x_ <= a - 2.2 and 2.2 <= y_ <= b - 2.2):
            return False
        return all(math.hypot(x_ - q[0], y_ - q[1]) >= 1.9 for q in kabul)
    havuz = [(2.2 + 0.25 * i_, 2.2 + 0.25 * j_) for i_ in range(int((a - 4.4) / 0.25) + 1) for j_ in range(int((b - 4.4) / 0.25) + 1)]
    yeni_pos = []
    for (x_, y_, ang_) in pos:
        if not gecerli(x_, y_, yeni_pos):
            adaylar = [q for q in havuz if gecerli(q[0], q[1], yeni_pos) and math.hypot(q[0] - kx, q[1] - ky) > 1.8]
            if not adaylar:
                uyar("Çizim 6: iç LED lambaları için yeterli yerleşim alanı bulunamadı.")
                adaylar = [(x_, y_)]
            x_, y_ = min(adaylar, key=lambda q: math.hypot(q[0] - x_, q[1] - y_))
            ang_ = math.degrees(math.atan2(x_ - kx, -(y_ - ky))) % 360
        yeni_pos.append((x_, y_, ang_ % 360))
    yeni_pos.sort(key=lambda t_: (round(t_[2] / 1.0), math.hypot(t_[0] - kx, t_[1] - ky)))
    for (x_, y_, ang_), e in zip(yeni_pos, leds):
        xp_, yp_ = pm(x_, y_)
        lamps(xp_, yp_, renk[e["hat"]])
        rad = math.hypot(x_ - kx, y_ - ky) or 1.0
        ux, uy = (x_ - kx) / rad, (y_ - ky) / rad
        c.t(xp_ + ux * 3.4, yp_ + uy * 3.0 + 0.6, e["poz"], size=1.45, bold=True, anchor="middle", color=renk[e["hat"]], poz=e["poz"], hat=e["hat"], halo=True)

    # ---- DIŞ PROJEKTÖRLER PR1-PRn + kablo yolu + JB
    prj = [e for e in M["elemanlar"] if e["tip"] == "projektor"]
    cev = 2 * (a + b)
    stp = cev / max(len(prj), 1)
    kenar = [(b, (-1, 0), 180.0), (a, (0, -1), 270.0), (b, (1, 0), 0.0), (a, (0, 1), 90.0)]

    def cevre(s_, off):
        s_ = s_ % cev
        # SB köşesinden başla: batı duvarı kuzeye
        for k_, (uz_, n_, ang_) in enumerate(kenar):
            if s_ <= uz_ + 1e-9:
                if k_ == 0:
                    bx_, by_ = 0.0, b - s_
                elif k_ == 1:
                    bx_, by_ = s_, 0.0
                elif k_ == 2:
                    bx_, by_ = a, s_
                else:
                    bx_, by_ = a - s_, b
                return bx_ + n_[0] * off, by_ + n_[1] * off, ang_, n_
            s_ -= uz_
    dis_loop = 1.0
    # dış hat kablo yolu (toprak altı / kablo kanalı) - genişletilmiş dikdörtgen
    loop = [pm(-dis_loop, b + dis_loop), pm(-dis_loop, -dis_loop), pm(a + dis_loop, -dis_loop), pm(a + dis_loop, b + dis_loop)]
    c.poly(loop + [loop[0]], sw=0.45, stroke="#555", dash="3,1.2")
    c.line(xa, ya + 2.0, xa, pm(0, b + dis_loop)[1], w=0.45, c="#555", dash="3,1.2")
    c.line(xa, pm(0, b + dis_loop)[1], loop[0][0], loop[0][1], w=0.45, c="#555", dash="3,1.2")
    hat_ilk = {}
    for k, e in enumerate(prj):
        s_ = (k + 0.5) * stp
        x_, y_, ang_, n_ = cevre(s_, 0.45)
        xp_, yp_ = pm(x_, y_)
        c.projektor(xp_, yp_, renk=renk[e["hat"]], yon=ang_, sc=0.9)
        # etiket: duvar boyunca bitişiğe
        if n_ == (-1, 0):
            lx_, ly_, anc = xp_, yp_ - 3.2, "middle"
        elif n_ == (1, 0):
            lx_, ly_, anc = xp_, yp_ + 4.8, "middle"
        elif n_ == (0, -1):
            lx_, ly_, anc = xp_ + 3.6, yp_ + 0.6, "start"
        else:
            lx_, ly_, anc = xp_ - 3.6, yp_ + 0.6, "end"
        c.t(lx_, ly_, e["poz"], size=1.6, bold=True, anchor=anc, color=renk[e["hat"]], poz=e["poz"], hat=e["hat"], halo=True)
        hat_ilk.setdefault(e["hat"], (s_, xp_, yp_))
    for i, (hid, (s_, xp_, yp_)) in enumerate(hat_ilk.items(), start=1):
        xj, yj, _, nj = cevre(s_ - 2.4, dis_loop)
        xjp, yjp = pm(xj, yj)
        c.rect(xjp - 1.4, yjp - 1.4, 2.8, 2.8, sw=0.4, stroke=renk[hid], fill="#fff")
        c.line(xjp, yjp, xp_, yp_, w=0.5, c=renk[hid])
        if nj[0] != 0:
            c.t(xjp, yjp - 2.4, f"JB{i}", size=1.6, bold=True, anchor="middle", poz=f"JB{i}", hat=hid, halo=True)
        else:
            c.t(xjp + 3.2, yjp + 0.6, f"JB{i}", size=1.6, bold=True, poz=f"JB{i}", hat=hid, halo=True)

    # ---- ölçekler, kuzey oku
    c.olcu(*pm(a, 0), *pm(a, b), f"{tr(b)} m", ofs=14.0, size=2.0, yatay=False)
    c.t(*pm(a / 2, 0.0), f"{tr(a)} m", size=1.8, anchor="middle", color="#555") if False else None
    for k in range(4):
        c.rect(120.0 + k * S * 2.5, 253.0, S * 2.5, 1.8, sw=0.25, fill="#fff" if k % 2 else "#222")
    c.t(120.0, 252.2, "0", size=1.6, anchor="middle")
    c.t(120.0 + S * 5, 252.2, "5", size=1.6, anchor="middle")
    c.t(120.0 + S * 10, 252.2, "10 m", size=1.6, anchor="middle")
    c.poly([(402.0, 15.0), (399.8, 23.0), (402.0, 21.0), (404.2, 23.0)], closed=True, sw=0.3, fill="#222")
    c.t(406.5, 21.5, "K", size=2.0, bold=True)

    # ================== sağ panel: lejant, hat/faz tablosu, özetler
    RX, RW = 248.0, 160.0
    ry = 20.0
    c.t(RX, ry, "LEJANT", size=2.6, bold=True)
    ry += 2.0
    items = [("prizs", "Priz girişi 16 A, 230 V, topraklı (P)"), ("lamp", "İç mekan LED armatür/lamba (L)"), ("avize", "Büyük avize (AVZ)"),
             ("proj", "LED projektör IP65, ok = ışık yönü (PR)"), ("jb", "IP65 bağlantı kutusu (JB)"), ("tray", "Kablo tavası KT1"), ("loop", "Dış hat kablo yolu"), ("ses", "Ses kolu güzergâhı (SS), → hoparlör grubu")]
    for k, (tp_, ad_) in enumerate(items):
        col_ = k % 2
        row_ = k // 2
        x_ = RX + col_ * 80.0
        y_ = ry + 5.0 + row_ * 6.2
        if tp_ == "prizs":
            prizs(x_ + 3, y_ - 0.6, 1, 0, "#444")
        elif tp_ == "lamp":
            lamps(x_ + 3, y_ - 0.6, "#444")
        elif tp_ == "avize":
            c.avize(x_ + 3, y_ - 0.6, r=1.8, renk="#444")
        elif tp_ == "proj":
            c.projektor(x_ + 1.5, y_ - 0.6, renk="#444", yon=0, sc=0.9)
        elif tp_ == "jb":
            c.rect(x_ + 1.8, y_ - 2.0, 2.8, 2.8, sw=0.4, fill="#fff")
        elif tp_ == "ses":
            c.line(x_, y_ - 0.6, x_ + 4, y_ - 0.6, w=0.95, c="#444")
            c.hoparlor(x_ + 4.2, y_ - 0.6)
        elif tp_ == "tray":
            c.line(x_, y_ - 0.6, x_ + 6, y_ - 0.6, w=1.6, c="#c9ccd1")
            c.line(x_, y_ - 0.6, x_ + 6, y_ - 0.6, w=0.25, c="#7d8590", dash="2.2,1.2")
        else:
            c.line(x_, y_ - 0.6, x_ + 6, y_ - 0.6, w=0.45, c="#555", dash="3,1.2")
        c.t(x_ + 8.5, y_, ad_, size=1.8, maxw=70)
    ry = ry + 5.0 + ((len(items) + 1) // 2) * 6.2 + 2.0

    c.t(RX, ry, "HAT / FAZ / RENK TABLOSU", size=2.6, bold=True)
    ry += 2.2
    kol_ = [("", 5), ("Hat (RCBO)", 17), ("Elemanlar", 30), ("Adet", 10), ("Güç (W)", 15), ("Faz", 10), ("Kablo", 50), ("Kor.", 19)]
    xcur = RX
    xs_ = []
    for ad_, w_ in kol_:
        xs_.append((xcur, w_))
        c.rect(xcur, ry, w_, 5.0, sw=0.25, fill="#e3e8ef")
        c.t(xcur + 1.0, ry + 3.4, ad_, size=1.7, bold=True)
        xcur += w_
    ry += 5.0
    for hid, col in renk.items():
        h = hat[hid]
        c.rect(xs_[0][0], ry, xs_[0][1], 5.4, sw=0.2)
        c.rect(xs_[0][0] + 1.2, ry + 1.2, 3.0, 3.0, sw=0, fill=col)
        elemanlar_ = h.get("elemanlar") or []
        vals = [None, hid, poz_aralik(elemanlar_) if elemanlar_ else f"hoparlör kolu ({h.get('kol_ad', '')})", f"{len(elemanlar_)}" if elemanlar_ else "-", f"{tr(h['p_kw'] * 1000.0, 0)}", h["faz"],
                f"{h['kablo_no']}: {h['tip']} {h['damar']} mm², {tr(h['L'], 0)} m" + (" ↑" if h.get("kesit_artirildi") else ""), (f"MCB {P['rcbo_defaults']['egri']}{h['rcbo_In']}" if h.get("koruma_tur") == "mcb" else f"{h['koruma_ad'].split(',')[0].replace('RCBO 1P+N ', '')}")]
        for j in range(1, len(kol_)):
            xx, ww = xs_[j]
            c.rect(xx, ry, ww, 5.4, sw=0.2)
            if j == 1:
                c.t(xx + 1.0, ry + 3.7, vals[j], size=1.8, bold=True, poz=hid, maxw=ww - 1.2)
            elif j == 2:
                if elemanlar_:
                    c.t(xx + 1.0, ry + 3.7, vals[j], size=1.7, poz=elemanlar_, hat=hid, maxw=ww - 1.2)
                else:
                    c.t(xx + 1.0, ry + 3.7, vals[j], size=1.7, maxw=ww - 1.2)
            else:
                c.t(xx + 1.0, ry + 3.7, vals[j], size=1.7, maxw=ww - 1.2)
        ry += 5.4
    ry += 4.0

    c.t(RX, ry, "FAZ YÜKÜ (talep gücü, açgözlü atama sonucu)", size=2.6, bold=True)
    ry += 2.0
    mxv = max(M["faz_kw"].values())
    for k, ph in enumerate(("L1", "L2", "L3")):
        yy = ry + 4.0 + k * 5.2
        c.t(RX, yy, f"Faz {ph}", size=1.9, bold=True)
        c.rect(RX + 14, yy - 3.0, 90.0 * M["faz_kw"][ph] / mxv, 3.8, sw=0.2, fill=RENK[ph] if ph != "L3" else "#a8a8a8")
        c.t(RX + 16 + 90.0 * M["faz_kw"][ph] / mxv, yy, f"{tr(M['faz_kw'][ph], 2)} kW", size=1.8)
    ry += 4.0 + 3 * 5.2
    c.t(RX, ry, f"Faz dengesizliği %{tr(M['dengesizlik'])} (hedef ≤ %10); Ib = {tr(M['Ib'], 1)} A; Q0 Ir = {M['Ir']} A", size=1.9, bold=True, color="#1c6a31")
    ry += 6.0

    c.t(RX, ry, "ENVANTER ÖZETİ", size=2.6, bold=True)
    ry += 2.0
    gruplar_ = [("priz", "Priz girişi", "P"), ("avize", "Büyük avize", "AVZ"), ("led", "İç mekan LED lamba", "L"), ("projektor", "Dış projektör", "PR")]
    tot_w = 0.0
    for k, (tp_, ad_, pre) in enumerate(gruplar_):
        lst = [e for e in M["elemanlar"] if e["tip"] == tp_]
        if not lst:
            continue
        wsum = sum(e["w"] for e in lst)
        tot_w += wsum
        yy = ry + 4.0 + k * 4.4
        c.t(RX, yy, f"{ad_} ({poz_aralik([e['poz'] for e in lst])})", size=1.8)
        c.t(RX + 100, yy, f"{len(lst)} adet x {tr(lst[0]['w'], 0)} W", size=1.8, anchor="end")
        c.t(RX + 150, yy, f"{tr(wsum, 0)} W", size=1.8, bold=True, anchor="end")
    yy = ry + 4.0 + len(gruplar_) * 4.4
    c.line(RX, yy - 3.0, RX + 150, yy - 3.0, w=0.2)
    c.t(RX, yy, "Toplam aydınlatma + priz (kurulu)", size=1.9, bold=True)
    c.t(RX + 150, yy, f"{tr(tot_w, 0)} W", size=1.9, bold=True, anchor="end")

    ry = yy + 9.0
    if any(h["grup"] == "PROJ" for h in M["hatlar"]):
        c.t(RX, ry, "DIŞ AYDINLATMA KUMANDA MANTIĞI (PJ1-PJ3)", size=2.6, bold=True)
        ry += 2.0
        for ln in ("Q3 (MCB 4P C25) → KM1 (kontaktör 3P 25 A) → PJ1-PJ3 RCBO'ları; SPD2 Tip 2 koruması (Q3 sonrası).",
                   "SA3 (kapıda) MAN: KM1 sürekli enerjili; 0: devre dışı; OTO: FT1 (alacakaranlık) VE ZR1 (astronomik zaman) kontakları kapalıyken KM1 enerjilenir.",
                   "H4 (kapıda) dış aydınlatma açık durum lambası. Kumanda devresi Çizim 1'de gösterilmiştir."):
            ry = c.tw_(RX, ry + 3.2, ln, 150.0, size=1.8, lh=1.2) - 1.0

    # ================== alt bölge notları
    ny = 253.5
    c.t(14.0, ny, "NOTLAR VE VARSAYIMLAR", size=2.6, bold=True)
    A_ = P["aydinlatma_priz"]
    notlar = [
        f"1. Priz başına yük {tr(A_['priz']['birim_w'], 0)} W, kd={tr(A_['priz']['kd'], 1)}; projektör gücü {tr(PJ['PROJEKTOR_GUCU_W'], 0)} W LED (varsayım, parametre PROJEKTOR_GUCU_W); LED lamba {tr(A_['led_ic']['birim_w'], 0)} W.",
        f"2. Avize {tr(A_['avize']['w'], 0)} W, cosφ={tr(A_['avize']['cos'], 2)} (varsayım) → {tr(hat['AV1']['I'], 1)} A (cosφ=1 için 8,7 A); diğer yüklerde cosφ={tr(P['sistem']['cos_varsayilan'], 2)}.",
        "3. İç LED hatları C tipi RCBO (sürücü kalkış akımı); dış hatlar KM1 kontaktörü ile SA3 (MAN-0-OTO), FT1 fotosel ve ZR1 astronomik röle üzerinden kumanda edilir.",
        "4. Ses kolları (SS1-SS3) PN1'den kablo tavası boyunca: sağ kol doğu duvarı, sol kol batı ve kuzey duvarı, minare kolu kuzey duvarından minareye; hoparlör sayısı/konumu şematiktir.",
        "5. Konumlar şematiktir; uygulama projesinde mimari/aydınlatma hesabı ile netleştirilir. Pozlar Çizim 1, 2 ve tablolarla aynıdır.",
    ]
    yy = ny + 3.3
    for n_ in notlar:
        yy = c.tw_(14.0, yy, n_, 100.0, size=1.75, lh=1.18)

    c.cerceve(P, "AYDINLATMA VE PRİZ YERLEŞİM PLANI", 4, "1:100")
    return c


# =====================================================================
#  EXCEL : 04_hesap_tablolari.xlsx   (hesaplar Excel formülüdür)
# =====================================================================
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL

_THIN = Side(style="thin", color="999999")
_BRD = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HDR = PatternFill("solid", fgColor="D9E2F3")
_PAR = PatternFill("solid", fgColor="FFF2CC")
_OK = PatternFill("solid", fgColor="E2EFDA")
_TITLE = Font(name="Liberation Sans", size=14, bold=True)
_B = Font(name="Liberation Sans", size=10, bold=True)
_N = Font(name="Liberation Sans", size=10)


def _w(ws, r, c, v, bold=False, fill=None, fmt=None, wrap=False, align=None, border=True):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = _B if bold else _N
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    cell.alignment = Alignment(wrap_text=wrap, vertical="center", horizontal=align)
    if border:
        cell.border = _BRD
    return cell


def _baslik(ws, baslik, alt, ncol):
    ws.cell(row=1, column=1, value=baslik).font = _TITLE
    ws.cell(row=2, column=1, value=alt).font = Font(name="Liberation Sans", size=9, italic=True, color="555555")
    ws.sheet_view.showGridLines = False


def _genislik(ws, genisler):
    for i, g in enumerate(genisler, start=1):
        ws.column_dimensions[CL(i)].width = g


def _sayfa(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def excel_yaz(P, M, PN, T, KAYIT, yol):
    wb = Workbook()
    S, K, C = P["sistem"], P["kablo"], P["cihazlar"]
    pr = P["proje"]
    alt = f"{pr['ad']} | {pr['sistem_yazisi']} | Tarih: {pr['tarih']} | Rev: {pr['revizyon']} | Sarı hücreler PARAMETRELER sözlüğünden gelen girdilerdir; diğer hesaplar formüldür."
    hatlar = M["hatlar"]
    ref = {}

    # ===================================================== 1) YÜK DENGESİ
    ws = wb.active
    ws.title = "Yük Dengesi"
    _baslik(ws, "YÜK DENGESİ VE FAZ DENGESİZLİĞİ", alt, 13)
    _w(ws, 4, 1, "U_LN (V)", True); _w(ws, 4, 2, S["U_LN"], fill=_PAR)
    _w(ws, 5, 1, "U_LL (V)", True); _w(ws, 5, 2, S["U_LL"], fill=_PAR)
    _w(ws, 6, 1, "Q0 Ir (A)", True); _w(ws, 6, 2, C["Q0"]["Ir_A"], fill=_PAR)
    _w(ws, 7, 1, "Gelecek payı (%)", True); _w(ws, 7, 2, S["gelecek_payi_yuzde"], fill=_PAR)
    _w(ws, 4, 4, "Faz dengesizliği = (Pmax − Pmin) / Port × 100  (hedef ≤ %10)", border=False)
    _w(ws, 5, 4, "Hesap akımı Ib = en yüklü faz akımı;  kd = eşzamanlılık katsayısı;  Talep gücü = Kurulu × kd", border=False)
    hdr = ["Poz", "Hat adı", "Kurulu güç (kW)", "Faz", "cosφ", "Eşzamanlılık kd", "Talep gücü (kW)", "L1", "L2", "L3", "L1", "L2", "L3"]
    _w(ws, 8, 8, "Talep gücü dağılımı (kW)", True, _HDR, align="center"); ws.merge_cells(start_row=8, start_column=8, end_row=8, end_column=10)
    _w(ws, 8, 11, "Faz akımları (A)", True, _HDR, align="center"); ws.merge_cells(start_row=8, start_column=11, end_row=8, end_column=13)
    for j, h in enumerate(hdr, start=1):
        _w(ws, 9, j, h, True, _HDR, wrap=True, align="center")
    r0 = 10
    satirlar = [dict(h) for h in hatlar]
    for yd in P["yedek_rcbo"]:
        satirlar.append({"id": yd, "ad": "Yedek RCBO yuvası (boş)", "p_kw": 0.0, "faz": "-", "cos": 1.0, "kd": 1.0})
    for i, h in enumerate(satirlar):
        r = r0 + i
        _w(ws, r, 1, h["id"], True)
        _w(ws, r, 2, h["ad"])
        _w(ws, r, 3, h["p_kw"], fill=_PAR, fmt="0.00")
        _w(ws, r, 4, h["faz"] if h["faz"] != "3F" else "3F", fill=_PAR, align="center")
        _w(ws, r, 5, h["cos"], fill=_PAR, fmt="0.00")
        _w(ws, r, 6, h["kd"], fill=_PAR, fmt="0.00")
        _w(ws, r, 7, f"=C{r}*F{r}", fmt="0.00")
        for k, col in enumerate("HIJ"):
            _w(ws, r, 8 + k, f'=IF($D{r}="3F",$G{r}/3,IF($D{r}={col}$9,$G{r},0))', fmt="0.00")
        for k, col in enumerate("KLM"):
            ph = "HIJ"[k]
            src = f"{col}$9"
            _w(ws, r, 11 + k, f'=IF($D{r}="3F",$G{r}*1000/(3*$B$4*$E{r}),IF($D{r}={src},$G{r}*1000/($B$4*$E{r}),0))', fmt="0.0")
    rl = r0 + len(satirlar) - 1
    rt = rl + 1
    _w(ws, rt, 1, "TOPLAM", True, _HDR); _w(ws, rt, 2, "", fill=_HDR)
    _w(ws, rt, 3, f"=SUM(C{r0}:C{rl})", True, _HDR, "0.00")
    _w(ws, rt, 4, "", fill=_HDR); _w(ws, rt, 5, "", fill=_HDR); _w(ws, rt, 6, "", fill=_HDR)
    for j in range(7, 14):
        col = CL(j)
        _w(ws, rt, j, f"=SUM({col}{r0}:{col}{rl})", True, _HDR, "0.00" if j < 11 else "0.0")
    rs = rt + 2
    _w(ws, rs, 1, "ÖZET", True); ws.merge_cells(start_row=rs, start_column=1, end_row=rs, end_column=3)
    ozet = [
        ("Kurulu güç (kW)", f"=C{rt}", "0.00"),
        ("Talep gücü (kW)", f"=G{rt}", "0.00"),
        ("Faz dengesizliği (%)", f"=(MAX(H{rt}:J{rt})-MIN(H{rt}:J{rt}))/AVERAGE(H{rt}:J{rt})*100", "0.00"),
        ("Dengesizlik hedefi ≤ %10", f'=IF(B{rs + 3}<=10,"UYGUN","UYGUN DEĞİL")', None),
        ("Hesap akımı Ib (A) - en yüklü faz", f"=MAX(K{rt}:M{rt})", "0.0"),
        ("Ortalama faz akımı (A)", f"=AVERAGE(K{rt}:M{rt})", "0.0"),
        ("cosφ (ortalama)", f"=G{rt}*1000/(SQRT(3)*B5*AVERAGE(K{rt}:M{rt}))", "0.000"),
        ("Talep görünür gücü (kVA)", f"=SQRT(3)*B5*AVERAGE(K{rt}:M{rt})/1000", "0.0"),
        ("Ib ≤ Q0 Ir kontrolü", f'=IF(B{rs + 5}<=B6,"UYGUN","UYGUN DEĞİL")', None),
        ("Trafo yüklenme oranı (%)", f"=B{rs + 8}/{S['trafo_kva']}*100", "0.0"),
        ("Ib × (1 + gelecek payı) (A)", f"=B{rs + 5}*(1+B7/100)", "0.0"),
        ("Gelecek payı: Ib×(1+pay) ≤ Q0 Ir", f'=IF(B{rs + 11}<=B6,"UYGUN","UYGUN DEĞİL")', None),
        ("Q0 kalan yedek kapasite (%)", f"=(B6-B{rs + 5})/B6*100", "0.0"),
    ]
    for i, (a_, f_, fm) in enumerate(ozet, start=1):
        _w(ws, rs + i, 1, a_, True); ws.merge_cells(start_row=rs + i, start_column=1, end_row=rs + i, end_column=1)
        _w(ws, rs + i, 2, f_, True, _OK, fm)
    ref["YB_kurulu"] = f"'Yük Dengesi'!B{rs + 1}"
    ref["YB_talep"] = f"'Yük Dengesi'!B{rs + 2}"
    ref["YB_dengesizlik"] = f"'Yük Dengesi'!B{rs + 3}"
    ref["YB_Ib"] = f"'Yük Dengesi'!B{rs + 5}"
    ref["YB_cos"] = f"'Yük Dengesi'!B{rs + 7}"
    # açgözlü (greedy) faz atama adımları
    ra = rs + len(ozet) + 3
    _w(ws, ra - 1, 1, "TEK FAZ YÜKLERİN AÇGÖZLÜ (GREEDY) FAZ ATAMASI", True)
    ws.cell(row=ra, column=1).value = None
    ba = ["Sıra", "Hat", "Hat adı", "Talep gücü (kW)", "Atanan faz", "Gerekçe", "", "L1 toplamı (kW)", "L2 toplamı (kW)", "L3 toplamı (kW)"]
    for j, h_ in enumerate(ba, start=1):
        _w(ws, ra, j, h_, True, _HDR, wrap=True, align="center")
    for i_, ad_ in enumerate(M["faz_atama"], start=1):
        rr = ra + i_
        _w(ws, rr, 1, ad_["sira"], align="center"); _w(ws, rr, 2, ad_["id"], True); _w(ws, rr, 3, ad_["ad"])
        _w(ws, rr, 4, ad_["talep_kw"], fmt="0.000"); _w(ws, rr, 5, ad_["faz"], True, _OK, align="center")
        _w(ws, rr, 6, ad_["neden"]); ws.merge_cells(start_row=rr, start_column=6, end_row=rr, end_column=7)
        for j_, ph_ in enumerate(("L1", "L2", "L3")):
            _w(ws, rr, 8 + j_, ad_[ph_], fmt="0.00")
    ws.cell(row=ra + len(M["faz_atama"]) + 1, column=1,
            value="Algoritma: sabit faz verilenler önce atanır; kalanlar talep gücüne göre büyükten küçüğe sıralanır, her biri en az yüklü faza (aynı gruptaki hatlar farklı fazlara) atanır. 3 fazlı klima yükleri fazlara eşit dağılır.").font = Font(name="Liberation Sans", size=9, italic=True)
    _genislik(ws, [30, 34, 14, 8, 8, 14, 14, 10, 10, 10, 10, 10, 10])
    ws.freeze_panes = "C10"
    _sayfa(ws)

    # ===================================================== 2) KABLO ÇİZELGESİ
    ws = wb.create_sheet("Kablo Çizelgesi")
    _baslik(ws, "KABLO ÇİZELGESİ (akım, kesit, koruma, gerilim düşümü)", alt, 21)
    parm = [("κ iletkenlik (m/Ω·mm²)", K["kappa"]), ("X reaktans (mΩ/m)", K["X_mohm_m"]), ("U_LL (V)", S["U_LL"]), ("U_LN (V)", S["U_LN"]),
            ("Azaltma k (pano-alıcı)", K["azaltma_k"]), ("Azaltma k (gömülü)", K["azaltma_k_gomulu"]),
            ("ΔU limiti güç hatları (%)", K["limit_guc_yuzde"]), ("ΔU limiti aydınlatma (%)", K["limit_aydinlatma_yuzde"]),
            ("Besleme (W0) ΔU hedefi (%) (varsayım)", 1.0)]
    for i, (a_, v) in enumerate(parm):
        _w(ws, 3 + i, 1, a_, True); ws.merge_cells(start_row=3 + i, start_column=1, end_row=3 + i, end_column=3)
        _w(ws, 3 + i, 4, v, fill=_PAR)
    # Iz tablosu (W:Z)
    ws.cell(row=2, column=23, value="Akım taşıma kapasitesi Iz (A) - IEC 60364-5-52 Ek B, PVC, Cu (yaklaşık)").font = _B
    for j, h in enumerate(["Kesit (mm²)", "C-2", "C-3", "D2-3"], start=23):
        _w(ws, 3, j, h, True, _HDR, align="center")
    kes_list = sorted(K["iz_C_2yuklu"].keys())
    for i, kk in enumerate(kes_list):
        _w(ws, 4 + i, 23, kk, fmt="0.0")
        _w(ws, 4 + i, 24, K["iz_C_2yuklu"][kk], fill=_PAR)
        _w(ws, 4 + i, 25, K["iz_C_3yuklu"][kk], fill=_PAR)
        _w(ws, 4 + i, 26, K["iz_D2_3yuklu"][kk], fill=_PAR)
    iz_son = 4 + len(kes_list) - 1
    # tablo
    kolonlar = ["Kablo no", "Koruma poz", "Hat adı", "Sınıf", "Güç (kW)", "Faz tipi", "cosφ", "Akım Ib (A)", "Kablo tipi", "Kesit (mm²)",
                "Uzunluk (m)", "Döşeme", "Koruma cihazı", "In (A)", "Iz tablo (A)", "Iz düzeltilmiş (A)", "ΔU (V)", "ΔU hat (%)",
                "ΔU besleme (%)", "ΔU toplam (%)", "ΔU limit (%)", "Uygunluk"]
    hr = 14
    for j, h in enumerate(kolonlar, start=1):
        _w(ws, hr, j, h, True, _HDR, wrap=True, align="center")
    _w(ws, hr, 27, "ΔU eşdeğer boy katsayısı", True, _HDR, wrap=True, align="center")
    ws.column_dimensions["AA"].width = 14
    ws.row_dimensions[hr].height = 30
    # kolon harfleri
    L = {n: CL(i + 1) for i, n in enumerate(kolonlar)}
    r = hr + 1
    b = M["besleme"]
    # W0 satırı
    _w(ws, r, 1, b["kablo_no"], True); _w(ws, r, 2, "Q0", True); _w(ws, r, 3, "Trafo - Q0 besleme (TR1 → PN1)")
    _w(ws, r, 4, "besleme"); _w(ws, r, 5, f"={ref['YB_talep']}", fmt="0.00"); _w(ws, r, 6, "3F", align="center")
    _w(ws, r, 7, f"={ref['YB_cos']}", fmt="0.000"); _w(ws, r, 8, f"={ref['YB_Ib']}", fmt="0.0")
    _w(ws, r, 9, f"{b['tip']} {b['damar']} mm² Cu"); _w(ws, r, 10, b["kesit"], fill=_PAR, fmt="0.0")
    _w(ws, r, 11, b["L"], fill=_PAR, fmt="0.0"); _w(ws, r, 12, "D2-3", align="center"); _w(ws, r, 27, 1.0, fill=_PAR, fmt="0.00")
    _w(ws, r, 13, f"MCCB 4P Ir={C['Q0']['Ir_A']} A"); _w(ws, r, 14, C["Q0"]["Ir_A"], fill=_PAR)
    w0 = r
    first_row = r
    for h in hatlar:
        r += 1
        f3 = h["faz3"]
        _w(ws, r, 1, h["kablo_no"], True); _w(ws, r, 2, h["id"], True); _w(ws, r, 3, h["ad"])
        _w(ws, r, 4, "aydınlatma" if h["limit"] == "aydinlatma" else "güç", align="center")
        _w(ws, r, 5, h["p_kw"], fill=_PAR, fmt="0.00"); _w(ws, r, 6, "3F" if f3 else "1F", fill=_PAR, align="center")
        _w(ws, r, 7, h["cos"], fill=_PAR, fmt="0.00")
        _w(ws, r, 8, f'=IF(F{r}="3F",E{r}*1000/(SQRT(3)*$D$5*G{r}),E{r}*1000/($D$6*G{r}))', fmt="0.00")
        _w(ws, r, 9, f"{h['tip']} {h['damar']} mm² Cu"); _w(ws, r, 10, h["kesit"], fill=_PAR, fmt="0.0")
        _w(ws, r, 11, h["L"], fill=_PAR, fmt="0.0"); _w(ws, r, 12, "C-3" if f3 else "C-2", fill=_PAR, align="center")
        _w(ws, r, 27, h.get("du_yayilim", 1.0), fill=_PAR, fmt="0.00")
        _w(ws, r, 13, h["koruma_ad"].replace("\n", " ")); _w(ws, r, 14, h["In_koruma"], fill=_PAR)
    last_row = r
    for rr in range(first_row, last_row + 1):
        _w(ws, rr, 15, f"=INDEX($X$4:$Z${iz_son},MATCH(J{rr},$W$4:$W${iz_son},0),MATCH(L{rr},$X$3:$Z$3,0))")
        _w(ws, rr, 16, f"=O{rr}*IF(L{rr}=\"D2-3\",$D$8,$D$7)", fmt="0.0")
        _w(ws, rr, 17, f'=IF(F{rr}="3F",SQRT(3),2)*H{rr}*K{rr}*AA{rr}*(G{rr}/($D$3*J{rr})+($D$4/1000)*SQRT(1-G{rr}^2))', fmt="0.00")
        _w(ws, rr, 18, f'=Q{rr}/IF(F{rr}="3F",$D$5,$D$6)*100', fmt="0.00")
        if rr == w0:
            _w(ws, rr, 19, 0, fmt="0.00")
            _w(ws, rr, 20, f"=R{rr}+S{rr}", fmt="0.00")
            _w(ws, rr, 21, "=$D$11", fmt="0.00")
        else:
            _w(ws, rr, 19, f"=$R${w0}", fmt="0.00")
            _w(ws, rr, 20, f"=R{rr}+S{rr}", fmt="0.00")
            _w(ws, rr, 21, f'=IF(D{rr}="aydınlatma",$D$10,$D$9)', fmt="0.00")
        _w(ws, rr, 22, f'=IF(AND(H{rr}<=N{rr},N{rr}<=P{rr},T{rr}<=U{rr}),"UYGUN","UYGUN DEĞİL")', True, _OK, align="center")
    ws.cell(row=last_row + 2, column=1, value="Kontroller: Ib ≤ In ≤ Iz(düzeltilmiş) ve toplam ΔU = ΔU hat + ΔU besleme ≤ limit. "
            "ΔU = (√3 veya 2)·I·L·(cosφ/(κ·S) + X·sinφ).").font = Font(name="Liberation Sans", size=9, italic=True)
    ws.cell(row=last_row + 3, column=1, value="Yolverme akımı (klima kompresörleri) ΔU hesabına dahil değildir; üretici yolverme akımı ile ayrıca kontrol edilmelidir.").font = Font(name="Liberation Sans", size=9, italic=True)
    ref["KC_w0"] = w0
    ref["KC_rows"] = {}
    for rr in range(first_row, last_row + 1):
        ref["KC_rows"][str(ws.cell(row=rr, column=2).value)] = rr
    ref["KC_first"], ref["KC_last"] = first_row, last_row
    _genislik(ws, [11, 11, 30, 11, 9, 8, 7, 11, 22, 9, 10, 9, 28, 8, 10, 11, 9, 10, 11, 11, 10, 14, 11, 8, 8, 8])
    ws.freeze_panes = ws.cell(row=hr + 1, column=4)
    _sayfa(ws)

    # ===================================================== 3) PANO MALZEME LİSTESİ
    ws = wb.create_sheet("Pano Malzeme Listesi")
    _baslik(ws, "PANO MALZEME LİSTESİ (PN1) VE SAHA EKİPMANI", alt, 6)
    for j, h in enumerate(["Poz", "Açıklama", "Adet", "Teknik değer", "Grup", "Çizim"], start=1):
        _w(ws, 4, j, h, True, _HDR, align="center")
    gr = satirlari_grupla(KAYIT)
    r = 5
    for kat_, kat_ad in (("pano", "PANO İÇİ / ÜZERİ MALZEME"), ("saha", "SAHA (PANO DIŞI) EKİPMAN VE KABLOLAR")):
        _w(ws, r, 1, kat_ad, True, _HDR); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        r += 1
        for g in gr:
            if g["kat"] != kat_:
                continue
            _w(ws, r, 1, g["poz"], True, wrap=True); _w(ws, r, 2, g["ad"], wrap=True); _w(ws, r, 3, g["adet"], align="center")
            _w(ws, r, 4, g["deger"], wrap=True); _w(ws, r, 5, "Pano" if kat_ == "pano" else "Saha", align="center")
            _w(ws, r, 6, ", ".join(f"Ç{n}" for n in sorted(g["cizim"])), align="center")
            r += 1
    _genislik(ws, [16, 48, 8, 70, 9, 12])
    ws.freeze_panes = "A5"
    _sayfa(ws)

    # ===================================================== 3b) AYDINLATMA VE PRİZ ENVANTERİ
    ws = wb.create_sheet("Aydınlatma ve Priz Envanteri")
    _baslik(ws, "AYDINLATMA VE PRİZ ENVANTERİ (cihaz → hat → faz → kablo)", alt, 9)
    _w(ws, 3, 1, "Fire/yedek payı (%)", True); ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=2)
    _w(ws, 3, 3, T["yedek_pay_yuzde"], fill=_PAR)
    for j, h_ in enumerate(["Poz", "Cihaz adı", "Adet", "Birim güç (W)", "Toplam güç (W)", "Hat (RCBO)", "Faz", "Kablo", "Not"], start=1):
        _w(ws, 5, j, h_, True, _HDR, wrap=True, align="center")
    r = 6
    r_ilk_env = r
    tip_ad = {"priz": "Priz girişi, 16 A, 230 V, topraklı", "led": "İç mekan LED armatür/lamba", "projektor": "LED projektör, IP65", "avize": "Büyük avize + bağlantı ünitesi"}
    for h in hatlar:
        if not h.get("elemanlar"):
            continue
        lst = [e for e in M["elemanlar"] if e["hat"] == h["id"]]
        _w(ws, r, 1, poz_aralik([e["poz"] for e in lst]), True, wrap=True); _w(ws, r, 2, tip_ad[lst[0]["tip"]], wrap=True)
        _w(ws, r, 3, len(lst), fill=_PAR, align="center"); _w(ws, r, 4, lst[0]["w"], fill=_PAR, fmt="0")
        _w(ws, r, 5, f"=C{r}*D{r}", fmt="0"); _w(ws, r, 6, h["id"], True, align="center")
        _w(ws, r, 7, h["faz"], align="center")
        _w(ws, r, 8, f"{h['kablo_no']}: {h['tip']} {h['damar']} mm², L={tr(h['L'], 0)} m")
        nt = {"led": "C eğrili RCBO (sürücü inrush)", "projektor": "PROJEKTOR_GUCU_W varsayım; IP65 kutular", "priz": "kd=%s" % tr(h["kd"], 1), "avize": "cosφ=%s (varsayım)" % tr(h["cos"], 2)}[lst[0]["tip"]]
        if h.get("kesit_artirildi"):
            nt += f"; kesit {kes_txt(h['kesit_ilk'])}→{kes_txt(h['kesit'])} mm² büyütüldü"
        _w(ws, r, 9, nt, wrap=True)
        r += 1
    r_son_env = r - 1
    _w(ws, r, 1, "TOPLAM", True, _HDR); _w(ws, r, 2, "", fill=_HDR)
    _w(ws, r, 3, f"=SUM(C{r_ilk_env}:C{r_son_env})", True, _HDR, "0"); _w(ws, r, 4, "", fill=_HDR)
    _w(ws, r, 5, f"=SUM(E{r_ilk_env}:E{r_son_env})", True, _HDR, "0")
    for j in range(6, 10):
        _w(ws, r, j, "", fill=_HDR)
    r_top = r
    r += 2
    _w(ws, r, 1, "EK KALEMLER", True, _HDR); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    r += 1
    pj_h = [h for h in hatlar if h["grup"] == "PROJ"]
    if pj_h:
        for i_, h in enumerate(pj_h, start=1):
            _w(ws, r, 1, f"JB{i_}", True); _w(ws, r, 2, "IP65 bağlantı kutusu (dış hat)"); _w(ws, r, 3, 1, fill=_PAR, align="center")
            _w(ws, r, 4, "", border=True); _w(ws, r, 5, "", border=True); _w(ws, r, 6, h["id"], True, align="center"); _w(ws, r, 7, h["faz"], align="center")
            _w(ws, r, 8, f"{h['kablo_no']}"); _w(ws, r, 9, "projektör grubu başına 1 kutu")
            r += 1
        sat_len = "+".join(f"'Kablo Çizelgesi'!K{ref['KC_rows'][h['id']]}" for h in pj_h)
        _w(ws, r, 1, "W14–W16" if len(pj_h) == 3 else "-", True)
        ws.cell(row=r, column=1).value = poz_aralik([h["kablo_no"] for h in pj_h])
        _w(ws, r, 2, "Dış hat kablosu metrajı (NYY, uzaklık parametreden, fire dahil)")
        _w(ws, r, 3, f"=ROUNDUP(({sat_len})*(1+$C$3/100),0)", True, _OK, "0"); _w(ws, r, 4, "m", align="center")
        _w(ws, r, 5, "", border=True); _w(ws, r, 6, "PJ", align="center"); _w(ws, r, 7, "-", align="center")
        _w(ws, r, 8, f"3x2,5 mm² (kesit hat çizelgesine göre), uzaklık {tr(pj_h[0]['L'], 0)} m")
        _w(ws, r, 9, "toprak altı veya kablo kanalı")
        r += 1
    _w(ws, r, 1, "KT1", True); _w(ws, r, 2, "Kablo tavası (iç mekan)"); _w(ws, r, 3, 1, fill=_PAR, align="center")
    for j in range(4, 8):
        _w(ws, r, j, "-" if j in (6, 7) else "", align="center")
    _w(ws, r, 8, "ana kablo yolu"); _w(ws, r, 9, "Çizim 6")
    _genislik(ws, [18, 44, 8, 12, 14, 12, 8, 46, 40])
    ws.freeze_panes = "A6"
    _sayfa(ws)
    ref["ENV_top"] = f"'Aydınlatma ve Priz Envanteri'!E{r_top}"

    # ===================================================== 4) TOPRAKLAMA VE YILDIRIM MALZEMESİ
    ws = wb.create_sheet("Topraklama ve Yıldırım Malzeme")
    _baslik(ws, "TOPRAKLAMA VE YILDIRIMDAN KORUNMA MALZEMESİ (metraj formüllerle)", alt, 6)
    tp = [
        ("Bina boyu a (m)", T["bina_a_m"]), ("Bina eni b (m)", T["bina_b_m"]), ("Halka ofseti (m)", T["halka_ofset_m"]),
        ("Halka derinliği (m)", T["halka_derinlik_m"]), ("Çubuk boyu (m)", T["cubuk_boy_m"]), ("Çubuk aralığı üst sınır (m)", T["cubuk_aralik_max_m"]),
        ("İniş iletkeni ortalama aralık sınırı (m)", T["inis_aralik"]), ("Yedek/fire payı (%)", T["yedek_pay_yuzde"]),
        ("Gövde yüksekliği (m)", T["govde_yuksekligi_m"]), ("Bina yüksekliği (m)", T["bina_yuksekligi_m"]),
        ("Minare yüksekliği (m)", T["minare_yuksekligi_m"]), ("Kubbe yarıçapı (m)", T["kubbe_cap_m"] / 2),
        ("Kontrol çukuru adedi", T["kontrol_cukuru_adet"]), ("PN1 BPE - EPB1 hattı (m)", T["pano_epb_hatti_m"]),
        ("EPB1 - halka hattı (m)", T["epb_halka_hatti_m"]), ("Trafo elektrodu çubuk adedi", T["trafo_elektrodu"]["cubuk_adet"]),
        ("Trafo elektrodu çubuk aralığı (m)", T["trafo_elektrodu"]["aralik_m"]), ("EB bağlantı iletkeni ort. boyu (m)", T["eb_ortalama_uzunluk_m"]),
        ("EB bağlantı noktası adedi", len(T["eb"])), ("Trafo - pano mesafesi (m)", S["trafo_pano_mesafe_m"]),
        ("YK1 boyu (m)", T["yk_boylari_m"]["YK1"]), ("YK2 boyu (m)", T["yk_boylari_m"]["YK2"]), ("YK3 boyu (m)", T["yk_boylari_m"]["YK3"]),
    ]
    pr_ = {}
    for i, (a_, v) in enumerate(tp):
        _w(ws, 3 + i, 1, a_, True); ws.merge_cells(start_row=3 + i, start_column=1, end_row=3 + i, end_column=2)
        _w(ws, 3 + i, 3, v, fill=_PAR)
        pr_[i] = f"$C${3 + i}"
    (pa, pb, po, pd_, pcb, pam, pia, pyp, pgh, pbh, pmh, pkr, pkc, ppe1, ppe2, ptn, pta, pebl, pebn, ptm, pyk1, pyk2, pyk3) = [pr_[i] for i in range(len(tp))]
    hr2 = 3 + len(tp) + 1
    for j, h in enumerate(["Poz", "Malzeme / iş kalemi", "Birim", "Miktar", "Hesap", "Teknik değer"], start=1):
        _w(ws, hr2, j, h, True, _HDR, align="center")
    r = hr2 + 1
    satir_ref = {}

    def sat(poz, ad, birim, formul, hesap, deger, anahtar=None, fmt="0"):
        nonlocal r
        _w(ws, r, 1, poz, True, wrap=True); _w(ws, r, 2, ad, wrap=True); _w(ws, r, 3, birim, align="center")
        _w(ws, r, 4, formul, True, _OK, fmt); _w(ws, r, 5, hesap, wrap=True); _w(ws, r, 6, deger, wrap=True)
        if anahtar:
            satir_ref[anahtar] = f"D{r}"
        r += 1

    cub_poz = poz_aralik([x["poz"] for x in T["cubuklar"]])
    il_poz = poz_aralik([x["poz"] for x in T["inis"]])
    ok_poz = poz_aralik([x["ok"] for x in T["inis"]])
    kc_poz = poz_aralik([x["poz"] for x in T["cukurlar"]])
    eb_poz = poz_aralik([x["poz"] for x in T["eb"]])
    sat("TH1", "Halka iletkeni (net)", "m", f"=2*(({pa}+2*{po})+({pb}+2*{po}))", "2·((a+2·ofs)+(b+2·ofs))", T["halka_iletkeni"], "halka")
    sat("TH1", "Halka iletkeni (fire payı dahil sipariş)", "m", f"=ROUNDUP({satir_ref['halka']}*(1+{pyp}/100),0)", "net × (1+pay)", T["halka_iletkeni"], "halka_sip")
    sat(cub_poz, "Topraklama çubuğu Ø%s mm, bakır kaplı" % tr(T["cubuk_cap_mm"], 0), "adet",
        f"=2*(ROUNDUP(({pa}+2*{po})/{pam},0)+ROUNDUP(({pb}+2*{po})/{pam},0))", "kenar başına ROUNDUP(L/aralık) toplamı", f"{tr(T['cubuk_boy_m'], 0)} m boy", "cubuk")
    sat(cub_poz, "Çubuk toplam boyu", "m", f"={satir_ref['cubuk']}*{pcb}", "adet × boy", "", "cubuk_boy")
    sat(cub_poz, "Gerçek çubuk aralığı (kontrol 6-10 m)", "m", f"=MAX(({pa}+2*{po})/ROUNDUP(({pa}+2*{po})/{pam},0),({pb}+2*{po})/ROUNDUP(({pb}+2*{po})/{pam},0))", "kenar uzunluğu / aralık sayısı", "6-10 m hedef", fmt="0.00")
    sat(kc_poz, "Kontrol (ölçüm) çukuru, kapaklı", "adet", f"={pkc}", "parametre", "300x300x300 mm", "cukur")
    sat(il_poz, "Yıldırım iniş iletkeni adedi", "adet", f"=MIN(8,MAX(2,ROUNDUP(2*({pa}+{pb})/{pia},0)))", "ROUNDUP(çevre/aralık), min 2", f"{tr(T['inis_kesit'], 0)} mm² Cu veya 30x3,5 mm galvaniz şerit", "inis")
    sat(ok_poz, "Ölçüm (test) ayırma noktası", "adet", f"={satir_ref['inis']}", "iniş iletkeni başına 1", "korumalı kutulu test klemensi", "ok")
    sat(il_poz, "İniş iletkeni metrajı (fire dahil)", "m",
        f"=ROUNDUP(({satir_ref['inis']}*({pgh}+{pd_})+({pmh}-{pgh})+(PI()*{pkr}/2+2))*(1+{pyp}/100),0)",
        "n·(h_gövde+derinlik) + minare boyu + kubbe iletkeni, +pay", f"{tr(T['inis_kesit'], 0)} mm² Cu", "inis_m")
    sat("ÇÇ1", "Çatı çevre iletkeni (30x3,5 galvaniz şerit, fire dahil)", "m",
        f"=ROUNDUP((2*({pa}+{pb})+{pb}/2+3)*(1+{pyp}/100),0)", "çevre + kubbe bağlantısı + minare bağlantısı, +pay", T["cati_iletkeni"], "cati")
    sat("YK1–YK3", "Franklin çubuğu (Ø16 mm, Cu/paslanmaz)", "adet", "=3", "minare ucu, kubbe tepesi, çatı kenarı", "boylar: aşağıda", "yk")
    sat("YK1–YK3", "Franklin çubuğu toplam boy", "m", f"={pyk1}+{pyk2}+{pyk3}", "YK1+YK2+YK3", "", fmt="0.0")
    sat("EPB1", "Esas eş potansiyel barası", "adet", "=1", "ana pano yanı", "Cu 40x5 mm, 12 klemens")
    sat("PE1", "PN1 BPE - EPB1 iletkeni 50 mm² Cu (sarı-yeşil)", "m", f"=ROUNDUP({ppe1}*(1+{pyp}/100),0)", "uzunluk + pay", "50 mm² Cu", fmt="0")
    sat("PE2", "EPB1 - halka iletkeni 50 mm² Cu (sarı-yeşil)", "m", f"=ROUNDUP({ppe2}*(1+{pyp}/100),0)", "uzunluk + pay", "50 mm² Cu", fmt="0")
    sat("PE3", "Trafo - PN1 PE iletkeni 1x50 mm² Cu", "m", f"=ROUNDUP({ptm}*(1+{pyp}/100),0)", "trafo-pano mesafesi + pay", "NYY 1x50 mm² sarı-yeşil", fmt="0")
    sat("TE1", "Trafo elektrodu çubuğu", "adet", f"={ptn}", "parametre", "Ø%s x %s m" % (tr(T["cubuk_cap_mm"], 0), tr(T["trafo_elektrodu"]["cubuk_boy_m"], 0)), "te")
    sat("TE1", "Trafo elektrodu bağlantı iletkeni 50 mm² Cu", "m", f"=({ptn}-1)*{pta}+6", "(n−1)·aralık + 6 m", "50 mm² Cu çıplak", fmt="0")
    sat(eb_poz, "Eş potansiyel bağlantı iletkeni 16 mm² Cu (sarı-yeşil)", "m", f"=ROUNDUP({pebn}*{pebl}*(1+{pyp}/100),0)", "adet × ort. boy + pay", "16 mm² Cu", fmt="0")
    sat(eb_poz, "Boru/kanal bağlantı kelepçesi (EB)", "adet", f"={pebn}", "EB başına 1", "bakır/paslanmaz kelepçe", "ebk")
    sat("İF1", "İzolasyon flanşı (gaz borusu)", "adet", "=1", "gaz girişi", "gaz şebekesi yönetmeliğine uygun")
    sat("EK1", "Eksotermik kaynak (kalıp + toz kartuşu)", "adet",
        f"={satir_ref['cubuk']}+{satir_ref['inis']}+2+1+{pebn}", "çubuk + iniş + PE1/PE2 + TE1 + EB", "50 mm² uyumlu kartuş", "ek")
    sat("EK1", "Eksotermik kaynak kalıbı (yeniden kullanılır)", "adet", "=3", "çubuk-halka, T bağlantı, kelepçe tipi", "grafit kalıp")
    sat("-", "Bitümlü koruma bandı (yeraltı bağlantılar)", "adet", f"=ROUNDUP({satir_ref['ek']}/10,0)", "10 bağlantıya 1 rulo", "")
    _genislik(ws, [16, 50, 8, 12, 46, 40])
    ws.freeze_panes = ws.cell(row=hr2 + 1, column=1)
    _sayfa(ws)
    ref["TP"] = dict(satir_ref)

    # ===================================================== 5) KESİCİ SELEKTİVİTE
    ws = wb.create_sheet("Kesici Selektivite")
    _baslik(ws, "KESİCİ SELEKTİVİTESİ: Q0 → RCR0 (RCD) → RCD1-2 / RCBO ve kısa devre kapasiteleri", alt, 10)
    sp = [
        ("Trafo gücü (kVA)", S["trafo_kva"]), ("U_LL (V)", S["U_LL"]), ("uk (%)", S["trafo_uk_yuzde"]), ("ur (%)", S["trafo_ur_yuzde"]),
        ("Trafo-pano mesafesi (m)", S["trafo_pano_mesafe_m"]), ("W0 kesiti (mm²)", M["besleme"]["kesit"]), ("X reaktans (mΩ/m)", K["X_mohm_m"]),
        ("Q0 Icu (kA)", C["Q0"]["Icu_kA"]), ("Q0 Icw (kA) (varsayım)", C["Q0"]["Icw_kA"]), ("Q0 Ir (A)", C["Q0"]["Ir_A"]),
        ("Q0 Isd (×Ir)", C["Q0"]["Isd_x"]), ("Q0 tsd (ms)", C["Q0"]["tsd_ms"]),
        ("RCR0 I∆n (mA)", C["RCR0"]["I_dn_A"] * 1000), ("RCR0 gecikme (ms)", C["RCR0"]["gecikme_ms"]),
        ("Min. I∆n oranı", C["secicilik"]["I_dn_oran_min"]), ("Zaman payı (ms)", C["secicilik"]["zaman_payi_ms"]),
        ("Alt kademe kısa devre açma süresi (ms) (varsayım)", C["secicilik"]["alt_kisa_devre_acma_ms"]),
        ("Gerekli minimum Q0 Icu (kA) (istek)", 25.0),
    ]
    sr = {}
    for i, (a_, v) in enumerate(sp):
        _w(ws, 3 + i, 1, a_, True); ws.merge_cells(start_row=3 + i, start_column=1, end_row=3 + i, end_column=3)
        _w(ws, 3 + i, 4, v, fill=_PAR)
        sr[i] = f"$D${3 + i}"
    (sS, sU, suk, sur, sL, sKes, sX, sIcu, sIcw, sIr, sIsd, stsd, sIdn, sTr, sOr, sPay, sTa, sIcuMin) = [sr[i] for i in range(len(sp))]
    r = 3 + len(sp) + 1
    _w(ws, r, 1, "A) KISA DEVRE HESABI", True, _HDR); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1
    kd = {}

    def kdw(ad, formul, birim, fmt="0.000", anahtar=None):
        nonlocal r
        _w(ws, r, 1, ad, True); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        _w(ws, r, 4, formul, True, _OK, fmt); _w(ws, r, 5, birim)
        if anahtar:
            kd[anahtar] = f"$D${r}"
        r += 1
    kdw("Trafo nominal akımı In", f"={sS}*1000/(SQRT(3)*{sU})", "A", "0.0", "In")
    kdw("Trafo uçları Ik'' (sonsuz şebeke)", f"={kd['In']}/({suk}/100)/1000", "kA", "0.00", "IkT")
    kdw("Z_T (trafo)", f"={suk}/100*{sU}^2/({sS}*1000)", "Ω", "0.00000", "ZT")
    kdw("R_T", f"={sur}/100*{sU}^2/({sS}*1000)", "Ω", "0.00000", "RT")
    kdw("X_T", f"=SQRT({kd['ZT']}^2-{kd['RT']}^2)", "Ω", "0.00000", "XT")
    kdw("W0 direnci R_c (20 °C, κ=56)", f"={sL}/(56*{sKes})", "Ω", "0.00000", "Rc")
    kdw("W0 reaktansı X_c", f"={sX}/1000*{sL}", "Ω", "0.00000", "Xc")
    kdw("Z toplam", f"=SQRT(({kd['RT']}+{kd['Rc']})^2+({kd['XT']}+{kd['Xc']})^2)", "Ω", "0.00000", "Zt")
    kdw("PN1 barası Ik'' (3 faz, maks.)", f"={sU}/(SQRT(3)*{kd['Zt']})/1000", "kA", "0.00", "Ik")
    r += 1
    _w(ws, r, 1, "B) KISA DEVRE KAPASİTESİ KONTROLÜ", True, _HDR); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1
    for j, h in enumerate(["Kademe", "Poz", "Cihaz", "In / Ir (A)", "I∆n (mA)", "Açma süresi / gecikme (ms)", "Icu / Icn / Inc (kA)", "Ik (kA)", "Kapasite ≥ Ik", "Not"], start=1):
        _w(ws, r, j, h, True, _HDR, wrap=True, align="center")
    ws.row_dimensions[r].height = 30
    r += 1
    rc, kk, rb, qg, rcd = C["RCR0"], C["K_klima"], C["RCBO"], C["Q_klima_grup"], C["RCD_klima"]
    ses_mcb_ = [h["id"] for h in hatlar if h.get("koruma_tur") == "mcb"]
    kd_rows = [
        ("1", "Q0", f"MCCB 4P {C['Q0']['govde_A']} A", f"={sIr}", "-", f"={stsd}", f"={sIcu}", "Q0 Icu ≥ 25 kA istek şartı ayrıca kontrol edilir"),
        ("2", "RCR0", "Gecikmeli kaçak akım rölesi (Q0 üzerinden açar)", "-", f"={sIdn}", f"={sTr}", f"={sIcu}", "Kesme Q0 ile yapılır"),
        ("3", "Q1, Q2", f"MCB 4P {qg['egri']}{qg['In_A']} A", qg["In_A"], "-", f"={sTa}", qg["Icu_kA"], "RCD1-2 yedek koruması"),
        ("3", "Q3", f"MCB 4P {C['Q3']['egri']}{C['Q3']['In_A']} A (dış aydınlatma)", C["Q3"]["In_A"], "-", f"={sTa}", C["Q3"]["Icu_kA"], "KM1 ve PJ RCBO'ları yedek koruması"),
        ("3", "KM1", f"Kontaktör 3P {C['KM1']['In_A']} A (Q3 ile koordineli)", C["KM1"]["In_A"], "-", "-", C["Q3"]["Icu_kA"], "Kısa devre koruması Q3 (koordinasyon tipi 1)"),
        ("3", "RCD1, RCD2", f"RCD 4P {rcd['In_A']} A, Tip {rcd['tip']}", rcd["In_A"], rcd["I_dn_mA"], rcd["t_acma_ms"], rcd["Inc_kA"], "Inc: yedek koruma (Q1/Q2) ile birlikte"),
        ("4", "K1–K6", f"MPCB 3P {kk['ayar_min_A']}-{kk['ayar_max_A']} A", kk["ayar_A"], "-", f"={sTa}", kk["Icu_kA"], ""),
        ("4", poz_aralik([h["id"] for h in hatlar if not klima_mi(h["id"]) and h.get("koruma_tur") != "mcb"]), f"RCBO 1P+N, {rb['I_dn_mA']} mA, Tip {rb['tip']}", "10-16", rb["I_dn_mA"], rb["t_acma_ms"], rb["Icn_kA"], "6 kA yetmez: Ik > 6 kA; LED hatlarında C eğrisi (inrush)"),
    ]
    if ses_mcb_:
        rc3 = C["RCD3"]
        kd_rows.append(("4", "RCD3", f"RCD 2P {rc3['In_A']} A, {rc3['I_dn_mA']} mA, Tip {rc3['tip']} (ses ortak)", rc3["In_A"], rc3["I_dn_mA"], rc3["t_acma_ms"], rc3["Inc_kA"],
                        "Inc: altındaki MCB'lerle birlikte yedek koruma"))
        kd_rows.append(("5", poz_aralik(ses_mcb_), f"MCB 1P+N {C['MCB1']['egri']}{[h['rcbo_In'] for h in hatlar if h['id'] in ses_mcb_][0]} A (RCD3 altında)",
                        [h["rcbo_In"] for h in hatlar if h["id"] in ses_mcb_][0], "-", f"={sTa}", C["MCB1"]["Icu_kA"], "Ses kolları aşırı akım koruması"))
    for kdm in kd_rows:
        _w(ws, r, 1, kdm[0], align="center"); _w(ws, r, 2, kdm[1], True, wrap=True); _w(ws, r, 3, kdm[2], wrap=True)
        for j, v in enumerate(kdm[3:7], start=4):
            _w(ws, r, j, v, align="center")
        _w(ws, r, 8, f"={kd['Ik']}", fmt="0.00")
        _w(ws, r, 9, f'=IF(G{r}>=H{r},"UYGUN","UYGUN DEĞİL")', True, _OK, align="center")
        _w(ws, r, 10, kdm[7], wrap=True)
        r += 1
    _w(ws, r, 2, "Q0 ≥ 25 kA şartı", True); _w(ws, r, 3, "Q0 Icu ≥ istenen minimum")
    _w(ws, r, 7, f"={sIcu}"); _w(ws, r, 8, f"={sIcuMin}"); _w(ws, r, 9, f'=IF(G{r}>=H{r},"UYGUN","UYGUN DEĞİL")', True, _OK, align="center")
    r += 2
    _w(ws, r, 1, "C) KAÇAK AKIM SELEKTİVİTESİ (üst kademe → alt kademe)", True, _HDR); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1
    for j, h in enumerate(["Üst", "Alt", "Üst I∆n (mA)", "Alt I∆n (mA)", "I∆n oranı", "Oran ≥ min", "Üst gecikme (ms)", "Alt açma (ms)", "Zaman ≥ alt+pay", "Sonuç"], start=1):
        _w(ws, r, j, h, True, _HDR, wrap=True, align="center")
    ws.row_dimensions[r].height = 30
    r += 1
    sel_satirlari = [("RCR0", "RCD1 / RCD2", rcd["I_dn_mA"], rcd["t_acma_ms"]), ("RCR0", "RCBO (AY, PR, KR, SS, AC, YA)", rb["I_dn_mA"], rb["t_acma_ms"])]
    if ses_mcb_:
        sel_satirlari.append(("RCR0", "RCD3 (ses ortak)", C["RCD3"]["I_dn_mA"], C["RCD3"]["t_acma_ms"]))
    for ust, alt_, adn, ams in sel_satirlari:
        _w(ws, r, 1, ust, True); _w(ws, r, 2, alt_, wrap=True)
        _w(ws, r, 3, f"={sIdn}"); _w(ws, r, 4, adn); _w(ws, r, 5, f"=C{r}/D{r}", fmt="0.0")
        _w(ws, r, 6, f'=IF(E{r}>={sOr},"UYGUN","UYGUN DEĞİL")', align="center")
        _w(ws, r, 7, f"={sTr}"); _w(ws, r, 8, ams); _w(ws, r, 9, f'=IF(G{r}>=H{r}+{sPay},"UYGUN","UYGUN DEĞİL")', align="center")
        _w(ws, r, 10, f'=IF(AND(F{r}="UYGUN",I{r}="UYGUN"),"SEÇİCİ","SEÇİCİ DEĞİL")', True, _OK, align="center")
        r += 1
    r += 1
    _w(ws, r, 1, "D) AŞIRI AKIM SEÇİCİLİĞİ (Q0 zaman gecikmeli, alt kademeler anlık)", True, _HDR); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1
    for j, h in enumerate(["Parametre", "", "", "Değer", "Birim", "Kontrol", "", "", "", "Not"], start=1):
        _w(ws, r, j, h, True, _HDR, align="center")
    r += 1
    d_rows = [
        ("Q0 kısa zaman eşiği Isd = Ir × katsayı", f"={sIr}*{sIsd}", "A", None, "Q0 Ii (anlık) kapalı"),
        ("Q0 kısa zaman gecikmesi tsd", f"={stsd}", "ms", f'=IF(D{{r}}>={sTa}+{sPay},"UYGUN","UYGUN DEĞİL")', "tsd ≥ alt kademe açma süresi + pay"),
        ("Q0 Icw ≥ PN1 barası Ik", f"={sIcw}", "kA", f'=IF(D{{r}}>={kd["Ik"]},"UYGUN","UYGUN DEĞİL")', "Icw değeri üretici kataloğundan doğrulanmalı"),
        ("PN1 barası Ik (alt kademe max. kısa devre)", f"={kd['Ik']}", "kA", None, "Alt kademe cihazlar Ik'yı kesebilmeli (bölüm B)"),
    ]
    for ad, f_, bir, kon, nt in d_rows:
        _w(ws, r, 1, ad, True); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        _w(ws, r, 4, f_, True, _OK, "0.00"); _w(ws, r, 5, bir)
        if kon:
            _w(ws, r, 6, kon.replace("{r}", str(r)), True, _OK, align="center")
        _w(ws, r, 10, nt, wrap=True)
        r += 1
    r += 1
    for t_ in ("Not: Tam seçicilik için üretici seçicilik (kademelendirme) tabloları ile Q0 → Q1/Q2 ve Q0 → RCBO çiftleri ayrıca doğrulanmalıdır.",
               "Not: Q0 Ii (anlık) kapalı/yüksek ayarlıdır; bu durumda Q0 Icw ≥ Ik ve kablo/bara termik dayanımı sağlanmalıdır."):
        ws.cell(row=r, column=1, value=t_).font = Font(name="Liberation Sans", size=9, italic=True)
        r += 1
    _genislik(ws, [14, 22, 40, 12, 11, 18, 16, 12, 16, 44])
    _sayfa(ws)
    ref["SEL_Ik"] = kd["Ik"]

    wb.calculation.fullCalcOnLoad = True
    wb.save(yol)
    return ref


def excel_yeniden_hesapla(yol):
    """LibreOffice ile formül değerlerini önbelleğe alır (varsa). Başarısızsa orijinal dosya kalır."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        uyar("LibreOffice bulunamadı: xlsx formül değerleri Excel açıldığında hesaplanacak.")
        return False
    tmp = tempfile.mkdtemp()
    try:
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, yol],
                       check=True, capture_output=True, timeout=180)
        out = os.path.join(tmp, os.path.basename(yol))
        if os.path.exists(out):
            shutil.copyfile(out, yol)
            return True
    except Exception as e:                                  # noqa
        uyar(f"LibreOffice yeniden hesaplama başarısız: {e}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return False


# =====================================================================
#  DOĞRULAMA (poz tutarlılığı, sayısal tutarlılık, PDF boyutu)
# =====================================================================
def topraklama_metraj(P, T):
    """Excel formülleri ile aynı sonuçları üretir (karşılaştırma ve rapor için)."""
    a, b, o = T["bina_a_m"], T["bina_b_m"], T["halka_ofset_m"]
    pay = 1 + T["yedek_pay_yuzde"] / 100.0
    halka = 2 * ((a + 2 * o) + (b + 2 * o))
    n_inis = len(T["inis"])
    kr = T["kubbe_cap_m"] / 2
    out = {
        "halka": halka, "halka_sip": math.ceil(halka * pay), "cubuk": len(T["cubuklar"]),
        "cukur": int(T["kontrol_cukuru_adet"]), "inis": n_inis, "ok": n_inis,
        "inis_m": math.ceil((n_inis * (T["govde_yuksekligi_m"] + T["halka_derinlik_m"]) + (T["minare_yuksekligi_m"] - T["govde_yuksekligi_m"])
                             + (math.pi * kr / 2 + 2)) * pay),
        "cati": math.ceil((2 * (a + b) + b / 2 + 3) * pay),
        "ek": len(T["cubuklar"]) + n_inis + 2 + 1 + len(T["eb"]),
    }
    out["cubuk_boy"] = out["cubuk"] * T["cubuk_boy_m"]
    return out


def svg_pozlari(yol):
    kok = ET.parse(yol).getroot()
    kume = set()
    for e in kok.iter():
        v = e.attrib.get("data-poz")
        if v:
            kume.update(v.split())
    return kume


def xlsx_pozlari(yol):
    wb = load_workbook(yol)
    sonuc = {}
    # Yük Dengesi: A sütunu (poz) -> "TOPLAM"a kadar
    ws = wb["Yük Dengesi"]
    s = set()
    for r in range(10, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v == "TOPLAM":
            break
        if v:
            s.add(str(v))
    sonuc["Yük Dengesi"] = s
    ws = wb["Kablo Çizelgesi"]
    s = set()
    for r in range(15, ws.max_row + 1):
        for c_ in (1, 2):
            v = ws.cell(row=r, column=c_).value
            if v and re.fullmatch(r"[A-ZÇĞİÖŞÜ]+\d+", str(v)):
                s.add(str(v))
    sonuc["Kablo Çizelgesi"] = s
    for ad in ("Pano Malzeme Listesi", "Topraklama ve Yıldırım Malzeme"):
        ws = wb[ad]
        s = set()
        basla = 1
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == "Poz":
                basla = r + 1
                break
        for r in range(basla, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if v and isinstance(v, str) and not v.startswith(("PANO", "SAHA", "-")):
                s.update(poz_genislet(v))
        sonuc[ad] = s
    ws = wb["Aydınlatma ve Priz Envanteri"]
    s, env_hat = set(), {}
    basla = 1
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == "Poz":
            basla = r + 1
            break
    for r in range(basla, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v and isinstance(v, str) and v not in ("TOPLAM", "EK KALEMLER"):
            hat_ = ws.cell(row=r, column=6).value
            for pz in poz_genislet(v):
                s.add(pz)
                env_hat[pz] = str(hat_)
    sonuc["Aydınlatma ve Priz Envanteri"] = s
    ws = wb["Kesici Selektivite"]
    s = set()
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=2).value
        if v and isinstance(v, str) and re.fullmatch(r"[A-ZÇĞİÖŞÜ0-9–, \-]+", v):
            try:
                s.update(poz_genislet(v))
            except Exception:
                pass
    sonuc["Kesici Selektivite"] = s
    return sonuc, env_hat


def svg_hat_haritasi(yol):
    """data-poz ve data-hat birlikte taşıyan öğeler: poz -> hat."""
    kok = ET.parse(yol).getroot()
    h = {}
    for e in kok.iter():
        v, ht = e.attrib.get("data-poz"), e.attrib.get("data-hat")
        if v and ht:
            for pz in v.split():
                h[pz] = ht
    return h


def dogrula_hepsi(P, M, PN, T, KAYIT, svgler, xlsx, pdfler):
    """Dönüş: dict(hatalar=[...], matris=[...], sayisal=[...], pdf=[...])"""
    hatalar = []
    bilinen = []
    bilinen_hat = {h["id"] for h in M["hatlar"] if h.get("sabit_kesit_uygunsuz")}
    kayit_poz = {k["poz"]: k for k in KAYIT}
    # ---- çizimlerdeki pozlar
    cz = {n: svg_pozlari(y) for n, y in svgler.items()}
    for n, kume in cz.items():
        bilinmeyen = sorted(p for p in kume if p not in kayit_poz)
        if bilinmeyen:
            hatalar.append(f"Çizim {n}: kayıtta olmayan poz(lar): {bilinmeyen}")
    for k in KAYIT:
        for n in k["cizim"]:
            if k["poz"] not in cz[n]:
                hatalar.append(f"Çizim {n}: poz {k['poz']} çizimde bulunamadı ({k['ad']})")
    # ---- tablolardaki pozlar
    tb, env_hat = xlsx_pozlari(xlsx)
    tum_tablo = set().union(*tb.values())
    for sayfa, kume in tb.items():
        bil = sorted(p for p in kume if p not in kayit_poz)
        if bil:
            hatalar.append(f"Tablo '{sayfa}': kayıtta olmayan poz(lar): {bil}")
    for k in KAYIT:
        hedef = "Topraklama ve Yıldırım Malzeme" if k["kat"] == "topraklama" else "Pano Malzeme Listesi"
        if k["poz"] not in tb[hedef]:
            hatalar.append(f"Tablo '{hedef}': poz {k['poz']} bulunamadı")
    # envanter kalemleri + çizim 6 eleman -> hat -> tek hat -> tablo zinciri
    hat6 = svg_hat_haritasi(svgler[6]) if 6 in svgler else {}
    for k in KAYIT:
        if k.get("envanter"):
            if k["poz"] not in tb["Aydınlatma ve Priz Envanteri"]:
                hatalar.append(f"Envanter sayfasında poz {k['poz']} bulunamadı")
            elif env_hat.get(k["poz"]) != k["hat"]:
                hatalar.append(f"Envanter hat uyuşmazlığı: {k['poz']} tablo={env_hat.get(k['poz'])} kayıt={k['hat']}")
    for pz, ht in hat6.items():
        kk = kayit_poz.get(pz)
        if kk is None or kk.get("hat") != ht:
            hatalar.append(f"Çizim 6: {pz} etiketindeki hat {ht} kayıtla uyuşmuyor")
        if ht not in cz[1] or ht not in tb["Kablo Çizelgesi"] or ht not in tb["Yük Dengesi"]:
            hatalar.append(f"Çizim 6: {pz} → hat {ht} tek hat şemasında/kablo/yük tablosunda yok")
    for k in KAYIT:
        if k.get("envanter") and k["hat"] != "-" and 6 in k["cizim"] and k["poz"] not in hat6:
            hatalar.append(f"Çizim 6: {k['poz']} etiketinde hat bilgisi (data-hat) yok")
    for h in M["hatlar"]:
        if h["id"] not in tb["Yük Dengesi"] or h["id"] not in tb["Kablo Çizelgesi"] or h["kablo_no"] not in tb["Kablo Çizelgesi"]:
            hatalar.append(f"Yük/Kablo çizelgesinde eksik: {h['id']}")
    # ---- poz matrisi (grup düzeyinde)
    matris = []
    for g in satirlari_grupla(KAYIT):
        satir = {"poz": g["poz"], "ad": g["ad"], "kat": g["kat"]}
        for n in (1, 2, 3, 6):
            satir[f"c{n}"] = all(p in cz[n] for p in g["pozlar"])
            satir[f"c{n}_any"] = any(p in cz[n] for p in g["pozlar"])
        hedef = "Topraklama ve Yıldırım Malzeme" if g["kat"] == "topraklama" else "Pano Malzeme Listesi"
        satir["tablo"] = all(p in tb[hedef] for p in g["pozlar"])
        matris.append(satir)
    # ---- sayısal doğrulama (xlsx önbellek değerleri)
    sayisal = []
    try:
        wbv = load_workbook(xlsx, data_only=True)
        ws = wbv["Yük Dengesi"]
        deg = {}
        for r in range(1, ws.max_row + 1):
            a_ = ws.cell(row=r, column=1).value
            if isinstance(a_, str):
                deg[a_] = ws.cell(row=r, column=2).value
        kontroller = [("Kurulu güç", deg.get("Kurulu güç (kW)"), M["p_kurulu"]), ("Talep gücü", deg.get("Talep gücü (kW)"), M["p_talep"]),
                      ("Faz dengesizliği", deg.get("Faz dengesizliği (%)"), M["dengesizlik"]), ("Hesap akımı Ib", deg.get("Hesap akımı Ib (A) - en yüklü faz"), M["Ib"]),
                      ("cosφ ort.", deg.get("cosφ (ortalama)"), M["cos_ort"])]
        wk = wbv["Kablo Çizelgesi"]
        kab = {}
        for r in range(15, wk.max_row + 1):
            kn = wk.cell(row=r, column=1).value
            if kn:
                kab[str(kn)] = r
        for h in M["hatlar"]:
            r = kab.get(h["kablo_no"])
            if r:
                kontroller.append((f"{h['kablo_no']} akım", wk.cell(row=r, column=8).value, h["I"]))
                kontroller.append((f"{h['kablo_no']} ΔU hat %", wk.cell(row=r, column=18).value, h["du_hat"]))
                kontroller.append((f"{h['kablo_no']} ΔU toplam %", wk.cell(row=r, column=20).value, h["du_toplam"]))
        r = kab.get(P["besleme"]["kablo_no"])
        if r:
            kontroller.append(("W0 ΔU %", wk.cell(row=r, column=18).value, M["besleme"]["du"]))
        wsl = wbv["Kesici Selektivite"]
        for r in range(1, wsl.max_row + 1):
            if wsl.cell(row=r, column=1).value == "PN1 barası Ik'' (3 faz, maks.)":
                kontroller.append(("PN1 Ik kA", wsl.cell(row=r, column=4).value, M["Ik_pano_kA"]))
        wt = wbv["Topraklama ve Yıldırım Malzeme"]
        mt = topraklama_metraj(P, T)
        kmap = {"Halka iletkeni (net)": "halka", "Halka iletkeni (fire payı dahil sipariş)": "halka_sip", "Yıldırım iniş iletkeni adedi": "inis",
                "İniş iletkeni metrajı (fire dahil)": "inis_m", "Eksotermik kaynak (kalıp + toz kartuşu)": "ek",
                "Kontrol (ölçüm) çukuru, kapaklı": "cukur", "Çatı çevre iletkeni (30x3,5 galvaniz şerit, fire dahil)": "cati",
                "Topraklama çubuğu Ø%s mm, bakır kaplı" % tr(T["cubuk_cap_mm"], 0): "cubuk"}
        for r in range(1, wt.max_row + 1):
            ad = wt.cell(row=r, column=2).value
            if ad in kmap:
                kontroller.append((f"Metraj: {ad}", wt.cell(row=r, column=4).value, mt[kmap[ad]]))
        for ad, xv, pv in kontroller:
            ok = xv is not None and abs(float(xv) - float(pv)) <= 1e-6 * max(1.0, abs(float(pv))) + 1e-9
            sayisal.append((ad, xv, pv, ok))
            if not ok:
                hatalar.append(f"Sayısal uyuşmazlık: {ad}: xlsx={xv} python={pv}")
        # formül sayısı ve UYGUN DEĞİL taraması
        wbf = load_workbook(xlsx)
        nform = sum(1 for w in wbf.worksheets for row in w.iter_rows() for c_ in row if isinstance(c_.value, str) and c_.value.startswith("="))
        sayisal.append(("Excel formül sayısı", nform, ">0", nform > 0))
        for w in wbv.worksheets:
            for row in w.iter_rows():
                for c_ in row:
                    if isinstance(c_.value, str) and c_.value in ("UYGUN DEĞİL", "SEÇİCİ DEĞİL"):
                        hid = w.cell(row=c_.row, column=2).value if w.title == "Kablo Çizelgesi" else None
                        if hid in bilinen_hat:
                            bilinen.append(f"Excel 'Kablo Çizelgesi' {c_.coordinate}: {hid} ΔU şartı sağlanmıyor (kesit parametreyle sabit)")
                        else:
                            hatalar.append(f"Excel '{w.title}' {c_.coordinate}: {c_.value}")
                    if isinstance(c_.value, str) and c_.value.startswith("#"):
                        hatalar.append(f"Excel '{w.title}' {c_.coordinate}: hata değeri {c_.value}")
    except Exception as e:                                  # noqa
        hatalar.append(f"Sayısal doğrulama yapılamadı: {e}")
    # ---- PDF boyutu
    pdf = []
    from pypdf import PdfReader
    for ad, y in pdfler.items():
        rd = PdfReader(y)
        sayfa = rd.pages[0]
        w_, h_ = float(sayfa.mediabox.width), float(sayfa.mediabox.height)
        a3 = abs(w_ - 1190.55) < 2 and abs(h_ - 841.89) < 2
        pdf.append((ad, len(rd.pages), round(w_ / 72 * 25.4), round(h_ / 72 * 25.4), a3))
        if ad.startswith("0") and ad[:2] in ("01", "02", "03", "06") and not (a3 and len(rd.pages) == 1):
            hatalar.append(f"{ad}: A3 yatay tek sayfa değil")
    return {"hatalar": hatalar, "bilinen": bilinen, "matris": matris, "sayisal": sayisal, "pdf": pdf, "cizim_pozlari": cz, "tablo_pozlari": tb}


# =====================================================================
#  PROJE RAPORU (reportlab)
# =====================================================================
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle, Image,
                                PageBreak, KeepTogether)
from xml.sax.saxutils import escape as _esc


def _fontlar():
    d = "/usr/share/fonts/truetype/liberation/"
    pdfmetrics.registerFont(TTFont("LS", d + "LiberationSans-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("LS-B", d + "LiberationSans-Bold.ttf"))
    pdfmetrics.registerFont(TTFont("LS-I", d + "LiberationSans-Italic.ttf"))
    pdfmetrics.registerFont(TTFont("DVS", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
    pdfmetrics.registerFontFamily("LS", normal="LS", bold="LS-B", italic="LS-I", boldItalic="LS-B")


def varsayimlar(P, M, PN, T):
    S, K, C = P["sistem"], P["kablo"], P["cihazlar"]
    h0 = M["hatlar"][0]
    hat_ = {h["id"]: h for h in M["hatlar"]}
    A_ = P["aydinlatma_priz"]
    mt = topraklama_metraj(P, T)
    L = lambda ids: sorted({tr(h["L"], 0) for h in M["hatlar"] if h["id"] in ids})
    return [
        ("Trafo", f"{tr(S['trafo_kva'], 0)} kVA, {S['trafo_grubu']}, uk=%{tr(S['trafo_uk_yuzde'])}, ur=%{tr(S['trafo_ur_yuzde'])}. OG tarafı yalnız blok gösterilmiştir.",
         "Trafo etiket/test verisi verilmedi; tipik değerler alındı. Kısa devre hesabı bu değerlere bağlıdır."),
        ("Kablo hesap sabitleri", f"κ(Cu, 70 °C)={tr(K['kappa'], 0)} m/Ω·mm², X={tr(K['X_mohm_m'], 2)} mΩ/m; Iz değerleri IEC 60364-5-52 Ek B (PVC, Cu) yaklaşık; azaltma k={tr(K['azaltma_k'], 2)} (pano-alıcı), {tr(K['azaltma_k_gomulu'], 2)} (gömülü).",
         "Gerçek döşeme koşulları (gruplama, sıcaklık, toprak ısıl direnci) ve üretici katalogları ile doğrulanmalıdır."),
        ("Güç faktörü", f"Hat bazında verilmeyen yüklerde cosφ = {tr(S['cos_varsayilan'], 2)} (talimat); avize için cosφ = {tr(P['aydinlatma_priz']['avize']['cos'], 2)}; acil/yangın hatlarında 0,95.",
         f"Avize: 2000 W / (230 V × 0,95) = {tr(hat_['AV1']['I'], 1)} A. Talimattaki ≈ 8,7 A değeri cosφ = 1 karşılığıdır; C16 her iki durumda yeterlidir. Önceki sürümdeki klima cosφ 0,80 varsayımı 0,90'a çekilmiştir."),
        ("Klima yükü", f"6 adet x {tr(h0['p_kw'])} kW elektrik giriş gücü, cosφ={tr(h0['cos'], 2)}, In={tr(h0['I'], 1)} A; MPCB ayarı {C['K_klima']['ayar_A']} A (alternatif C20).",
         "Klima plaka verisi yok. Kompresör yolverme akımı ΔU/seçicilik hesabına dahil edilmedi; üretici verisi ile kontrol edilmelidir."),
        ("Priz", f"{A_['priz']['adet']} adet priz girişi (16 A, 230 V, topraklı), priz başına {tr(A_['priz']['birim_w'], 0)} W (parametre), kd={tr(A_['priz']['kd'], 1)}; {'+'.join(str(x) for x in A_['priz']['dagilim'])} dağılımı ile PZ1-PZ3 (RCBO C16 30 mA); toplam kablo {tr(A_['priz']['toplam_kablo_m'], 0)} m (kullanıcı verisi) → hat başına {tr(hat_['PZ1']['L'], 0)} m, NYY 3x{kes_txt(hat_['PZ1']['kesit'])}.",
         f"Toplam kurulu {tr(A_['priz']['adet'] * A_['priz']['birim_w'] / 1000, 1)} kW. Priz başına {tr(A_['priz']['birim_w'], 0)} W tahmini yüktür. ΔU'da yük hat boyunca yayılı kabul edilmiştir (eşdeğer boy katsayısı {tr(A_['priz']['du_yayilim'], 1)}); ΔU şartı için kesit {kes_txt(A_['priz']['kesit'])} → {kes_txt(hat_['PZ1']['kesit'])} mm² büyütülmüştür. Gerçek güzergâh boyları ve priz yerleri ile doğrulanmalıdır."),
        ("Büyük avize", f"{A_['avize']['adet']} adet, {tr(A_['avize']['w'], 0)} W; ayrı hat AV1 (RCBO C16 30 mA, NYY 3x{kes_txt(hat_['AV1']['kesit'])}), kablo boyu {tr(A_['avize']['toplam_kablo_m'], 0)} m (kullanıcı verisi); faz parametresi {A_['avize']['faz']}; dimmer/kontaktör (DM1) bölümü {'gösterilmiştir' if A_['avize']['dimmer_goster'] else 'gösterilmemiştir'}.",
         f"Avizeye {tr(A_['avize']['toplam_kablo_m'], 0)} m kablo çekileceği için gerilim düşümü şartıyla kesit {kes_txt(A_['avize']['kesit'])} → {kes_txt(hat_['AV1']['kesit'])} mm² büyütülmüştür (ΔU toplam %{tr(hat_['AV1']['du_toplam'], 2)}). Avize tipi (halojen/LED) ve dimmer uyumu üretici verisiyle doğrulanmalıdır; LED sürücülerde kalkış akımı için dimmer/RCBO seçimi gözden geçirilmelidir."),
        ("İç mekan LED", f"{A_['led_ic']['adet']} adet x {tr(A_['led_ic']['birim_w'], 0)} W, {'+'.join(str(x) for x in A_['led_ic']['dagilim'])} lamba dağılımı ile LD1-LD3 (RCBO C10 30 mA); toplam kablo {tr(A_['led_ic']['toplam_kablo_m'], 0)} m (kullanıcı verisi) → hat başına {tr(hat_['LD1']['L'], 0)} m, NYY 3x{kes_txt(hat_['LD1']['kesit'])}.",
         "LED sürücü kalkış akımı (inrush) için C tipi eğri kullanıldı. ΔU ≤ %2 sağlanmazsa kesit 2,5 mm²'ye otomatik büyütülür (bu hesapta: " + ("büyütüldü" if any(h.get('kesit_artirildi') and h['grup'] == 'LED' for h in M['hatlar']) else "büyütme gerekmedi") + ")."),
        ("Dış projektör", f"{A_['projektor']['adet']} adet LED projektör; birim güç PROJEKTOR_GUCU_W = {tr(A_['projektor']['PROJEKTOR_GUCU_W'], 0)} W (projektör gücü belirtilmediği için VARSAYIM); {'+'.join(str(x) for x in A_['projektor']['dagilim'])} dağılımı ile PJ1-PJ3; toplam {tr(A_['projektor']['adet'] * A_['projektor']['PROJEKTOR_GUCU_W'] / 1000, 2)} kW.",
         f"Dış hat toplam kablo metrajı {tr(A_['projektor']['toplam_kablo_m'], 0)} m (kullanıcı verisi) → hat başına PROJEKTOR_MESAFE_M = {tr(A_['projektor']['PROJEKTOR_MESAFE_M'], 0)} m. ΔU ≤ %3 sağlanmazsa kesit otomatik büyütülür (bu hesapta: " + ("büyütüldü" if any(h.get('kesit_artirildi') and h['grup'] == 'PROJ' for h in M['hatlar']) else "büyütme gerekmedi") + "). Gerçek projektör gücü ve güzergâh uzunluğu teyit edilmelidir."),
        ("Dış aydınlatma kumandası", "Q3 (MCB 4P C25) → KM1 (3P 25 A kontaktör); SA3 MAN-0-OTO anahtarı, FT1 fotosel ve ZR1 astronomik zaman rölesi (OTO'da FT1 VE ZR1 kontakları seri); H4 durum lambası; SPD2 Tip 2 parafudr " + ("gösterilmiştir" if A_['projektor']['spd2_goster'] else "gösterilmemiştir") + ".",
         "KM1'in AC-5a (LED) yük kategorisi ve SPD2 yedek koruması (Q3 ≤ 32 A) üretici verisiyle doğrulanmalıdır. Dış hatlar toprak altı/kablo kanalında, IP65 bağlantı kutuları ile (JB1-JB3)."),
        ("Dekoratif aydınlatma", f"Minare/kubbe dekoratif aydınlatması: {tr(A_['dekoratif']['kw'], 2)} kW (varsayılan 0 = yok).",
         "Eklenirse DK1 hattı otomatik oluşur; yük dengesi ve pano modülleri yeniden hesaplanır."),
        ("Eşzamanlılık", f"Klima 0,95; avize {tr(A_['avize']['kd'], 1)}; priz {tr(A_['priz']['kd'], 1)}; iç LED {tr(A_['led_ic']['kd'], 1)}; projektör {tr(A_['projektor']['kd'], 1)}; kamera 1,00; ses 0,90; acil/yangın 1,00.",
         f"Talep gücü {tr(M['p_talep'])} kW, kurulu {tr(M['p_kurulu'])} kW. Faz dengesizliği %{tr(M['dengesizlik'])}."),
        ("Faz ataması", "Tek faz yüklerin fazları açgözlü (greedy) algoritma ile otomatik atanır (avize fazı parametre: " + A_['avize']['faz'] + "); aynı gruptaki priz/LED/projektör hatları farklı fazlara bağlanır.",
         "Atama adımları 04_hesap_tablolari.xlsx → Yük Dengesi sayfasındadır."),
        ("Hesap akımı ve gelecek payı", f"Ib = en yüklü faz akımı = {tr(M['Ib'], 1)} A (talep görünür gücü {tr(M['S_talep_kva'], 1)} kVA). Q0 Ir={M['Ir']} A; Ib×{tr(1 + S['gelecek_payi_yuzde'] / 100, 2)} = {tr(M['Ib_gelecek'], 0)} A {'≤' if M['gelecek_uygun'] else '>'} Ir (%{tr(S['gelecek_payi_yuzde'], 0)} gelecek payı {'sağlanıyor' if M['gelecek_uygun'] else 'SAĞLANMIYOR'}).",
         "Boyler/ısıtıcı eklenirse yeniden hesaplanır."),
        ("Hat uzunlukları", f"Trafo-pano {tr(S['trafo_pano_mesafe_m'], 0)} m; pano-klima {', '.join(L(['K1']))} m; avize {', '.join(L(['AV1']))} m; priz {', '.join(L(['PZ1']))} m; iç LED {', '.join(L(['LD1']))} m; projektör {', '.join(L(['PJ1']))} m; kamera {', '.join(L(['KR1']))} m; ses kolları {', '.join(L([h['id'] for h in M['hatlar'] if h['grup'] == 'SES']))} m; acil/yangın {', '.join(L(['AC1']))} m.",
         "Trafo-pano ve pano-klima mesafeleri ile priz, iç LED, avize, projektör ve ses kolu kablo metrajları verilmiştir; kamera ve acil/yangın mesafeleri tahmindir. "
         "Priz, iç LED ve projektör için hat başına boy, toplam metrajın hat sayısına bölünmesiyle bulunmuştur."),
        ("Ses sistemi", f"{len(A_['ses']['kollar'])} kol ({', '.join(k['ad'] + ' ' + tr(k['L'], 0) + ' m' for k in A_['ses']['kollar'])}); toplam {tr(A_['ses']['toplam_kw'], 1)} kW kollara eşit bölünmüş "
         f"({tr(A_['ses']['toplam_kw'] / len(A_['ses']['kollar']) * 1000, 0)} W/kol, kd={tr(A_['ses']['kd'], 2)}); koruma: {'; '.join(k_['ad'] + ' kol → ' + ('ortak RCD3 (2P 25 A, 30 mA) altında MCB C' + str(A_['ses']['rcbo_In']) if k_.get('koruma') == 'ortak_rcd' else 'ayrı RCBO C' + str(A_['ses']['rcbo_In']) + ' 30 mA') for k_ in A_['ses']['kollar'])} (hatlar {poz_aralik([h['id'] for h in M['hatlar'] if h['grup'] == 'SES'])}). Ortak RCD altındaki kollar aynı faza bağlanır (açgözlü atama); minare kolu ayrı RCBO ile korunur, böylece minare kolundaki bir kaçak/yıldırım kaynaklı arıza iç mekân seslendirmesini kesmez.",
         f"Kol başına yük dağılımı ve amfi/hoparlör güçleri verilmediği için eşit bölme VARSAYIMDIR. Kablo NYY 3x{kes_txt(hat_['SS1']['kesit'])}; ΔU toplam en fazla %{tr(max(hat_[i]['du_toplam'] for i in hat_ if hat_[i]['grup'] == 'SES'), 2)} (limit %{tr(hat_['SS1']['du_limit'], 1)}). Gerçek güçler, amfi yeri ve kol güzergâhları teyit edilmelidir."),
        ("Besleme kablosu", f"NYY 4x95 mm² gömülü (yöntem D2), PE iletkeni ayrı 1x50 mm² Cu (PE3, aynı hendek); W0 için ΔU payı hedefi ≤ %1 (varsayım).",
         "4 damarlı kablo PE içermediğinden TN-S için ayrı PE iletkeni öngörüldü."),
        ("Q0 ve seçicilik", f"Q0: 4P MCCB {C['Q0']['govde_A']} A gövde, Ir={C['Q0']['Ir_A']} A, Icu={tr(C['Q0']['Icu_kA'], 0)} kA (≥25 kA), Isd={C['Q0']['Isd_x']}xIr / {C['Q0']['tsd_ms']} ms, Ii kapalı, Icw≥{tr(C['Q0']['Icw_kA'], 0)} kA (varsayım). RCR0: I∆n={tr(C['RCR0']['I_dn_A'], 1)} A, {C['RCR0']['gecikme_ms']} ms (ayar aralığı 0,3-1 A / 100-300 ms içinden seçildi).",
         "Alt kademe RCD'ler 300 mA ve RCBO 30 mA, anlık (açma süresi 40 ms tipik varsayım). Seçicilik üretici tablolarıyla doğrulanmalıdır."),
        ("RCBO kısa devre kapasitesi", f"RCBO'lar {tr(C['RCBO']['Icn_kA'], 0)} kA seçildi; PN1 barası Ik={tr(M['Ik_pano_kA'], 1)} kA.",
         "6 kA'lık RCBO bu Ik değerinde yetersiz kalırdı."),
        ("Öncelikli hat", "Acil aydınlatma ve yangın alarm hattı Q0 çıkışında, ayrı RCBO ve FE180 kablo ile beslenir; Q0 açıldığında kesilir.",
         "Cihazların dahili aküsü (≥3 saat) ve jeneratör/OG önü ayrı besleme ihtiyacı proje yönetmeliklerine göre netleştirilmelidir."),
        ("Pano yerleşimi", f"Modül 17,5 mm; 1P+N=2, 3P=3, 4P=4 modül; ray boyu {int(PN['ray_uzunlugu_mm'])} mm ({PN['modul_kapasite']} modül); satır adımı {P['pano']['ray_adim_mm']} mm.",
         "Gerçek cihaz boyutları (ör. MPCB 45 mm) marka kataloğuna göre kontrol edilmelidir."),
        ("Isıl hesap", f"Pv≈{tr(PN['isil']['Pv'], 0)} W; Ae={tr(PN['isil']['Ae'], 2)} m², ΔT={tr(PN['isil']['dT'], 0)} K, k={tr(P['pano']['yuzey_isi_katsayisi_W_m2K'])} W/m²K; Pd≈{tr(PN['isil']['Pd'], 0)} W.",
         "Kayıp güçler tipik katalog değerleridir; yaz koşulu için ortam {0:.0f} °C varsayıldı.".format(S["ortam_sicakligi_C"])),
        ("Bina geometrisi", f"Gövde {tr(T['govde_yuksekligi_m'], 0)} m, kubbe tepesi {tr(T['bina_yuksekligi_m'], 0)} m (16 m), kubbe Ø{tr(T['kubbe_cap_m'], 0)} m merkezde, minare {tr(T['minare_taban_m'])}x{tr(T['minare_taban_m'])} m, h={tr(T['minare_yuksekligi_m'], 0)} m KD köşesinde; elektrik odası GB köşesinde.",
         "Mimari proje verilmedi; yerleşim ve yükseklikler varsayımdır."),
        ("Yıldırımdan korunma", f"Seviye {T['ypk_seviye']} (R={tr(T['kure_R'], 0)} m); {len(T['inis'])} iniş iletkeni (ort. {tr(T['inis_ort_aralik'], 1)} m ≤ {tr(T['inis_aralik'], 0)} m); Franklin çubukları YK1 (minare), YK2 (kubbe), YK3 (çatı kenarı) + ÇÇ1 çatı çevre iletkeni hava sonlandırma kabul edildi.",
         f"Yuvarlanan küre 2B kesitte sayısal hesaplandı; en büyük sapma {tr(max(0.0, T['yuvarlanan_kure']['ihlal_m']), 2)} m (ağ toleransı 0,25 m). Risk analizi (IEC 62305-2) yapılmadı."),
        ("Topraklama", f"Halka {tr(T['halka_cevre'], 0)} m ({T['halka_iletkeni']}), {len(T['cubuklar'])} çubuk Ø{tr(T['cubuk_cap_mm'], 0)}x{tr(T['cubuk_boy_m'], 0)} m, {len(T['cukurlar'])} kontrol çukuru; trafo için ayrı TE1 elektrodu.",
         "Zemin özdirenci bilinmiyor; ≤1-2 Ω hedefi hesapla garanti edilemez, sahada ölçülmeli (kuru mevsimde tekrar) ve gerekirse ek çubuk/zemin iyileştirme yapılmalıdır."),
        ("Metraj", f"Halka net {tr(mt['halka'], 0)} m, sipariş {mt['halka_sip']} m (%{tr(T['yedek_pay_yuzde'], 0)} fire); {mt['ek']} eksotermik kaynak.",
         "Güzergâh uzunlukları (PE1, PE2, EB bağlantıları) yaklaşıktır."),
    ]


def rapor_yaz(P, M, PN, T, KAYIT, dog, svgler, yol, gecici):
    _fontlar()
    pr = P["proje"]
    st = {
        "n": ParagraphStyle("n", fontName="LS", fontSize=8.6, leading=11.2),
        "s": ParagraphStyle("s", fontName="LS", fontSize=7.4, leading=9.2),
        "b": ParagraphStyle("b", fontName="LS-B", fontSize=8.6, leading=11.2),
        "h1": ParagraphStyle("h1", fontName="LS-B", fontSize=13, leading=16, spaceBefore=6, spaceAfter=5, textColor=colors.HexColor("#1f3a5f")),
        "h2": ParagraphStyle("h2", fontName="LS-B", fontSize=10, leading=13, spaceBefore=6, spaceAfter=3, textColor=colors.HexColor("#1f3a5f")),
        "k": ParagraphStyle("k", fontName="LS-B", fontSize=24, leading=30, alignment=TA_CENTER),
        "k2": ParagraphStyle("k2", fontName="LS", fontSize=13, leading=18, alignment=TA_CENTER),
        "w": ParagraphStyle("w", fontName="LS-B", fontSize=9, leading=12, textColor=colors.HexColor("#8a1c1c")),
    }
    P_ = lambda t, s="n": Paragraph(t, st[s])
    E = _esc

    def tablo(veri, genisler, baslik=True, fs=7.4, zebra=True, hizala=None):
        satirlar = []
        for i, r in enumerate(veri):
            satirlar.append([Paragraph(E(str(c)).replace("✓", '<font name="DVS">✓</font>').replace("✗", '<font name="DVS">✗</font>'), ParagraphStyle("t", fontName="LS-B" if (baslik and i == 0) else "LS", fontSize=fs, leading=fs + 1.8)) for c in r])
        t = Table(satirlar, colWidths=[g * mm for g in genisler], repeatRows=1 if baslik else 0)
        stl = [("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#999999")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
               ("TOPPADDING", (0, 0), (-1, -1), 1.6), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6), ("LEFTPADDING", (0, 0), (-1, -1), 2.5), ("RIGHTPADDING", (0, 0), (-1, -1), 2.5)]
        if baslik:
            stl.append(("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E2F3")))
        if zebra:
            for i in range(1, len(veri)):
                if i % 2 == 0:
                    stl.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F5F7FA")))
        t.setStyle(TableStyle(stl))
        return t

    def sayfa_ust_alt(cv, doc):
        cv.saveState()
        cv.setFont("LS", 7.2)
        cv.setFillColor(colors.HexColor("#555555"))
        if doc.page > 1:
            cv.drawString(15 * mm, 287 * mm, f"{pr['ad']} — Proje Raporu — {pr['sistem_yazisi']}")
            cv.drawRightString(195 * mm, 287 * mm, f"Rev {pr['revizyon']} / {pr['tarih']}")
            cv.line(15 * mm, 285.5 * mm, 195 * mm, 285.5 * mm)
            cv.drawString(15 * mm, 9 * mm, "Uygulamadan önce yetkili elektrik mühendisi ve yerel şebeke (dağıtım şirketi) şartnamesi ile onaylanmalıdır.")
            cv.drawRightString(195 * mm, 9 * mm, f"Sayfa {doc.page}")
        cv.restoreState()

    doc = BaseDocTemplate(yol, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=17 * mm, bottomMargin=16 * mm,
                          title=f"{pr['ad']} — Proje Raporu", author=pr["cizen"])
    doc.addPageTemplates([PageTemplate(id="p", frames=[Frame(15 * mm, 16 * mm, 180 * mm, 264 * mm, id="f", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)],
                                       onPage=sayfa_ust_alt)])
    S_ = []
    # ------------------------------------------------ kapak
    S_.append(Spacer(1, 45 * mm))
    S_.append(P_(E(pr["ad"].upper()), "k"))
    S_.append(Spacer(1, 6 * mm))
    S_.append(P_("ELEKTRİK DAĞITIM PROJESİ — PROJE RAPORU", "k2"))
    S_.append(P_(E(f"{pr['sistem_yazisi']}, {P['sistem']['frekans']} Hz  |  Kaynak: {tr(P['sistem']['trafo_kva'], 0)} kVA trafo"), "k2"))
    S_.append(Spacer(1, 18 * mm))
    kv = [["Konu", "Bilgi"],
          ["Bina", f"Cami, {tr(T['bina_a_m'], 0)} m x {tr(T['bina_b_m'], 0)} m taban, {tr(T['bina_yuksekligi_m'], 0)} m yükseklik, kubbe ve minare"],
          ["Kurulu / talep gücü", f"{tr(M['p_kurulu'])} kW / {tr(M['p_talep'])} kW; hesap akımı Ib = {tr(M['Ib'], 1)} A"],
          ["Ana pano", f"PN1, {PN['G']}x{PN['Y']}x{P['pano']['D']} mm, {P['pano']['IP']}; Q0 MCCB 4P {P['cihazlar']['Q0']['govde_A']} A, Ir={P['cihazlar']['Q0']['Ir_A']} A"],
          ["Çizim no / revizyon", f"{pr['cizim_no'][1]}, {pr['cizim_no'][2]}, {pr['cizim_no'][3]}, {pr['cizim_no'][6]}  /  {pr['revizyon']}"],
          ["Tarih", pr["tarih"]], ["Çizen", pr["cizen"]], ["Kontrol", pr["kontrol"]]]
    S_.append(tablo(kv, [48, 132], fs=9))
    S_.append(Spacer(1, 10 * mm))
    S_.append(P_("Paket içeriği: 01 Tek hat şeması, 02 Pano ön görünüş ve yerleşim, 03 Topraklama ve yıldırımdan korunma, 06 Aydınlatma ve priz yerleşim planı (A3 yatay, SVG+PDF); 04 Hesap tabloları (xlsx); 05 Bu rapor.", "n"))
    S_.append(Spacer(1, 8 * mm))
    S_.append(P_("DİKKAT: Bu paket, verilen veriler ve açıkça belirtilen varsayımlarla otomatik üretilmiş bir ön/uygulama taslağıdır. Uygulamadan önce yetkili elektrik mühendisi ve yerel şebeke (dağıtım şirketi) şartnamesi ile onaylanmalıdır.", "w"))
    S_.append(PageBreak())

    # ------------------------------------------------ 1 sistem verileri
    S_.append(P_("1. Sistem verileri ve kapsam", "h1"))
    sv = [["Parametre", "Değer"],
          ["Sistem", f"{P['sistem']['U_LL']:.0f}/{P['sistem']['U_LN']:.0f} V, {P['sistem']['frekans']} Hz, {P['sistem']['sebeke']}"],
          ["Kaynak", f"{tr(P['sistem']['trafo_kva'], 0)} kVA, {P['sistem']['trafo_grubu']} (OG tarafı blok)"],
          ["Trafo - pano mesafesi", f"{tr(P['sistem']['trafo_pano_mesafe_m'])} m (parametre)"],
          ["Pano - klima mesafeleri", f"{tr(M['hatlar'][0]['L'])} m (parametre)"],
          ["Besleme kablosu", f"{M['besleme']['tip']} {M['besleme']['damar']} mm² Cu, {M['besleme']['kablo_no']}"],
          ["Kısa devre", f"Trafo uçları Ik''={tr(M['Ik_trafo_kA'], 2)} kA; PN1 barası Ik''={tr(M['Ik_pano_kA'], 2)} kA"]]
    S_.append(tablo(sv, [60, 120], fs=8.2))

    # ------------------------------------------------ 2 varsayımlar
    S_.append(P_("2. Varsayımlar", "h1"))
    S_.append(P_("Aşağıdaki maddeler girdi verisinde bulunmayan ve hesaplarda kullanılan varsayımlardır. Hiçbiri sessizce uydurulmamıştır; tamamı PARAMETRELER sözlüğünde değiştirilebilir.", "n"))
    S_.append(Spacer(1, 2 * mm))
    vr = [["No", "Konu", "Varsayım / değer", "Not / doğrulama"]]
    for i, (a_, b_, c_) in enumerate(varsayimlar(P, M, PN, T), start=1):
        vr.append([i, a_, b_, c_])
    S_.append(tablo(vr, [8, 28, 76, 68], fs=7.0))
    S_.append(PageBreak())

    # ------------------------------------------------ 3 hesap özeti
    S_.append(P_("3. Hesap özeti", "h1"))
    S_.append(P_("3.1 Yük dengesi", "h2"))
    yb = [["Poz", "Hat", "Kurulu (kW)", "kd", "Talep (kW)", "Faz", "I (A)"]]
    for h in M["hatlar"]:
        yb.append([h["id"], h["ad"], tr(h["p_kw"]), tr(h["kd"], 2), tr(h["p_talep"], 2), faz_adi(h), tr(h["I"], 1)])
    yb.append(["", "TOPLAM", tr(M["p_kurulu"]), "", tr(M["p_talep"], 2), "", ""])
    S_.append(tablo(yb, [14, 62, 24, 14, 24, 24, 18], fs=7.4))
    S_.append(Spacer(1, 2 * mm))
    S_.append(P_(E(f"Faz toplamları (talep): L1={tr(M['faz_kw']['L1'], 2)} kW, L2={tr(M['faz_kw']['L2'], 2)} kW, L3={tr(M['faz_kw']['L3'], 2)} kW. "
                   f"Faz dengesizliği = (Pmax−Pmin)/Port = %{tr(M['dengesizlik'], 2)} (hedef ≤ %10: {'UYGUN' if M['dengesizlik'] <= 10 else 'UYGUN DEĞİL'}). "
                   f"Faz akımları: L1={tr(M['faz_I']['L1'], 1)} A, L2={tr(M['faz_I']['L2'], 1)} A, L3={tr(M['faz_I']['L3'], 1)} A; Ib={tr(M['Ib'], 1)} A; cosφ(ort.)={tr(M['cos_ort'], 2)}; Q0 Ir={P['cihazlar']['Q0']['Ir_A']} A ≥ Ib."), "n"))
    S_.append(P_(E(f"Gelecek payı: Ib × {tr(1 + P['sistem']['gelecek_payi_yuzde'] / 100, 2)} = {tr(M['Ib_gelecek'], 1)} A {'≤' if M['gelecek_uygun'] else '>'} Q0 Ir = {M['Ir']} A → "
                   f"%{tr(P['sistem']['gelecek_payi_yuzde'], 0)} gelecek payı {'SAĞLANIYOR' if M['gelecek_uygun'] else 'SAĞLANMIYOR'}; Q0'ın kalan yedek kapasitesi %{tr(M['yedek_yuzde'], 1)} ({tr(M['yedek_A'], 1)} A). "
                   f"Talep görünür gücü {tr(M['S_talep_kva'], 1)} kVA (trafo yüklenmesi %{tr(M['S_talep_kva'] / P['sistem']['trafo_kva'] * 100, 1)})."), "n"))
    S_.append(Spacer(1, 1.5 * mm))
    ev = [["Hat", "Elemanlar", "Adet", "Birim (W)", "Toplam (W)", "Faz", "Kablo"]]
    for h in M["hatlar"]:
        if h.get("elemanlar"):
            lst = [e for e in M["elemanlar"] if e["hat"] == h["id"]]
            ev.append([h["id"], poz_aralik([e["poz"] for e in lst]), len(lst), tr(lst[0]["w"], 0), tr(sum(e["w"] for e in lst), 0), h["faz"], f"{h['kablo_no']} {h['tip']} {h['damar']}"])
    S_.append(P_("Aydınlatma ve priz envanteri (Çizim 6 ile aynı pozlar):", "b"))
    S_.append(tablo(ev, [14, 34, 14, 20, 24, 14, 60], fs=7.2))
    S_.append(P_("3.2 Kablo ve gerilim düşümü özeti", "h2"))
    kb = [["Kablo", "Hat", "Kablo", "L (m)", "I (A)", "In (A)", "Iz* (A)", "ΔU toplam %", "Limit %", "Sonuç"]]
    b0 = M["besleme"]
    kb.append([b0["kablo_no"], "Q0", f"{b0['tip']} {b0['damar']}", tr(b0["L"], 0), tr(M["Ib"], 1), P["cihazlar"]["Q0"]["Ir_A"], tr(b0["iz"], 0), tr(b0["du"], 2), "1,00", "UYGUN" if b0["du"] <= 1.0 else "UYGUN DEĞİL"])
    for h in M["hatlar"]:
        kb.append([h["kablo_no"], h["id"], f"{h['tip']} {h['damar']}", tr(h["L"], 0), tr(h["I"], 1), h["In_koruma"], tr(h["iz"], 1), tr(h["du_toplam"], 2), tr(h["du_limit"], 2),
                   "UYGUN" if h["uygun"] else "UYGUN DEĞİL"])
    S_.append(tablo(kb, [12, 12, 36, 12, 14, 14, 15, 24, 16, 25], fs=7.0))
    S_.append(P_("* Iz: düzeltilmiş akım taşıma kapasitesi. ΔU toplam = hat + besleme (W0). Ayrıntı: 04_hesap_tablolari.xlsx → Kablo Çizelgesi.", "s"))
    S_.append(P_("3.3 Kısa devre ve seçicilik", "h2"))
    rc, rb_, rcd = P["cihazlar"]["RCR0"], P["cihazlar"]["RCBO"], P["cihazlar"]["RCD_klima"]
    sl = [["Kontrol", "Değer", "Sonuç"],
          ["PN1 barası Ik''", f"{tr(M['Ik_pano_kA'], 2)} kA", ""],
          ["Q0 Icu ≥ 25 kA", f"{tr(P['cihazlar']['Q0']['Icu_kA'], 0)} kA", "UYGUN" if P["cihazlar"]["Q0"]["Icu_kA"] >= 25 else "UYGUN DEĞİL"],
          ["Q0 Icu ≥ Ik", f"{tr(P['cihazlar']['Q0']['Icu_kA'], 0)} ≥ {tr(M['Ik_pano_kA'], 1)} kA", "UYGUN" if P["cihazlar"]["Q0"]["Icu_kA"] >= M["Ik_pano_kA"] else "UYGUN DEĞİL"],
          ["RCBO Icn ≥ Ik", f"{tr(rb_['Icn_kA'], 0)} ≥ {tr(M['Ik_pano_kA'], 1)} kA", "UYGUN" if rb_["Icn_kA"] >= M["Ik_pano_kA"] else "UYGUN DEĞİL"],
          ["RCR0 → RCD1-2: I∆n oranı", f"{tr(rc['I_dn_A'] * 1000, 0)}/{rcd['I_dn_mA']} = {tr(rc['I_dn_A'] * 1000 / rcd['I_dn_mA'], 1)} (≥3)", "UYGUN" if rc["I_dn_A"] * 1000 / rcd["I_dn_mA"] >= 3 else "UYGUN DEĞİL"],
          ["RCR0 → RCBO: I∆n oranı", f"{tr(rc['I_dn_A'] * 1000, 0)}/{rb_['I_dn_mA']} = {tr(rc['I_dn_A'] * 1000 / rb_['I_dn_mA'], 1)} (≥3)", "UYGUN"],
          *([["RCR0 → RCD3 (ses ortak): I∆n oranı", f"{tr(rc['I_dn_A'] * 1000, 0)}/{P['cihazlar']['RCD3']['I_dn_mA']} = {tr(rc['I_dn_A'] * 1000 / P['cihazlar']['RCD3']['I_dn_mA'], 1)} (≥3)", "UYGUN"]]
            if any(h.get("koruma_tur") == "mcb" for h in M["hatlar"]) else []),
          ["RCR0 gecikmesi", f"{rc['gecikme_ms']} ms ≥ {rb_['t_acma_ms']} ms + {P['cihazlar']['secicilik']['zaman_payi_ms']} ms", "UYGUN" if rc["gecikme_ms"] >= rb_["t_acma_ms"] + P["cihazlar"]["secicilik"]["zaman_payi_ms"] else "UYGUN DEĞİL"],
          ["Q0 tsd / Icw", f"{P['cihazlar']['Q0']['tsd_ms']} ms; Icw={tr(P['cihazlar']['Q0']['Icw_kA'], 0)} kA ≥ Ik", "UYGUN" if P["cihazlar"]["Q0"]["Icw_kA"] >= M["Ik_pano_kA"] else "UYGUN DEĞİL"]]
    S_.append(tablo(sl, [70, 80, 30], fs=7.6))
    S_.append(P_("3.4 Pano yerleşimi ve ısıl kontrol", "h2"))
    pl = [["Satır", "Gruplar", "Modül", "Kapasite", "Doluluk"]]
    for i, sat in enumerate(PN["satir_hesap"], start=1):
        pl.append([i, " + ".join(g for g, _ in sat["gruplar"]), sat["modul"], sat["kapasite"], f"%{tr(100.0 * sat['modul'] / sat['kapasite'], 0)}"])
    pl.append(["Üst", " + ".join(PN["ust_ray_grup"]), PN["ust_ray_modul_toplam"], PN["ust_ray_kapasite"], f"%{tr(100.0 * PN['ust_ray_modul_toplam'] / PN['ust_ray_kapasite'], 0)}"])
    S_.append(tablo(pl, [16, 90, 24, 24, 26], fs=7.6))
    isl = PN["isil"]
    S_.append(Spacer(1, 1.5 * mm))
    S_.append(P_(E(f"Toplam kayıp güç Pv ≈ {tr(isl['Pv'], 0)} W; pano yüzeyi Ae={tr(isl['Ae'], 2)} m², ΔT={tr(isl['dT'], 0)} K → doğal ısı atımı Pd≈{tr(isl['Pd'], 0)} W. "
                   + ("Pv > 0,8·Pd: fan gereklidir (≥ %s m³/h)." % tr(isl["hava_m3h"], 0) if isl["fan_gerekli"] else "Pv < 0,8·Pd: ek fan gerekmez; IP54 filtreli havalandırma ızgarası (GR1) yeterlidir. Ortam sıcaklığı > 35 °C ise termostatlı fan eklenmelidir.")), "n"))
    if PN["uyarilar"]:
        S_.append(P_(E("UYARI: " + " ".join(PN["uyarilar"])), "w"))
    S_.append(P_("3.5 Topraklama ve yıldırımdan korunma", "h2"))
    mt = topraklama_metraj(P, T)
    tm = [["Kalem", "Miktar", "Not"],
          ["Halka iletkeni TH1", f"{tr(mt['halka'], 0)} m (sipariş {mt['halka_sip']} m)", T["halka_iletkeni"] + f", derinlik {tr(T['halka_derinlik_m'])} m, bina dışında {tr(T['halka_ofset_m'])} m"],
          ["Topraklama çubuğu", f"{mt['cubuk']} adet", f"Ø{tr(T['cubuk_cap_mm'], 0)}x{tr(T['cubuk_boy_m'], 0)} m, aralık {tr(T['cubuk_aralik_min_gercek'], 1)}-{tr(T['cubuk_aralik_gercek'], 1)} m"],
          ["Kontrol çukuru", f"{mt['cukur']} adet", ", ".join(c['poz'] + "→" + c['cubuk'] for c in T['cukurlar'])],
          ["İniş iletkeni / ölçüm ayırma", f"{mt['inis']} / {mt['ok']} adet", f"toplam {mt['inis_m']} m; ort. aralık {tr(T['inis_ort_aralik'], 1)} m"],
          ["Çatı çevre iletkeni ÇÇ1", f"{mt['cati']} m", T["cati_iletkeni"]],
          ["Franklin çubuğu", "3 adet (YK1-YK3)", "minare ucu, kubbe tepesi, çatı kenarı"],
          ["Eksotermik kaynak", f"{mt['ek']} adet", "çubuk, iniş, PE1/PE2, TE1, EB bağlantıları"],
          ["Yuvarlanan küre", f"Seviye {T['ypk_seviye']}, R={tr(T['kure_R'], 0)} m", "Yapı korunan hacim içinde: " + ("EVET" if T["yuvarlanan_kure"]["korunuyor"] else "HAYIR")]]
    S_.append(tablo(tm, [44, 52, 84], fs=7.4))

    # ------------------------------------------------ 4 küçük görüntüler
    S_.append(P_("4. Çizimlerin küçük görüntüleri", "h1"))
    ad_ = {1: "Çizim 1 — Tek hat şeması (01_tek_hat_semasi)", 2: "Çizim 2 — Pano ön görünüş ve yerleşim (02_pano_on_gorunus)", 3: "Çizim 3 — Topraklama ve yıldırımdan korunma (03_topraklama_ve_yildirimdan_korunma)",
           6: "Çizim 6 — Aydınlatma ve priz yerleşim planı (06_aydinlatma_priz_yerlesim)"}
    for n in (1, 2, 3, 6):
        png = os.path.join(gecici, f"kucuk{n}.png")
        cairosvg.svg2png(url=svgler[n], write_to=png, output_width=1900)
        from PIL import Image as PILImage
        w_, h_ = PILImage.open(png).size
        img = Image(png, width=180 * mm, height=180 * mm * h_ / w_)
        blok = [P_(E(ad_[n]), "h2"), img, Spacer(1, 3 * mm)]
        S_.append(KeepTogether(blok))
        if n == 2:
            S_.append(PageBreak())
    # ------------------------------------------------ 5 notlar
    S_.append(P_("5. Notlar ve doğrulanması gereken noktalar", "h1"))
    notlar = [
        "Uygulamadan önce yetkili elektrik mühendisi ve yerel şebeke (dağıtım şirketi) şartnamesi ile onaylanmalıdır; bu paket tek başına uygulama onayı yerine geçmez.",
        "Trafo uk/ur değerleri, Q0 Icw/Isd ayarları ve RCD/RCBO açma süreleri üretici kataloğu ile doğrulanmalı; Q0→Q1/Q2→RCBO seçiciliği üretici tablolarıyla teyit edilmelidir.",
        "Klima kompresör yolverme akımı, motor koruma ayarı (11-16 A aralığında 14 A) ve alternatif C20 seçeneği klima plaka verisiyle netleştirilmelidir.",
        "Topraklama direnci (≤1-2 Ω) sahada ölçülmeli, kuru mevsimde tekrar ölçülmelidir; zemin özdirenci bilinmediği için ek çubuk/derin elektrot gerekebilir.",
        "Yıldırımdan korunma için IEC 62305-2 risk değerlendirmesi ve yapı mimari verileri ile koruma seviyesi (III/IV) onaylanmalıdır.",
        "Acil aydınlatma/yangın alarm hatlarının besleme önceliği, akü süreleri ve yangına dayanım sınıfı (FE180/E90) ilgili yangın yönetmeliğine göre teyit edilmelidir.",
        "Gaz tesisatında izolasyon flanşı (İF1) ve gaz boru bonding şartları gaz dağıtım şirketi ile netleştirilmelidir.",
        "Boyler/ısıtıcı ve benzeri ek yüklerde yük dengesi, kablo, kesici ve pano ısıl hesabı yeniden yapılmalıdır.",
        "Hesaplar 5 sayfalık xlsx içinde formüllüdür; PARAMETRELER değişince `python3 proje_uret.py` ile tüm paket yeniden üretilir.",
    ]
    for i, n_ in enumerate(notlar, start=1):
        S_.append(P_(E(f"{i}. {n_}"), "n"))
        S_.append(Spacer(1, 1 * mm))
    S_.append(P_("Poz numarası tutarlılığı (otomatik doğrulama)", "h2"))
    S_.append(P_("Her poz için ✓: ilgili çizimde etiket olarak bulundu / tabloda satır olarak bulundu. “–”: ilgili çizimde beklenmiyor.", "s"))
    pm_ = [["Poz", "Açıklama", "Ç1", "Ç2", "Ç3", "Ç6", "Tablo"]]
    kayit_ciz = {}
    for g in satirlari_grupla(KAYIT):
        kayit_ciz[g["poz"]] = g["cizim"]
    for r in dog["matris"]:
        exp = kayit_ciz[r["poz"]]
        row = [r["poz"], r["ad"]]
        for n in (1, 2, 3, 6):
            row.append("✓" if r[f"c{n}"] else ("–" if n not in exp else "✗"))
        row.append("✓" if r["tablo"] else "✗")
        pm_.append(row)
    S_.append(tablo(pm_, [26, 94, 10, 10, 10, 10, 14], fs=6.6))
    S_.append(Spacer(1, 2 * mm))
    S_.append(P_(E("Otomatik denetim sonucu: " + (("tüm kontroller geçti." if not dog["hatalar"] else "; ".join(dog["hatalar"][:6])) + ("" if not dog["bilinen"] else " Bilinen sapma: " + "; ".join(dog["bilinen"])))), "n"))
    doc.build(S_)


# =====================================================================
#  README
# =====================================================================
README = """# Cami Elektrik Dağıtım Projesi — çizim paketi üreteci

400/230 V, 50 Hz, TN-S sistemli, 250 kVA trafo ile beslenen cami binasının elektrik dağıtım projesi.
Tüm çizim, tablo ve rapor tek bir Python dosyasından (`proje_uret.py`) üretilir; tüm girdiler dosyanın
başındaki **`PARAMETRELER`** sözlüğündedir. Bir değer değişince çizimler, tablolar ve rapor birlikte güncellenir
(poz numaraları dört çizimde ve tablolarda aynıdır).

## Üretilen dosyalar (`cikti/`)

| Dosya | İçerik |
|---|---|
| `01_tek_hat_semasi.svg/.pdf` | Tek hat şeması (A3 yatay, IEC 60617 sembolleri, faz renkleri, çıkış tabloları, antet) |
| `02_pano_on_gorunus.svg/.pdf` | Pano ön görünüş + iç yerleşim (1:10), DETAY-A (1:5), DIN ray modül hesabı, ısıl not, poz listesi |
| `03_topraklama_ve_yildirimdan_korunma.svg/.pdf` | Kat planı (1:150), bina kesiti (1:200), yuvarlanan küre (1:500), detaylar, EPB şeması |
| `04_hesap_tablolari.xlsx` | Yük Dengesi (gelecek payı, açgözlü faz ataması), Kablo Çizelgesi (ΔU formüllü), Pano Malzeme Listesi, Aydınlatma ve Priz Envanteri, Topraklama ve Yıldırım Malzeme, Kesici Selektivite (formüllü) |
| `05_proje_raporu.pdf` | Kapak, varsayımlar, hesap özeti, çizimlerin küçük görüntüsü, notlar, poz tutarlılık matrisi |
| `06_aydinlatma_priz_yerlesim.svg/.pdf` | Aydınlatma ve priz yerleşim planı (A3 yatay, 1:100): priz P, iç LED L, avize AVZ, dış projektör PR, kablo tavası, hat/faz renk tablosu |

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

## Önemli not

Bu paket verilen veri ve raporun “Varsayımlar” bölümündeki kabullerle üretilmiştir. Uygulamadan önce **yetkili elektrik
mühendisi ve yerel şebeke (dağıtım şirketi) şartnamesi ile onaylanmalıdır.**
"""


# =====================================================================
#  ANA AKIŞ
# =====================================================================
def main():
    P = hazirla(PARAMETRELER)
    kok = os.path.dirname(os.path.abspath(__file__))
    cikti = os.path.join(kok, P["proje"]["cikti_klasoru"])
    os.makedirs(cikti, exist_ok=True)
    gecici = tempfile.mkdtemp()
    print("1) Hesap modeli ...")
    M = hesapla(P)
    PN = pano_hesapla(P, M)
    T = topraklama_hesapla(P)
    KAYIT = poz_kayit(P, M, PN, T)
    ad = {1: "01_tek_hat_semasi", 2: "02_pano_on_gorunus", 3: "03_topraklama_ve_yildirimdan_korunma", 6: "06_aydinlatma_priz_yerlesim"}
    svgler = {n: os.path.join(cikti, a + ".svg") for n, a in ad.items()}
    pdfler = {}
    print("2) Çizimler ...")
    ciz = {1: cizim1(P, M, PN, T, svgler[1]), 2: cizim2(P, M, PN, T, KAYIT, svgler[2]), 3: cizim3(P, M, PN, T, KAYIT, svgler[3]),
           6: cizim6(P, M, PN, T, KAYIT, svgler[6])}
    for n, c in ciz.items():
        pdfler[ad[n] + ".pdf"] = c.kaydet()
        for x in c.cakismalar():
            uyar(f"[Çizim {n}] metin çakışması: '{x[0]}' ↔ '{x[1]}' ({x[2]}, {x[3]})")
    print("3) Excel ...")
    xlsx = os.path.join(cikti, "04_hesap_tablolari.xlsx")
    excel_yaz(P, M, PN, T, KAYIT, xlsx)
    excel_yeniden_hesapla(xlsx)
    print("4) Doğrulama ...")
    dog = dogrula_hepsi(P, M, PN, T, KAYIT, svgler, xlsx, pdfler)
    print("5) Rapor ...")
    rapor = os.path.join(cikti, "05_proje_raporu.pdf")
    rapor_yaz(P, M, PN, T, KAYIT, dog, svgler, rapor, gecici)
    from pypdf import PdfReader
    n_rapor = len(PdfReader(rapor).pages)
    dog["pdf"].append(("05_proje_raporu.pdf", n_rapor, 210, 297, None))
    for dosya in (os.path.join(cikti, "README.md"), os.path.join(kok, "README.md")):
        with open(dosya, "w", encoding="utf-8") as f:
            f.write(README)
    shutil.rmtree(gecici, ignore_errors=True)

    # ---------------- özet
    print("\n=== ÖZET ===")
    print(f"Kurulu {tr(M['p_kurulu'])} kW | talep {tr(M['p_talep'])} kW | Ib {tr(M['Ib'], 1)} A | dengesizlik %{tr(M['dengesizlik'], 1)} | Ik(PN1) {tr(M['Ik_pano_kA'], 2)} kA")
    print(f"Pano {PN['G']}x{PN['Y']}x{P['pano']['D']} mm | Pv {tr(PN['isil']['Pv'], 0)} W | Pd {tr(PN['isil']['Pd'], 0)} W | fan {'GEREKLİ' if PN['isil']['fan_gerekli'] else 'gerekmez'}")
    print(f"Halka {tr(T['halka_cevre'], 0)} m | çubuk {len(T['cubuklar'])} | iniş iletkeni {len(T['inis'])} | küre koruması {'EVET' if T['yuvarlanan_kure']['korunuyor'] else 'HAYIR'}")
    for a_, n_, w_, h_, ok in dog["pdf"]:
        print(f"  PDF {a_}: {n_} sayfa, {w_}x{h_} mm" + ("" if ok is None else (" (A3 yatay ✓)" if ok else " (A3 DEĞİL ✗)")))
    ok_say = sum(1 for x in dog["sayisal"] if x[3])
    print(f"Sayısal doğrulama: {ok_say}/{len(dog['sayisal'])} uyumlu")
    ok_poz = sum(1 for r in dog["matris"] if r["tablo"])
    print(f"Poz doğrulama: {len(dog['matris'])} poz grubu, tablo eşleşmesi {ok_poz}/{len(dog['matris'])}")
    if UYARILAR:
        print(f"Uyarı sayısı: {len(UYARILAR)}")
    for b_ in dog["bilinen"]:
        print("  [BİLİNEN SAPMA]", b_)
    if dog["hatalar"]:
        print("HATALAR:")
        for h in dog["hatalar"]:
            print("  -", h)
        sys.exit(1)
    print("Tüm otomatik kontroller geçti.")


if __name__ == "__main__":
    main()
