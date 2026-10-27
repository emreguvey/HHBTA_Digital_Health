"""Gerçek veriyle AYNI değişken adları ve kodlamayla sahte HHBTA dosyaları üretir.

Amaç: boru hattını gerçek mikro veriye dokunmadan uçtan uca test etmek.
Sayılar anlamsızdır. 2026'da etiket kayması 'HB' senaryosuna göre kurgulanmıştır:
RANDEVU_AL = sağlık bilgisi kopyası, SONUC_AL = MHRS, TIBKONS = kayıt, MALHIZMET_SATIS = konsültasyon.
Çalıştır: python tests/make_synthetic.py
"""
import pathlib

import numpy as np
import pandas as pd

OUT = pathlib.Path(__file__).resolve().parent / "synthetic"
OUT.mkdir(exist_ok=True)


def yn(rng, p):
    return np.where(rng.random(len(p)) < p, 1, 2)


def make_year(yr, rng, n_hh=2600):
    hh = pd.DataFrame({"REFERANS_YIL": yr, "BULTEN_NO": np.arange(10011, 10011 + n_hh)})
    hh["IBBS_1"] = rng.choice([f"TR{i}" for i in range(1, 10)] + ["TRA", "TRB", "TRC"], n_hh)
    # gerçek veride IBBS_1 metin kodudur (TR1 ... TRC)
    hh["BLOKNO"] = hh.groupby("IBBS_1").cumcount() // 12 + 1001
    hh["MAH_DEG_L1"] = rng.choice([1, 2, 3], n_hh, p=[.2, .25, .55])
    inc = rng.choice([1, 2, 3, 4, 5, 99], n_hh, p=[.2, .2, .2, .2, .17, .03])
    hh["HANE_AYLIK_GELIR_GRP_5"] = inc
    hh["HHB"] = rng.choice([1, 2, 3, 4, 5, 6], n_hh, p=[.12, .25, .22, .22, .12, .07])
    hh["BT_BILG_MASAUSTU"] = yn(rng, np.full(n_hh, .2))
    hh["BT_BILG_DIZUSTU"] = yn(rng, 0.15 + 0.08 * inc.clip(0, 5))
    hh["BT_TABLET"] = yn(rng, np.full(n_hh, .3))
    hh["HANE_INTERNET_ERISIM_DURUMU"] = yn(rng, np.full(n_hh, .95))
    hh["HANE_FAKTOR"] = rng.uniform(800, 3000, n_hh)
    # fertler
    sizes = hh["HHB"].to_numpy()
    f = hh.loc[hh.index.repeat(sizes), ["BULTEN_NO", "BLOKNO", "IBBS_1", "MAH_DEG_L1", "HANE_AYLIK_GELIR_GRP_5"]].reset_index(drop=True)
    f["FERTNO"] = f.groupby("BULTEN_NO").cumcount() + 1
    f["HANE_SIRANO"] = f.groupby("BLOKNO").ngroup() % 50 + 1
    n = len(f)
    f["REFERANS_YIL"] = yr
    f["YAS"] = rng.integers(0, 90, n)
    f["CINSIYET"] = rng.choice([1, 2], n)
    f["OKUL_BITEN"] = rng.choice([1, 2, 3, 4, 511, 512, 52, 53], n, p=[.06, .2, .2, .27, .07, .16, .03, .01])
    f["OKUMA_YAZMA"] = np.where(f["OKUL_BITEN"] == 1, rng.choice([1, 2], n), 1)
    f["GORUSME_YONTEM_FERT"] = rng.choice([1, 3], n, p=[.7, .3])
    f["FERT_FAKTOR"] = rng.uniform(500, 4000, n)
    age = f["YAS"].to_numpy()
    edu_hi = f["OKUL_BITEN"].isin([511, 512, 52, 53]).to_numpy()
    fem = (f["CINSIYET"] == 2).to_numpy()
    # GALI
    pg = 0.04 + 0.004 * np.clip(age - 30, 0, None)
    gali_any = rng.random(n) < pg
    severe = gali_any & (rng.random(n) < .35)
    dur = np.where(severe, 4, np.where(gali_any, 5, 6))
    sure = np.where(dur < 6, np.where(rng.random(n) < .9, 1, 2), np.nan)
    if yr == 2026:
        f["FAALIYET_SINIRLAMA_DURUM"] = dur
        f["FAALIYET_SINIRLAMA_SURE"] = sure
    # internet
    p_int = 1 / (1 + np.exp(-(4.5 - 0.07 * np.clip(age - 30, 0, None) + 1.0 * edu_hi - 0.9 * severe - 0.3 * gali_any)))
    user = rng.random(n) < p_int
    f["SON_KULLANIM_INTERNET"] = np.where(user, 1, rng.choice([2, 3, 90], n))
    f["KULLANIM_SIKLIK_INTERNET"] = np.where(user, rng.choice([16, 17, 13, 9], n), np.nan)

    def act(p):
        return np.where(user, yn(rng, np.clip(p, 0.01, 0.99)), np.nan)

    for v, p in [("EPOSTA", .4), ("TELEFON", .8), ("SOSYAL_GRKATILIM", .7), ("MESAJ", .95),
                 ("ONLINE_HABER", .6), ("MALHIZ_BILGI", .5), ("WEBSITE_PAYLASIM", .1),
                 ("OYL_KTLM", .08), ("ISARAMA", .15), ("BANKA_ISLEM", .6)]:
        f[f"INT_FAAL_{v}"] = act(np.full(n, p) + .15 * edu_hi)
    gen = act(.55 + .1 * fem + .1 * edu_hi + .12 * gali_any)
    f["INT_FAAL_SAGLIK_BILGI_ARM"] = gen
    f["INT_FAAL_SAGLIK_BILGI_ARM_F"] = np.where(gen == 1, yn(rng, np.full(n, .9)), np.nan)
    f["INT_FAAL_SAGLIK_BILGI_ARM_R"] = np.where(gen == 1, yn(rng, .35 + .1 * fem + .05 * gali_any), np.nan)
    mhrs = act(.45 + .004 * np.clip(age - 25, 0, None) + .05 * gali_any - .1 * severe)
    rec = act(.35 + .1 * edu_hi)
    cons = act(np.full(n, .06))
    f["EDEVLET_RANDEVU_REZRVSYN"] = np.where(user, np.where(mhrs == 1, yn(rng, np.full(n, .93)), yn(rng, np.full(n, .25))), np.nan)
    f["EDEVLET_KSL_BILGI_ERISME"] = np.where(user, np.where(rec == 1, yn(rng, np.full(n, .9)), yn(rng, np.full(n, .5))), np.nan)
    f["EDEVLET_TALEP_HIC_BASKASI"] = act(.05 + .15 * severe)
    f["SON_KULLANIM_ETICARET"] = np.where(user, rng.choice([1, 2, 3, 90], n, p=[.55, .15, .1, .2]), np.nan)
    if yr == 2026:
        f["INT_FAAL_SAGLIK_RANDEVU_AL"] = gen                 # kopya
        f["INT_FAAL_SAGLIK_SONUC_AL"] = mhrs                  # MHRS
        f["INT_FAAL_SAGLIK_TIBKONS_RECETE"] = rec             # kayıt
        f["INT_FAAL_MALHIZMET_SATIS"] = cons                  # konsültasyon
        f["NESN_INT_AKILL_SGLK_CHZ"] = act(np.full(n, .05) + .03 * gali_any)
        f["NESN_INT_AKILL_SAAT"] = act(.15 + .1 * edu_hi - .05 * gali_any)
        f["EDEVLET_TALEP_HIC_ENGEL"] = act(.01 + .08 * severe)
    else:
        f["INT_FAAL_MALHIZMET_SATIS"] = act(np.full(n, .08))  # satış
    f["YAPAYZEKA_KULLAN"] = act(.15 + (.05 if yr == 2026 else 0) + .25 * edu_hi
                               - .004 * np.clip(age - 25, 0, None) - .05 * gali_any)
    for v in ["BLGMOBIL_DOSYA_AKT", "BLGMOBIL_UYGLM_YUKLM", "BLGMOBIL_SISTEM_DEG", "YZLM_KLM_ISLEMCI",
              "YZLM_DOKUMN_OLSTRM", "YZLM_HZPRG_KULLNM", "YZLM_HZPRG_FORML", "YZLM_FOTO_VIDEO", "YZLM_KODYZM"]:
        f[f"EYETENEK_{v}"] = act(.3 + .3 * edu_hi - .004 * np.clip(age - 25, 0, None))
    # gerçek veride görülen koşullu sorular
    f.loc[f["EYETENEK_YZLM_HZPRG_KULLNM"] == 2, "EYETENEK_YZLM_HZPRG_FORML"] = np.nan
    asked = rng.random(n) < .06
    tal = [c for c in f.columns if c.startswith("EDEVLET_TALEP_HIC")]
    f.loc[~asked, tal] = np.nan
    # bireysel sorular yalnızca 16-74
    ind_cols = [c for c in f.columns if c.startswith(("INT_", "EDEVLET", "SON_", "KULLANIM", "YAPAY",
                                                         "EYETENEK", "NESN", "FAALIYET", "OKUL", "OKUMA"))]
    out_age = ~((age >= 16) & (age <= 74))
    f.loc[out_age, ind_cols] = np.nan
    fert_cols = [c for c in f.columns if c not in ("IBBS_1", "MAH_DEG_L1", "HANE_AYLIK_GELIR_GRP_5")]
    if yr == 2025:
        fert_cols = [c for c in fert_cols if c not in ("BLOKNO", "HANE_SIRANO")]
    hh_cols = [c for c in hh.columns if c != "BLOKNO"]
    hh = hh[hh_cols].rename(columns={"HANE_AYLIK_GELIR_GRP_5": "hane_aylik_gelir_grp_5"})
    return f[fert_cols], hh


if __name__ == "__main__":
    rng = np.random.default_rng(1)
    for yr in (2025, 2026):
        f, h = make_year(yr, rng)
        f.to_excel(OUT / f"hbtka_{yr}_fert.xlsx", index=False)
        h.to_excel(OUT / f"hbtka_{yr}_hh.xlsx", index=False)
        print(yr, len(f), "fert,", len(h), "hane")
