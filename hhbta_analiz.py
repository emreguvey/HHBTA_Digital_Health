from __future__ import annotations
# =============================================================================
# HHBTA 2025–2026 · Activity limitation and digital health: individual and household digital resources
# SINGLE-FILE ANALYSIS
#
# HOW TO USE
#   1) Edit only the "SETTINGS" block below (stages, replicates, data paths, decisions).
#   2) Run the file: python hhbta_analiz.py (or "Run Python File" in an editor).
#   3) Each stage writes outputs/<stage>/report.md; all reports are combined in outputs/SON_RAPOR.md.
#   4) Intermediate reports are in Turkish; all manuscript outputs (stage 11) are in English.
#
# Required packages: pandas numpy scipy patsy openpyxl pyyaml (statsmodels optional, for validation)
#   (stage 11 installs matplotlib and python-docx if they are missing)
# =============================================================================

# ============================== SETTINGS =====================================
# (Turkish identifiers: ASAMALAR = stages, TASLAK = draft, VERI = data, KARARLAR = decisions,
#  SENTETIK = synthetic. Glossary of Turkish column names: see README.md)

# --- 1) Which stages to run? ----------------------------------------------
#   "00" load (reads Excel; needed once)   "01" data audit
#   "02" derive variables   "03" design and power   "04" descriptive cascade
#   "05" models (H1–H4)   "06" access/conditional-use decomposition   "07" KHB
#   "08" 2025→2026 change   "09" crude and adjusted gaps   "10" sensitivity analyses
#   "12" relative scale (risk ratios) and age strata
#   "13" relational (household) access and digital isolation   "14" within-household comparison (fixed effects)
#   "15" household need spillover and negative control
#   "11" manuscript outputs (English tables, figures, supplementary files; Healthcare/MDPI) — run AFTER 12–15
#   all stages: ["hepsi"]
ASAMALAR = ["hepsi"]

# Quick test: True = 100 bootstrap replicates, False = 500 (final).
# If you change this, re-run stage 03 as well.
TASLAK = False

# --- 2) Data files ---------------------------------------------------------
VERI = {
    # TurkStat microdata (not distributed). Place the four Excel files in ./data/
    "fert_2025": "data/hbtka_2025_fert.xlsx",
    "hh_2025":   "data/hbtka_2025_hh.xlsx",
    "fert_2026": "data/hbtka_2026_fert.xlsx",
    "hh_2026":   "data/hbtka_2026_hh.xlsx",
}

# --- 3) Analytic decisions (checkpoints) -------------------------------------
KARARLAR = {
    # Checkpoint A (after stage 01) — decided 2026-10-26
    "label_mapping": "HA",          # RANDEVU_AL = appointment, SONUC_AL = records, TIBKONS = consultation (confirmed with the phi matrix)
    "activity_blank_is_no": False,  # no blanks among internet users; coded 1/2
    "fr_structural_zero": True,     # physical/mental items are not asked if general health information = no
    "yz_denominator": "users",      # generative AI asked only of internet users in the last 3 months
    "gali_rule": "eurostat",        # 722 respondents with duration < 6 months (SURE = 2) classified as 'not limited'
    # Checkpoint B (after stage 03)
    "delta_pp": 5,                  # 2026-10-26: hardest comparison (65–74 vs 25–34, users) needs ≥4.9 pp for TOST
    "deft_ratio_2025": 1.24,        # 2026-10-26: median 2026 block/household SE ratio (1.235)
    # Checkpoint C (after stage 04)
    "gali_merge_levels": False,     # 2026-10-26: no small cells in the transaction step; three levels kept
    "drop_outcomes": ["SAGLIK_CIHAZ"],  # 2026-10-26: < 50 "yes" among severely limited (pre-specified rule); secondary, descriptive only
}

SENTETIK = False   # True: run on synthetic data (python tests/make_synthetic.py first)

# ======================= DO NOT EDIT BELOW THIS LINE ==========================

CONFIG_YAML = r"""
bootstrap:
  B: 500
  B_draft: 100
  seed: 20261022
age:
  min: 16
  max: 74
  spline_df: 4
  groups:
  - 16
  - 25
  - 35
  - 45
  - 55
  - 65
  - 75
education_map:
  1: 1_ilkokul_alti
  2: 1_ilkokul_alti
  3: 2_ortaokul
  4: 3_lise
  511: 4_yuksekogretim
  512: 4_yuksekogretim
  52: 4_yuksekogretim
  53: 4_yuksekogretim
coding:
  evet:
  - 1
  hayir:
  - 2
  internet_recent: 1
  income_valid:
  - 1
  - 2
  - 3
  - 4
  - 5
vars:
  activities_common:
  - INT_FAAL_EPOSTA
  - INT_FAAL_TELEFON
  - INT_FAAL_SOSYAL_GRKATILIM
  - INT_FAAL_MESAJ
  - INT_FAAL_ONLINE_HABER
  - INT_FAAL_SAGLIK_BILGI_ARM
  - INT_FAAL_MALHIZ_BILGI
  - INT_FAAL_WEBSITE_PAYLASIM
  - INT_FAAL_OYL_KTLM
  - INT_FAAL_ISARAMA
  - INT_FAAL_MALHIZMET_SATIS
  - INT_FAAL_BANKA_ISLEM
  repertoire:
  - INT_FAAL_EPOSTA
  - INT_FAAL_TELEFON
  - INT_FAAL_SOSYAL_GRKATILIM
  - INT_FAAL_MESAJ
  - INT_FAAL_ONLINE_HABER
  - INT_FAAL_MALHIZ_BILGI
  - INT_FAAL_WEBSITE_PAYLASIM
  - INT_FAAL_OYL_KTLM
  - INT_FAAL_ISARAMA
  - INT_FAAL_BANKA_ISLEM
  skills:
  - EYETENEK_BLGMOBIL_DOSYA_AKT
  - EYETENEK_BLGMOBIL_UYGLM_YUKLM
  - EYETENEK_BLGMOBIL_SISTEM_DEG
  - EYETENEK_YZLM_KLM_ISLEMCI
  - EYETENEK_YZLM_DOKUMN_OLSTRM
  - EYETENEK_YZLM_HZPRG_KULLNM
  - EYETENEK_YZLM_HZPRG_FORML
  - EYETENEK_YZLM_FOTO_VIDEO
  - EYETENEK_YZLM_KODYZM
  health_2026_candidates:
  - INT_FAAL_SAGLIK_RANDEVU_AL
  - INT_FAAL_SAGLIK_SONUC_AL
  - INT_FAAL_SAGLIK_TIBKONS_RECETE
  - INT_FAAL_MALHIZMET_SATIS
  computer:
  - BT_BILG_MASAUSTU
  - BT_BILG_DIZUSTU
  - BT_TABLET
  skill_conditional:
    EYETENEK_YZLM_HZPRG_FORML: EYETENEK_YZLM_HZPRG_KULLNM
label_mappings:
  HA:
    MHRS: INT_FAAL_SAGLIK_RANDEVU_AL
    KAYIT: INT_FAAL_SAGLIK_SONUC_AL
    KONSULT: INT_FAAL_SAGLIK_TIBKONS_RECETE
  HB:
    MHRS: INT_FAAL_SAGLIK_SONUC_AL
    KAYIT: INT_FAAL_SAGLIK_TIBKONS_RECETE
    KONSULT: INT_FAAL_MALHIZMET_SATIS
models:
  total_rhs: cr(AGE, df=4, constraints='center') + FEMALE + C(EDU) + C(URBAN) + C(REGION)
  direct_extra: C(INCOME) + COMPUTER + C(HHSIZE)
  user_outcomes:
  - BILGI_F
  - BILGI_R
  - MHRS
  - KAYIT
  - KONSULT
  - YZ
  - SAGLIK_CIHAZ
  - AKILLI_SAAT
  steps:
    bilgi:
    - BILGI_F
    - BILGI_R
    islem:
    - MHRS
    - KAYIT
    sinir:
    - YZ
    - AKILLI_SAAT
  primary_family:
  - INTERNET
  - BILGI_F
  - BILGI_R
  - MHRS
  - KAYIT
  - YZ
  - H3a_bilgi_eksi_islem
  - H3b_bilgi_eksi_sinir
  hurdle_outcomes:
  - BILGI_F
  - BILGI_R
  - MHRS
  - KAYIT
  - KONSULT
  - YZ
  khb_outcomes:
  - BILGI_F
  - BILGI_R
  - MHRS
  - KAYIT
  - YZ
  khb_mediators:
  - REPERTUAR
  - BECERI
trend:
  outcomes:
    INTERNET: all
    BILGI_F: users
    BILGI_R: users
    YZ: users
  rhs: C(YEAR) * (C(EDU) + FEMALE + C(AGEGRP)) + C(REGION) + C(URBAN)
  contrasts:
    EDU:
    - 4_yuksekogretim
    - 1_ilkokul_alti
    FEMALE:
    - 1
    - 0
    AGEGRP:
    - 65-74
    - 25-34
  dfl_vars: C(AGEGRP) + FEMALE + C(EDU) + C(REGION) + C(URBAN) + C(HHSIZE)
  assumed_cluster_deff: 1.5
min_cell_n: 50
"""



import datetime as _dt
import pathlib
import pickle
import sys
import time
from typing import Callable

import numpy as np
import pandas as pd
import patsy
import yaml
from scipy import stats

ROOT = pathlib.Path(__file__).resolve().parent
CONFIG = yaml.safe_load(CONFIG_YAML)


class DecisionMissing(Exception):
    """A required decision has not yet been set in KARARLAR."""


class Ctx:
    """Run context: configuration, decisions, paths, number of bootstrap replicates."""

    def __init__(self, synthetic: bool = False, draft: bool = False):
        self.cfg = CONFIG
        self.dec = dict(KARARLAR)
        self.synthetic = synthetic
        if synthetic:
            self.paths = {f"{k}_{y}": f"tests/synthetic/hbtka_{y}_{'fert' if k == 'fert' else 'hh'}.xlsx"
                          for k in ("fert", "hh") for y in (2025, 2026)}
            self.paths.update(cache_dir="cache_synthetic", output_dir="outputs_synthetic")
            self.dec.update(label_mapping="HB", delta_pp=5, deft_ratio_2025=1.2, gali_merge_levels=False)
        else:
            self.paths = dict(VERI, cache_dir="cache", output_dir="outputs")
        self.B = self.cfg["bootstrap"]["B_draft"] if draft else self.cfg["bootstrap"]["B"]
        self.draft = draft
        self.seed = self.cfg["bootstrap"]["seed"]

    def _abs(self, p: str) -> pathlib.Path:
        p = pathlib.Path(p)
        return p if p.is_absolute() else ROOT / p

    def input_path(self, key: str) -> pathlib.Path:
        return self._abs(self.paths[key])

    def cache(self, name: str) -> pathlib.Path:
        d = self._abs(self.paths["cache_dir"])
        d.mkdir(parents=True, exist_ok=True)
        return d / name

    def out(self, stage: str) -> pathlib.Path:
        d = self._abs(self.paths["output_dir"]) / stage
        d.mkdir(parents=True, exist_ok=True)
        return d

    def require(self, *keys: str) -> None:
        missing = [k for k in keys if self.dec.get(k) is None]
        if missing:
            raise DecisionMissing(
                "Dosyanın en üstündeki KARARLAR içinde şu alanlar henüz boş (None): "
                + ", ".join(missing)
                + ". Bu raporu Claude'a gönderin; kararı birlikte verelim."
            )

    def log_decisions(self, stage: str) -> None:
        log = self._abs(self.paths["output_dir"]) / "decision_log.md"
        log.parent.mkdir(parents=True, exist_ok=True)
        stamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
        with open(log, "a", encoding="utf-8") as f:
            f.write(f"\n## {stamp} — {stage} (B={self.B}, sentetik={self.synthetic})\n\n")
            f.write("```yaml\n" + yaml.safe_dump(self.dec, allow_unicode=True, sort_keys=False) + "```\n")

# --- intermediate files ---
def save(obj, path: pathlib.Path) -> None:
    with open(path, "wb") as f:
        pickle.dump(obj, f, protocol=pickle.HIGHEST_PROTOCOL)


def load(path: pathlib.Path):
    if not path.exists():
        raise FileNotFoundError(f"{path} yok. Önceki aşamayı çalıştırın.")
    with open(path, "rb") as f:
        return pickle.load(f)


# --- weighted helpers ---
def wmean(x, w) -> float:
    x = np.asarray(x, float)
    w = np.asarray(w, float)
    m = ~np.isnan(x) & ~np.isnan(w)
    return float(np.sum(x[m] * w[m]) / np.sum(w[m])) if m.any() else np.nan


def kish_deff(w) -> float:
    w = np.asarray(w, float)
    w = w[~np.isnan(w)]
    return float(len(w) * np.sum(w**2) / np.sum(w) ** 2)


# --- reports ---
def md_table(df: pd.DataFrame, digits: int = 3, index: bool = False) -> str:
    if df is None or len(df) == 0:
        return "_(boş tablo)_\n"
    d = df.reset_index() if index else df.copy()

    def fmt(v):
        if isinstance(v, (float, np.floating)):
            if np.isnan(v):
                return ""
            return f"{v:.{digits}f}"
        return str(v)

    head = "| " + " | ".join(map(str, d.columns)) + " |"
    sep = "| " + " | ".join("---" for _ in d.columns) + " |"
    rows = ["| " + " | ".join(fmt(v) for v in r) + " |" for r in d.itertuples(index=False)]
    return "\n".join([head, sep, *rows]) + "\n"


class Report:
    """report.md of each stage, ending with a checklist."""

    def __init__(self, ctx: Ctx, stage: str, title: str):
        self.ctx, self.stage = ctx, stage
        self.dir = ctx.out(stage)
        self.lines = [f"# {title}", "",
                      f"_Oluşturma: {_dt.datetime.now():%Y-%m-%d %H:%M} · B={ctx.B}"
                      + (" · SENTETİK VERİ" if ctx.synthetic else "")
                      + (" · TASLAK (B_draft)" if ctx.draft else "") + "_", ""]

    def h(self, text: str, level: int = 2):
        self.lines += [f"{'#' * level} {text}", ""]

    def p(self, text: str):
        self.lines += [text, ""]

    def table(self, df: pd.DataFrame, name: str | None = None, digits: int = 3, index=False):
        self.lines.append(md_table(df, digits, index))
        if name:
            path = self.dir / f"{name}.csv"
            (df.reset_index() if index else df).to_csv(path, index=False, encoding="utf-8-sig")
            self.lines += [f"_CSV: `{path.name}`_", ""]

    def checkpoint(self, items: list[str]):
        self.h("KONTROL NOKTASI — devam etmeden önce", 2)
        self.lines += [f"- [ ] {i}" for i in items] + [""]

    def write(self):
        path = self.dir / "report.md"
        path.write_text("\n".join(self.lines), encoding="utf-8")
        print(f"  → rapor: {path}")
        return path



# =============================================================================
# ESTIMATION
# =============================================================================
# Estimation engine.
# 
# Principle: every estimate is a function "weight vector → result (pd.Series)".
# The point estimate uses the full-sample weights; the variance comes from calling the
# same function with each set of Rao–Wu bootstrap replicate weights.
# AMEs, contrasts, two-part decompositions, KHB and DFL therefore all get the same
# design-based variance, and the covariance between outputs is preserved automatically.




# ------------------------------------------------------------------ replicates
def rao_wu_weights(strata, psu, w, B: int, seed: int) -> np.ndarray:
    """Rao–Wu rescaled bootstrap (m_h = n_h - 1). Returns an (n, B) float32 array."""
    rng = np.random.default_rng(seed)
    strata = np.asarray(strata)
    psu = np.asarray(psu)
    w = np.asarray(w, float)
    R = np.repeat(w[:, None], B, axis=1).astype(np.float32)
    for h in pd.unique(strata):
        rows = np.flatnonzero(strata == h)
        codes, uniq = pd.factorize(psu[rows])
        n_h = len(uniq)
        if n_h < 2:
            continue  # stratum with a single PSU: no contribution to variance (flagged in the report)
        draws = rng.integers(0, n_h, size=(B, n_h - 1))
        counts = np.zeros((B, n_h))
        np.add.at(counts, (np.arange(B)[:, None], draws), 1)
        factor = counts * n_h / (n_h - 1)          # (B, n_h)
        R[rows, :] = (w[rows][:, None] * factor[:, codes].T).astype(np.float32)
    return R


def replicate(fn: Callable[[np.ndarray], pd.Series], W: np.ndarray, R: np.ndarray,
              progress: str | None = None) -> tuple[pd.Series, pd.DataFrame]:
    point = fn(W)
    reps = []
    B = R.shape[1]
    for b in range(B):
        reps.append(fn(R[:, b].astype(float)))
        if progress and (b + 1) % max(1, B // 5) == 0:
            print(f"    {progress}: {b + 1}/{B}")
    return point, pd.DataFrame(reps)


def summarize(point: pd.Series, reps: pd.DataFrame, level: float = 0.95,
              scale: float = 1.0) -> pd.DataFrame:
    """Bootstrap summary: SE (around the point estimate), normal-approximation CI, z, p."""
    dev = reps[point.index] - point
    se = np.sqrt((dev**2).mean()) * scale
    z = stats.norm.ppf(0.5 + level / 2)
    out = pd.DataFrame({"tahmin": point, "se": se,
                        "ga_alt": point - z * se, "ga_ust": point + z * se})
    out["z"] = out["tahmin"] / out["se"]
    out["p"] = 2 * stats.norm.sf(np.abs(out["z"]))
    out["gecersiz_rep"] = reps[point.index].isna().sum()
    return out


# ------------------------------------------------------------------ logistic
def expit(x):
    return 1.0 / (1.0 + np.exp(-np.clip(x, -35, 35)))


def _deviance(y, mu, w):
    return -2 * np.sum(w * (y * np.log(mu) + (1 - y) * np.log(1 - mu)))


def irls_logit(X: np.ndarray, y: np.ndarray, w: np.ndarray, beta0=None,
               tol: float = 1e-10, maxit: int = 50, ridge: float = 1e-8) -> np.ndarray:
    """Weighted logistic regression (pseudo-ML). A tiny ridge term is for numerical stability only.

    Stopping rule: relative change in deviance < tol (same logic as statsmodels GLM).
    Returns NaN if not converged; such replicates are counted as `gecersiz_rep` (invalid).
    """
    k = X.shape[1]
    beta = np.zeros(k) if beta0 is None else beta0.copy()
    mu = np.clip(expit(X @ beta), 1e-10, 1 - 1e-10)
    dev = _deviance(y, mu, w)
    for _ in range(maxit):
        eta = X @ beta
        v = mu * (1 - mu)
        z = eta + (y - mu) / v
        XtW = X.T * (w * v)
        beta = np.linalg.solve(XtW @ X + ridge * np.eye(k), XtW @ z)
        mu = np.clip(expit(X @ beta), 1e-10, 1 - 1e-10)
        dev_new = _deviance(y, mu, w)
        if abs(dev_new - dev) / (abs(dev_new) + 0.1) < tol:
            return beta
        dev = dev_new
    return np.full(k, np.nan)


def wls(X, y, w):
    XtW = X.T * w
    return np.linalg.solve(XtW @ X + 1e-10 * np.eye(X.shape[1]), XtW @ y)


class Logit:
    """Weighted logistic model built from a formula; used for counterfactual predictions.

    data: full analysis table (RangeIndex). subset: rows entering the model (boolean).
    Rows with missing values are dropped by patsy; self.rows holds the remaining positions.
    """

    def __init__(self, data: pd.DataFrame, formula: str, subset=None):
        d = data if subset is None else data[np.asarray(subset, bool)]
        y, X = patsy.dmatrices(formula, d, return_type="dataframe")
        self.rows = X.index.to_numpy()          # positions within data
        self.X = X.to_numpy(float)
        self.y = y.iloc[:, 0].to_numpy(float)
        self.names = list(X.columns)
        self.design_info = X.design_info
        self.data = data
        self._beta_full = None

    def fit(self, w_full: np.ndarray) -> np.ndarray:
        w = w_full[self.rows]
        beta = irls_logit(self.X, self.y, w, self._beta_full)
        if self._beta_full is None:
            if not np.all(np.isfinite(beta)):
                raise RuntimeError(f"Tam örneklem modeli yakınsamadı: {self.names[:3]}... "
                                   "Seyrek hücre/ayrışma olabilir; çıktıyı drop_outcomes'a ekleyin "
                                   "veya gali_merge_levels: true deneyin.")
            self._beta_full = beta
        return beta

    def design(self, newdata: pd.DataFrame) -> np.ndarray:
        return np.asarray(patsy.build_design_matrices([self.design_info], newdata)[0], float)

    def cf_matrix(self, rows: np.ndarray, **setvals) -> np.ndarray:
        nd = self.data.iloc[rows].copy()
        for k, v in setvals.items():
            nd[k] = v
        return self.design(nd)

    def coef(self, beta, name: str) -> float:
        return float(beta[self.names.index(name)])


def ame_fn(model: Logit, exposure: str, levels: list, ref, pop_rows=None):
    """AME(l) = Σ w [p(E=l) - p(E=ref)] / Σ w — over pop_rows (default: model rows)."""
    rows = model.rows if pop_rows is None else pop_rows
    Xref = model.cf_matrix(rows, **{exposure: ref})
    Xl = {l: model.cf_matrix(rows, **{exposure: l}) for l in levels}

    def fn(w_full):
        beta = model.fit(w_full)
        w = w_full[rows]
        p0 = expit(Xref @ beta)
        out = {}
        for l, Xm in Xl.items():
            out[f"{l}"] = np.sum(w * (expit(Xm @ beta) - p0)) / np.sum(w)
        return pd.Series(out)

    return fn


def holm(p: pd.Series) -> pd.Series:
    from statsmodels.stats.multitest import multipletests
    p = p.dropna()
    return pd.Series(multipletests(p.values, method="holm")[1], index=p.index)


def bh(p: pd.Series) -> pd.Series:
    from statsmodels.stats.multitest import multipletests
    p = p.dropna()
    return pd.Series(multipletests(p.values, method="fdr_bh")[1], index=p.index)


# =============================================================================
# S00_LOAD
# =============================================================================
# Stage 00 — read the Excel files, merge individual and household files, cache.
# 
# Checks: row counts versus the codebook, unmatched records in the merge,
# variables common to both years.



EXPECTED = {2025: dict(hh=11924, fert=25493), 2026: dict(hh=12178, fert=26143)}


def _read(path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df


def asama_00(ctx: Ctx):
    rep = Report(ctx, "00_load", "Aşama 00 — Yükleme ve birleştirme")
    rows, colsets = [], {}
    for yr in (2025, 2026):
        print(f"  {yr} okunuyor (Excel okuma birkaç dakika sürebilir)...")
        fert = _read(ctx.input_path(f"fert_{yr}"))
        hh = _read(ctx.input_path(f"hh_{yr}"))
        hh_dups = int(hh["BULTEN_NO"].duplicated().sum())
        overlap = (set(fert.columns) & set(hh.columns)) - {"BULTEN_NO"}
        df = fert.merge(hh.drop(columns=list(overlap)), on="BULTEN_NO",
                        how="left", validate="m:1", indicator=True)
        unmatched = int((df["_merge"] != "both").sum())
        df = df.drop(columns="_merge")
        df["YEAR"] = yr
        df = df.reset_index(drop=True)
        save(df, ctx.cache(f"raw_{yr}.pkl"))
        colsets[yr] = set(df.columns)
        age = pd.to_numeric(df.get("YAS"), errors="coerce")
        rows.append(dict(
            yil=yr, fert_satir=len(fert), hane_satir=len(hh),
            beklenen_fert=EXPECTED[yr]["fert"], beklenen_hane=EXPECTED[yr]["hh"],
            yas_16_74=int(age.between(16, 74).sum()),
            hane_dosyasinda_tekrar_BULTEN_NO=hh_dups,
            fertte_eslesmeyen=unmatched,
            ortak_kolonlar_hh_ve_fert=", ".join(sorted(overlap)),
        ))

    rep.h("Satır sayıları")
    rep.table(pd.DataFrame(rows), "satir_sayilari")
    rep.p("Not: Fert dosyası tüm hane üyelerini içeriyorsa satır sayısı beklenenden büyük olabilir; "
          "beklenen sayı 'cevaplı fert'tir. `yas_16_74` sütunu analiz evrenini gösterir.")

    common = sorted(colsets[2025] & colsets[2026])
    only25 = sorted(colsets[2025] - colsets[2026])
    only26 = sorted(colsets[2026] - colsets[2025])
    rep.h("Değişken kümeleri")
    rep.p(f"İki yılda ortak: {len(common)} · yalnızca 2025: {len(only25)} · yalnızca 2026: {len(only26)}")
    rep.p("**Yalnızca 2025:** " + ", ".join(only25))
    rep.p("**Yalnızca 2026:** " + ", ".join(only26))
    pd.DataFrame({"degisken": common}).to_csv(rep.dir / "ortak_degiskenler.csv", index=False)

    rep.checkpoint([
        "Fert ve hane sayıları kılavuzdaki (2025: 25.493 / 11.924; 2026: 26.143 / 12.178) değerlerle uyumlu mu?",
        "`fertte_eslesmeyen` = 0 mı? Değilse BULTEN_NO tipini (sayı/metin) kontrol edin.",
        "Yalnızca-2025 / yalnızca-2026 listelerinde beklenmedik bir değişken var mı "
        "(ör. büyük/küçük harf ya da yazım farkı yüzünden ortak görünmeyen)?",
        "Bu aşama karar gerektirmez; doğrudan 01_audit çalıştırılabilir.",
    ])
    rep.write()


# =============================================================================
# S01_AUDIT
# =============================================================================
# Stage 01 — data audit (WITHOUT looking at outcome distributions by GALI).
# 
# Produces: coding tables, skip patterns, label-shift tests, design variables,
# weight summaries, GALI distribution. Feeds the checkpoint A decisions.




def num(s):
    return pd.to_numeric(s, errors="coerce")


def asama_01(ctx: Ctx):
    cfg = ctx.cfg
    yes = cfg["coding"]["evet"]
    rep = Report(ctx, "01_audit", "Aşama 01 — Veri denetimi")
    d = {yr: load(ctx.cache(f"raw_{yr}.pkl")) for yr in (2025, 2026)}
    for yr, df in d.items():
        age = num(df["YAS"])
        d[yr] = df[age.between(cfg["age"]["min"], cfg["age"]["max"])].reset_index(drop=True)

    def is_user(df):
        return num(df["SON_KULLANIM_INTERNET"]) == cfg["coding"]["internet_recent"]

    # ------------------------------------------------------------------ 1
    rep.h("1. Ağırlıklar")
    rows = []
    for yr, df in d.items():
        w = num(df["FERT_FAKTOR"])
        rows.append(dict(yil=yr, n=len(df), eksik_agirlik=int(w.isna().sum()),
                         toplam=w.sum(), min=w.min(), medyan=w.median(), maks=w.max(),
                         maks_min_orani=w.max() / w.min(), kish_deff=kish_deff(w)))
    rep.table(pd.DataFrame(rows), "agirliklar", digits=2)

    # ------------------------------------------------------------------ 2
    rep.h("2. Tasarım değişkenleri")
    df = d[2026]
    if "BLOKNO" in df.columns:
        blk = df[["IBBS_1", "BLOKNO", "BULTEN_NO"]].dropna()
        n_ibbs_per_blk = blk.groupby("BLOKNO")["IBBS_1"].nunique()
        psu = blk.groupby(["IBBS_1", "BLOKNO"])
        per_stratum = blk.drop_duplicates(["IBBS_1", "BLOKNO"]).groupby("IBBS_1").size()
        rep.table(pd.DataFrame([dict(
            tekil_BLOKNO=blk["BLOKNO"].nunique(),
            tekil_IBBS1_BLOKNO=psu.ngroups,
            birden_fazla_bolgede_gorulen_BLOKNO=int((n_ibbs_per_blk > 1).sum()),
            tabaka_basina_PSU_min=per_stratum.min(), tabaka_basina_PSU_medyan=per_stratum.median(),
            PSU_basina_hane_medyan=psu["BULTEN_NO"].nunique().median(),
            eksik_BLOKNO=int(d[2026]["BLOKNO"].isna().sum()),
        )]), "tasarim_2026")
        rep.p("`birden_fazla_bolgede_gorulen_BLOKNO` > 0 ise BLOKNO tek başına tekil değildir; "
              "boru hattı PSU kimliğini zaten YIL+IBBS_1+BLOKNO olarak kurar.")
    else:
        rep.p("**UYARI:** 2026 dosyasında BLOKNO bulunamadı.")
    rep.p(f"2025 hane sayısı (PSU olarak kullanılacak): {d[2025]['BULTEN_NO'].nunique()}")

    # ------------------------------------------------------------------ 3
    rep.h("3. Ham kodlar (internet kullanıcısı / kullanmayan)")
    rep.p("Her madde için gözlenen değerler ve sayılar. Karar: `activity_blank_is_no`. "
          "Kullanıcılarda 2 (hayır) kodu hiç yoksa ve boşlar çoksa, boş = hayır demektir.")
    check_vars = (cfg["vars"]["activities_common"]
                  + ["INT_FAAL_SAGLIK_BILGI_ARM_F", "INT_FAAL_SAGLIK_BILGI_ARM_R"]
                  + cfg["vars"]["health_2026_candidates"]
                  + ["YAPAYZEKA_KULLAN", "NESN_INT_AKILL_SGLK_CHZ", "NESN_INT_AKILL_SAAT",
                     "EDEVLET_RANDEVU_REZRVSYN", "EDEVLET_KSL_BILGI_ERISME",
                     "EDEVLET_TALEP_HIC_BASKASI", "EDEVLET_TALEP_HIC_ENGEL"]
                  + cfg["vars"]["skills"])
    check_vars = list(dict.fromkeys(check_vars))
    rows = []
    for yr, df in d.items():
        u = is_user(df)
        for v in check_vars:
            if v not in df.columns:
                continue
            for grp, mask in (("kullanici", u), ("kullanmayan", ~u)):
                vc = df.loc[mask, v].value_counts(dropna=False)
                rows.append(dict(yil=yr, degisken=v, grup=grp,
                                 kodlar="; ".join(f"{k}:{n}" for k, n in vc.items())))
    codes = pd.DataFrame(rows)
    codes.to_csv(rep.dir / "ham_kodlar.csv", index=False, encoding="utf-8-sig")
    # summary: share of blanks among users
    rows = []
    for yr, df in d.items():
        u = is_user(df)
        for v in check_vars:
            if v not in df.columns:
                continue
            x = num(df[v])
            rows.append(dict(yil=yr, degisken=v,
                             kullanicida_bos=float(x[u].isna().mean()),
                             kullanicida_evet=float(x[u].isin(yes).mean()),
                             kullanicida_kod2=float((x[u] == 2).mean()),
                             kullanmayanda_bos=float(x[~u].isna().mean()) if (~u).any() else np.nan))
    rep.table(pd.DataFrame(rows), "atlama_ozeti")
    rep.p("Tüm değer listesi: `ham_kodlar.csv`.")

    # ------------------------------------------------------------------ 4
    rep.h("4. Sağlık bilgisi alt maddeleri (F/R) atlama deseni")
    rep.p("Genel soruya cevap kodu × F ve R cevap kodu (internet kullanıcıları). "
          "Genel = 2 olanlarda F/R boşsa `fr_structural_zero: true`.")
    for yr, df in d.items():
        u = df[is_user(df)]
        for sub in ("INT_FAAL_SAGLIK_BILGI_ARM_F", "INT_FAAL_SAGLIK_BILGI_ARM_R"):
            ct = pd.crosstab(u["INT_FAAL_SAGLIK_BILGI_ARM"].fillna(-9), u[sub].fillna(-9))
            ct.index.name = "genel"
            rep.p(f"**{yr} — {sub}** (−9 = boş)")
            rep.table(ct, index=True, digits=0)

    # ------------------------------------------------------------------ 5
    rep.h("5. Üretken YZ sorusu kime soruluyor?")
    rep.p("Satır: SON_KULLANIM_INTERNET; sütun: YAPAYZEKA_KULLAN (−9 = boş). "
          "Kullanmayanlarda tamamen boşsa `yz_denominator: users`.")
    for yr, df in d.items():
        ct = pd.crosstab(df["SON_KULLANIM_INTERNET"].fillna(-9), df["YAPAYZEKA_KULLAN"].fillna(-9))
        rep.p(f"**{yr}**")
        rep.table(ct, index=True, digits=0)

    # ------------------------------------------------------------------ 6
    rep.h("6. Etiket kayması testleri (2026)")
    df = d[2026]
    u = df[is_user(df)].copy()
    w = num(u["FERT_FAKTOR"])

    def yesb(v):
        return num(u[v]).isin(yes).astype(float).where(num(u[v]).notna())

    gen = yesb("INT_FAAL_SAGLIK_BILGI_ARM")
    rows = []
    for v in cfg["vars"]["health_2026_candidates"]:
        if v not in u.columns:
            continue
        x = yesb(v)
        both = x.notna() & gen.notna()
        m1 = x == 1
        rows.append(dict(
            degisken=v,
            yayginlik_kullanicilarda=wmean(x, w),
            genel_saglik_bilgisi_ile_uyum=float((x[both] == gen[both]).mean()),
            P_EDEVLET_RANDEVU_given_evet=wmean(yesb("EDEVLET_RANDEVU_REZRVSYN")[m1], w[m1]),
            P_EDEVLET_KAYIT_given_evet=wmean(yesb("EDEVLET_KSL_BILGI_ERISME")[m1], w[m1]),
            P_ETICARET_3ay_given_evet=wmean((num(u["SON_KULLANIM_ETICARET"]) == 1).astype(float)[m1], w[m1]),
            ort_yas_evet=wmean(num(u["YAS"])[m1], w[m1]),
            ort_yas_hayir=wmean(num(u["YAS"])[x == 0], w[x == 0]),
        ))
    ref = dict(degisken="(referans) tüm kullanıcılar",
               yayginlik_kullanicilarda=np.nan, genel_saglik_bilgisi_ile_uyum=np.nan,
               P_EDEVLET_RANDEVU_given_evet=wmean(yesb("EDEVLET_RANDEVU_REZRVSYN"), w),
               P_EDEVLET_KAYIT_given_evet=wmean(yesb("EDEVLET_KSL_BILGI_ERISME"), w),
               P_ETICARET_3ay_given_evet=wmean((num(u["SON_KULLANIM_ETICARET"]) == 1).astype(float), w),
               ort_yas_evet=wmean(num(u["YAS"]), w), ort_yas_hayir=np.nan)
    rows.append(ref)
    lab = pd.DataFrame(rows)
    rep.table(lab, "etiket_kaymasi_2026")
    d25 = d[2025]
    u25 = d25[is_user(d25)]
    sat25 = wmean(num(u25["INT_FAAL_MALHIZMET_SATIS"]).isin(yes).astype(float).where(
        num(u25["INT_FAAL_MALHIZMET_SATIS"]).notna()), num(u25["FERT_FAKTOR"]))
    # weighted phi correlation between item pairs
    pair_vars = {"BILGI_ARM": "INT_FAAL_SAGLIK_BILGI_ARM",
                 **{v.replace("INT_FAAL_", ""): v for v in cfg["vars"]["health_2026_candidates"]},
                 "EDV_RANDEVU": "EDEVLET_RANDEVU_REZRVSYN", "EDV_KAYIT": "EDEVLET_KSL_BILGI_ERISME"}
    M = pd.DataFrame({k: yesb(v) for k, v in pair_vars.items() if v in u.columns})
    M["ETICARET_3AY"] = (num(u["SON_KULLANIM_ETICARET"]) == 1).astype(float)
    ok = M.notna().all(axis=1)
    Mw, ww = M[ok].to_numpy(float), w[ok].to_numpy(float)
    mu = (Mw * ww[:, None]).sum(0) / ww.sum()
    C = ((Mw - mu) * ww[:, None]).T @ (Mw - mu) / ww.sum()
    sd = np.sqrt(np.diag(C))
    phi = pd.DataFrame(C / np.outer(sd, sd), index=M.columns, columns=M.columns)
    rep.p("**Çiftler arası ağırlıklı phi korelasyonu** (internet kullanıcıları). "
          "Aynı kavramı ölçen iki madde birbiriyle ve kendi doğrulama göstergesiyle (EDV_*) "
          "diğerlerinden daha güçlü ilişkili olmalıdır.")
    rep.table(phi, "etiket_phi_2026", index=True)
    rep.p(f"Karşılaştırma: 2025'te INT_FAAL_MALHIZMET_SATIS (mal/hizmet satışı) yaygınlığı, "
          f"kullanıcılarda: **{sat25:.3f}**.")
    rep.p(
        "**Nasıl okunur?**\n\n"
        "- `genel_saglik_bilgisi_ile_uyum` ≈ 1,00 olan değişken B3.15'in kopyasıdır; belirgin altındaysa ayrı bir sorudur.\n"
        "- MHRS (3 ay) kullanıcısı 12 ayda kamu randevusu da bildirmeli: MHRS adayında "
        "`P_EDEVLET_RANDEVU_given_evet` tüm kullanıcılardaki referanstan **belirgin yüksek** ve ideali ≥ 0,80.\n"
        "- Kayıt erişimi adayında `P_EDEVLET_KAYIT_given_evet` (e-Nabız dahil) belirgin yüksek olmalı.\n"
        "- Satış maddesiyse yaygınlığı 2025 satış oranına yakın olur ve e-ticaret ile, sağlık değil, ilişkilidir.\n"
        "- HA: RANDEVU_AL=MHRS, SONUC_AL=kayıt, TIBKONS=konsültasyon. "
        "HB: SONUC_AL=MHRS, TIBKONS=kayıt, MALHIZMET_SATIS=konsültasyon."
    )

    # ------------------------------------------------------------------ 7
    rep.h("7. Faaliyet kısıtlılığı (GALI), 2026")
    df = d[2026]
    w = num(df["FERT_FAKTOR"])
    g = df["FAALIYET_SINIRLAMA_DURUM"].fillna(-9)
    s = df["FAALIYET_SINIRLAMA_SURE"].fillna(-9)
    tab = pd.DataFrame({"DURUM": g, "SURE": s, "w": w}).groupby(["DURUM", "SURE"]).agg(
        n=("w", "size"), agirlikli_pay=("w", "sum"))
    tab["agirlikli_pay"] /= w.sum()
    rep.table(tab, "gali_dagilim", index=True)
    ct = pd.crosstab(g, df["GORUSME_YONTEM_FERT"].fillna(-9), normalize="index")
    ct.index.name = "GALI_DURUM"
    rep.p("GALI × görüşme modu (satır yüzdesi; 1 = yüz yüze, 3 = telefon):")
    rep.table(ct, "gali_x_mod", index=True)

    rep.checkpoint([
        "Bölüm 3: kodlar gerçekten 1 = evet, 2 = hayır mı? Değilse Claude'a bildirin.",
        "Bölüm 3 → `activity_blank_is_no` (true/false)",
        "Bölüm 4 → `fr_structural_zero` (true/false)",
        "Bölüm 5 → `yz_denominator` (users/all)",
        "Bölüm 6 → `label_mapping` (HA/HB). İki eşleme de net değilse TÜİK'e yazın ve "
        "ana analizde HA ile HB'yi sağlamlık olarak yan yana çalıştırın (09 aşamasında).",
        "Bölüm 7: GALI eksik oranı ve 'SURE = 2' sayısı. `gali_rule` önerilen: eurostat.",
        "Bölüm 7: ciddi kısıtlıların telefonla görüşülme oranı farklı mı? (mod etkisi notu)",
        "Bölüm 2: tabaka başına en az 2 PSU var mı?",
    ])
    rep.write()


# =============================================================================
# S02_DERIVE
# =============================================================================
# Stage 02 — derive analysis variables (according to checkpoint A decisions).
# 
# Output: cache/analysis.pkl — both years, ages 16–74, one table.
# Rule: activity variables are defined only for internet users (NaN otherwise).
# Whole-population versions (non-users = 0) are created where needed.



OUTCOME_LABELS = {
    "INTERNET": "Son 3 ayda internet",
    "BILGI_GENEL": "Sağlık bilgisi (genel)",
    "BILGI_F": "Fiziksel sağlık bilgisi",
    "BILGI_R": "Ruhsal sağlık bilgisi",
    "MHRS": "Çevrimiçi doktor randevusu",
    "KAYIT": "Kişisel sağlık kayıtlarına erişim",
    "KONSULT": "Çevrimiçi konsültasyon / e-reçete",
    "YZ": "Üretken YZ",
    "SAGLIK_CIHAZ": "Bağlantılı sağlık cihazı",
    "AKILLI_SAAT": "Akıllı saat / giyilebilir",
    "EDEVLET_RANDEVU": "e-Devlet randevu (12 ay)",
    "EDEVLET_KAYIT": "e-Devlet kişisel bilgi erişimi (12 ay)",
    "VEKALET": "e-Devlet işlemini başkası yaptı",
    "ENGEL": "e-Devlet erişilebilirlik engeli",
}


def num(s):
    return pd.to_numeric(s, errors="coerce")


def as_code(s: pd.Series) -> pd.Series:
    """Convert a categorical code to text; missing values stay true NaN.

    Numeric codes (1, 2.0, "3") → "1", "2", "3". Text codes (TR1, TRA) unchanged.
    Independent of the pandas version (never produces "<NA>" or "nan" strings).
    """
    x = pd.to_numeric(s, errors="coerce")
    n_obs = int(s.notna().sum())
    if n_obs and int(x.notna().sum()) == n_obs:
        out = x.map(lambda v: str(int(v)) if pd.notna(v) else np.nan)
    else:
        out = s.map(lambda v: str(v).strip() if pd.notna(v) and str(v).strip() != "" else np.nan)
    return out.astype(object)


def asama_02(ctx: Ctx):
    ctx.require("label_mapping", "activity_blank_is_no", "fr_structural_zero",
                "yz_denominator", "gali_rule")
    ctx.log_decisions("02_derive")
    cfg, dec = ctx.cfg, ctx.dec
    yes, no = cfg["coding"]["evet"], cfg["coding"]["hayir"]
    mapping = cfg["label_mappings"][dec["label_mapping"]]
    rep = Report(ctx, "02_derive", "Aşama 02 — Değişken türetme")

    frames = []
    for yr in (2025, 2026):
        raw = load(ctx.cache(f"raw_{yr}.pkl"))
        age = num(raw["YAS"])
        raw = raw[age.between(cfg["age"]["min"], cfg["age"]["max"])].reset_index(drop=True)
        a = pd.DataFrame(index=raw.index)
        a["YEAR"] = yr
        a["BULTEN_NO"] = raw["BULTEN_NO"]
        a["FERTNO"] = raw.get("FERTNO")
        # ---- design
        w = num(raw["FERT_FAKTOR"])
        a["W_RAW"] = w
        a["W"] = w / w.mean()
        a["REGION"] = as_code(raw["IBBS_1"])
        a["STRATUM"] = f"{yr}_" + a["REGION"]
        a["PSU_HH"] = f"{yr}_" + raw["BULTEN_NO"].astype(str)
        if yr == 2026 and "BLOKNO" in raw.columns:
            a["PSU"] = f"{yr}_" + a["REGION"] + "_" + raw["BLOKNO"].astype(str)
        else:
            a["PSU"] = a["PSU_HH"]
        # ---- internet
        last = num(raw["SON_KULLANIM_INTERNET"])
        user = last == cfg["coding"]["internet_recent"]
        a["INTERNET"] = np.where(last.isna(), np.nan, user.astype(float))

        def yn(v, user_only=True, blank_is_no=None):
            """1 = yes, 0 = no, NaN = undefined. Non-users are NaN."""
            if v not in raw.columns:
                return pd.Series(np.nan, index=raw.index)
            x = num(raw[v])
            out = pd.Series(np.nan, index=raw.index)
            out[x.isin(yes)] = 1.0
            out[x.isin(no)] = 0.0
            bin_ = dec["activity_blank_is_no"] if blank_is_no is None else blank_is_no
            if bin_:
                out[x.isna() & user] = 0.0
            if user_only:
                out[~user] = np.nan
            return out

        # ---- information
        a["BILGI_GENEL"] = yn("INT_FAAL_SAGLIK_BILGI_ARM")
        for k, v in (("BILGI_F", "INT_FAAL_SAGLIK_BILGI_ARM_F"), ("BILGI_R", "INT_FAAL_SAGLIK_BILGI_ARM_R")):
            s = yn(v)
            if dec["fr_structural_zero"]:
                s[(a["BILGI_GENEL"] == 0)] = 0.0
            a[k] = s
        # ---- transactions (2026 only)
        for k, v in mapping.items():
            a[k] = yn(v) if yr == 2026 else np.nan
        # ---- frontier
        a["YZ"] = yn("YAPAYZEKA_KULLAN", user_only=(dec["yz_denominator"] == "users"), blank_is_no=False)
        a["SAGLIK_CIHAZ"] = yn("NESN_INT_AKILL_SGLK_CHZ") if yr == 2026 else np.nan
        a["AKILLI_SAAT"] = yn("NESN_INT_AKILL_SAAT") if yr == 2026 else np.nan
        # ---- e-government (12 months) and descriptive indicators
        a["EDEVLET_RANDEVU"] = yn("EDEVLET_RANDEVU_REZRVSYN", blank_is_no=False)
        a["EDEVLET_KAYIT"] = yn("EDEVLET_KSL_BILGI_ERISME")
        a["VEKALET"] = yn("EDEVLET_TALEP_HIC_BASKASI")
        a["ENGEL"] = yn("EDEVLET_TALEP_HIC_ENGEL") if yr == 2026 else np.nan
        # ---- activity repertoire and digital skills (users)
        rep_items = pd.concat([yn(v) for v in cfg["vars"]["repertoire"]], axis=1)
        a["REPERTUAR"] = rep_items.sum(axis=1, min_count=len(cfg["vars"]["repertoire"]))
        sk_items = pd.concat([yn(v) for v in cfg["vars"]["skills"]], axis=1)
        sk_items.columns = cfg["vars"]["skills"]
        for child, parent in (cfg["vars"].get("skill_conditional") or {}).items():
            if child in sk_items.columns and parent in sk_items.columns:
                sk_items.loc[sk_items[parent] == 0, child] = 0.0
        a["BECERI"] = sk_items.sum(axis=1, min_count=len(cfg["vars"]["skills"]))
        for c in sk_items.columns:
            a["SK_" + c] = sk_items[c]
        # ---- GALI
        if yr == 2026 and "FAALIYET_SINIRLAMA_DURUM" in raw.columns:
            g = num(raw["FAALIYET_SINIRLAMA_DURUM"])
            s = num(raw["FAALIYET_SINIRLAMA_SURE"])
            gali = pd.Series(np.nan, index=raw.index)
            gali[g == 6] = 0
            if dec["gali_rule"] == "eurostat":
                gali[(g == 5) & (s == 1)] = 1
                gali[(g == 4) & (s == 1)] = 2
                gali[g.isin([4, 5]) & (s == 2)] = 0
            else:
                gali[g == 5] = 1
                gali[g == 4] = 2
            a["GALI3"] = gali
        else:
            a["GALI3"] = np.nan
        a["GALI_ANY"] = (a["GALI3"] > 0).astype(float).where(a["GALI3"].notna())
        # ---- kovaryatlar
        a["AGE"] = num(raw["YAS"])
        a["AGEGRP"] = pd.cut(a["AGE"], bins=cfg["age"]["groups"], right=False,
                             labels=[f"{lo}-{hi - 1}" for lo, hi in zip(cfg["age"]["groups"][:-1],
                                                                       cfg["age"]["groups"][1:])]).astype(str)
        a["FEMALE"] = (num(raw["CINSIYET"]) == 2).astype(float)
        a["EDU"] = num(raw["OKUL_BITEN"]).map({int(k): v for k, v in cfg["education_map"].items()})
        a["ILLITERATE"] = (num(raw.get("OKUMA_YAZMA")) == 2).astype(float)
        inc_col = "HANE_AYLIK_GELIR_GRP_5"
        inc = num(raw[inc_col]) if inc_col in raw.columns else pd.Series(np.nan, index=raw.index)
        a["INCOME"] = as_code(inc.where(inc.isin(cfg["coding"]["income_valid"])))
        a["URBAN"] = as_code(raw["MAH_DEG_L1"])
        a["COMPUTER"] = pd.concat([num(raw[v]).isin(yes) for v in cfg["vars"]["computer"]
                                   if v in raw.columns], axis=1).any(axis=1).astype(float)
        hhb = num(raw["HHB"])
        a["HHSIZE"] = pd.cut(hhb, bins=[0, 1, 2, 4, 99], labels=["1", "2", "3-4", "5+"]).astype(str)
        a["MODE"] = as_code(raw["GORUSME_YONTEM_FERT"])
        frames.append(a)

    A = pd.concat(frames, ignore_index=True)
    # design variables must be complete; otherwise stop and show raw values
    bad = {c: int(A[c].isna().sum()) for c in ("REGION", "STRATUM", "PSU", "W")}
    if any(bad.values()):
        raise ValueError(f"Tasarım değişkenlerinde eksik var: {bad}. IBBS_1 ham değer örnekleri: "
                         f"{load(ctx.cache('raw_2026.pkl'))['IBBS_1'].drop_duplicates().head(15).tolist()}")
    for c in ("REGION", "URBAN", "MODE"):
        A[c] = A[c].replace("<NA>", np.nan)
    save(A, ctx.cache("analysis.pkl"))

    # ---- report: weighted prevalence and missingness — NO breakdown by GALI
    rows = []
    for k, lab in OUTCOME_LABELS.items():
        for yr in (2025, 2026):
            s = A[A.YEAR == yr]
            x = s[k]
            if x.notna().sum() == 0:
                continue
            rows.append(dict(cikti=k, etiket=lab, yil=yr, n_tanimli=int(x.notna().sum()),
                             agirlikli_yayginlik=wmean(x, s["W"])))
    rep.h("Türetilmiş çıktıların genel yaygınlığı (tanımlı payda)")
    rep.table(pd.DataFrame(rows), "yayginlik")
    rep.p("Etkinlik çıktılarının paydası internet kullanıcılarıdır (INTERNET hariç). "
          "Tüm nüfus paydası için aşamalar `INTERNET=0 → 0` dönüşümünü kendisi yapar.")

    miss = []
    for c in ["GALI3", "EDU", "INCOME", "URBAN", "REGION", "HHSIZE", "MODE", "REPERTUAR", "BECERI", "W"]:
        for yr in (2025, 2026):
            s = A[A.YEAR == yr]
            miss.append(dict(degisken=c, yil=yr, eksik_oran=float(s[c].isna().mean())))
    rep.h("Kovaryat eksiklik oranları")
    rep.table(pd.DataFrame(miss), "eksiklik")
    rep.p("Not: VEKALET ve ENGEL (EDEVLET_TALEP_HIC_*) yalnızca e-devlette form göndermeyenlerin "
          "alt kümesine sorulur; paydaları küçüktür ve yalnızca betimsel kullanılır. "
          "REPERTUAR ve BECERI internet kullanmayanlarda yapısal olarak eksiktir. "
          "GALI3 2025'te tamamen eksiktir (soru yok).")

    # --- comparison of internet definitions with the TurkStat bulletin
    rows = []
    bulten = {2025: 90.9, 2026: 92.3}
    for yr in (2025, 2026):
        raw = load(ctx.cache(f"raw_{yr}.pkl"))
        age = num(raw["YAS"])
        raw = raw[age.between(cfg["age"]["min"], cfg["age"]["max"])]
        w = num(raw["FERT_FAKTOR"])
        son = num(raw["SON_KULLANIM_INTERNET"])
        rows.append(dict(
            yil=yr,
            son_3_ay_agirlikli=100 * wmean((son == 1).astype(float), w),
            son_12_ay_agirlikli=100 * wmean(son.isin([1, 2]).astype(float), w),
            hic_kullanmis_agirlikli=100 * wmean(son.isin([1, 2, 3]).astype(float), w),

            son_3_ay_agirliksiz=100 * float((son == 1).mean()),
            tuik_bulteni=bulten[yr],
        ))
    rep.h("İnternet kullanım oranı: TÜİK bülteni ile karşılaştırma (16–74, %)")
    rep.table(pd.DataFrame(rows), "internet_tanim_kontrol", digits=1)
    rep.p("Bülten değeri hangi sütuna eşitse TÜİK'in başlık tanımı odur. "
          "Ağırlıklı ve ağırlıksız sütunlar arasındaki fark, ağırlıklandırmanın etkisini gösterir.")
    rep.h("Kategorik kodlar (kontrol)")
    for c in ("REGION", "URBAN", "INCOME", "EDU", "HHSIZE", "MODE", "AGEGRP"):
        vc = A.groupby("YEAR")[c].value_counts(dropna=False).rename("n").reset_index()
        rep.p(f"**{c}:** " + "; ".join(f"{r.YEAR}·{r[c]}={r.n}" for _, r in vc.iterrows()))
    rep.checkpoint([
        "REGION 12 bölge, URBAN 3, MODE 2 düzey gösteriyor mu? 'nan' düzeyi olmamalı.",
        "2026 genel internet yaygınlığı TÜİK bülteniyle (%92,3) uyumlu mu? (ağırlıklı, 16–74)",
        "MHRS / KAYIT yaygınlıkları makul mü? (TÜİK bülteni: e-devlet kişisel bilgi %65,6, kamu randevusu %52,6)",
        "EDU, INCOME, GALI3 eksiklik oranı %5'i aşıyor mu? Aşıyorsa not edin (ön kayıttaki çoklu atama kuralı).",
        "Bu aşama yeni karar gerektirmez; 03_design çalıştırılabilir.",
    ])
    rep.write()


# =============================================================================
# S03_DESIGN
# =============================================================================
# Stage 03 — design: Rao–Wu replicate weights, block vs household deft ratio, MDE.
# 
# Checkpoint B: delta_pp (equivalence margin) and deft_ratio_2025.
# The MDE calculation does NOT use outcomes (p = 0.5 worst case),
# so the margin can be fixed before results are seen.




def _eq(series: pd.Series, level) -> pd.Series:
    """Level equality: numeric levels compared as numbers, others as text."""
    if isinstance(level, (int, float)) and not isinstance(level, bool):
        return pd.to_numeric(series, errors="coerce") == float(level)
    return series.astype(object).map(str) == str(level)


def asama_03(ctx: Ctx):
    rep = Report(ctx, "03_design", "Aşama 03 — Tasarım ve güç")
    A = load(ctx.cache("analysis.pkl"))
    W = A["W"].to_numpy(float)
    B = ctx.B

    # ---- main replicates: each year with its own strata/PSUs and an independent seed
    R = np.empty((len(A), B), np.float32)
    for i, yr in enumerate((2025, 2026)):
        m = (A.YEAR == yr).to_numpy()
        R[m] = rao_wu_weights(A.loc[m, "STRATUM"], A.loc[m, "PSU"], W[m], B, ctx.seed + i)
    save(R, ctx.cache(f"rep_B{B}.pkl"))

    # tek PSU'lu tabakalar
    psu_n = A.groupby(["YEAR", "STRATUM"])["PSU"].nunique()
    if len(psu_n) < 2:
        raise ValueError("Tabaka bulunamadı (STRATUM boş). 02'yi yeniden çalıştırın ve REGION kodlarını kontrol edin.")
    single = psu_n[psu_n < 2]
    rep.h("Tabakalar")
    rep.p(f"Tabaka sayısı: {len(psu_n)} · tek PSU'lu tabaka: {len(single)} "
          + ("(bunlar varyansa katkı vermez!)" if len(single) else ""))
    t = psu_n.rename("PSU_sayisi").reset_index()
    rep.table(t, "tabaka_psu")

    # ---- 2026: block vs household clusters (covariate means only — independent of outcomes)
    m26 = (A.YEAR == 2026).to_numpy()
    A26 = A[m26].reset_index(drop=True)
    W26 = W[m26]
    R_blk = R[m26]
    R_hh = rao_wu_weights(A26["STRATUM"], A26["PSU_HH"], W26, B, ctx.seed + 99)
    probe = {
        "kadin": A26["FEMALE"].to_numpy(float),
        "yuksekogretim": (A26["EDU"] == "4_yuksekogretim").astype(float).to_numpy(),
        "yas_55plus": (A26["AGE"] >= 55).astype(float).to_numpy(),
        "kentsel_yogun": (A26["URBAN"] == "3").astype(float).to_numpy(),
        "hanede_bilgisayar": A26["COMPUTER"].to_numpy(float),
    }

    def means(w):
        return pd.Series({k: np.nansum(v * w) / np.sum(w[~np.isnan(v)]) for k, v in probe.items()})

    pt, rb = replicate(means, W26, R_blk)
    _, rh = replicate(means, W26, R_hh)
    sb, sh = summarize(pt, rb), summarize(pt, rh)
    cmp = pd.DataFrame({"ortalama": pt, "se_blok": sb["se"], "se_hane": sh["se"]})
    cmp["oran_blok_hane"] = cmp["se_blok"] / cmp["se_hane"]
    rep.h("2026: blok ve hane kümesiyle standart hata oranı")
    rep.p("Sonuç değişkenleri değil, kovaryat ortalamaları kullanıldı. Oran > 1 ise hane "
          "kümesi SE'yi küçük tahmin ediyor; 2025 (yalnızca hane kümesi) için bu oran "
          "`deft_ratio_2025` olarak sağlamlıkta kullanılabilir. Kovaryatlar küme içi "
          "benzerliği yüksek değişkenler (bölge, kentsellik) için oranı abartabilir; medyanı esas alın.")
    rep.table(cmp, "deft_orani", index=True)
    rep.p(f"**Medyan oran: {cmp['oran_blok_hane'].median():.3f}**")

    # ---- MDE (group × year interaction), assuming p = 0.5
    deff = ctx.cfg["trend"]["assumed_cluster_deff"]
    rows = []
    for var, (lvl, ref) in ctx.cfg["trend"]["contrasts"].items():
        for denom in ("all", "users"):
            se2 = 0.0
            ns = {}
            for yr in (2025, 2026):
                s = A[A.YEAR == yr]
                if denom == "users":
                    s = s[s.INTERNET == 1]
                kd = kish_deff(s["W"])
                n1 = int(_eq(s[var], lvl).sum())
                n0 = int(_eq(s[var], ref).sum())
                ns[yr] = (n1, n0)
                if min(n1, n0) == 0:
                    se2 = np.nan
                    continue
                se2 += 0.25 * kd * deff * (1 / n1 + 1 / n0)
            se = np.sqrt(se2)
            rows.append(dict(karsilastirma=f"{var}: {lvl} vs {ref}", payda=denom,
                             n_2025=ns[2025], n_2026=ns[2026],
                             se_fark_farki_pp=100 * se, MDE_80guc_pp=100 * 2.8 * se,
                             TOST_icin_gerekli_min_delta_pp=100 * 1.645 * se))
    rep.h("Saptanabilir en küçük etki (fark-farkı, yüzde puan)")
    rep.p(f"Varsayım: p = 0,5 (en kötü durum), Kish deff gözlenen, küme deff = {deff} "
          "(config). Sonuç değişkeni kullanılmadı.")
    rep.table(pd.DataFrame(rows), "mde")
    rep.p("Eşdeğerlik iddiası ancak gerçek fark-farkı ≈ 0 iken, `delta` en az "
          "`TOST_icin_gerekli_min_delta_pp` kadar büyükse mümkündür. `delta_pp` değerini "
          "bu tablonun en kötü satırına ve kuramsal anlamlılığa göre seçin.")

    rep.checkpoint([
        "Tek PSU'lu tabaka var mı? Varsa TÜİK tasarımıyla uyumsuzluk olabilir; not edin.",
        "`deft_ratio_2025` → medyan oranı yazın (sağlamlıkta 2025 SE'leri bununla şişirilecek).",
        "`delta_pp` → MDE tablosuna bakarak, SONUÇ GÖRMEDEN belirleyin. Bu değeri OSF ön kaydına yazın.",
        "Ön kaydı şimdi OSF'ye yükleyin: buradan sonraki aşamalar sonuç değişkenlerini GALI ve gruplara göre gösterir.",
    ])
    rep.write()


# =============================================================================
# S04_DESCRIPTIVES
# =============================================================================
# Stage 04 — descriptive cascade (2026): each step by GALI level, two denominators.
# 
# Checkpoint C: gali_merge_levels, drop_outcomes.



CASCADE = ["INTERNET", "BILGI_F", "BILGI_R", "MHRS", "KAYIT", "KONSULT",
           "YZ", "SAGLIK_CIHAZ", "AKILLI_SAAT", "EDEVLET_RANDEVU", "EDEVLET_KAYIT"]
DESCRIPTIVE_ONLY = ["VEKALET", "ENGEL"]


def all_denominator(A: pd.DataFrame, k: str) -> np.ndarray:
    """Whole 16–74 denominator: internet non-users = 0."""
    x = A[k].to_numpy(float).copy()
    if k != "INTERNET":
        x[(A["INTERNET"] == 0).to_numpy()] = 0.0
    return x


def asama_04(ctx: Ctx):
    ctx.log_decisions("04_descriptives")
    rep = Report(ctx, "04_descriptives", "Aşama 04 — Betimsel kaskad (2026)")
    A_all = load(ctx.cache("analysis.pkl"))
    R_all = load(ctx.cache(f"rep_B{ctx.B}.pkl"))
    m = (A_all.YEAR == 2026).to_numpy()
    A = A_all[m].reset_index(drop=True)
    R = R_all[m]
    W = A["W"].to_numpy(float)
    gali = A["GALI3"].to_numpy(float)

    cells = []   # (label, value vector, mask)
    for k in CASCADE + DESCRIPTIVE_ONLY:
        if A[k].notna().sum() == 0:
            continue
        for denom in ("kullanicilar", "tum_16_74"):
            if (denom == "tum_16_74" and k in DESCRIPTIVE_ONLY) or (denom == "kullanicilar" and k == "INTERNET"):
                continue
            x = A[k].to_numpy(float) if denom == "kullanicilar" else all_denominator(A, k)
            for g in (0, 1, 2, "toplam"):
                mask = ~np.isnan(x) & ((~np.isnan(gali)) if g == "toplam" else (gali == g))
                cells.append((f"{k}|{denom}|{g}", x, mask))

    def fn(w):
        return pd.Series({lab: np.sum(w[mk] * x[mk]) / np.sum(w[mk]) for lab, x, mk in cells})

    point, reps = replicate(fn, W, R, progress="betimsel")
    s = summarize(point, reps)
    idx = s.index.str.split("|", expand=True)
    s["cikti"], s["payda"], s["GALI"] = idx.get_level_values(0), idx.get_level_values(1), idx.get_level_values(2)
    s["n_payda"] = [int(mk.sum()) for _, _, mk in cells]
    s["n_evet"] = [int(np.nansum(x[mk])) for _, x, mk in cells]
    s = s.reset_index(drop=True)
    for c in ("tahmin", "se", "ga_alt", "ga_ust"):
        s[c] = 100 * s[c]
    s["kucuk_hucre"] = s["n_evet"] < ctx.cfg["min_cell_n"]
    cols = ["cikti", "payda", "GALI", "n_payda", "n_evet", "tahmin", "se", "ga_alt", "ga_ust", "kucuk_hucre"]
    s[cols].to_csv(rep.dir / "kaskad_uzun.csv", index=False, encoding="utf-8-sig")

    rep.p("GALI: 0 = kısıtlı değil, 1 = kısıtlı (ciddi değil), 2 = ciddi kısıtlı. "
          "Değerler ağırlıklı yüzde; GA bootstrap %95.")
    for denom in ("kullanicilar", "tum_16_74"):
        t = s[s.payda == denom].copy()
        t["hucre"] = t.apply(lambda r: f"{r.tahmin:.1f} [{r.ga_alt:.1f}–{r.ga_ust:.1f}]"
                             + (" †" if r.kucuk_hucre else ""), axis=1)
        wide = t.pivot(index="cikti", columns="GALI", values="hucre")
        nn = t.pivot(index="cikti", columns="GALI", values="n_payda")
        wide = wide.join(nn.add_prefix("n_"))
        wide.insert(0, "etiket", [OUTCOME_LABELS.get(i, i) for i in wide.index])
        order = [c for c in CASCADE + DESCRIPTIVE_ONLY if c in wide.index]
        rep.h(f"Kaskad — payda: {'internet kullanıcıları' if denom == 'kullanicilar' else 'tüm 16–74'}")
        rep.table(wide.loc[order], f"kaskad_{denom}", index=True)
    rep.p(f"† = evet sayısı < {ctx.cfg['min_cell_n']} (ağırlıksız); tahmin kararsız.")

    # GALI distribution and interview mode
    gd = pd.DataFrame({"GALI": gali, "W": W, "MODE": A["MODE"]}).dropna(subset=["GALI"])
    dist = gd.groupby("GALI").agg(n=("W", "size"), agirlikli_yuzde=("W", "sum"))
    dist["agirlikli_yuzde"] = 100 * dist["agirlikli_yuzde"] / dist["agirlikli_yuzde"].sum()
    dist["telefon_orani"] = gd.groupby("GALI").apply(
        lambda d: np.sum(d.W * (d.MODE == "3")) / np.sum(d.W), include_groups=False)
    rep.h("GALI dağılımı ve görüşme modu")
    rep.table(dist, "gali_dagilim", index=True)

    rep.checkpoint([
        "İşlem basamağında (MHRS, KAYIT, KONSULT) ciddi kısıtlı hücrelerde † var mı? "
        "Çoksa → `gali_merge_levels: true` (ön kayıttaki koşullu kural).",
        "KONSULT, SAGLIK_CIHAZ gibi düşük yaygınlıklı çıktılarda hücre çok küçükse → `drop_outcomes` listesine ekleyin "
        "(birincil ailedeki bir çıktıyı düşürmek ön kayıttan sapmadır: gerekçesi Claude ile birlikte KARARLAR yorumuna yazılır).",
        "VEKALET (başkası yaptı) ciddi kısıtlılarda belirgin yüksek mi? Tartışma bölümü için not alın.",
        "Mod: ciddi kısıtlılarda telefon oranı farklıysa 09'daki mod sağlamlığını öne çıkarın.",
    ])
    rep.write()


# =============================================================================
# S05_MODELS
# =============================================================================
# Stage 05 — H1–H4: access model and per-outcome models among users (2026).
# 
# A fully interacted stacked logit is equivalent to separate models per outcome.
# Covariance between outcomes is preserved because all models are re-estimated
# with the same weights in every replicate, so cross-outcome contrasts (H3) get valid SEs.
# Two specifications: 'toplam' (confounders only) and 'dogrudan' (+ resource mediators).




def exposure(ctx: Ctx):
    if ctx.dec.get("gali_merge_levels"):
        return "GALI_ANY", [1.0], 0.0
    return "GALI3", [1.0, 2.0], 0.0


def asama_05(ctx: Ctx):
    ctx.require("gali_merge_levels")
    ctx.log_decisions("05_models")
    cfg = ctx.cfg["models"]
    rep = Report(ctx, "05_models", "Aşama 05 — Erişim ve basamak modelleri (H1–H4)")
    A_all = load(ctx.cache("analysis.pkl"))
    R_all = load(ctx.cache(f"rep_B{ctx.B}.pkl"))
    m = (A_all.YEAR == 2026).to_numpy()
    A = A_all[m].reset_index(drop=True)
    R = R_all[m]
    W = A["W"].to_numpy(float)
    E, levels, ref = exposure(ctx)
    drop = set(ctx.dec.get("drop_outcomes") or [])
    outcomes = [k for k in cfg["user_outcomes"] if k not in drop and A[k].notna().sum() > 0]

    specs = {"toplam": cfg["total_rhs"], "dogrudan": cfg["total_rhs"] + " + " + cfg["direct_extra"]}
    models, ns = {}, []
    users = (A["INTERNET"] == 1).to_numpy()
    for spec, rhs in specs.items():
        models[(spec, "INTERNET")] = Logit(A, f"INTERNET ~ C({E}) + {rhs}")
        for k in outcomes:
            models[(spec, k)] = Logit(A, f"{k} ~ C({E}) + {rhs}", subset=users)
    for (spec, k), mod in models.items():
        ns.append(dict(belirtim=spec, cikti=k, n=len(mod.rows), n_evet=int(mod.y.sum()), parametre=len(mod.names)))
    fns = {key: ame_fn(mod, E, levels, ref) for key, mod in models.items()}

    steps = {s: [k for k in ks if k in outcomes] for s, ks in cfg["steps"].items()}

    def fn(w):
        out = {}
        for (spec, k), f in fns.items():
            for l, v in f(w).items():
                out[f"{spec}|{k}|{l}"] = v
        for spec in specs:
            for l in levels:
                mean = {s: np.mean([out[f"{spec}|{k}|{l}"] for k in ks]) if ks else np.nan
                        for s, ks in steps.items()}
                out[f"{spec}|H3a_bilgi_eksi_islem|{l}"] = mean["bilgi"] - mean["islem"]
                out[f"{spec}|H3b_bilgi_eksi_sinir|{l}"] = mean["bilgi"] - mean["sinir"]
        return pd.Series(out)

    point, reps = replicate(fn, W, R, progress="modeller")
    s = summarize(point, reps)
    idx = s.index.str.split("|", expand=True)
    s["belirtim"], s["cikti"], s["GALI_duzey"] = (idx.get_level_values(i) for i in range(3))
    s = s.reset_index(drop=True)

    # multiple testing: Holm for the primary family (total specification), BH for the rest
    prim = s[(s.belirtim == "toplam") & s.cikti.isin(cfg["primary_family"])]
    s["p_holm"] = np.nan
    s.loc[prim.index, "p_holm"] = holm(prim["p"]).values
    sec = s[~s.index.isin(prim.index)]
    s["p_bh"] = np.nan
    s.loc[sec.index, "p_bh"] = bh(sec["p"]).values
    s["aile"] = np.where(s.index.isin(prim.index), "birincil", "ikincil")
    for c in ("tahmin", "se", "ga_alt", "ga_ust"):
        s[c] = 100 * s[c]

    rep.p(f"Maruziyet: {E} (referans {ref}); düzeyler {levels}. AME = yüzde puan. "
          "INTERNET tüm 16–74 yaşta, diğer çıktılar internet kullanıcılarında.")
    rep.h("Model örneklemleri")
    rep.table(pd.DataFrame(ns), "model_n", digits=0)
    for spec in specs:
        t = s[s.belirtim == spec].copy()
        t.insert(0, "etiket", t["cikti"].map(lambda k: OUTCOME_LABELS.get(k, k)))
        rep.h(f"AME — belirtim: {spec}")
        rep.table(t[["etiket", "cikti", "GALI_duzey", "tahmin", "ga_alt", "ga_ust", "p",
                     "p_holm", "p_bh", "aile", "gecersiz_rep"]], f"ame_{spec}")
    t = s[s.cikti.str.startswith("H3")]
    rep.h("H3 kontrastları (merkezî hipotez)")
    rep.p("H3a = ort.AME(bilgi) − ort.AME(işlem); H3b = ort.AME(bilgi) − ort.AME(sınır). "
          "Ön kayıtlı yön: > 0.")
    rep.table(t[["belirtim", "cikti", "GALI_duzey", "tahmin", "ga_alt", "ga_ust", "p", "p_holm"]], "h3")

    # total minus direct (share via mediators, descriptive)
    piv = s.pivot_table(index=["cikti", "GALI_duzey"], columns="belirtim", values="tahmin")
    piv["fark_pp"] = piv["toplam"] - piv["dogrudan"]
    rep.h("Toplam − doğrudan (kaynak aracılarıyla açıklanan kısım, betimsel)")
    rep.table(piv, "toplam_dogrudan", index=True)

    rep.checkpoint([
        "`gecersiz_rep` > 0 olan satır var mı? (Bootstrap'ta yakınsamama; ilgili çıktıyı inceleyin.)",
        "H1: INTERNET AME'leri negatif mi? H2: BILGI_F / BILGI_R pozitif mi? Yön ön kayıtla uyumlu mu?",
        "H3: kontrastların GA'sı 0'ı dışlıyor mu? Holm düzeltmeli p'yi raporlayın.",
        "Toplam − doğrudan farkı büyükse tartışmada 'kaynak aracılığı' diye yorumlayın; nedensel dil kullanmayın.",
        "Bu aşama karar gerektirmez; 06_hurdle ve 07_khb bağımsız çalışır.",
    ])
    rep.write()


# =============================================================================
# S06_HURDLE
# =============================================================================
# Stage 06 — decomposition of the GALI gap into access and conditional-use components.
# 
# g-computation: the same population (2026, 16–74, complete covariates) is "assigned" each GALI level.
#   pA(g) = P(internet | X, g),  pC(g) = P(outcome | internet, X, g)
#   P(g)  = Σ w pA(g) pC(g) / Σ w
# Decomposition (exact identity at the individual level; order: access first):
#   ΔP = Σ w (pA1 − pA0) pC0 / Σw   [access]  +  Σ w pA1 (pC1 − pC0) / Σw   [conditional]
#   ln RR = ln(P* / P0) [access] + ln(P1 / P*) [conditional],  P* = Σ w pA1 pC0 / Σ w
# The reverse order (conditional first) is reported as a robustness check.




def asama_06(ctx: Ctx):
    ctx.require("gali_merge_levels")
    ctx.log_decisions("06_hurdle")
    cfg = ctx.cfg["models"]
    rep = Report(ctx, "06_hurdle", "Aşama 06 — Erişim / koşullu kullanım ayrıştırması (T1)")
    A_all = load(ctx.cache("analysis.pkl"))
    R_all = load(ctx.cache(f"rep_B{ctx.B}.pkl"))
    m = (A_all.YEAR == 2026).to_numpy()
    A = A_all[m].reset_index(drop=True)
    R = R_all[m]
    W = A["W"].to_numpy(float)
    E, levels, ref = exposure(ctx)
    drop = set(ctx.dec.get("drop_outcomes") or [])
    outcomes = [k for k in cfg["hurdle_outcomes"] if k not in drop and A[k].notna().sum() > 0]
    rhs = cfg["total_rhs"]

    acc = Logit(A, f"INTERNET ~ C({E}) + {rhs}")
    pop = acc.rows                                   # standardisation population
    XA = {g: acc.cf_matrix(pop, **{E: g}) for g in [ref] + levels}
    users = (A["INTERNET"] == 1).to_numpy()
    cond = {}
    for k in outcomes:
        mod = Logit(A, f"{k} ~ C({E}) + {rhs}", subset=users)
        cond[k] = (mod, {g: mod.cf_matrix(pop, **{E: g}) for g in [ref] + levels})

    def fn(w_full):
        w = w_full[pop]
        sw = w.sum()
        bA = acc.fit(w_full)
        pA = {g: expit(X @ bA) for g, X in XA.items()}
        out = {}
        for g in levels:
            out[f"erisim|P_erisim|{g}"] = np.sum(w * pA[g]) / sw
        out[f"erisim|P_erisim|{ref}"] = np.sum(w * pA[ref]) / sw
        for k, (mod, XC) in cond.items():
            bC = mod.fit(w_full)
            pC = {g: expit(X @ bC) for g, X in XC.items()}
            P0 = np.sum(w * pA[ref] * pC[ref]) / sw
            out[f"{k}|P_toplam|{ref}"] = P0
            for g in levels:
                P1 = np.sum(w * pA[g] * pC[g]) / sw
                Pstar = np.sum(w * pA[g] * pC[ref]) / sw         # access changed, conditional fixed
                Pstar_r = np.sum(w * pA[ref] * pC[g]) / sw       # reverse order
                out[f"{k}|P_toplam|{g}"] = P1
                out[f"{k}|dP_toplam|{g}"] = P1 - P0
                out[f"{k}|dP_erisim|{g}"] = Pstar - P0
                out[f"{k}|dP_kosullu|{g}"] = P1 - Pstar
                out[f"{k}|lnRR_toplam|{g}"] = np.log(P1 / P0)
                out[f"{k}|lnRR_erisim|{g}"] = np.log(Pstar / P0)
                out[f"{k}|lnRR_kosullu|{g}"] = np.log(P1 / Pstar)
                out[f"{k}|pay_erisim|{g}"] = np.log(Pstar / P0) / np.log(P1 / P0)
                out[f"{k}|pay_erisim_ters_sira|{g}"] = np.log(P1 / Pstar_r) / np.log(P1 / P0)
                out[f"{k}|P_kosullu_std|{g}"] = np.sum(w * pC[g]) / sw
            out[f"{k}|P_kosullu_std|{ref}"] = np.sum(w * pC[ref]) / sw
        return pd.Series(out)

    point, reps = replicate(fn, W, R, progress="hurdle")
    s = summarize(point, reps)
    idx = s.index.str.split("|", expand=True)
    s["cikti"], s["olcu"], s["GALI"] = (idx.get_level_values(i) for i in range(3))
    s = s.reset_index(drop=True)
    s.to_csv(rep.dir / "hurdle_uzun.csv", index=False, encoding="utf-8-sig")

    rep.p(f"Standartlaştırma nüfusu: 2026, 16–74, n = {len(pop)}. Model: toplam etki belirtimi. "
          "Yüzde puan değerleri ×100. `pay_erisim` = ln RR erişim / ln RR toplam; "
          "toplam açık 0'a yakınsa veya bileşenler zıt işaretliyse pay anlamsızdır; "
          "bu durumda yalnızca bileşenlerin pp değerlerini yorumlayın.")
    for g in levels:
        rows = []
        for k in outcomes:
            def get(o):
                r = s[(s.cikti == k) & (s.olcu == o) & (s.GALI == str(g))].iloc[0]
                return r
            dt, da, dc = get("dP_toplam"), get("dP_erisim"), get("dP_kosullu")
            pe, pr = get("pay_erisim"), get("pay_erisim_ters_sira")
            rt = get("lnRR_toplam")
            rows.append(dict(
                etiket=OUTCOME_LABELS.get(k, k),
                acik_toplam_pp=f"{100 * dt.tahmin:.1f} [{100 * dt.ga_alt:.1f}; {100 * dt.ga_ust:.1f}]",
                erisim_pp=f"{100 * da.tahmin:.1f} [{100 * da.ga_alt:.1f}; {100 * da.ga_ust:.1f}]",
                kosullu_pp=f"{100 * dc.tahmin:.1f} [{100 * dc.ga_alt:.1f}; {100 * dc.ga_ust:.1f}]",
                RR_toplam=f"{np.exp(rt.tahmin):.2f} [{np.exp(rt.ga_alt):.2f}; {np.exp(rt.ga_ust):.2f}]",
                pay_erisim=f"{pe.tahmin:.2f} [{pe.ga_alt:.2f}; {pe.ga_ust:.2f}]",
                pay_erisim_ters_sira=f"{pr.tahmin:.2f}",
            ))
        rep.h(f"GALI = {g} vs {ref}")
        rep.table(pd.DataFrame(rows), f"hurdle_ozet_GALI{g}")

    rep.checkpoint([
        "Erişim ve koşullu bileşenler toplam açığa eşit mi? (Yapı gereği eşit olmalı; değilse hata bildirin.)",
        "Pay ile ters sıra payı çok farklıysa (> 0,15), ayrıştırmanın sıraya duyarlı olduğunu raporlayın.",
        "Toplam açığın GA'sı 0'ı içeriyorsa payı yorumlamayın; yalnızca bileşenlerin pp değerlerini verin.",
    ])
    rep.write()


# =============================================================================
# S07_KHB
# =============================================================================
# Stage 07 — KHB (Karlson–Holm–Breen) decomposition, internet users, 2026.
# 
# Exposure: GALI_ANY (limited = 1). Mediators: REPERTUAR (repertoire), BECERI (skills).
#   b_reduced : Y ~ GALI + C + resid(Z | GALI, C)    (same scale)
#   b_full    : Y ~ GALI + C + Z
#   diff = b_reduced − b_full = Σ_j θ_j δ_j   (θ_j: coefficient of Z_j in the full model,
#                                             δ_j: GALI coefficient in the regression Z_j ~ GALI + C)
# Interpretation: "share explained"; not a causal mediation claim with cross-sectional data.




def asama_07(ctx: Ctx):
    ctx.log_decisions("07_khb")
    cfg = ctx.cfg["models"]
    rep = Report(ctx, "07_khb", "Aşama 07 — KHB ayrıştırması (T2)")
    A_all = load(ctx.cache("analysis.pkl"))
    R_all = load(ctx.cache(f"rep_B{ctx.B}.pkl"))
    m = (A_all.YEAR == 2026).to_numpy()
    A = A_all[m].reset_index(drop=True)
    R = R_all[m]
    W = A["W"].to_numpy(float)
    Zn = cfg["khb_mediators"]
    drop = set(ctx.dec.get("drop_outcomes") or [])
    outcomes = [k for k in cfg["khb_outcomes"] if k not in drop and A[k].notna().sum() > 0]
    users = (A["INTERNET"] == 1).to_numpy()

    prepared = {}
    for k in outcomes:
        full = Logit(A, f"{k} ~ GALI_ANY + {cfg['total_rhs']} + " + " + ".join(Zn), subset=users)
        rows = full.rows
        dm = patsy.dmatrix(f"GALI_ANY + {cfg['total_rhs']}", A.iloc[rows])
        g_idx = dm.design_info.column_names.index("GALI_ANY")
        Xc = np.asarray(dm, float)
        Z = A.iloc[rows][Zn].to_numpy(float)
        prepared[k] = (rows, Xc, Z, full.y, g_idx)

    def fn(w_full):
        out = {}
        for k, (rows, Xc, Z, y, g_idx) in prepared.items():
            w = w_full[rows]
            deltas, resid = [], np.empty_like(Z)
            for j in range(Z.shape[1]):
                gam = wls(Xc, Z[:, j], w)
                deltas.append(gam[g_idx])
                resid[:, j] = Z[:, j] - Xc @ gam
            b_red = irls_logit(np.column_stack([Xc, resid]), y, w)
            b_full = irls_logit(np.column_stack([Xc, Z]), y, w)
            b_naive = irls_logit(Xc, y, w)
            red, ful = b_red[g_idx], b_full[g_idx]
            out[f"{k}|b_naif"] = b_naive[g_idx]
            out[f"{k}|b_azaltilmis"] = red
            out[f"{k}|b_tam"] = ful
            out[f"{k}|fark"] = red - ful
            out[f"{k}|aciklanan_pay"] = (red - ful) / red
            out[f"{k}|olcek_faktoru"] = red / b_naive[g_idx]
            thetas = b_full[Xc.shape[1]:]
            for j, zn in enumerate(Zn):
                out[f"{k}|pay_{zn}"] = thetas[j] * deltas[j] / red
        return pd.Series(out)

    point, reps = replicate(fn, W, R, progress="KHB")
    s = summarize(point, reps)
    idx = s.index.str.split("|", expand=True)
    s["cikti"], s["olcu"] = idx.get_level_values(0), idx.get_level_values(1)
    s = s.reset_index(drop=True)
    s.insert(0, "etiket", s["cikti"].map(lambda k: OUTCOME_LABELS.get(k, k)))
    rep.p("Katsayılar log-odds ölçeğinde. `aciklanan_pay` = (b_azaltılmış − b_tam) / b_azaltılmış. "
          "`olcek_faktoru` ≠ 1 ise naif karşılaştırma (b_naif − b_tam) ölçek değişimiyle karışırdı. "
          "Toplam etki 0'a yakınsa pay kararsızdır.")
    rep.table(s[["etiket", "olcu", "tahmin", "ga_alt", "ga_ust", "p", "gecersiz_rep"]], "khb")
    rep.checkpoint([
        "b_azaltılmış'ın GA'sı 0'ı içeriyorsa açıklanan payı yorumlamayın.",
        "Pay > 1 veya < 0 ise aracılar ters yönde etkiliyor (tutarsız aracılık); metinde açıkça belirtin.",
        "BECERI ve REPERTUAR etkinlik ölçüsüdür: 'mekanizma' değil 'açıklanan pay' dili kullanın (AGReMA).",
    ])
    rep.write()


# =============================================================================
# S08_TREND
# =============================================================================
# Stage 08 — H5–H6 (secondary): change in gaps 2025→2026, pooled cross-sections.
# 
# For each outcome: y ~ C(YEAR) * (EDU + FEMALE + AGEGRP) + REGION + URBAN.
# Gap(t) = mean over year-t rows of [p(group = l) − p(group = ref)].
# Difference-in-differences = gap(2026) − gap(2025). Equivalence: TOST (90% CI within ±delta).
# Also DFL: reweighting 2025 to the 2026 composition separates composition from behaviour.
# Variance: years independent → var = var(2025 replicates) + var(2026 replicates);
# for robustness the 2025 component is inflated by deft_ratio_2025².




def asama_08(ctx: Ctx):
    ctx.require("delta_pp", "deft_ratio_2025", "yz_denominator")
    ctx.log_decisions("08_trend")
    tcfg = ctx.cfg["trend"]
    rep = Report(ctx, "08_trend", "Aşama 08 — 2025→2026 değişim (H5–H6, ikincil)")
    A = load(ctx.cache("analysis.pkl"))
    R = load(ctx.cache(f"rep_B{ctx.B}.pkl"))
    W = A["W"].to_numpy(float)
    y25 = (A.YEAR == 2025).to_numpy()
    R25, R26 = R.copy(), R.copy()
    R25[~y25] = W[~y25, None]        # only 2025 changes
    R26[y25] = W[y25, None]          # only 2026 changes
    del R

    outcomes = dict(tcfg["outcomes"])
    if ctx.dec["yz_denominator"] == "all":
        outcomes["YZ"] = "all"
    contrasts = tcfg["contrasts"]
    users = (A["INTERNET"] == 1).to_numpy()

    # derived columns for the whole-population denominator
    D = A.copy()
    for k, denom in outcomes.items():
        if denom == "all" and k != "INTERNET":
            D[k] = D[k].where(D["INTERNET"] != 0, 0.0)

    prepared = {}
    for k, denom in outcomes.items():
        sub = None if denom == "all" else users
        mod = Logit(D, f"{k} ~ {tcfg['rhs']}", subset=sub)
        rows = mod.rows
        yr_rows = {yr: rows[(D["YEAR"].to_numpy()[rows] == yr)] for yr in (2025, 2026)}
        cf = {}
        for var, (lvl, ref) in contrasts.items():
            for yr in (2025, 2026):
                cf[(var, yr)] = (mod.cf_matrix(yr_rows[yr], **{var: lvl}),
                                 mod.cf_matrix(yr_rows[yr], **{var: ref}))
        # DFL
        dfl_d = patsy.dmatrix(tcfg["dfl_vars"], D.iloc[rows], return_type="dataframe")
        dfl_rows = rows[np.isin(rows, dfl_d.index.to_numpy())]
        Xd = dfl_d.loc[dfl_rows].to_numpy(float)
        yd = (D["YEAR"].to_numpy()[dfl_rows] == 2026).astype(float)
        prepared[k] = (mod, yr_rows, cf, dfl_rows, Xd, yd)

    def fn(w_full):
        out = {}
        for k, (mod, yr_rows, cf, dfl_rows, Xd, yd) in prepared.items():
            beta = mod.fit(w_full)
            for yr in (2025, 2026):
                r = yr_rows[yr]
                out[f"{k}|yayginlik|{yr}"] = np.sum(w_full[r] * D[k].to_numpy(float)[r]) / np.sum(w_full[r])
            for var in contrasts:
                gaps = {}
                for yr in (2025, 2026):
                    X1, X0 = cf[(var, yr)]
                    w = w_full[yr_rows[yr]]
                    gaps[yr] = np.sum(w * (expit(X1 @ beta) - expit(X0 @ beta))) / np.sum(w)
                    out[f"{k}|acik_{var}|{yr}"] = gaps[yr]
                out[f"{k}|acik_{var}|fark_farki"] = gaps[2026] - gaps[2025]
            # DFL
            wd = w_full[dfl_rows]
            g = irls_logit(Xd, yd, wd)
            p = expit(Xd @ g)
            m25 = yd == 0
            psi = p[m25] / (1 - p[m25])
            yv = D[k].to_numpy(float)[dfl_rows]
            p25 = np.sum(wd[m25] * yv[m25]) / np.sum(wd[m25])
            p26 = np.sum(wd[~m25] * yv[~m25]) / np.sum(wd[~m25])
            p25cf = np.sum(wd[m25] * psi * yv[m25]) / np.sum(wd[m25] * psi)
            out[f"{k}|DFL|toplam_degisim"] = p26 - p25
            out[f"{k}|DFL|bilesim"] = p25cf - p25
            out[f"{k}|DFL|davranis"] = p26 - p25cf
        return pd.Series(out)

    point, reps25 = replicate(fn, W, R25, progress="trend (2025 bileşeni)")
    _, reps26 = replicate(fn, W, R26, progress="trend (2026 bileşeni)")
    v25 = ((reps25[point.index] - point) ** 2).mean()
    v26 = ((reps26[point.index] - point) ** 2).mean()
    r = float(ctx.dec["deft_ratio_2025"])
    se = np.sqrt(v25 + v26)
    se_adj = np.sqrt(r**2 * v25 + v26)
    z = stats.norm.ppf(0.975)
    s = pd.DataFrame({"tahmin": point, "se": se, "se_deft": se_adj})
    s["ga_alt"], s["ga_ust"] = s.tahmin - z * s.se, s.tahmin + z * s.se
    s["p"] = 2 * stats.norm.sf(np.abs(s.tahmin / s.se))
    s["p_deft"] = 2 * stats.norm.sf(np.abs(s.tahmin / s.se_deft))
    delta = float(ctx.dec["delta_pp"]) / 100
    s["TOST_esdeger"] = (s.tahmin.abs() + 1.645 * s.se) < delta
    s["TOST_esdeger_deft"] = (s.tahmin.abs() + 1.645 * s.se_deft) < delta
    idx = s.index.str.split("|", expand=True)
    s["cikti"], s["olcu"], s["donem"] = (idx.get_level_values(i) for i in range(3))
    not_did = (s["donem"] != "fark_farki").to_numpy()
    for c in ("TOST_esdeger", "TOST_esdeger_deft"):   # TOST only for difference-in-differences
        s[c] = s[c].astype(object)
        s.loc[not_did, c] = ""
    s = s.reset_index(drop=True)
    for c in ("tahmin", "se", "se_deft", "ga_alt", "ga_ust"):
        s[c] = 100 * s[c]
    s.to_csv(rep.dir / "trend_uzun.csv", index=False, encoding="utf-8-sig")

    rep.p(f"Değerler yüzde puan. Eşdeğerlik eşiği delta = {ctx.dec['delta_pp']} pp; "
          f"deft oranı (2025) = {r}. Paydalar: {outcomes}.")
    rep.h("Yaygınlık ve DFL ayrıştırması")
    t = s[s.olcu.isin(["yayginlik", "DFL"])].copy()
    t.insert(0, "etiket", t.cikti.map(lambda k: OUTCOME_LABELS.get(k, k)))
    rep.table(t[["etiket", "olcu", "donem", "tahmin", "ga_alt", "ga_ust", "p"]], "yayginlik_dfl")
    rep.h("Açıklar ve fark-farkı (H5–H6)")
    t = s[s.olcu.str.startswith("acik_")].copy()
    t.insert(0, "etiket", t.cikti.map(lambda k: OUTCOME_LABELS.get(k, k)))
    rep.table(t[["etiket", "olcu", "donem", "tahmin", "ga_alt", "ga_ust", "p", "p_deft",
                 "TOST_esdeger", "TOST_esdeger_deft"]], "aciklar")
    rep.p("Karar kuralı (ön kayıt): fark-farkının GA'sı 0'ı dışlarsa değişim var; "
          "işaret normalleşme mi tabakalaşma mı olduğunu gösterir (açık işaretine göre yorumlayın). "
          "GA 0'ı içeriyor ve TOST_esdeger = True ise 'değişmedi'; ikisi de değilse 'belirsiz'. "
          "Not: sağlık bilgisi sorusunun liste bağlamı yıllar arasında değiştiği için değişim "
          "soru formu etkisinden ayrıştırılamaz.")
    rep.checkpoint([
        "p ve p_deft farklı sonuç veriyor mu? Veriyorsa ana metinde temkinli (p_deft) sürümü raporlayın.",
        "DFL bileşim payı ≈ 0 olmalı (bir yılda bileşim az değişir); büyükse ağırlık kalibrasyon farkını inceleyin.",
        "Eşdeğerlik sonuçlarını ön kayıttaki delta ile raporlayın; delta'yı sonuç gördükten sonra DEĞİŞTİRMEYİN.",
    ])
    rep.write()


# =============================================================================
# S09_HAM_AYARLI
# =============================================================================
# Stage 09 — from crude to adjusted gaps: how much does each covariate group explain?
# Denominator: all aged 16–74 (non-users = 0); this is the population-level ("inverse care") gap.
# Covariates are added in sequence; shares depend on the order (age first). All use the same
# bootstrap replicates, so differences between steps also have confidence intervals.

HAM_SPECS = [
    ("S0_ham", ""),
    ("S1_yas_cinsiyet", "cr(AGE, df=4, constraints='center') + FEMALE"),
    ("S2_egitim", "cr(AGE, df=4, constraints='center') + FEMALE + C(EDU)"),
    ("S3_kir_kent_bolge", "cr(AGE, df=4, constraints='center') + FEMALE + C(EDU) + C(URBAN) + C(REGION)"),
    ("S4_kaynaklar", "cr(AGE, df=4, constraints='center') + FEMALE + C(EDU) + C(URBAN) + C(REGION)"
                     " + C(INCOME) + COMPUTER + C(HHSIZE)"),
]
HAM_OUTCOMES = ["INTERNET", "BILGI_F", "BILGI_R", "MHRS", "KAYIT", "KONSULT", "YZ"]


def asama_09(ctx: Ctx):
    ctx.require("gali_merge_levels")
    ctx.log_decisions("09_ham_ayarli")
    rep = Report(ctx, "09_ham_ayarli", "Aşama 09 — Ham ve ayarlı faaliyet kısıtlılığı açığı (tüm 16–74)")
    A_all = load(ctx.cache("analysis.pkl"))
    R_all = load(ctx.cache(f"rep_B{ctx.B}.pkl"))
    m = (A_all.YEAR == 2026).to_numpy()
    A = A_all[m].reset_index(drop=True).copy()
    R = R_all[m]
    W = A["W"].to_numpy(float)
    E, levels, ref = exposure(ctx)
    drop = set(ctx.dec.get("drop_outcomes") or [])
    outs = [k for k in HAM_OUTCOMES if k not in drop and A[k].notna().sum() > 0]
    users = (A["INTERNET"] == 1).to_numpy()
    gl = [ref] + levels

    # Two-part model (as in stage 06): P(outcome) = P(access) × P(outcome | access).
    # In every specification the standardisation population = rows of the access model (2026, 16–74).
    parts = {}
    for spec, rhs in HAM_SPECS:
        tail = f" + {rhs}" if rhs else ""
        acc = Logit(A, f"INTERNET ~ C({E}){tail}")
        pop = acc.rows
        XA = {g: acc.cf_matrix(pop, **{E: g}) for g in gl}
        cond = {}
        for k in outs:
            if k == "INTERNET":
                continue
            mod = Logit(A, f"{k} ~ C({E}){tail}", subset=users)
            cond[k] = (mod, {g: mod.cf_matrix(pop, **{E: g}) for g in gl})
        parts[spec] = (acc, pop, XA, cond)
    specs = [s for s, _ in HAM_SPECS]

    def fn(w_full):
        out = {}
        for spec, (acc, pop, XA, cond) in parts.items():
            w = w_full[pop]
            sw = w.sum()
            bA = acc.fit(w_full)
            pA = {g: expit(X @ bA) for g, X in XA.items()}
            for l in levels:
                out[f"INTERNET|{spec}|{l}"] = np.sum(w * (pA[l] - pA[ref])) / sw
            for k, (mod, XC) in cond.items():
                bC = mod.fit(w_full)
                pC = {g: expit(XC[g] @ bC) for g in gl}
                P0 = np.sum(w * pA[ref] * pC[ref]) / sw
                for l in levels:
                    out[f"{k}|{spec}|{l}"] = np.sum(w * pA[l] * pC[l]) / sw - P0
        for k in outs:
            for l in levels:
                ham = out[f"{k}|S0_ham|{l}"]
                for s in specs[1:]:
                    out[f"{k}|pay_{s}|{l}"] = (ham - out[f"{k}|{s}|{l}"]) / ham
        return pd.Series(out)

    point, reps = replicate(fn, W, R, progress="ham/ayarlı")
    s = summarize(point, reps)
    idx = s.index.str.split("|", expand=True)
    s["cikti"], s["olcu"], s["GALI"] = (idx.get_level_values(i) for i in range(3))
    s = s.reset_index(drop=True)
    s.to_csv(rep.dir / "ham_ayarli_uzun.csv", index=False, encoding="utf-8-sig")

    rep.p("Her hücre: GALI düzeyinin kısıtlı olmayanlara göre standartlaştırılmış farkı (yüzde puan) "
          "[%95 GA]. İki parçalı model: P(erişim) × P(çıktı | erişim), 06 ile aynı yapı; "
          "S3 sütunu 06'daki 'acik_toplam_pp' ile birebir aynıdır, S0 ise ağırlıklı ham farktır (04 ile aynı). "
          "Payda tüm 16–74 yaş; internet kullanmayanlar her çıktıda 0. "
          "Kovaryatlar soldan sağa eklenir. S3 = ana analizdeki 'toplam etki' belirtimi; "
          "S4 = 'doğrudan' belirtim (gelir, bilgisayar, hane büyüklüğü aracı olabilir).")
    for l in levels:
        rows = []
        for k in outs:
            r = {"etiket": OUTCOME_LABELS.get(k, k)}
            for sp in specs:
                x = s[(s.cikti == k) & (s.olcu == sp) & (s.GALI == str(l))].iloc[0]
                r[sp] = f"{100 * x.tahmin:.1f} [{100 * x.ga_alt:.1f}; {100 * x.ga_ust:.1f}]"
            for sp in ("S2_egitim", "S3_kir_kent_bolge"):
                x = s[(s.cikti == k) & (s.olcu == f"pay_{sp}") & (s.GALI == str(l))].iloc[0]
                r[f"aciklanan_{sp}"] = f"%{100 * x.tahmin:.0f} [{100 * x.ga_alt:.0f}; {100 * x.ga_ust:.0f}]"
            rows.append(r)
        rep.h(f"GALI = {l} vs {ref}")
        rep.table(pd.DataFrame(rows), f"ham_ayarli_GALI{l}")
    rep.p("`aciklanan_S2_egitim` = ham açığın yaş, cinsiyet ve eğitimle açıklanan payı; "
          "`aciklanan_S3_kir_kent_bolge` = buna kır-kent ve bölge de eklenince. "
          "Pay sıraya bağlıdır (önce yaş girer); sırasız (Shapley) ayrıştırma istenirse eklenebilir. "
          "**Ham açığın (S0) GA'sı 0'ı içeriyorsa veya ham açık küçükse payı yorumlamayın** "
          "(payda sıfıra yakın olduğu için GA anlamsız genişler); yalnızca yüzde puan sütunlarını okuyun.")
    rep.checkpoint([
        "Ham açıkların (S0) 04'teki betimsel farklarla tutarlı olması gerekir (aynı payda).",
        "Açığın büyük kısmı S1'de (yaş) mı kapanıyor, S2'de (eğitim) mi? Politika mesajı buna göre kurulur.",
    ])
    rep.write()


# =============================================================================
# S10_SAGLAMLIK
# =============================================================================
# Stage 10 — sensitivity: primary AMEs and H3 contrasts from stage 05 under alternative assumptions.
# Variance uses the first B_SAGLAMLIK replicates (for speed; point estimates unaffected).

B_SAGLAMLIK = 200


def _inv_link(link):
    if link == "logit":
        return expit
    if link == "probit":
        return lambda e: stats.norm.cdf(np.clip(e, -8, 8))
    if link == "cloglog":
        return lambda e: 1 - np.exp(-np.exp(np.clip(e, -30, 3.5)))
    return lambda e: e


def irls_binomial(X, y, w, link, beta0=None, tol=1e-10, maxit=100):
    """Weighted IRLS for Bernoulli GLMs (probit / cloglog)."""
    k = X.shape[1]
    inv = _inv_link(link)
    if beta0 is None:
        mu = np.clip((y + 0.5) / 2, 0.01, 0.99)
        eta0 = stats.norm.ppf(mu) if link == "probit" else np.log(-np.log(1 - mu))
        beta = wls(X, eta0, w)
    else:
        beta = beta0.copy()
    dev = np.inf
    for _ in range(maxit):
        eta = X @ beta
        if link == "probit":
            mu = np.clip(inv(eta), 1e-10, 1 - 1e-10)
            d = np.maximum(stats.norm.pdf(np.clip(eta, -8, 8)), 1e-12)
        else:
            e = np.clip(eta, -30, 3.5)
            mu = np.clip(inv(e), 1e-10, 1 - 1e-10)
            d = np.maximum(np.exp(e - np.exp(e)), 1e-12)
        v = mu * (1 - mu)
        z = eta + (y - mu) / d
        ww = w * d**2 / v
        XtW = X.T * ww
        beta = np.linalg.solve(XtW @ X + 1e-8 * np.eye(k), XtW @ z)
        mu = np.clip(inv(X @ beta), 1e-10, 1 - 1e-10)
        dev_new = _deviance(y, mu, w)
        if abs(dev_new - dev) / (abs(dev_new) + 0.1) < tol:
            return beta
        dev = dev_new
    return np.full(k, np.nan)


class BinGLM(Logit):
    """Same interface as Logit; link: logit, probit, cloglog, identity (linear probability)."""

    def __init__(self, data, formula, subset=None, link="logit"):
        super().__init__(data, formula, subset)
        self.link = link
        self.inv = _inv_link(link)

    def fit(self, w_full):
        if self.link == "logit":
            return super().fit(w_full)
        w = w_full[self.rows]
        if self.link == "identity":
            return wls(self.X, self.y, w)
        beta = irls_binomial(self.X, self.y, w, self.link, self._beta_full)
        if self._beta_full is None:
            if not np.all(np.isfinite(beta)):
                raise RuntimeError(f"{self.link} modeli yakınsamadı.")
            self._beta_full = beta
        return beta


def ame_fn_g(model: BinGLM, exposure_: str, levels: list, ref):
    rows = model.rows
    Xref = model.cf_matrix(rows, **{exposure_: ref})
    Xl = {l: model.cf_matrix(rows, **{exposure_: l}) for l in levels}

    def fn(w_full):
        beta = model.fit(w_full)
        w = w_full[rows]
        p0 = model.inv(Xref @ beta)
        return pd.Series({f"{l}": np.sum(w * (model.inv(Xm @ beta) - p0)) / np.sum(w)
                          for l, Xm in Xl.items()})

    return fn


def _raw_2026_aligned(ctx, A26):
    raw = load(ctx.cache("raw_2026.pkl"))
    age = num(raw["YAS"])
    raw = raw[age.between(ctx.cfg["age"]["min"], ctx.cfg["age"]["max"])].reset_index(drop=True)
    if not (raw["BULTEN_NO"].to_numpy() == A26["BULTEN_NO"].to_numpy()).all():
        raise RuntimeError("Ham 2026 verisi ile analiz tablosu satır sırası uyuşmuyor; 02'yi yeniden çalıştırın.")
    return raw


def asama_10(ctx: Ctx):
    ctx.require("gali_merge_levels", "label_mapping")
    ctx.log_decisions("10_saglamlik")
    cfg = ctx.cfg["models"]
    rep = Report(ctx, "10_saglamlik", "Aşama 10 — Sağlamlık analizleri (H1–H3)")
    A_all = load(ctx.cache("analysis.pkl"))
    R_all = load(ctx.cache(f"rep_B{ctx.B}.pkl"))
    m = (A_all.YEAR == 2026).to_numpy()
    A = A_all[m].reset_index(drop=True).copy()
    Bs = min(B_SAGLAMLIK, R_all.shape[1])
    R = R_all[m][:, :Bs]
    W = A["W"].to_numpy(float)
    E, levels, ref = exposure(ctx)
    yes, no = ctx.cfg["coding"]["evet"], ctx.cfg["coding"]["hayir"]
    users = (A["INTERNET"] == 1).to_numpy()
    rhs = cfg["total_rhs"]
    drop = set(ctx.dec.get("drop_outcomes") or [])
    prim = [k for k in ["BILGI_F", "BILGI_R", "MHRS", "KAYIT", "YZ", "AKILLI_SAAT"] if k not in drop]
    steps = {s: [k for k in ks if k in prim] for s, ks in cfg["steps"].items()}

    # --- raw variables for the alternative item mapping and GALI rule
    raw = _raw_2026_aligned(ctx, A)
    alt_map_name = "HB" if ctx.dec["label_mapping"] == "HA" else "HA"
    for k, v in ctx.cfg["label_mappings"][alt_map_name].items():
        x = num(raw[v]) if v in raw.columns else pd.Series(np.nan, index=raw.index)
        s_ = pd.Series(np.nan, index=raw.index)
        s_[x.isin(yes)] = 1.0
        s_[x.isin(no)] = 0.0
        s_[~users] = np.nan
        A[k + "_ALT"] = s_.to_numpy()
    g = num(raw["FAALIYET_SINIRLAMA_DURUM"])
    alt_rule = "ignore_sure" if ctx.dec["gali_rule"] == "eurostat" else "eurostat"
    galt = pd.Series(np.nan, index=raw.index)
    galt[g == 6] = 0
    if alt_rule == "ignore_sure":
        galt[g == 5], galt[g == 4] = 1, 2
    else:
        sure = num(raw["FAALIYET_SINIRLAMA_SURE"])
        galt[(g == 5) & (sure == 1)], galt[(g == 4) & (sure == 1)] = 1, 2
        galt[g.isin([4, 5]) & (sure == 2)] = 0
    A["GALI3_ALT"] = galt.to_numpy()
    A["GALI_ANY_ALT"] = (A["GALI3_ALT"] > 0).astype(float).where(A["GALI3_ALT"].notna())

    # --- weight scenarios (replicate multipliers preserved)
    cap = np.quantile(W, 0.99)
    W_trim = np.minimum(W, cap)
    W_trim = W_trim / W_trim.mean()
    W_one = np.ones_like(W)
    fac = R / W[:, None]

    capi = (A["MODE"] == "1").to_numpy()
    scen = [
        ("ana (logit, toplam)", dict()),
        ("probit", dict(link="probit")),
        ("cloglog", dict(link="cloglog")),
        ("dogrusal olasilik", dict(link="identity")),
        ("+ gorusme modu", dict(extra=" + C(MODE)")),
        ("yalnizca yuz yuze", dict(sub=capi)),
        ("agirlik kirpma %99", dict(weights="trim")),
        ("agirliksiz", dict(weights="none")),
        (f"etiket esleme {alt_map_name}", dict(alt_map=True)),
        (f"GALI kurali {alt_rule}", dict(alt_gali=True)),
    ]

    blocks = {}
    for name, o in scen:
        expo = E + "_ALT" if o.get("alt_gali") else E
        models = {}
        f_int = f"INTERNET ~ C({expo}) + {rhs}{o.get('extra', '')}"
        sub_all = o.get("sub")
        models["INTERNET"] = BinGLM(A, f_int, subset=sub_all, link=o.get("link", "logit"))
        for k in prim:
            yk = k + "_ALT" if (o.get("alt_map") and k in ("MHRS", "KAYIT", "KONSULT")) else k
            sub = users if sub_all is None else (users & sub_all)
            models[k] = BinGLM(A, f"{yk} ~ C({expo}) + {rhs}{o.get('extra', '')}", subset=sub,
                               link=o.get("link", "logit"))
        blocks[name] = (o.get("weights"), {k: ame_fn_g(md, expo, levels, ref) for k, md in models.items()})

    rows_out = []
    for name, (wmode, fns) in blocks.items():
        if wmode == "trim":
            W0, R0 = W_trim, W_trim[:, None] * fac
        elif wmode == "none":
            W0, R0 = W_one, fac
        else:
            W0, R0 = W, R

        def fn(w):
            out = {}
            for k, f in fns.items():
                for l, v in f(w).items():
                    out[f"{k}|{l}"] = v
            for l in levels:
                mean = {s: np.mean([out[f"{k}|{l}"] for k in ks]) if ks else np.nan for s, ks in steps.items()}
                out[f"H3a_bilgi_eksi_islem|{l}"] = mean["bilgi"] - mean["islem"]
                out[f"H3b_bilgi_eksi_sinir|{l}"] = mean["bilgi"] - mean["sinir"]
            return pd.Series(out)

        print(f"    senaryo: {name}")
        point, reps = replicate(fn, W0, R0)
        sm = summarize(point, reps)
        for key, r in sm.iterrows():
            k, l = key.split("|")
            rows_out.append(dict(senaryo=name, cikti=k, GALI=l, tahmin=100 * r.tahmin,
                                 ga_alt=100 * r.ga_alt, ga_ust=100 * r.ga_ust, p=r.p,
                                 gecersiz_rep=int(r.gecersiz_rep)))
    T = pd.DataFrame(rows_out)
    T.to_csv(rep.dir / "saglamlik_uzun.csv", index=False, encoding="utf-8-sig")

    rep.p(f"AME'ler yüzde puan [%95 GA]; varyans ilk {Bs} replikasyonla. "
          "INTERNET tüm 16–74, diğerleri internet kullanıcıları; tüm senaryolarda 'toplam etki' kovaryatları. "
          f"Alternatif eşleme: {alt_map_name} (yalnızca MHRS/KAYIT değişir). "
          f"Alternatif GALI kuralı: {alt_rule}. 'agirliksiz' tasarım ağırlığı olmadan, kümelenme korunarak.")
    order = ["INTERNET"] + prim + ["H3a_bilgi_eksi_islem", "H3b_bilgi_eksi_sinir"]
    for l in levels:
        t = T[T.GALI == str(l)].copy()
        t["hucre"] = t.apply(lambda r: f"{r.tahmin:.1f} [{r.ga_alt:.1f}; {r.ga_ust:.1f}]"
                             + ("*" if r.p < 0.05 else ""), axis=1)
        wide = t.pivot(index="senaryo", columns="cikti", values="hucre")
        wide = wide.reindex([n for n, _ in scen])[[c for c in order if c in wide.columns]]
        rep.h(f"GALI = {l} vs {ref}")
        rep.table(wide, f"saglamlik_GALI{l}", index=True)
    rep.p("* = düzeltilmemiş p < 0,05. Yorum: ana satırla aynı yönde ve benzer büyüklükteyse sonuç sağlamdır.")
    bad = T[T.gecersiz_rep > 0]
    if len(bad):
        rep.p(f"**Uyarı:** {len(bad)} tahminde yakınsamayan replikasyon var (gecersiz_rep > 0).")
    rep.checkpoint([
        "Ana bulguların (H1 negatif, BILGI_F hafif kısıtlı pozitif, H3b hafif kısıtlı pozitif) yönü tüm senaryolarda korunuyor mu?",
        "HB eşlemesiyle MHRS/KAYIT sonuçları değişiyorsa, etiket belirsizliği makalede ana sınırlılık olarak yazılır.",
        "Yüz yüze alt örneklemde fark büyükse mod etkisi tartışılır.",
    ])
    rep.write()


# =============================================================================
# S11_MANUSCRIPT — English tables, figures and supplementary files for Healthcare (MDPI)
# =============================================================================
# Reads the CSV outputs of earlier stages; does NOT estimate anything new (except Table 1:
# weighted descriptive percentages). Produces (outputs/11_manuscript/):
#   Main_Tables.docx         Tables 1–5 (editable Word tables, MDPI style)
#   Figure1–5 .png (600 dpi) + .pdf (vector), Graphical_Abstract, FigureS1–S7
#   Supplementary_Materials.docx   Tables S1–S16, Figures S1–S7
#   Supplementary_Data.xlsx        all long-format result tables (one workbook)

import importlib
import subprocess

GALI_EN = {"0": "Not limited", "1": "Limited but not severely", "2": "Severely limited",
           "toplam": "Total"}
GALI_COL = {"0": "#2a78d6", "1": "#eb6834", "2": "#1baf7a"}   # validated: distinguishable under colour-vision deficiency simulation
GALI_MRK = {"0": "o", "1": "s", "2": "^"}
OUT_EN = {
    "INTERNET": "Internet use (last 3 months)",
    "BILGI_F": "Physical health information seeking",
    "BILGI_R": "Mental health information seeking",
    "MHRS": "Online appointment booking",
    "KAYIT": "Online access to personal health records",
    "KONSULT": "Online consultation / e-prescription",
    "YZ": "Generative AI use",
    "AKILLI_SAAT": "Smartwatch / wearable use",
    "SAGLIK_CIHAZ": "Connected health device use",
    "EDEVLET_RANDEVU": "e-Government appointment (12 months)",
    "EDEVLET_KAYIT": "e-Government access to personal records (12 months)",
    "VEKALET": "e-Government task done by someone else",
    "ENGEL": "e-Government accessibility barrier",
    "H3a_bilgi_eksi_islem": "H3a: information − transaction",
    "H3b_bilgi_eksi_sinir": "H3b: information − frontier",
}
STEP_EN = {"INTERNET": "Access", "BILGI_F": "Information", "BILGI_R": "Information",
           "MHRS": "Transaction", "KAYIT": "Transaction", "KONSULT": "Transaction",
           "YZ": "Frontier", "AKILLI_SAAT": "Frontier", "SAGLIK_CIHAZ": "Frontier"}
SPEC_EN = {"S0_ham": "Crude", "S1_yas_cinsiyet": "+ Age, sex", "S2_egitim": "+ Education",
           "S3_kir_kent_bolge": "+ Urbanisation, region", "S4_kaynaklar": "+ Income, computer, household size"}
SCEN_EN = {"ana (logit, toplam)": "Main (logit)", "probit": "Probit", "cloglog": "Complementary log-log",
           "dogrusal olasilik": "Linear probability", "+ gorusme modu": "+ Interview mode",
           "yalnizca yuz yuze": "Face-to-face interviews only", "agirlik kirpma %99": "Weights trimmed at 99th pct",
           "agirliksiz": "Unweighted", "etiket esleme HB": "Alternative item mapping (HB)",
           "etiket esleme HA": "Alternative item mapping (HA)",
           "GALI kurali ignore_sure": "GALI ignoring duration", "GALI kurali eurostat": "GALI Eurostat rule"}
EDU_EN = {"1_ilkokul_alti": "Primary or less", "2_ortaokul": "Lower secondary",
          "3_lise": "Upper secondary", "4_yuksekogretim": "Tertiary"}
URB_EN = {"1": "Thinly populated", "2": "Intermediate density", "3": "Densely populated"}
FONT = "Palatino Linotype"


def _ensure(pkgs):
    """Install missing packages into this Python environment."""
    for mod, pipname in pkgs:
        try:
            importlib.import_module(mod)
        except ImportError:
            print(f"    '{pipname}' kuruluyor (bir kez)...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", pipname])


def _ci(t, lo, hi, d=1):
    z = lambda v: 0.0 if round(v, d) == 0 else v   # avoid printing "−0.0"
    return f"{z(t):.{d}f} ({z(lo):.{d}f} to {z(hi):.{d}f})".replace("-", "−")


def _p(p):
    if pd.isna(p):
        return ""
    return "<0.001" if p < 0.001 else f"{p:.3f}"


def _rcsv(ctx, stage, name):
    path = ctx._abs(ctx.paths["output_dir"]) / stage / name
    if not path.exists():
        raise FileNotFoundError(f"{path} yok. Önce {stage[:2]} aşamasını çalıştırın.")
    return pd.read_csv(path, encoding="utf-8-sig")


def _gkey(x):
    """'1.0' / 1 / '1' → '1'."""
    try:
        return str(int(float(x)))
    except (TypeError, ValueError):
        return str(x)


# ------------------------------------------------------------------ Word helpers
def _doc(landscape=True):
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.shared import Cm, Pt
    d = Document()
    st = d.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(10)
    st.paragraph_format.space_after = Pt(3)
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Cm(2.0))
    d._caps = []
    return d


def _caption(doc, label, text):
    from docx.shared import Pt
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(10)
    r = p.add_run(label + ". ")
    r.bold = True
    r.font.size = Pt(9)
    r2 = p.add_run(text)
    r2.font.size = Pt(9)
    if hasattr(doc, "_caps"):
        doc._caps.append((label, text))


def _heading(doc, text, size=12):
    from docx.shared import Pt
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    return p


def _figure(doc, path, label, text, width_cm=24.0):
    from docx.shared import Cm
    doc.add_picture(str(path), width=Cm(width_cm))
    doc.paragraphs[-1].paragraph_format.keep_with_next = True
    _caption(doc, label, text)


def _note(doc, text):
    from docx.shared import Pt
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(8)


def _table(doc, df: pd.DataFrame, bold_first_col=False):
    """MDPI style: three horizontal rules (top, below header, bottom), no vertical rules, 9 pt."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
    t = doc.add_table(rows=1, cols=len(df.columns))
    for i, c in enumerate(df.columns):
        cell = t.rows[0].cells[i]
        cell.text = ""
        run = cell.paragraphs[0].add_run(str(c))
        run.bold = True
        run.font.size = Pt(8.5)
    # repeat the header row on every page
    trPr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trPr.append(th)
    for _, row in df.iterrows():
        cells = t.add_row().cells
        vals = ["" if pd.isna(v) else str(v) for v in row]
        section_row = bold_first_col and vals[0] != "" and all(v == "" for v in vals[1:])
        for i, v in enumerate(vals):
            cells[i].text = ""
            run = cells[i].paragraphs[0].add_run(v)
            run.font.size = Pt(8.5)
            if section_row and i == 0:
                run.bold = True
                run.italic = True

    def border(cell, side):
        tcPr = cell._tc.get_or_add_tcPr()
        b = tcPr.find(qn("w:tcBorders"))
        if b is None:
            b = OxmlElement("w:tcBorders")
            tcPr.append(b)
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "8")
        el.set(qn("w:color"), "000000")
        b.append(el)

    for c in t.rows[0].cells:
        border(c, "top")
        border(c, "bottom")
    for c in t.rows[-1].cells:
        border(c, "bottom")
    doc.add_paragraph()
    return t


# ------------------------------------------------------------------ Tablo 1
def _table1(ctx):
    A = load(ctx.cache("analysis.pkl"))
    A = A[A.YEAR == 2026].copy()
    if ctx.dec.get("gali_merge_levels"):
        A["G"] = A["GALI_ANY"].map(lambda v: _gkey(v) if pd.notna(v) else np.nan)
    else:
        A["G"] = A["GALI3"].map(lambda v: _gkey(v) if pd.notna(v) else np.nan)
    groups = [g for g in ["0", "1", "2"] if (A["G"] == g).any()] + ["toplam"]
    rows = []

    def add(label, mask_fn, header=False):
        r = {"Characteristic": label}
        for g in groups:
            s = A if g == "toplam" else A[A["G"] == g]
            if header:
                r[GALI_EN[g]] = ""
            else:
                m = mask_fn(s)
                r[GALI_EN[g]] = f"{100 * np.sum(s['W'][m]) / np.sum(s['W']):.1f}"
        rows.append(r)

    r = {"Characteristic": "Unweighted n"}
    for g in groups:
        r[GALI_EN[g]] = f"{len(A if g == 'toplam' else A[A['G'] == g]):,}"
    rows.append(r)
    r = {"Characteristic": "Weighted share of population, %"}
    for g in groups:
        s = A if g == "toplam" else A[A["G"] == g]
        r[GALI_EN[g]] = f"{100 * s['W'].sum() / A['W'].sum():.1f}"
    rows.append(r)
    add("Age group, years", None, True)
    for ag in sorted(A["AGEGRP"].dropna().unique()):
        add(f"  {ag.replace('-', '–')}", lambda s, ag=ag: (s["AGEGRP"] == ag).to_numpy())
    add("Women", lambda s: (s["FEMALE"] == 1).to_numpy())
    add("Educational attainment", None, True)
    for k, v in EDU_EN.items():
        add(f"  {v}", lambda s, k=k: (s["EDU"] == k).to_numpy())
    add("Degree of urbanisation", None, True)
    for k, v in URB_EN.items():
        add(f"  {v}", lambda s, k=k: (s["URBAN"] == k).to_numpy())
    add("Household income quintile", None, True)
    for q in ["1", "2", "3", "4", "5"]:
        add(f"  Q{q}" + (" (lowest)" if q == "1" else " (highest)" if q == "5" else ""),
            lambda s, q=q: (s["INCOME"] == q).to_numpy())
    add("Household size", None, True)
    for h in ["1", "2", "3-4", "5+"]:
        add(f"  {h.replace('-', '–')}", lambda s, h=h: (s["HHSIZE"] == h).to_numpy())
    add("Computer or tablet in household", lambda s: (s["COMPUTER"] == 1).to_numpy())
    add("Internet use, last 3 months", lambda s: (s["INTERNET"] == 1).to_numpy())
    add("Telephone interview (CATI)", lambda s: (s["MODE"] == "3").to_numpy())
    return pd.DataFrame(rows)


# ------------------------------------------------------------------ figures
def _fig_setup():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({
        "font.family": "sans-serif", "font.sans-serif": ["Arial", "Helvetica", "DejaVu Sans"],
        "font.size": 9, "axes.titlesize": 9, "axes.labelsize": 9, "xtick.labelsize": 8,
        "ytick.labelsize": 8, "legend.fontsize": 8, "axes.spines.top": False,
        "axes.spines.right": False, "axes.edgecolor": "#52514e", "axes.linewidth": 0.8,
        "xtick.color": "#52514e", "ytick.color": "#52514e", "axes.labelcolor": "#0b0b0b",
        "savefig.dpi": 600, "figure.dpi": 150, "pdf.fonttype": 42,
    })
    return plt


def _save(fig, out, name):
    for ext in ("png", "pdf"):
        fig.savefig(out / f"{name}.{ext}", bbox_inches="tight", facecolor="white")
    import matplotlib.pyplot as plt
    plt.close(fig)


SHORT = {"INTERNET": "Internet\nuse", "BILGI_F": "Physical\nhealth info.", "BILGI_R": "Mental\nhealth info.",
         "MHRS": "Online\nappointment", "KAYIT": "Health\nrecords", "KONSULT": "Online\nconsultation",
         "YZ": "Generative\nAI", "AKILLI_SAAT": "Smartwatch/\nwearable", "SAGLIK_CIHAZ": "Connected\nhealth device",
         "H3a_bilgi_eksi_islem": "H3a: information\n− transaction", "H3b_bilgi_eksi_sinir": "H3b: information\n− frontier"}


def _grid(ax, axis="y"):
    (ax.yaxis if axis == "y" else ax.xaxis).grid(True, color="#e6e5e0", lw=0.6)
    ax.set_axisbelow(True)


def _fig1(ctx, out, plt):
    """Cascade: whole 16–74 denominator, by GALI level, point + 95% CI."""
    K = _rcsv(ctx, "04_descriptives", "kaskad_uzun.csv")
    K = K[K.payda == "tum_16_74"].copy()
    K["G"] = K["GALI"].map(_gkey)
    order = [k for k in ["INTERNET", "BILGI_F", "BILGI_R", "MHRS", "KAYIT", "KONSULT", "YZ", "AKILLI_SAAT"]
             if k in set(K.cikti)]
    gl = [g for g in ["0", "1", "2"] if g in set(K.G)]
    fig, ax = plt.subplots(figsize=(7.0, 3.8))
    off = {g: (i - (len(gl) - 1) / 2) * 0.24 for i, g in enumerate(gl)}
    for g in gl:
        d = K[K.G == g].set_index("cikti").loc[order]
        x = np.arange(len(order)) + off[g]
        ax.errorbar(x, d.tahmin, yerr=[d.tahmin - d.ga_alt, d.ga_ust - d.tahmin], fmt="none",
                    ecolor=GALI_COL[g], elinewidth=0.9, capsize=2.2, capthick=0.9, zorder=2)
        ax.plot(x, d.tahmin, ls="none", marker=GALI_MRK[g], color=GALI_COL[g], ms=4.6, mec=GALI_COL[g],
                zorder=3, label=GALI_EN[g])
    ax.set_xticks(np.arange(len(order)))
    ax.set_xticklabels([SHORT[k] for k in order], fontsize=7.5)
    ax.set_xlim(-0.6, len(order) - 0.4)
    ax.set_ylabel("Weighted prevalence, % (95% CI)")
    ax.set_ylim(0, 100)
    _grid(ax)
    steps = [STEP_EN[k] for k in order]
    i = 0
    while i < len(steps):
        j = i
        while j + 1 < len(steps) and steps[j + 1] == steps[i]:
            j += 1
        ax.annotate(steps[i], xy=((i + j) / 2, 1.02), xycoords=("data", "axes fraction"),
                    ha="center", fontsize=8, color="#52514e")
        if i > 0:
            ax.axvline(i - 0.5, color="#c3c2b7", lw=0.6)
        i = j + 1
    ax.legend(frameon=False, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.18))
    _save(fig, out, "Figure2")


def _fig2(ctx, out, plt):
    """Crude → adjusted gap across specifications, by GALI level; one panel per outcome (2×3)."""
    H = _rcsv(ctx, "09_ham_ayarli", "ham_ayarli_uzun.csv")
    H["G"] = H["GALI"].map(_gkey)
    specs = list(SPEC_EN)[:4]
    H = H[H.olcu.isin(specs)]
    order = [k for k in ["INTERNET", "BILGI_F", "BILGI_R", "MHRS", "KAYIT", "YZ"] if k in set(H.cikti)]
    gl = [g for g in ["1", "2"] if g in set(H.G)]
    nrow = 2
    ncol = int(np.ceil(len(order) / nrow))
    fig, axes = plt.subplots(nrow, ncol, figsize=(7.0, 4.8), sharey=True)
    axes = np.atleast_1d(axes).ravel()
    for ax, k in zip(axes, order):
        for g in gl:
            d = H[(H.cikti == k) & (H.G == g)].set_index("olcu").loc[specs]
            x = np.arange(len(specs)) + (0.07 if g == "2" else -0.07)
            ax.plot(x, 100 * d.tahmin, color=GALI_COL[g], lw=1.6, marker=GALI_MRK[g], ms=5,
                    mec="white", mew=0.7, label=GALI_EN[g])
            ax.vlines(x, 100 * d.ga_alt, 100 * d.ga_ust, color=GALI_COL[g], lw=1)
        ax.axhline(0, color="#52514e", lw=0.8)
        ax.set_title(SHORT[k].replace("\n", " "), fontsize=8.5)
        ax.set_xticks(range(len(specs)))
        ax.set_xticklabels(["Crude", "+ Age,\nsex", "+ Educ.", "+ Area"], fontsize=7)
        _grid(ax)
    for ax in axes[len(order):]:
        ax.axis("off")
    for r in range(nrow):
        axes[r * ncol].set_ylabel("Gap vs not limited, pp")
    h, l = axes[0].get_legend_handles_labels()
    fig.legend(h, l, frameon=False, ncol=2, loc="lower center", bbox_to_anchor=(0.5, -0.02))
    fig.tight_layout(rect=(0, 0.05, 1, 1))
    _save(fig, out, "Figure3")


def _figS1(ctx, out, plt):
    """Access and conditional-use components (stage 06), by GALI level (supplementary figure)."""
    U = _rcsv(ctx, "06_hurdle", "hurdle_uzun.csv")
    U["G"] = U["GALI"].map(_gkey)
    order = [k for k in ["BILGI_F", "BILGI_R", "MHRS", "KAYIT", "KONSULT", "YZ"] if k in set(U.cikti)]
    gl = [g for g in ["1", "2"] if g in set(U.G)]
    comp = [("dP_toplam", "Total gap", "#0b0b0b", "D"), ("dP_erisim", "Access component", "#2a78d6", "o"),
            ("dP_kosullu", "Conditional-use component", "#8a8984", "s")]
    sub = U[U.olcu.isin([c[0] for c in comp]) & U.cikti.isin(order)]
    lim = 100 * max(abs(sub.ga_alt).max(), abs(sub.ga_ust).max()) * 1.08
    fig, axes = plt.subplots(1, len(gl), figsize=(7.0, 3.3), sharey=True)
    axes = np.atleast_1d(axes)
    y = np.arange(len(order))[::-1]
    for ax, g in zip(axes, gl):
        for j, (olc, lab, col, mk) in enumerate(comp):
            d = U[(U.olcu == olc) & (U.G == g)].set_index("cikti").loc[order]
            yy = y + (0.24 - 0.24 * j)
            ax.errorbar(100 * d.tahmin, yy, xerr=[100 * (d.tahmin - d.ga_alt), 100 * (d.ga_ust - d.tahmin)],
                        fmt="none", ecolor=col, elinewidth=1.0, capsize=0, zorder=2)
            ax.plot(100 * d.tahmin, yy, ls="none", marker=mk, color=col, ms=4.6, mec=col, zorder=3, label=lab)
        ax.axvline(0, color="#52514e", lw=0.8)
        ax.set_xlim(-lim, lim)
        ax.set_title(f"{GALI_EN[g]} vs not limited", fontsize=8.5)
        ax.set_xlabel("Percentage points (95% CI)")
        _grid(ax, "x")
    axes[0].set_yticks(y)
    axes[0].set_yticklabels([SHORT[k].replace("\n", " ") for k in order])
    h, l = axes[0].get_legend_handles_labels()
    fig.legend(h, l, frameon=False, ncol=3, loc="lower center", bbox_to_anchor=(0.55, -0.02))
    fig.tight_layout(rect=(0, 0.07, 1, 1), w_pad=1.5)
    _save(fig, out, "FigureS1")


def _figS2(ctx, out, plt):
    """Sensitivity: scenario × primary AME (forest plot), 2×3 panels."""
    T = _rcsv(ctx, "10_saglamlik", "saglamlik_uzun.csv")
    T["G"] = T["GALI"].map(_gkey)
    outs = [k for k in ["INTERNET", "BILGI_F", "MHRS", "KAYIT", "YZ", "H3b_bilgi_eksi_sinir"] if k in set(T.cikti)]
    scen = list(dict.fromkeys(T.senaryo))
    gl = [g for g in ["1", "2"] if g in set(T.G)]
    fig, axes = plt.subplots(2, 3, figsize=(7.2, 6.2), sharey=True)
    axes = axes.ravel()
    y0 = np.arange(len(scen))[::-1]
    for ax, k in zip(axes, outs):
        for g in gl:
            d = T[(T.cikti == k) & (T.G == g)].set_index("senaryo").loc[scen]
            y = y0 + (0.16 if g == "1" else -0.16)
            ax.errorbar(d.tahmin, y, xerr=[d.tahmin - d.ga_alt, d.ga_ust - d.tahmin], fmt="none",
                        ecolor=GALI_COL[g], elinewidth=0.9, capsize=0, zorder=2)
            ax.plot(d.tahmin, y, ls="none", marker=GALI_MRK[g], color=GALI_COL[g], ms=3.8, mec=GALI_COL[g],
                    zorder=3, label=GALI_EN[g])
        ax.axvline(0, color="#52514e", lw=0.7)
        ax.set_title(SHORT[k].replace("\n", " "), fontsize=8)
        _grid(ax, "x")
        ax.tick_params(labelsize=7)
    for ax in axes[len(outs):]:
        ax.axis("off")
    for r in (0, 3):
        axes[r].set_yticks(y0)
        axes[r].set_yticklabels([SCEN_EN.get(s, s) for s in scen], fontsize=7)
    h, l = axes[0].get_legend_handles_labels()
    fig.legend(h, l, frameon=False, ncol=2, loc="lower center", bbox_to_anchor=(0.5, -0.01))
    fig.supxlabel("Average marginal effect, percentage points (95% CI)", fontsize=8, y=0.04)
    fig.tight_layout(rect=(0, 0.06, 1, 1))
    _save(fig, out, "FigureS2")


# ------------------------------------------------------------------ supplementary figure helpers
COL_ONLINE, COL_PROXY, COL_ISOL = "#c9c8c2", "#7c7b76", "#b03a2e"


def _box(ax, x, y, w, h, title, body, fc="#f4f3ee", ec="#52514e", tc="#0b0b0b"):
    from matplotlib.patches import FancyBboxPatch
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.012,rounding_size=0.02",
                                fc=fc, ec=ec, lw=0.9, zorder=2))
    ax.text(x + w / 2, y + h - 0.035, title, ha="center", va="top", fontsize=8, weight="bold", color=tc, zorder=3)
    ax.text(x + w / 2, y + h - 0.105, body, ha="center", va="top", fontsize=6.8, color="#3d3c39",
            linespacing=1.35, zorder=3)


def _arrow(ax, p0, p1, lab=None, ls="-", col="#52514e", lab_xy=None, rad=0.0, fs=6.8, lab_col=None):
    from matplotlib.patches import FancyArrowPatch
    ax.add_patch(FancyArrowPatch(p0, p1, arrowstyle="-|>", mutation_scale=9, lw=1.0, color=col, ls=ls,
                                 connectionstyle=f"arc3,rad={rad}", zorder=4, shrinkA=1, shrinkB=1))
    if lab:
        x, y = lab_xy if lab_xy else ((p0[0] + p1[0]) / 2, (p0[1] + p1[1]) / 2)
        ax.text(x, y, lab, ha="center", va="center", fontsize=fs, color=lab_col or col, zorder=4,
                bbox=dict(boxstyle="round,pad=0.18", fc="white", ec="none"), linespacing=1.25)


def _key_numbers(ctx):
    """Key numbers for the graphical abstract — all read from CSV files (never typed in)."""
    k = {}
    try:
        H = _rcsv(ctx, "09_ham_ayarli", "ham_ayarli_uzun.csv")
        H["G"] = H["GALI"].map(_gkey)
        sh = []
        for c in ["INTERNET", "MHRS", "KAYIT"]:
            a = H[(H.cikti == c) & (H.G == "2") & (H.olcu == "S0_ham")].tahmin.iloc[0]
            b = H[(H.cikti == c) & (H.G == "2") & (H.olcu == "S3_kir_kent_bolge")].tahmin.iloc[0]
            sh.append(1 - b / a)
        k["share"] = f"{100 * min(sh):.0f}–{100 * max(sh):.0f}%"
        k["crude_int"] = 100 * H[(H.cikti == "INTERNET") & (H.G == "2") & (H.olcu == "S0_ham")].tahmin.iloc[0]
    except Exception:
        pass
    try:
        F = _rcsv(ctx, "14_hane_ici", "hane_ici_uzun.csv")
        F["G"] = F["GALI"].map(_gkey)
        g = lambda o: 100 * F[(F.olcu == o) & (F.cikti == "INTERNET") & (F.G == "2")].tahmin.iloc[0]
        k["fe"], k["fe_old"], k["fe_young"] = g("hane_FE"), g("FE_55-74"), g("FE_16-54")
    except Exception:
        pass
    try:
        Hh = _rcsv(ctx, "13_hane", "hane_uzun.csv")
        Hh["G"] = Hh["GALI"].map(_gkey)
        g = lambda o, c, G="2": 100 * Hh[(Hh.olcu == o) & (Hh.cikti == c) & (Hh.G == G)].tahmin.iloc[0]
        k["comp_int"], k["comp_mhrs"] = g("ham_telafi", "INTERNET"), g("ham_telafi", "MHRS_TUM")
        k["isol_old"], k["isol_ref_old"] = g("oran_55-74", "YALITIK"), g("oran_55-74", "YALITIK", "0")
        k["cores"], k["cores_ref"] = g("ham_oran", "HH_DIGER_INT"), g("ham_oran", "HH_DIGER_INT", "0")
    except Exception:
        pass
    try:
        k["hh"] = int(_rcsv(ctx, "13_hane", "kapsam.csv").hane_sayisi.iloc[0])
        k["n"] = int(_rcsv(ctx, "13_hane", "kapsam.csv").fert_16_74_veride.iloc[0])
    except Exception:
        pass
    try:
        M = _rcsv(ctx, "05_models", "ame_toplam.csv")
        M["G"] = M["GALI_duzey"].map(_gkey)
        k["need"] = M[(M.cikti == "BILGI_F") & (M.G == "1")].tahmin.iloc[0]
    except Exception:
        pass
    return k


def _fig_framework(ctx, out, plt, graphical=False):
    """Figure 1: extended Andersen model (Q1–Q4). graphical=True: graphical abstract with key numbers."""
    K = _key_numbers(ctx) if graphical else {}
    top = 1.10 if graphical else 0.95
    fig = plt.figure(figsize=(7.2, 4.0 if graphical else 3.5))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, top)
    ax.axis("off")
    W = 0.21
    _box(ax, 0.02, 0.60, W, 0.28, "Predisposing", "Age, sex, education\n(area of residence)")
    _box(ax, 0.02, 0.14, W, 0.28, "Need", "Activity limitation (GALI):\nmoderate or severe", fc="#e9f1fb", ec="#2a78d6")
    _box(ax, 0.41, 0.60, W, 0.28, "Individual enabling", "Own internet access,\ndevices, digital skills")
    _box(ax, 0.41, 0.14, W, 0.28, "Relational enabling", "Adult co-resident online\n(warm expert, proxy use)")
    body = "Health information seeking;\nonline appointments, records;\ngenerative AI, wearables"
    _box(ax, 0.77, 0.30, W, 0.42, "Digital health use", body, fc="#e8f6ef", ec="#1baf7a")
    if graphical:
        ax.text(0.875, 0.47, f"{K.get('isol_old', 0):.0f}% of severely limited\nadults aged 55–74 are\n"
                f"digitally isolated\n(vs {K.get('isol_ref_old', 0):.0f}% without limitation)", ha="center", va="top",
                fontsize=6.8, color=COL_ISOL, weight="bold", zorder=4, linespacing=1.3)
        L = {"q1": f"Q1  {K.get('share', '')} of\ncrude gaps\ncompositional",
             "q2": f"Q2  gap within the\nsame household\n{K.get('fe', 0):.1f} pp\n(55–74: {K.get('fe_old', 0):.1f})".replace("-", "−"),
             "q3": f"Q3  household closes\ninternet gap {K.get('comp_int', 0):.0f}%,\nappointments {K.get('comp_mhrs', 0):.0f}%",
             "q4": f"Q4  need: +{K.get('need', 0):.1f} pp health information seeking among moderately limited users "
                   "(attenuated within households)"}
    else:
        L = {"q1": "Q1  composition", "q2": "Q2  individual\nbarrier", "q3": "Q3  household\ncompensation",
             "q4": "Q4  need-driven use among internet users"}
    kw = dict(fs=6.6)
    _arrow(ax, (0.245, 0.74), (0.397, 0.74))
    from matplotlib.patches import FancyArrowPatch
    ax.add_patch(FancyArrowPatch((0.125, 0.587), (0.125, 0.433), arrowstyle="<|-|>", mutation_scale=8, lw=1.0,
                                 color="#52514e", ls="--", zorder=4, shrinkA=1, shrinkB=1))
    ax.text(0.125, 0.51, L["q1"], ha="center", va="center", fontsize=6.6, color="#52514e", zorder=5,
            bbox=dict(boxstyle="round,pad=0.18", fc="white", ec="none"), linespacing=1.25)
    _arrow(ax, (0.245, 0.40), (0.397, 0.60), L["q2"], col="#2a78d6", lab_xy=(0.321, 0.52), **kw)
    _arrow(ax, (0.245, 0.30), (0.397, 0.30), L["q3"], col="#2a78d6", lab_xy=(0.321, 0.205), **kw)
    _arrow(ax, (0.633, 0.74), (0.757, 0.64))
    _arrow(ax, (0.633, 0.26), (0.757, 0.36), ls="--")
    _arrow(ax, (0.515, 0.433), (0.515, 0.587), ls=":", col="#8a8984")
    # Q4: need → use, elbow line along the bottom edge
    ax.plot([0.125, 0.125, 0.875], [0.127, 0.05, 0.05], color="#1baf7a", lw=1.0, zorder=4, solid_capstyle="butt")
    _arrow(ax, (0.875, 0.05), (0.875, 0.287), col="#1baf7a")
    ax.text(0.50, 0.05, L["q4"], ha="center", va="center", fontsize=6.6, color="#1baf7a", zorder=5,
            bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none"))
    if graphical:
        ax.text(0.5, 1.07, "Who is left offline when health care goes digital?", ha="center", va="top",
                fontsize=10.5, weight="bold")
        ax.text(0.5, 1.0, f"Activity limitation and digital health in Türkiye, 2026 (n = {K.get('n', 0):,} adults "
                f"in {K.get('hh', 0):,} households)",
                ha="center", va="top", fontsize=7.5, color="#52514e")
        _save(fig, out, "Graphical_Abstract")
    else:
        _save(fig, out, "Figure1")


def _fig_left_offline(ctx, out, plt):
    """Figure 5: who is left offline? Online themselves / offline but an online adult at home / isolated."""
    Hh = _rcsv(ctx, "13_hane", "hane_uzun.csv")
    Hh["G"] = Hh["GALI"].map(_gkey)
    g = lambda o, c, G: 100 * Hh[(Hh.olcu == o) & (Hh.cikti == c) & (Hh.G == G)].tahmin.iloc[0]
    rows = []
    for sn, sl in (("16-54", "Aged 16–54"), ("55-74", "Aged 55–74")):
        for G in ["0", "1", "2"]:
            on, iso = g(f"oran_{sn}", "INTERNET", G), g(f"oran_{sn}", "YALITIK", G)
            rows.append((sl, GALI_EN[G], on, max(100 - on - iso, 0), iso))
    fig, ax = plt.subplots(figsize=(7.0, 3.2))
    ypos, y = [], 0
    for i, r in enumerate(rows):
        if i == 3:
            y += 0.7
        ypos.append(y)
        y += 1
    ypos = -np.array(ypos)
    for yy, (_, gl, on, px, iso) in zip(ypos, rows):
        ax.barh(yy, on, color=COL_ONLINE, height=0.72, edgecolor="white", lw=0.6)
        ax.barh(yy, px, left=on, color=COL_PROXY, height=0.72, edgecolor="white", lw=0.6)
        ax.barh(yy, iso, left=on + px, color=COL_ISOL, height=0.72, edgecolor="white", lw=0.6)
        ax.text(2, yy, f"{on:.0f}%", va="center", ha="left", fontsize=7, color="#0b0b0b")
        if px >= 5:
            ax.text(on + px / 2, yy, f"{px:.0f}%", va="center", ha="center", fontsize=7, color="white")
        ax.text(101.5, yy, f"{iso:.1f}%", va="center", ha="left", fontsize=7.5, color=COL_ISOL, weight="bold")
    ax.set_yticks(ypos)
    ax.set_yticklabels([r[1] for r in rows], fontsize=7.5)
    for i, lab in ((0, "Aged 16–54"), (3, "Aged 55–74")):
        ax.text(-1.5, ypos[i] + 0.75, lab, ha="right", va="center", fontsize=8, weight="bold",
                transform=ax.transData)
    ax.set_xlim(0, 100)
    ax.set_xlabel("Weighted % of adults")
    ax.tick_params(axis="y", length=0)
    ax.spines["left"].set_visible(False)
    ax.text(101.5, ypos[0] + 0.8, "Isolated", fontsize=7, color=COL_ISOL, ha="left", va="center")
    from matplotlib.patches import Patch
    h = [Patch(color=COL_ONLINE, label="Online themselves"),
         Patch(color=COL_PROXY, label="Offline, lives with an online adult"),
         Patch(color=COL_ISOL, label="Digitally isolated (no adult online)")]
    fig.legend(handles=h, frameon=False, ncol=3, loc="lower center", bbox_to_anchor=(0.55, -0.01), fontsize=7.5)
    fig.tight_layout(rect=(0, 0.08, 0.97, 1))
    _save(fig, out, "Figure5")


def _forest_panel(ax, labels, series, xlab=None, ref=0.0, log=False):
    """series: [(name, color, marker, filled, [(t, lo, hi), ...]), ...] — offset within each row."""
    n = len(series)
    y0 = np.arange(len(labels))[::-1].astype(float)
    for j, (nm, col, mk, filled, vals) in enumerate(series):
        yy = y0 + (n - 1) / 2 * 0.22 - j * 0.22
        for y, v in zip(yy, vals):
            if v is None or not np.isfinite(v[0]):
                continue
            t, lo, hi = v
            ax.plot([lo, hi], [y, y], color=col, lw=0.9, zorder=2)
            ax.plot(t, y, ls="none", marker=mk, ms=4.2, color=col, mfc=col if filled else "white",
                    mec=col, mew=1.0, zorder=3, label=nm)
    ax.axvline(ref, color="#52514e", lw=0.8)
    if log:
        ax.set_xscale("log")
        from matplotlib.ticker import FixedLocator, FuncFormatter, NullLocator
        lo, hi = ax.get_xlim()
        tk = [t for t in (0.25, 0.33, 0.5, 0.67, 0.8, 1, 1.25, 1.5, 2, 3, 4) if lo <= t <= hi]
        ax.xaxis.set_major_locator(FixedLocator(tk))
        ax.xaxis.set_minor_locator(NullLocator())
        ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:g}"))
    ax.set_yticks(y0)
    ax.set_yticklabels(labels, fontsize=7.5)
    ax.set_ylim(-0.6, len(labels) - 0.4)
    if xlab:
        ax.set_xlabel(xlab, fontsize=8)
    _grid(ax, "x")


def _uniq_legend(fig, axes, ncol, y=-0.01):
    h, l = [], []
    for ax in np.atleast_1d(axes):
        for hh, ll in zip(*ax.get_legend_handles_labels()):
            if ll not in l:
                h.append(hh)
                l.append(ll)
    fig.legend(h, l, frameon=False, ncol=ncol, loc="lower center", bbox_to_anchor=(0.5, y), fontsize=7.5)


def _figS3_rr(ctx, out, plt):
    """Figure S3: standardised risk ratios (log scale)."""
    R = _rcsv(ctx, "12_goreli_yas", "goreli_yas_uzun.csv")
    R["G"] = R["GALI"].map(_gkey)
    order = [k for k in ["INTERNET", "BILGI_F", "BILGI_R", "MHRS", "KAYIT", "KONSULT", "YZ", "AKILLI_SAAT",
                         "H3a_bilgi_eksi_islem", "H3b_bilgi_eksi_sinir"] if k in set(R.cikti)]
    ser = []
    for G in ["1", "2"]:
        vals = []
        for k in order:
            x = R[(R.olcu == "RR") & (R.cikti == k) & (R.G == G)]
            vals.append(None if not len(x) else tuple(np.exp([x.iloc[0].tahmin, x.iloc[0].ga_alt, x.iloc[0].ga_ust])))
        ser.append((GALI_EN[G], GALI_COL[G], GALI_MRK[G], True, vals))
    fig, ax = plt.subplots(figsize=(6.0, 4.0))
    _forest_panel(ax, [SHORT[k].replace("\n", " ") for k in order], ser,
                  "Standardised risk ratio vs not limited (log scale, 95% CI)", ref=1.0, log=True)
    ax.axhline(1.5, color="#c3c2b7", lw=0.6, ls="--")
    _uniq_legend(fig, ax, 2)
    fig.tight_layout(rect=(0, 0.06, 1, 1))
    _save(fig, out, "FigureS3")


def _figS4_age(ctx, out, plt):
    """Figure S4: AMEs within age strata (logistic)."""
    R = _rcsv(ctx, "12_goreli_yas", "goreli_yas_uzun.csv")
    R["G"] = R["GALI"].map(_gkey)
    order = [k for k in ["INTERNET", "BILGI_F", "BILGI_R", "MHRS", "KAYIT", "YZ"] if k in set(R.cikti)]
    fig, axes = plt.subplots(1, 2, figsize=(7.0, 3.4), sharey=True, sharex=True)
    for ax, G in zip(axes, ["1", "2"]):
        ser = []
        for sn, sl, filled in (("16-54", "Aged 16–54", True), ("55-74", "Aged 55–74", False)):
            vals = []
            for k in order:
                x = R[(R.olcu == f"YAS_{sn}") & (R.cikti == k) & (R.G == G)]
                vals.append(None if not len(x) else (100 * x.iloc[0].tahmin, 100 * x.iloc[0].ga_alt, 100 * x.iloc[0].ga_ust))
            ser.append((sl, GALI_COL[G], GALI_MRK[G], filled, vals))
        _forest_panel(ax, [SHORT[k].replace("\n", " ") for k in order], ser, "AME, percentage points (95% CI)")
        ax.set_title(f"{GALI_EN[G]} vs not limited", fontsize=8.5)
    from matplotlib.lines import Line2D
    h = [Line2D([], [], ls="none", marker="o", color="#52514e", ms=5, label="Aged 16–54"),
         Line2D([], [], ls="none", marker="o", mfc="white", mec="#52514e", ms=5, label="Aged 55–74")]
    fig.legend(handles=h, frameon=False, ncol=2, loc="lower center", bbox_to_anchor=(0.5, -0.01), fontsize=7.5)
    fig.tight_layout(rect=(0, 0.07, 1, 1))
    _save(fig, out, "FigureS4")


def _figS5_fe(ctx, out, plt):
    """Figure S5: pooled and within-household estimates, all outcomes."""
    F = _rcsv(ctx, "14_hane_ici", "hane_ici_uzun.csv")
    F["G"] = F["GALI"].map(_gkey)
    lab = {"INTERNET": "Internet use", "MHRS_TUM": "Online appointment (all)", "KAYIT_TUM": "Health records (all)",
           "BILGI_F_TUM": "Physical health info. (all)", "YZ_TUM": "Generative AI (all)",
           "BILGI_F": "Physical health info. (users)", "MHRS": "Online appointment (users)",
           "KAYIT": "Health records (users)", "YZ": "Generative AI (users)"}
    order = [k for k in lab if k in set(F.cikti)]
    fig, axes = plt.subplots(1, 2, figsize=(7.0, 3.9), sharey=True, sharex=True)
    for ax, G in zip(axes, ["1", "2"]):
        ser = []
        for o, nm, filled in (("havuz", "Pooled (between and within households)", False),
                              ("hane_FE", "Within household (fixed effects)", True)):
            vals = []
            for k in order:
                x = F[(F.olcu == o) & (F.cikti == k) & (F.G == G)]
                vals.append(None if not len(x) else (100 * x.iloc[0].tahmin, 100 * x.iloc[0].ga_alt, 100 * x.iloc[0].ga_ust))
            ser.append((nm, GALI_COL[G], GALI_MRK[G], filled, vals))
        _forest_panel(ax, [lab[k] for k in order], ser, "Percentage points (95% CI)")
        ax.axhline(3.5, color="#c3c2b7", lw=0.6, ls="--")
        ax.set_title(f"{GALI_EN[G]} vs not limited", fontsize=8.5)
    from matplotlib.lines import Line2D
    h = [Line2D([], [], ls="none", marker="o", mfc="white", mec="#52514e", ms=5, label="Pooled"),
         Line2D([], [], ls="none", marker="o", color="#52514e", ms=5, label="Within household (fixed effects)")]
    fig.legend(handles=h, frameon=False, ncol=2, loc="lower center", bbox_to_anchor=(0.5, -0.01), fontsize=7.5)
    fig.tight_layout(rect=(0, 0.07, 1, 1))
    _save(fig, out, "FigureS5")


def _figS6_spill(ctx, out, plt):
    """Figure S6: household spillover and negative control."""
    import re
    Y = _rcsv(ctx, "15_yayilim", "yayilim.csv")
    order = [k for k in ["BILGI_F", "BILGI_R", "MHRS", "KAYIT", "KONSULT", "YZ"] if k in set(Y.cikti)]

    def parse(sv):
        m = re.findall(r"-?\d+\.?\d*", str(sv))
        return tuple(float(v) for v in m[:3]) if len(m) >= 3 else None
    ser = []
    for h, G in (("orta kısıtlı", "1"), ("ağır kısıtlı", "2")):
        vals = []
        for k in order:
            x = Y[(Y.cikti == k) & (Y.hanede == h)]
            vals.append(parse(x.iloc[0].AME_pp) if len(x) else None)
        ser.append((f"Lives with a {'moderately' if G == '1' else 'severely'} limited adult", GALI_COL[G],
                    GALI_MRK[G], True, vals))
    fig, ax = plt.subplots(figsize=(6.0, 3.3))
    labels = [SHORT[k].replace("\n", " ") + (" (negative control)" if k == "YZ" else "") for k in order]
    _forest_panel(ax, labels, ser, "AME among internet users without limitation, pp (95% CI)")
    if "YZ" in order:
        ax.axhspan(-0.5, 0.5, color="#f4f3ee", zorder=0)
    _uniq_legend(fig, ax, 2)
    fig.tight_layout(rect=(0, 0.08, 1, 1))
    _save(fig, out, "FigureS6")


def _figS7_trend(ctx, out, plt):
    """Figure S7: 2025→2026 prevalence and sociodemographic gaps (dumbbell)."""
    T = _rcsv(ctx, "08_trend", "trend_uzun.csv")
    order = [k for k in ["INTERNET", "BILGI_F", "BILGI_R", "YZ"] if k in set(T.cikti)]
    meas = [("yayginlik", "Prevalence, %"), ("acik_EDU", "Education gap,\ntertiary − primary, pp"),
            ("acik_AGEGRP", "Age gap,\n65–74 − 25–34, pp"), ("acik_FEMALE", "Sex gap,\nfemale − male, pp")]
    fig, axes = plt.subplots(1, 4, figsize=(7.2, 2.9), sharey=True)
    y = np.arange(len(order))[::-1]
    c25, c26 = "#9fc3ee", "#2a78d6"
    for ax, (m, ttl) in zip(axes, meas):
        for yy, k in zip(y, order):
            a = T[(T.cikti == k) & (T.olcu == m) & (T.donem.astype(str) == "2025")]
            b = T[(T.cikti == k) & (T.olcu == m) & (T.donem.astype(str) == "2026")]
            if not (len(a) and len(b)):
                continue
            a, b = float(a.tahmin.iloc[0]), float(b.tahmin.iloc[0])
            ax.plot([a, b], [yy, yy], color="#c3c2b7", lw=1.6, zorder=1)
            ax.plot(a, yy, "o", color=c25, ms=5, zorder=2, label="2025")
            ax.plot(b, yy, "o", color=c26, ms=5, zorder=3, label="2026")
        if m != "yayginlik":
            ax.axvline(0, color="#52514e", lw=0.7)
        ax.set_title(ttl, fontsize=7.5)
        ax.tick_params(labelsize=7)
        _grid(ax, "x")
    axes[0].set_yticks(y)
    lab = {"INTERNET": "Internet use (all)", "BILGI_F": "Physical health info.", "BILGI_R": "Mental health info.",
           "YZ": "Generative AI"}
    axes[0].set_yticklabels([lab[k] for k in order], fontsize=7.5)
    _uniq_legend(fig, axes, 2)
    fig.tight_layout(rect=(0, 0.08, 1, 1), w_pad=0.8)
    _save(fig, out, "FigureS7")


def _fig4(ctx, out, plt):
    """Figure 4: (A) own vs household-inclusive use, (B) within-household internet gap by age."""
    Hh = _rcsv(ctx, "13_hane", "hane_uzun.csv")
    Hh["G"] = Hh["GALI"].map(_gkey)
    F = _rcsv(ctx, "14_hane_ici", "hane_ici_uzun.csv")
    F["G"] = F["GALI"].map(_gkey)
    from matplotlib.lines import Line2D

    def v(df, o, k, g):
        x = df[(df.olcu == o) & (df.cikti == k) & (df.G == g)].iloc[0]
        return 100 * x.tahmin, 100 * x.ga_alt, 100 * x.ga_ust

    fig, axes = plt.subplots(1, 2, figsize=(7.2, 3.2), gridspec_kw={"width_ratios": [1.9, 1]})
    gl = ["0", "1", "2"]
    off = {"0": -0.24, "1": 0.0, "2": 0.24}
    ax = axes[0]
    pairs = [("INTERNET", "ERISIM_ILISKI", "Internet\nuse"), ("MHRS_TUM", "MHRS_ILISKI", "Online\nappointment"),
             ("KAYIT_TUM", "KAYIT_ILISKI", "Health\nrecords")]
    for i, (own, rel, lab) in enumerate(pairs):
        for g in gl:
            o, r = v(Hh, "ham_oran", own, g)[0], v(Hh, "ham_oran", rel, g)[0]
            x = i + off[g]
            ax.plot([x, x], [o, r], color=GALI_COL[g], lw=1.1, zorder=2)
            ax.plot(x, o, ls="none", marker=GALI_MRK[g], ms=4.8, color=GALI_COL[g], mec=GALI_COL[g], zorder=3)
            ax.plot(x, r, ls="none", marker=GALI_MRK[g], ms=4.8, mfc="white", mec=GALI_COL[g], mew=1.1, zorder=3)
        try:
            c = Hh[(Hh.olcu == "ham_telafi") & (Hh.cikti == own) & (Hh.G == "2")].tahmin.iloc[0]
            ax.text(i, 3, f"Severe: {100 * c:.0f}% of gap\nclosed by household", ha="center", va="bottom",
                    fontsize=6.3, color=GALI_COL["2"])
        except IndexError:
            pass
    ax.set_xticks(range(len(pairs)))
    ax.set_xticklabels([p[2] for p in pairs], fontsize=7.5)
    ax.set_xlim(-0.55, len(pairs) - 0.45)
    ax.set_ylim(0, 100)
    ax.set_ylabel("Weighted prevalence, %")
    ax.set_title("A  Own use and household-inclusive use", fontsize=8.5, loc="left")
    _grid(ax)
    ax = axes[1]
    for j, sn in enumerate(["16-54", "55-74"]):
        for g in ["1", "2"]:
            t, lo, hi = v(F, f"FE_{sn}", "INTERNET", g)
            x = j + (-0.12 if g == "1" else 0.12)
            ax.errorbar(x, t, yerr=[[t - lo], [hi - t]], fmt="none", ecolor=GALI_COL[g], elinewidth=0.9,
                        capsize=2, zorder=2)
            ax.plot(x, t, ls="none", marker=GALI_MRK[g], ms=4.8, color=GALI_COL[g], mec=GALI_COL[g], zorder=3)
    ax.axhline(0, color="#52514e", lw=0.8)
    ax.set_xticks([0, 1])
    ax.set_xticklabels(["16–54", "55–74"], fontsize=7.5)
    ax.set_xlim(-0.55, 1.55)
    ax.set_xlabel("Age, years", fontsize=8)
    ax.set_ylabel("Gap vs not limited, pp")
    ax.set_title("B  Internet gap within households", fontsize=8.5, loc="left")
    _grid(ax)
    h = [Line2D([], [], ls="none", marker=GALI_MRK[g], color=GALI_COL[g], ms=5, label=GALI_EN[g]) for g in gl]
    h += [Line2D([], [], ls="none", marker="o", color="#52514e", ms=5, label="Own use"),
          Line2D([], [], ls="none", marker="o", mfc="white", mec="#52514e", ms=5, label="Own or household member (A)")]
    fig.legend(handles=h, frameon=False, ncol=5, loc="lower center", bbox_to_anchor=(0.5, -0.01), fontsize=7,
               handletextpad=0.3, columnspacing=1.2)
    fig.tight_layout(rect=(0, 0.07, 1, 1), w_pad=1.5)
    _save(fig, out, "Figure4")


# ------------------------------------------------------------------ main stage
def asama_11(ctx: Ctx):
    _ensure([("docx", "python-docx"), ("matplotlib", "matplotlib")])
    plt = _fig_setup()
    rep = Report(ctx, "11_manuscript", "Aşama 11 — Makale çıktıları (Healthcare, MDPI; İngilizce)")
    out = rep.dir

    # ---------------- FIGURES
    made = []
    figs = ((_fig_framework, "Figure1"), (_fig1, "Figure2"), (_fig2, "Figure3"), (_fig4, "Figure4"),
            (_fig_left_offline, "Figure5"), (lambda c, o, p: _fig_framework(c, o, p, graphical=True), "Graphical_Abstract"),
            (_figS1, "FigureS1"), (_figS2, "FigureS2"), (_figS3_rr, "FigureS3"), (_figS4_age, "FigureS4"),
            (_figS5_fe, "FigureS5"), (_figS6_spill, "FigureS6"), (_figS7_trend, "FigureS7"))
    for fn, nm in figs:
        try:
            fn(ctx, out, plt)
            made.append(nm)
        except (FileNotFoundError, IndexError, KeyError) as e:
            rep.p(f"**{nm} üretilemedi:** {e}")

    # ---------------- ANA TABLOLAR
    doc = _doc()
    _heading(doc, "Main tables", 13)
    T1 = _table1(ctx)
    _caption(doc, "Table 1", "Weighted characteristics of the study population by global activity limitation "
             "(GALI), individuals aged 16–74 years, Türkiye, 2026.")
    _table(doc, T1)
    _note(doc, "Values are weighted column percentages unless stated otherwise. GALI: Global Activity "
          "Limitation Indicator (Eurostat definition: limitation for at least 6 months). CATI: computer-"
          "assisted telephone interviewing. Source: TurkStat Survey on ICT Usage in Households and by "
          "Individuals, 2026 microdata; authors' calculations.")

    def _g(df, o, k, g, col="olcu"):
        x = df[(df[col] == o) & (df.cikti == k) & (df.G == g)]
        return x.iloc[0] if len(x) else None

    def _c100(x):
        return "" if x is None else _ci(100 * x.tahmin, 100 * x.ga_alt, 100 * x.ga_ust)

    # Table 2: crude → adjusted (stage 09)
    try:
        H = _rcsv(ctx, "09_ham_ayarli", "ham_ayarli_uzun.csv")
        H["G"] = H["GALI"].map(_gkey)
        rows = []
        for g in [g for g in ["2", "1"] if g in set(H.G)]:
            rows.append({"Outcome": f"{GALI_EN[g]} vs not limited"})
            for k in [o for o in ["INTERNET", "BILGI_F", "BILGI_R", "MHRS", "KAYIT", "YZ"] if o in set(H.cikti)]:
                x0, x1, x2, x3 = (_g(H, sp, k, g) for sp in ("S0_ham", "S1_yas_cinsiyet", "S2_egitim",
                                                             "S3_kir_kent_bolge"))
                sh = 1 - x3.tahmin / x0.tahmin if x0.tahmin != 0 else np.nan
                rows.append({"Outcome": OUT_EN[k], "Crude, pp (95% CI)": _c100(x0),
                             "+ Age, sex": f"{100 * x1.tahmin:.1f}".replace("-", "−"),
                             "+ Education": f"{100 * x2.tahmin:.1f}".replace("-", "−"),
                             "Adjusted, pp (95% CI)": _c100(x3),
                             "Share explained": (f"{100 * sh:.0f}%" if (x3.tahmin < 0 and 0 <= sh <= 1) else "—")})
        _caption(doc, "Table 2", "Crude and adjusted gaps in digital health engagement by activity limitation, "
                 "population aged 16–74, Türkiye, 2026 (percentage points).")
        _table(doc, pd.DataFrame(rows).fillna(""), bold_first_col=True)
        _note(doc, "Two-part model (internet use × use among internet users) with g-computation; non-users coded 0 "
              "for every online outcome. Covariates are added cumulatively; the adjusted gap also includes degree of "
              "urbanisation and NUTS-1 region. Share explained = 1 − adjusted/crude; not shown (—) when adjustment "
              "reverses the sign. 95% CIs from Rao–Wu rescaled bootstrap (500 replicates).")
    except FileNotFoundError as e:
        rep.p(f"**Tablo 2 üretilemedi:** {e}")

    # Table 3: within household (stage 14)
    try:
        F = _rcsv(ctx, "14_hane_ici", "hane_ici_uzun.csv")
        F["G"] = F["GALI"].map(_gkey)
        rows = []
        for k in ["INTERNET", "MHRS_TUM", "KAYIT_TUM"]:
            r = {"Outcome": {"INTERNET": "Internet use", "MHRS_TUM": "Online appointment booking",
                             "KAYIT_TUM": "Online access to health records"}[k]}
            for g in ["1", "2"]:
                lab = "Moderate" if g == "1" else "Severe"
                r[f"{lab}: pooled"] = _c100(_g(F, "havuz", k, g))
                r[f"{lab}: within household"] = _c100(_g(F, "hane_FE", k, g))
            r["Severe, within household: 16–54"] = _c100(_g(F, "FE_16-54", k, "2"))
            r["Severe, within household: 55–74"] = _c100(_g(F, "FE_55-74", k, "2"))
            rows.append(r)
        n = _rcsv(ctx, "14_hane_ici", "orneklem.csv").iloc[0]
        _caption(doc, "Table 3", "Gaps in internet use and online health services by activity limitation: pooled "
                 "and within-household estimates, adults aged 16–74 in households with at least two respondents "
                 "(percentage points, 95% CI).")
        _table(doc, pd.DataFrame(rows), bold_first_col=True)
        _note(doc, "Weighted linear probability models adjusted for age (restricted cubic spline), sex and education; "
              "pooled models also for degree of urbanisation and NUTS-1 region. Within-household estimates use "
              f"household fixed effects and are identified from {int(n.n_bilgi_tasiyan_hane):,} households in which "
              f"limitation status differed ({int(n.n_kisitli_bilgi_tasiyan):,} adults with limitation). Online "
              "services coded 0 for internet non-users. Moderate = limited but not severely. Rao–Wu bootstrap "
              "(500 replicates).")
    except FileNotFoundError as e:
        rep.p(f"**Tablo 3 üretilemedi:** {e}")

    # Table 4: relational access (stage 13)
    try:
        Hh = _rcsv(ctx, "13_hane", "hane_uzun.csv")
        Hh["G"] = Hh["GALI"].map(_gkey)

        def pct(o, k, g):
            x = _g(Hh, o, k, g)
            return "" if x is None else f"{100 * x.tahmin:.1f}"

        spec = [("Internet use: own", "INTERNET", True), ("Internet use: own or household member", "ERISIM_ILISKI", False),
                ("Online appointment: own", "MHRS_TUM", True),
                ("Online appointment: own or household member", "MHRS_ILISKI", False),
                ("Health record access: own", "KAYIT_TUM", True),
                ("Health record access: own or household member", "KAYIT_ILISKI", False),
                ("Lives with an online adult", "HH_DIGER_INT", None), ("Digitally isolated", "YALITIK", None)]
        own_of = {"ERISIM_ILISKI": "INTERNET", "MHRS_ILISKI": "MHRS_TUM", "KAYIT_ILISKI": "KAYIT_TUM"}
        rows = []
        for lab, k, is_own in spec:
            r = {"Measure": lab}
            for g in ["0", "1", "2"]:
                r[GALI_EN[g] + ", %"] = pct("ham_oran", k, g)
            r["Crude gap, severe, pp (95% CI)"] = _c100(_g(Hh, "ham_acik", k, "2"))
            if is_own is False:
                x = _g(Hh, "ham_telafi", own_of[k], "2")
                r["Household compensation, severe (95% CI)"] = (
                    "" if x is None else f"{100 * x.tahmin:.0f}% ({100 * x.ga_alt:.0f} to {100 * x.ga_ust:.0f})")
            rows.append(r)
        for sn, lab in (("16-54", "Digitally isolated, aged 16–54"), ("55-74", "Digitally isolated, aged 55–74")):
            r = {"Measure": lab}
            for g in ["0", "1", "2"]:
                r[GALI_EN[g] + ", %"] = pct(f"oran_{sn}", "YALITIK", g)
            rows.append(r)
        _caption(doc, "Table 4", "Individual and household-inclusive (relational) digital access and digital "
                 "isolation by activity limitation, population aged 16–74, Türkiye, 2026.")
        _table(doc, pd.DataFrame(rows).fillna(""), bold_first_col=False)
        _note(doc, "Own or household member = the respondent or any other household member aged 16–74 used the "
              "internet or the service in the last three months (an upper bound of potential help, as the survey "
              "records each person's own use). Digitally isolated = neither the respondent nor any other member aged "
              "16–74 used the internet. Household compensation = 1 − relational gap / individual gap. Weighted "
              "percentages; 95% CIs from Rao–Wu bootstrap (500 replicates).")
    except FileNotFoundError as e:
        rep.p(f"**Tablo 4 üretilemedi:** {e}")

    # Table 5: AME + RR among users (stages 05, 12)
    M = _rcsv(ctx, "05_models", "ame_toplam.csv")
    M["G"] = M["GALI_duzey"].map(_gkey)
    try:
        RRd = _rcsv(ctx, "12_goreli_yas", "goreli_yas_uzun.csv")
        RRd["G"] = RRd["GALI"].map(_gkey)
    except FileNotFoundError:
        RRd = None
    order = ["INTERNET", "BILGI_F", "BILGI_R", "MHRS", "KAYIT", "KONSULT", "YZ", "AKILLI_SAAT",
             "H3a_bilgi_eksi_islem", "H3b_bilgi_eksi_sinir"]
    rows = []
    for k in [o for o in order if o in set(M.cikti)]:
        r = {"Outcome": OUT_EN[k]}
        for g in [g for g in ["1", "2"] if g in set(M.G)]:
            lab = "Moderate" if g == "1" else "Severe"
            x = M[(M.cikti == k) & (M.G == g)].iloc[0]
            padj = x.p_holm if pd.notna(x.p_holm) else x.p_bh
            r[f"{lab}: AME, pp (95% CI)"] = _ci(x.tahmin, x.ga_alt, x.ga_ust)
            r[f"{lab}: adj. p"] = _p(padj) + ("" if pd.notna(x.p_holm) else "†")
            if RRd is not None:
                y = RRd[(RRd.olcu == "RR") & (RRd.cikti == k) & (RRd.G == g)]
                if len(y):
                    y = y.iloc[0]
                    r[f"{lab}: RR (95% CI)"] = f"{np.exp(y.tahmin):.2f} ({np.exp(y.ga_alt):.2f} to {np.exp(y.ga_ust):.2f})"
        rows.append(r)
    _caption(doc, "Table 5", "Adjusted associations of activity limitation with internet use and with digital "
             "health engagement among internet users (reference: not limited), Türkiye, 2026.")
    _table(doc, pd.DataFrame(rows))
    _note(doc, "AME: average marginal effect in percentage points (pp); RR: standardised risk ratio. Survey-weighted "
          "logistic models adjusted for age (restricted cubic spline), sex, education, degree of urbanisation and "
          "NUTS-1 region. Internet use modelled among all individuals aged 16–74; other outcomes among internet users. "
          "H3a/H3b: mean AME of the information step minus that of the transaction/frontier step (RR rows: ratio of "
          "geometric mean RRs). Adjusted p: Holm within the primary family; † Benjamini–Hochberg (secondary outcomes). "
          "Moderate = limited but not severely. Rao–Wu bootstrap (500 replicates).")
    doc.save(out / "Main_Tables.docx")

    # ---------------- SUPPLEMENTARY FILES
    S = _doc()
    from docx.shared import Pt
    h = S.add_paragraph()
    rr = h.add_run("Supplementary Materials")
    rr.bold = True
    rr.font.size = Pt(16)
    q = S.add_paragraph()
    rq = q.add_run("Who is left offline when health care goes digital? Individual and household digital "
                   "resources among people with activity limitations in a nationally representative survey from Türkiye")
    rq.italic = True
    S.add_paragraph("Data: Turkish Statistical Institute (TurkStat), Survey on ICT Usage in Households and by "
                    "Individuals, 2026 (and 2025 for Table S9 and Figure S7), microdata; authors' calculations. "
                    "All estimates are survey-weighted; 95% confidence intervals (CI) from the Rao–Wu rescaled "
                    "bootstrap (500 replicates) respecting NUTS-1 strata and address-block clusters. "
                    "pp = percentage points; AME = average marginal effect; RR = risk ratio; GALI = Global "
                    "Activity Limitation Indicator. Moderate = limited but not severely.")
    _heading(S, "Contents", 11)
    toc_anchor = S.add_paragraph()
    S.add_page_break()
    _heading(S, "Supplementary tables", 12)

    # S1 variables
    s1 = pd.DataFrame([
        ["Internet use", "SON_KULLANIM_INTERNET (B1)", "When did you last use the internet? (1 = within the last 3 months)", "All 16–74", "Last 3 months"],
        ["Physical health information", "INT_FAAL_SAGLIK_BILGI_ARM_F (B3.15.1)", "Seeking information about physical health (e.g., flu symptoms, hypertension, painkillers, surgery, nutrition)", "Internet users", "Last 3 months"],
        ["Mental health information", "INT_FAAL_SAGLIK_BILGI_ARM_R (B3.15.2)", "Seeking information about mental health (e.g., burnout, depression, anxiety, eating disorders, stress reduction)", "Internet users", "Last 3 months"],
        ["Online appointment booking", "INT_FAAL_SAGLIK_RANDEVU_AL", "Making an appointment with a doctor via a website or app (hospital, health centre, e-Nabız, MHRS)", "Internet users", "Last 3 months"],
        ["Personal health records", "INT_FAAL_SAGLIK_SONUC_AL", "Accessing personal health records online (appointments, prescriptions, reports, test results)", "Internet users", "Last 3 months"],
        ["Online consultation / e-prescription", "INT_FAAL_SAGLIK_TIBKONS_RECETE", "Using other health services via a website or app instead of visiting a hospital or doctor (online consultation, online prescription)", "Internet users", "Last 3 months"],
        ["Generative AI use", "YAPAYZEKA_KULLAN", "Used generative AI tools to create text, images, code or video", "Internet users", "Last 3 months"],
        ["Smartwatch / wearable", "NESN_INT_AKILL_SAAT", "Used an internet-connected smartwatch, fitness band or other wearable", "Internet users", "Current use"],
        ["Connected health device", "NESN_INT_AKILL_SGLK_CHZ", "Used an internet-connected health device (blood pressure monitor, glucose meter, scale)", "Internet users", "Current use"],
        ["Activity limitation (GALI)", "FAALIYET_SINIRLAMA_DURUM, _SURE", "Limited because of a health problem in activities people usually do (severely / not severely / not at all); for at least the past 6 months", "All 16–74", "≥6 months"],
        ["Activity repertoire (0–10)", "10 INT_FAAL_* items", "Count of non-health internet activities (email, calls, social networks, messaging, news, goods information, opinions, consultations/voting, job search, banking)", "Internet users", "Last 3 months"],
        ["Digital skills (0–9)", "EYETENEK_* (9 items)", "Count of digital skill activities (file transfer, app installation, settings, word processing, presentations, spreadsheets, advanced spreadsheet functions*, photo/video editing, coding)", "Internet users", "Last 3 months"],
    ], columns=["Construct", "Variable (questionnaire item)", "Item wording (English translation)", "Denominator", "Reference period"])
    _caption(S, "Table S1", "Operationalisation of study variables (TurkStat ICT Usage Survey, 2026 questionnaire).")
    _table(S, s1)
    _note(S, "* Asked only of respondents who used spreadsheets; coded 0 otherwise. Item wordings are the authors' "
          "translations. Health-service items follow the corrected mapping described in Table S2.")

    # S2 label shift
    try:
        L = _rcsv(ctx, "01_audit", "etiket_kaymasi_2026.csv")
        P = _rcsv(ctx, "01_audit", "etiket_phi_2026.csv")
        P = P.set_index(P.columns[0])
        guide = {"INT_FAAL_SAGLIK_RANDEVU_AL": "B3.16 'Health information seeking' (duplicate of B3.15)",
                 "INT_FAAL_SAGLIK_SONUC_AL": "B3.17 Doctor appointment via website/app",
                 "INT_FAAL_SAGLIK_TIBKONS_RECETE": "B3.18 Access to personal health records",
                 "INT_FAAL_MALHIZMET_SATIS": "B3.20 Other online health services (consultation)"}
        used = {"INT_FAAL_SAGLIK_RANDEVU_AL": "Online appointment booking",
                "INT_FAAL_SAGLIK_SONUC_AL": "Access to personal health records",
                "INT_FAAL_SAGLIK_TIBKONS_RECETE": "Online consultation / e-prescription",
                "INT_FAAL_MALHIZMET_SATIS": "Selling goods or services (as in 2025 codebook)"}
        rows = []
        for _, x in L.iterrows():
            v = x["degisken"]
            if v not in guide:
                continue
            key = v.replace("INT_FAAL_", "")
            rows.append({"Variable": v, "Label in 2026 codebook": guide[v], "Meaning used in this study": used[v],
                         "Weighted prevalence, %": f"{100 * x.yayginlik_kullanicilarda:.1f}",
                         "Agreement with B3.15": f"{x.genel_saglik_bilgisi_ile_uyum:.2f}",
                         "φ with e-Gov appointment": f"{P.loc[key, 'EDV_RANDEVU']:.2f}",
                         "φ with e-Gov records": f"{P.loc[key, 'EDV_KAYIT']:.2f}",
                         "φ with e-commerce": f"{P.loc[key, 'ETICARET_3AY']:.2f}"})
        _caption(S, "Table S2", "Audit of the health-service item labels in the 2026 microdata codebook "
                 "(internet users, 2026).")
        _table(S, pd.DataFrame(rows))
        _note(S, "The 2026 codebook descriptions for items B3.16–B3.20 are shifted relative to the variable names: "
              "B3.16 repeats the wording of B3.15, and no item carries the selling-goods description, whereas the "
              "2025 codebook defines INT_FAAL_MALHIZMET_SATIS as selling goods or services online. Variable names, "
              "prevalence (selling: 9.8% in 2025, 11.5% in 2026), agreement with B3.15 (0.70, inconsistent with a "
              "duplicate) and weighted φ correlations with 12-month e-Government appointment and record-access "
              "items and with e-commerce all support the mapping used. Results under the codebook's literal "
              "mapping are reported in Table S8.")
    except FileNotFoundError as e:
        rep.p(f"**Table S2 üretilemedi:** {e}")

    # S3 design
    try:
        D = _rcsv(ctx, "03_design", "deft_orani.csv")
        D = D.rename(columns={D.columns[0]: "Covariate", "ortalama": "Weighted mean", "se_blok": "SE (block clusters)",
                              "se_hane": "SE (household clusters)", "oran_blok_hane": "Ratio"})
        D["Covariate"] = D["Covariate"].map({"kadin": "Women", "yuksekogretim": "Tertiary education",
                                             "yas_55plus": "Aged 55+", "kentsel_yogun": "Densely populated",
                                             "hanede_bilgisayar": "Computer in household"}).fillna(D["Covariate"])
        for c in D.columns[1:]:
            D[c] = D[c].map(lambda v: f"{v:.3f}")
        _caption(S, "Table S3", "Survey design and variance estimation: standard errors under block versus "
                 "household clustering (2026).")
        _table(S, D)
        _note(S, "Two-stage stratified cluster sample (strata: NUTS-1 regions; PSUs: enumeration blocks selected "
              "with probability proportional to size). Variance was estimated with the Rao–Wu rescaled bootstrap "
              "(m_h = n_h − 1; 500 replicates) applied within strata to PSUs. Because the 2025 file contains no "
              "block identifier, households were used as PSUs in 2025; as a sensitivity check, 2025 variances "
              "were inflated by the median 2026 block/household SE ratio "
              f"({ctx.dec.get('deft_ratio_2025')}). Weights: final person weights (FERT_FAKTOR) normalised to mean 1 within year.")
    except FileNotFoundError as e:
        rep.p(f"**Table S3 üretilemedi:** {e}")

    # S4 cascade: user denominator + all 16–74
    K = _rcsv(ctx, "04_descriptives", "kaskad_uzun.csv")
    K["G"] = K["GALI"].map(_gkey)
    rows = []
    for denom, lab in (("tum_16_74", "All aged 16–74"), ("kullanicilar", "Internet users")):
        rows.append({"Outcome": lab.upper()})
        for k in [o for o in OUT_EN if o in set(K[K.payda == denom].cikti)]:
            r = {"Outcome": OUT_EN[k]}
            for g in ["0", "1", "2", "toplam"]:
                x = K[(K.payda == denom) & (K.cikti == k) & (K.G == g)]
                if len(x):
                    x = x.iloc[0]
                    r[GALI_EN[g]] = _ci(x.tahmin, x.ga_alt, x.ga_ust) + ("‡" if x.kucuk_hucre else "")
            rows.append(r)
    _caption(S, "Table S4", "Weighted prevalence (%) of each cascade step by activity limitation, 2026.")
    _table(S, pd.DataFrame(rows).fillna(""), bold_first_col=True)
    _note(S, "95% CIs in parentheses (Rao–Wu bootstrap). ‡ fewer than 50 unweighted positive responses; "
          "interpret with caution. e-Government items refer to the last 12 months; delegation and accessibility "
          "barrier items were asked only of respondents who did not submit forms online.")

    # S5 full AMEs: total + direct
    rows = []
    for spec, lab in (("toplam", "Total-effect specification"), ("dogrudan", "Direct (resource-adjusted) specification")):
        try:
            M = _rcsv(ctx, "05_models", f"ame_{spec}.csv")
        except FileNotFoundError:
            continue
        M["G"] = M["GALI_duzey"].map(_gkey)
        rows.append({"Outcome": lab.upper()})
        for k in [o for o in order if o in set(M.cikti)]:
            r = {"Outcome": OUT_EN[k]}
            for g in ["1", "2"]:
                x = M[(M.cikti == k) & (M.G == g)]
                if len(x):
                    x = x.iloc[0]
                    r[f"{GALI_EN[g]}: AME (95% CI)"] = _ci(x.tahmin, x.ga_alt, x.ga_ust)
                    r[f"{GALI_EN[g]}: p"] = _p(x.p)
            rows.append(r)
    _caption(S, "Table S5", "Average marginal effects (pp) under the total-effect and direct specifications.")
    _table(S, pd.DataFrame(rows).fillna(""), bold_first_col=True)
    _note(S, "The direct specification additionally adjusts for household income quintile, computer in the "
          "household and household size, which may lie on the pathway from activity limitation to digital "
          "engagement; it is therefore reported as a sensitivity analysis. Unadjusted p values.")

    # S6 hurdle tam
    U = _rcsv(ctx, "06_hurdle", "hurdle_uzun.csv")
    U["G"] = U["GALI"].map(_gkey)
    rows = []
    for g in ["1", "2"]:
        if g not in set(U.G):
            continue
        rows.append({"Outcome": GALI_EN[g].upper()})
        for k in [o for o in OUT_EN if o in set(U.cikti) and o != "INTERNET"]:
            r = {"Outcome": OUT_EN[k]}
            for olc, lab, sc in (("dP_toplam", "Total gap, pp", 100), ("dP_erisim", "Access, pp", 100),
                                 ("dP_kosullu", "Conditional use, pp", 100), ("lnRR_toplam", "Risk ratio", None)):
                x = U[(U.cikti == k) & (U.G == g) & (U.olcu == olc)]
                if len(x):
                    x = x.iloc[0]
                    r[lab] = (_ci(np.exp(x.tahmin), np.exp(x.ga_alt), np.exp(x.ga_ust), 2) if sc is None
                              else _ci(sc * x.tahmin, sc * x.ga_alt, sc * x.ga_ust))
            rows.append(r)
    _caption(S, "Table S6", "Two-part decomposition of adjusted gaps into access and conditional-use components.")
    _table(S, pd.DataFrame(rows).fillna(""), bold_first_col=True)
    _note(S, "Standardised to the 2026 population aged 16–74. Component shares are not reported because "
          "total gaps were close to zero for several outcomes, making ratio-based shares unstable.")

    # S7 KHB
    try:
        KH = _rcsv(ctx, "07_khb", "khb.csv")
        lab_map = {v: k for k, v in OUTCOME_LABELS.items()}
        KH["cikti"] = KH["etiket"].map(lab_map)
        meas = {"b_azaltilmis": "Total effect (reduced model), log-odds", "b_tam": "Direct effect (full model), log-odds",
                "fark": "Indirect effect via mediators, log-odds", "pay_REPERTUAR": "Share via activity repertoire",
                "pay_BECERI": "Share via digital skills"}
        rows = []
        for k in [o for o in OUT_EN if o in set(KH.cikti)]:
            r = {"Outcome": OUT_EN[k]}
            for m_, lab in meas.items():
                x = KH[(KH.cikti == k) & (KH.olcu == m_)]
                if len(x):
                    x = x.iloc[0]
                    r[lab] = _ci(x.tahmin, x.ga_alt, x.ga_ust, 2)
            rows.append(r)
        _caption(S, "Table S7", "Karlson–Holm–Breen decomposition: association between any activity limitation "
                 "and engagement among internet users, with digital activity repertoire and skills as mediators.")
        _table(S, pd.DataFrame(rows))
        _note(S, "Negative indirect effects indicate suppression: activity limitation is associated with a narrower "
              "activity repertoire, which in turn is associated with lower engagement. Because repertoire and "
              "skills are measured concurrently and partly reflect use, estimates describe statistical "
              "decomposition, not causal mediation. Shares are unstable when the total effect is near zero.")
    except FileNotFoundError as e:
        rep.p(f"**Table S7 üretilemedi:** {e}")

    # S8 sensitivity
    try:
        T = _rcsv(ctx, "10_saglamlik", "saglamlik_uzun.csv")
        T["G"] = T["GALI"].map(_gkey)
        rows = []
        for g in ["1", "2"]:
            if g not in set(T.G):
                continue
            rows.append({"Scenario": GALI_EN[g].upper()})
            for sc in dict.fromkeys(T.senaryo):
                r = {"Scenario": SCEN_EN.get(sc, sc)}
                for k in ["INTERNET", "BILGI_F", "BILGI_R", "MHRS", "KAYIT", "YZ", "H3a_bilgi_eksi_islem", "H3b_bilgi_eksi_sinir"]:
                    x = T[(T.senaryo == sc) & (T.cikti == k) & (T.G == g)]
                    if len(x):
                        x = x.iloc[0]
                        r[OUT_EN[k].replace("Online access to personal health records", "Health records").replace(
                            " (last 3 months)", "")] = f"{x.tahmin:.1f} ({x.ga_alt:.1f}, {x.ga_ust:.1f})".replace("-", "−")
                rows.append(r)
        _caption(S, "Table S8", "Robustness of the main average marginal effects (pp, 95% CI) across "
                 "alternative specifications.")
        _table(S, pd.DataFrame(rows).fillna(""), bold_first_col=True)
        _note(S, "Variance from the first 200 bootstrap replicates. See also Figure S2.")
    except FileNotFoundError as e:
        rep.p(f"**Table S8 üretilemedi:** {e}")

    # S9 trend
    try:
        TR = _rcsv(ctx, "08_trend", "trend_uzun.csv")
        cn = {"acik_EDU": "Education gap (tertiary − primary or less)", "acik_FEMALE": "Sex gap (women − men)",
              "acik_AGEGRP": "Age gap (65–74 − 25–34)"}
        rows = []
        for k in [o for o in ["INTERNET", "BILGI_F", "BILGI_R", "YZ"] if o in set(TR.cikti)]:
            rows.append({"Measure": OUT_EN[k].upper()})
            pv = TR[(TR.cikti == k) & (TR.olcu == "yayginlik")].set_index("donem")
            rows.append({"Measure": "Prevalence, %", "2025": f"{pv.loc['2025', 'tahmin']:.1f}",
                         "2026": f"{pv.loc['2026', 'tahmin']:.1f}"})
            for a, lab in cn.items():
                d = TR[(TR.cikti == k) & (TR.olcu == a)].set_index("donem")
                if not len(d):
                    continue
                dd = d.loc["fark_farki"]
                eq = str(dd.get("TOST_esdeger", "")).strip()
                rows.append({"Measure": lab, "2025": f"{d.loc['2025', 'tahmin']:.1f}",
                             "2026": f"{d.loc['2026', 'tahmin']:.1f}",
                             "Change, pp (95% CI)": _ci(dd.tahmin, dd.ga_alt, dd.ga_ust),
                             "p (design-inflated)": _p(dd.p_deft),
                             f"Equivalent within ±{ctx.dec.get('delta_pp')} pp": "Yes" if eq == "True" else "No"})
            dfl = TR[(TR.cikti == k) & (TR.olcu == "DFL")].set_index("donem")
            if len(dfl):
                rows.append({"Measure": "Change due to composition (DFL), pp",
                             "Change, pp (95% CI)": _ci(dfl.loc["bilesim", "tahmin"], dfl.loc["bilesim", "ga_alt"],
                                                        dfl.loc["bilesim", "ga_ust"])})
        _caption(S, "Table S9", "Change in prevalence and in sociodemographic gaps between 2025 and 2026 "
                 "(pooled cross-sections).")
        _table(S, pd.DataFrame(rows).fillna(""), bold_first_col=True)
        _note(S, "Gaps are adjusted average marginal differences from models with year × covariate interactions. "
              "Equivalence: two one-sided tests with a margin fixed before outcome analysis. Design-inflated p uses "
              "2025 SEs multiplied by the 2026 block/household ratio. The 2026 questionnaire moved the generative AI "
              "item to a separate module and lengthened the activity list containing health information seeking; "
              "changes may partly reflect questionnaire effects.")
    except FileNotFoundError as e:
        rep.p(f"**Table S9 üretilemedi:** {e}")

    # S10–S11 relative scale and age strata
    try:
        G = _rcsv(ctx, "12_goreli_yas", "goreli_yas_uzun.csv")
        G["G"] = G["GALI"].map(_gkey)
        rows = []
        for k in [o for o in ["INTERNET", "BILGI_F", "BILGI_R", "MHRS", "KAYIT", "KONSULT", "YZ", "AKILLI_SAAT",
                              "H3a_bilgi_eksi_islem", "H3b_bilgi_eksi_sinir"] if o in set(G.cikti)]:
            r = {"Outcome": OUT_EN[k]}
            pr = G[(G.olcu == "P_ref") & (G.cikti == k)]
            if len(pr):
                r["Prevalence among not limited, %"] = f"{100 * pr.iloc[0].tahmin:.1f}"
            for g in ["1", "2"]:
                x = G[(G.olcu == "RR") & (G.cikti == k) & (G.G == g)]
                if len(x):
                    x = x.iloc[0]
                    r[f"{GALI_EN[g]}: RR (95% CI)"] = _ci(np.exp(x.tahmin), np.exp(x.ga_alt), np.exp(x.ga_ust), 2)
                    r[f"{GALI_EN[g]}: p"] = _p(x.p)
            rows.append(r)
        _caption(S, "Table S10", "Standardised risk ratios by activity limitation and relative-scale version "
                 "of the step contrasts (H3), Türkiye, 2026.")
        _table(S, pd.DataFrame(rows).fillna(""))
        _note(S, "RR: ratio of standardised predicted prevalences (limited / not limited) from the same models as "
              "Table 2. For H3a/H3b the entry is a ratio of risk ratios, exp[mean ln RR (information) − mean ln RR "
              "(transaction or frontier)]; values above 1 indicate a more favourable relative association in the "
              "information step. Because outcomes differ widely in baseline prevalence, contrasts on the "
              "percentage-point scale (Table 2) and the ratio scale are reported together.")
        rows = []
        for k in [o for o in ["INTERNET", "BILGI_F", "BILGI_R", "MHRS", "KAYIT", "YZ"] if o in set(G.cikti)]:
            for g in ["1", "2"]:
                r = {"Outcome": OUT_EN[k], "Limitation": GALI_EN[g]}
                for sn, lab in (("YAS_16-54", "Aged 16–54: AME, pp (95% CI)"), ("YAS_55-74", "Aged 55–74: AME, pp (95% CI)"),
                                ("YAS_fark", "Difference (55–74 minus 16–54), pp (95% CI)")):
                    x = G[(G.olcu == sn) & (G.cikti == k) & (G.G == g)]
                    if len(x):
                        x = x.iloc[0]
                        r[lab] = _ci(100 * x.tahmin, 100 * x.ga_alt, 100 * x.ga_ust)
                        if sn == "YAS_fark":
                            r["p"] = _p(x.p)
                rows.append(r)
        _caption(S, "Table S11", "Average marginal effects of activity limitation stratified by age group "
                 "(exploratory), Türkiye, 2026.")
        _table(S, pd.DataFrame(rows).fillna(""))
        _note(S, "Separate models within each age stratum with the covariates of Table 2. Internet use among all "
              "individuals; other outcomes among internet users. Unadjusted p values; exploratory analysis.")
    except FileNotFoundError as e:
        rep.p(f"**Table S10–S11 üretilemedi:** {e}")

    # S12 STROBE + PRICSSA checklist (with section names)
    strobe = [
        ("1", "Title and abstract", "Design (cross-sectional, nationally representative) indicated", "Title; Abstract"),
        ("2", "Background/rationale", "Scientific background and rationale", "Introduction"),
        ("3", "Objectives", "Specific objectives and hypotheses", "Introduction (Q1–Q4); Figure 1"),
        ("4", "Study design", "Key elements of design", "Methods 2.1"),
        ("5", "Setting", "Setting, dates of data collection", "Methods 2.1 (April 2025, April 2026)"),
        ("6", "Participants", "Eligibility, sources, selection", "Methods 2.1; Table 1"),
        ("7", "Variables", "Outcomes, exposures, confounders, mediators", "Methods 2.2–2.4; Table S1"),
        ("8", "Data sources/measurement", "Measurement and comparability", "Methods 2.2; Table S2"),
        ("9", "Bias", "Efforts to address potential bias", "Methods 2.6; Discussion 4.6"),
        ("10", "Study size", "Sample size", "Methods 2.1"),
        ("11", "Quantitative variables", "Handling of quantitative variables", "Methods 2.5 (age spline)"),
        ("12", "Statistical methods", "Models, missing data, sensitivity analyses, survey design", "Methods 2.6; Table S3"),
        ("13", "Participants (results)", "Numbers at each stage", "Results 3.1; Table 1"),
        ("14", "Descriptive data", "Characteristics of participants", "Table 1; Table S4"),
        ("15", "Outcome data", "Outcome summary measures", "Figure 2; Table S4"),
        ("16", "Main results", "Unadjusted and adjusted estimates with CIs", "Tables 2–5; Figures 2–5"),
        ("17", "Other analyses", "Subgroup, sensitivity analyses", "Tables S5–S11, S14–S16; Figures S1–S7"),
        ("18", "Key results", "Summary of key results", "Discussion 4.1"),
        ("19", "Limitations", "Sources of bias and imprecision", "Discussion 4.6"),
        ("20", "Interpretation", "Cautious overall interpretation", "Discussion"),
        ("21", "Generalisability", "External validity", "Discussion 4.6"),
        ("22", "Funding", "Funding sources", "Funding statement"),
        ("PRICSSA 1", "Survey design", "Weights, strata and clusters used", "Methods 2.6; Table S3"),
        ("PRICSSA 2", "Variance estimation", "Method and software", "Methods 2.6"),
        ("PRICSSA 3", "Subpopulation analysis", "Domain estimation rather than subsetting", "Methods 2.6"),
        ("PRICSSA 4", "Unweighted n and weighted %", "Both reported", "Table 1"),
        ("PRICSSA 5", "Code availability", "Analysis code shared", "Data Availability Statement"),
    ]
    _caption(S, "Table S12", "STROBE (cross-sectional) and PRICSSA reporting checklist.")
    _table(S, pd.DataFrame(strobe, columns=["Item", "Topic", "Recommendation", "Reported in"]))
    _note(S, "Section numbers refer to the manuscript draft and should be updated at final submission.")

    # S13 decision log
    log = ctx._abs(ctx.paths["output_dir"]) / "decision_log.md"
    dec_rows = [
        ["Item mapping", ctx.dec.get("label_mapping"), "Codebook descriptions shifted; see Table S2"],
        ["Blank activity responses", "Not applicable" if not ctx.dec.get("activity_blank_is_no") else "Coded no", "No blanks among internet users"],
        ["Physical/mental sub-items", "Structural zero", "Asked only if general health-information item = yes"],
        ["Generative AI denominator", ctx.dec.get("yz_denominator"), "Asked only of internet users"],
        ["GALI coding", ctx.dec.get("gali_rule"), "Limitation ≥6 months (Eurostat)"],
        ["Equivalence margin", f"±{ctx.dec.get('delta_pp')} pp", "Fixed from minimum detectable effects before outcome analysis"],
        ["2025 SE inflation (sensitivity)", ctx.dec.get("deft_ratio_2025"), "Median 2026 block/household SE ratio"],
        ["GALI levels", "Merged" if ctx.dec.get("gali_merge_levels") else "Three levels", "No small cells in transaction step"],
        ["Outcomes excluded from models", ", ".join(ctx.dec.get("drop_outcomes") or []) or "None",
         "<50 positive responses among severely limited"],
        ["Preregistration", "Not preregistered", "Analysis plan and dated decision log provided"],
    ]
    _caption(S, "Table S13", "Analytic decisions taken during the analysis (dated log available with the code).")
    _table(S, pd.DataFrame(dec_rows, columns=["Decision", "Choice", "Rationale"]))

    # S14 relational: adjusted gaps, potential proxies among non-users, e-government item
    try:
        Hh = _rcsv(ctx, "13_hane", "hane_uzun.csv")
        Hh["G"] = Hh["GALI"].map(_gkey)
        labs = {"INTERNET": "Internet use: own", "ERISIM_ILISKI": "Internet use: own or household member",
                "MHRS_TUM": "Online appointment: own", "MHRS_ILISKI": "Online appointment: own or household member",
                "KAYIT_TUM": "Health records: own", "KAYIT_ILISKI": "Health records: own or household member",
                "HH_DIGER_INT": "Lives with an online adult", "YALITIK": "Digitally isolated"}
        rows = []
        for k, lab in labs.items():
            r = {"Measure": lab}
            for g in ["1", "2"]:
                L = "Moderate" if g == "1" else "Severe"
                for o, on in (("ham_acik", "crude"), ("ayarli_acik", "adjusted")):
                    x = Hh[(Hh.olcu == o) & (Hh.cikti == k) & (Hh.G == g)]
                    r[f"{L}: {on} gap, pp"] = "" if not len(x) else _ci(100 * x.iloc[0].tahmin, 100 * x.iloc[0].ga_alt, 100 * x.iloc[0].ga_ust)
            rows.append(r)
        _caption(S, "Table S14", "Relational (household) digital resources: crude and adjusted gaps by activity "
                 "limitation (percentage points, 95% CI).")
        _table(S, pd.DataFrame(rows), bold_first_col=True)
        rows = []
        for k, lab in (("HH_DIGER_INT", "Has an online co-resident"), ("HH_DIGER_MHRS", "Has a co-resident who booked online")):
            for sn, sl in (("tum", "All 16–74"), ("16-54", "16–54"), ("55-74", "55–74")):
                r = {"Among internet non-users": lab, "Age": sl}
                for g in ["0", "1", "2"]:
                    x = Hh[(Hh.olcu == f"kullanmayan_vekil_{sn}") & (Hh.cikti == k) & (Hh.G == g)]
                    r[GALI_EN[g] + ", % (95% CI)"] = "" if not len(x) else _ci(100 * x.iloc[0].tahmin, 100 * x.iloc[0].ga_alt, 100 * x.iloc[0].ga_ust)
                rows.append(r)
        _table(S, pd.DataFrame(rows))
        vk = _rcsv(ctx, "13_hane", "vekalet.csv")
        vtxt = "; ".join(f"{GALI_EN[_gkey(r.GALI)]}: {r.baskasi_yapti_yuzde:.1f}% (n asked = {int(r.n_soruldu)})"
                         for r in vk.itertuples())
        _note(S, "Adjusted gaps: survey-weighted logistic regression with g-computation, adjusted for age (spline), sex, "
              "education, degree of urbanisation and NUTS-1 region; household size deliberately not adjusted. "
              "Co-residents are other household members aged 16–74. e-Government item 'someone else submitted the form "
              f"on my behalf' (asked only of a subsample): {vtxt}.")
    except FileNotFoundError as e:
        rep.p(f"**Table S14 üretilemedi:** {e}")

    # S15 within household, all outcomes
    try:
        F = _rcsv(ctx, "14_hane_ici", "hane_ici_uzun.csv")
        F["G"] = F["GALI"].map(_gkey)
        flab = {"INTERNET": "Internet use", "MHRS_TUM": "Online appointment (all)", "KAYIT_TUM": "Health records (all)",
                "BILGI_F_TUM": "Physical health information (all)", "YZ_TUM": "Generative AI (all)",
                "BILGI_F": "Physical health information (users)", "MHRS": "Online appointment (users)",
                "KAYIT": "Health records (users)", "YZ": "Generative AI (users)"}
        rows = []
        for k, lab in flab.items():
            for g in ["1", "2"]:
                r = {"Outcome": lab, "Limitation": "Moderate" if g == "1" else "Severe"}
                for o, on in (("havuz", "Pooled"), ("hane_FE", "Within household"), ("FE_16-54", "Within, 16–54"),
                              ("FE_55-74", "Within, 55–74"), ("FE_yas_fark", "Age difference")):
                    x = F[(F.olcu == o) & (F.cikti == k) & (F.G == g)]
                    r[on] = "" if not len(x) else _ci(100 * x.iloc[0].tahmin, 100 * x.iloc[0].ga_alt, 100 * x.iloc[0].ga_ust)
                rows.append(r)
        _caption(S, "Table S15", "Pooled and within-household (household fixed-effects) gaps for all outcomes, overall "
                 "and by age (percentage points, 95% CI).")
        _table(S, pd.DataFrame(rows))
        _note(S, "Weighted linear probability models; (all) = population aged 16–74 with non-users coded 0; (users) = "
              "internet users in households with at least two internet users. Age difference = (55–74) − (16–54).")
    except FileNotFoundError as e:
        rep.p(f"**Table S15 üretilemedi:** {e}")

    # S16 spillover and negative control
    try:
        Y = _rcsv(ctx, "15_yayilim", "yayilim.csv")
        Y["Outcome"] = Y["cikti"].map(OUT_EN)
        Y["Lives with"] = Y["hanede"].map({"orta kısıtlı": "Moderately limited adult", "ağır kısıtlı": "Severely limited adult"})
        Y["AME, pp (95% CI)"] = Y["AME_pp"].astype(str).str.replace("[", "(", regex=False).str.replace("]", ")", regex=False) \
            .str.replace("; ", " to ", regex=False).str.replace("-", "−", regex=False)
        Y["p"] = Y["p"].map(_p)
        _caption(S, "Table S16", "Household spillover: association of living with a limited adult with the outcomes of "
                 "internet users without limitation; generative AI as negative-control outcome.")
        _table(S, Y[["Outcome", "Lives with", "AME, pp (95% CI)", "p"]])
        _note(S, "Internet users without limitation in households with at least two adults aged 16–74. Exposure: highest "
              "limitation level among other adult household members. Adjusted for age (spline), sex, education, degree "
              "of urbanisation, NUTS-1 region and household size. Unadjusted p.")
    except FileNotFoundError as e:
        rep.p(f"**Table S16 üretilemedi:** {e}")

    S.add_page_break()
    _heading(S, "Supplementary figures", 12)
    sfig = [
        ("FigureS1", 22, "Decomposition of the adjusted population-level gap into an access component (lower internet "
         "use) and a conditional-use component (use among internet users), two-part model with g-computation "
         "(percentage points, 95% CI)."),
        ("FigureS2", 20, "Robustness of average marginal effects across nine alternative specifications "
         "(percentage points, 95% CI; variance from the first 200 bootstrap replicates)."),
        ("FigureS3", 17, "Standardised risk ratios of each outcome by activity limitation (reference: not limited; "
         "log scale, 95% CI). Internet use among all individuals aged 16–74, other outcomes among internet users; "
         "H3a/H3b rows are ratios of geometric mean risk ratios."),
        ("FigureS4", 22, "Average marginal effects of activity limitation estimated separately at ages 16–54 and "
         "55–74 (percentage points, 95% CI). Logistic models adjusted for age (spline), sex, education, degree of "
         "urbanisation and NUTS-1 region."),
        ("FigureS5", 22, "Pooled and within-household (household fixed-effects) gaps for all outcomes (weighted linear "
         "probability models, percentage points, 95% CI). (all) = population aged 16–74 with non-users coded 0; "
         "(users) = internet users."),
        ("FigureS6", 17, "Household spillover test: outcomes of internet users without limitation who live with a "
         "moderately or severely limited adult, versus those who do not (AME, percentage points, 95% CI). Generative "
         "AI use is a negative-control outcome."),
        ("FigureS7", 22, "Change between 2025 and 2026 in prevalence and in sociodemographic gaps (education: "
         "tertiary minus primary or less; age: 65–74 minus 25–34; sex: female minus male). Internet use among all "
         "individuals aged 16–74; other outcomes among internet users."),
    ]
    for nm, w, cap in sfig:
        if nm in made:
            _figure(S, out / f"{nm}.png", nm.replace("FigureS", "Figure S"), cap, width_cm=w)
    # table of contents
    from docx.shared import Pt as _Pt
    for lab, txt in S._caps:
        first = txt.split(". ")[0].rstrip(".")
        para = toc_anchor.insert_paragraph_before()
        para.paragraph_format.space_after = _Pt(1)
        r1 = para.add_run(lab + ". ")
        r1.bold = True
        r1.font.size = _Pt(9)
        r2 = para.add_run(first + ".")
        r2.font.size = _Pt(9)
    S.save(out / "Supplementary_Materials.docx")

    # Excel: all long-format tables
    sheets = [("S4_cascade", "04_descriptives", "kaskad_uzun.csv"), ("S5_AME_total", "05_models", "ame_toplam.csv"),
              ("S5_AME_direct", "05_models", "ame_dogrudan.csv"), ("S6_two_part", "06_hurdle", "hurdle_uzun.csv"),
              ("S7_KHB", "07_khb", "khb.csv"), ("S9_trend", "08_trend", "trend_uzun.csv"),
              ("T3_crude_adjusted", "09_ham_ayarli", "ham_ayarli_uzun.csv"),
              ("S8_robustness", "10_saglamlik", "saglamlik_uzun.csv"),
              ("S2_item_audit", "01_audit", "etiket_kaymasi_2026.csv"), ("S2_phi", "01_audit", "etiket_phi_2026.csv"),
              ("S3_mde", "03_design", "mde.csv"), ("S10_S11_relative_age", "12_goreli_yas", "goreli_yas_uzun.csv"),
              ("S14_relational", "13_hane", "hane_uzun.csv"), ("S15_within_household", "14_hane_ici", "hane_ici_uzun.csv"),
              ("S15_sample", "14_hane_ici", "orneklem.csv"), ("S16_spillover", "15_yayilim", "yayilim.csv")]
    with pd.ExcelWriter(out / "Supplementary_Data.xlsx") as xw:
        for nm, stg, f in sheets:
            try:
                _rcsv(ctx, stg, f).to_excel(xw, sheet_name=nm, index=False)
            except FileNotFoundError:
                pass
        T1.to_excel(xw, sheet_name="Table1", index=False)
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        wb = load_workbook(out / "Supplementary_Data.xlsx")
        info = {
            "Table1": "Weighted characteristics by activity limitation (Table 1).",
            "S2_item_audit": "Audit of 2026 health-item labels (Table S2).", "S2_phi": "Phi correlations among items (Table S2).",
            "S3_mde": "Minimum detectable effects (Table S3).", "S4_cascade": "Cascade prevalences (Table S4, Figure 2).",
            "S5_AME_total": "AMEs, total specification (Tables 5, S5).", "S5_AME_direct": "AMEs, direct specification (Table S5).",
            "S6_two_part": "Two-part decomposition (Table S6, Figure S1).", "S7_KHB": "KHB decomposition (Table S7).",
            "S8_robustness": "Sensitivity analyses (Table S8, Figure S2).", "S9_trend": "2025–2026 change (Table S9, Figure S7).",
            "T3_crude_adjusted": "Crude to adjusted gaps (Table 2, Figure 3).",
            "S10_S11_relative_age": "Risk ratios and age strata (Tables 5, S10, S11; Figures S3, S4).",
            "S14_relational": "Relational access, isolation (Tables 4, S14; Figures 4, 5).",
            "S15_within_household": "Within-household estimates (Tables 3, S15; Figures 4, S5).",
            "S15_sample": "Within-household analytic sample sizes.", "S16_spillover": "Household spillover (Table S16, Figure S6).",
        }
        ws = wb.create_sheet("README", 0)
        ws.append(["Sheet", "Content"])
        ws.append(["(all)", "Long-format estimates. Columns: tahmin = estimate, se = bootstrap SE, ga_alt/ga_ust = 95% CI "
                   "lower/upper, p = two-sided p value, cikti = outcome, olcu = measure, GALI = activity limitation "
                   "(0 not limited, 1 moderate, 2 severe). Proportions (0–1) or percentage points as indicated in the "
                   "source table. TurkStat ICT Usage Survey microdata, authors' calculations."])
        for sh in wb.sheetnames[1:]:
            ws.append([sh, info.get(sh, "")])
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="E9F1FB")
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 110
        for row in ws.iter_rows(min_row=2):
            row[1].alignment = Alignment(wrap_text=True, vertical="top")
        for sh in wb.sheetnames[1:]:
            w = wb[sh]
            w.freeze_panes = "A2"
            for c in w[1]:
                c.font = Font(bold=True)
            for col in w.columns:
                L = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                w.column_dimensions[col[0].column_letter].width = min(max(10, L + 2), 50)
        wb.save(out / "Supplementary_Data.xlsx")
    except Exception as e:  # formatting is optional
        rep.p(f"Excel biçimlendirme atlandı: {e}")

    rep.p("Üretilen dosyalar (`outputs/11_manuscript/`):")
    rep.p("- `Main_Tables.docx` — Table 1–5 (düzenlenebilir Word tabloları)\n"
          f"- Şekiller (PNG 600 dpi + PDF vektör): {', '.join(made)}\n"
          "- `Graphical_Abstract.png` — dergi sistemine ayrı yüklenir\n"
          "- `Supplementary_Materials.docx` — Table S1–S16, Figure S1–S7 (içindekiler dahil)\n"
          "- `Supplementary_Data.xlsx` — tüm tahminlerin uzun tabloları")
    rep.p("**Tablo 1 önizleme:**")
    rep.table(T1, "Table1")
    rep.checkpoint([
        "Şekilleri açıp etiket çakışması / taşma var mı bakın (özellikle Figure 1 ve 3 eksen etiketleri).",
        "Main_Tables.docx'teki tablolar Word'de düzgün görünüyor mu?",
        "Table S1 soru çevirilerini kılavuzla karşılaştırın; terim tercihleri (ör. 'activity limitation') uygun mu?",
    ])
    rep.write()


# =============================================================================
# S12_GORELI_YAS — relative scale (risk ratios) and age strata
# =============================================================================
# (a) Standardised risk ratios: RR(l) = Σw p(E=l) / Σw p(E=ref), same models as stage 05
#     ("toplam" specification). Relative H3: mean ln RR(information) − mean ln RR(transaction/frontier);
#     exponentiated and reported as a "ratio of risk ratios" (> 1 = more favourable for information).
# (b) Age strata: separate models for <55 and 55+; GALI AMEs and the difference between strata.
# All in one replicate function → common design-based CIs.

YAS_ESIK = 55


def asama_12(ctx: Ctx):
    ctx.require("gali_merge_levels")
    ctx.log_decisions("12_goreli_yas")
    cfg = ctx.cfg["models"]
    rep = Report(ctx, "12_goreli_yas", "Aşama 12 — Göreli ölçek (risk oranı) ve yaş katmanları")
    A_all = load(ctx.cache("analysis.pkl"))
    R_all = load(ctx.cache(f"rep_B{ctx.B}.pkl"))
    m = (A_all.YEAR == 2026).to_numpy()
    A = A_all[m].reset_index(drop=True)
    R = R_all[m]
    W = A["W"].to_numpy(float)
    E, levels, ref = exposure(ctx)
    rhs = cfg["total_rhs"]
    drop = set(ctx.dec.get("drop_outcomes") or [])
    users = (A["INTERNET"] == 1).to_numpy()
    outs = [k for k in ["BILGI_F", "BILGI_R", "MHRS", "KAYIT", "KONSULT", "YZ", "AKILLI_SAAT"]
            if k not in drop and A[k].notna().sum() > 0]
    steps = {s: [k for k in ks if k in outs] for s, ks in cfg["steps"].items()}

    # ---- (a) risk ratios
    rr_models = {"INTERNET": Logit(A, f"INTERNET ~ C({E}) + {rhs}")}
    for k in outs:
        rr_models[k] = Logit(A, f"{k} ~ C({E}) + {rhs}", subset=users)
    rr_X = {k: {g: md.cf_matrix(md.rows, **{E: g}) for g in [ref] + levels} for k, md in rr_models.items()}

    # ---- (b) age strata
    young = (A["AGE"] < YAS_ESIK).to_numpy()
    strata = {"16-54": young, "55-74": ~young}
    age_outs = ["INTERNET"] + [k for k in ["BILGI_F", "BILGI_R", "MHRS", "KAYIT", "YZ"] if k in outs]
    age_fns, age_skip = {}, []
    for sn, sm in strata.items():
        for k in age_outs:
            sub = sm if k == "INTERNET" else (sm & users)
            try:
                md = Logit(A, f"{k} ~ C({E}) + {rhs}", subset=sub)
                md.fit(W)
                age_fns[(sn, k)] = ame_fn(md, E, levels, ref)
            except Exception as e:  # sparse cell / non-convergence
                age_skip.append(f"{sn}·{k}: {e}")

    def fn(w):
        out = {}
        lr = {}
        for k, md in rr_models.items():
            b = md.fit(w)
            ww = w[md.rows]
            p = {g: np.sum(ww * expit(X @ b)) / np.sum(ww) for g, X in rr_X[k].items()}
            for l in levels:
                lr[(k, l)] = np.log(p[l] / p[ref])
                out[f"RR|{k}|{l}"] = lr[(k, l)]
                out[f"P_ref|{k}|{l}"] = p[ref]
        for l in levels:
            mean = {s: np.mean([lr[(k, l)] for k in ks]) if ks else np.nan for s, ks in steps.items()}
            out[f"RR|H3a_bilgi_eksi_islem|{l}"] = mean["bilgi"] - mean["islem"]
            out[f"RR|H3b_bilgi_eksi_sinir|{l}"] = mean["bilgi"] - mean["sinir"]
        for (sn, k), f in age_fns.items():
            for l, v in f(w).items():
                out[f"YAS_{sn}|{k}|{l}"] = v
        for k in age_outs:
            for l in [str(x) for x in levels]:
                a, b_ = out.get(f"YAS_55-74|{k}|{l}"), out.get(f"YAS_16-54|{k}|{l}")
                if a is not None and b_ is not None:
                    out[f"YAS_fark|{k}|{l}"] = a - b_
        return pd.Series(out)

    point, reps = replicate(fn, W, R, progress="göreli/yaş")
    s = summarize(point, reps)
    idx = s.index.str.split("|", expand=True)
    s["olcu"], s["cikti"], s["GALI"] = (idx.get_level_values(i) for i in range(3))
    s = s.reset_index(drop=True)
    s.to_csv(rep.dir / "goreli_yas_uzun.csv", index=False, encoding="utf-8-sig")

    # ---- report: risk ratios
    rep.p(f"Maruziyet: {E}; referans {ref}. Modeller 05'teki 'toplam' belirtim. "
          "RR = standartlaştırılmış risk oranı (kısıtlı / kısıtlı değil). INTERNET tüm 16–74, diğerleri internet kullanıcıları.")
    rows = []
    for k in ["INTERNET"] + outs + ["H3a_bilgi_eksi_islem", "H3b_bilgi_eksi_sinir"]:
        r = {"cikti": k, "etiket": OUTCOME_LABELS.get(k, k)}
        for l in levels:
            x = s[(s.olcu == "RR") & (s.cikti == k) & (s.GALI == str(l))]
            if len(x):
                x = x.iloc[0]
                r[f"RR_GALI{l}"] = f"{np.exp(x.tahmin):.2f} [{np.exp(x.ga_alt):.2f}; {np.exp(x.ga_ust):.2f}]"
                r[f"p_GALI{l}"] = round(float(x.p), 4)
            pr = s[(s.olcu == "P_ref") & (s.cikti == k) & (s.GALI == str(l))]
            if len(pr):
                r["temel_oran_%"] = round(100 * float(pr.iloc[0].tahmin), 1)
        rows.append(r)
    rep.h("Standartlaştırılmış risk oranları ve göreli H3")
    rep.table(pd.DataFrame(rows), "risk_oranlari")
    rep.p("H3 satırları 'risk oranlarının oranı'dır: exp[ort. ln RR(bilgi) − ort. ln RR(işlem veya sınır)]. "
          "> 1 ise kısıtlılığın göreli ilişkisi bilgi basamağında daha olumlu. Yüzde puan H3 ile aynı yönde "
          "ve anlamlıysa bulgu ölçekten bağımsızdır; göreli ölçekte kayboluyorsa ölçek artefaktı olabilir.")

    # ---- report: age strata
    rows = []
    for k in age_outs:
        for l in levels:
            r = {"cikti": k, "etiket": OUTCOME_LABELS.get(k, k), "GALI": l}
            for sn in list(strata) + ["fark"]:
                x = s[(s.olcu == f"YAS_{sn}") & (s.cikti == k) & (s.GALI == str(l))]
                if len(x):
                    x = x.iloc[0]
                    r[f"{sn}_AME_pp"] = f"{100 * x.tahmin:.1f} [{100 * x.ga_alt:.1f}; {100 * x.ga_ust:.1f}]"
                    if sn == "fark":
                        r["fark_p"] = round(float(x.p), 4)
            rows.append(r)
    rep.h(f"Yaş katmanları: <{YAS_ESIK} ve {YAS_ESIK}+ (AME, yüzde puan)")
    rep.table(pd.DataFrame(rows), "yas_katmanlari")
    rep.p("`fark` = (55–74 AME) − (16–54 AME). Pozitif fark: kısıtlılığın ilişkisi yaşlılarda daha olumlu "
          "(veya daha az olumsuz). Düzeltilmemiş p; keşifsel analiz.")
    if age_skip:
        rep.p("**Tahmin edilemeyen katman/çıktı:** " + "; ".join(age_skip))
    bad = s[s.gecersiz_rep > 0]
    if len(bad):
        rep.p(f"**Uyarı:** {len(bad)} tahminde yakınsamayan replikasyon var.")
    rep.checkpoint([
        "Göreli H3 yüzde puan H3 ile aynı yönde mi? (Ölçek artefaktı kontrolü)",
        "Yaş katmanları arasında GALI etkisi farklı mı? Fark varsa tartışmada 'yaşla etkileşim' olarak ele alınır.",
    ])
    rep.write()


# =============================================================================
# S13_HANE — relational (household) digital access and digital isolation
# =============================================================================
# Framework: Andersen behavioral model — predisposing (age, education),
# individual enabling (own internet access), RELATIONAL enabling
# (an online adult in the same household: 'warm expert' / proxy use) and need (GALI).
# Household links use the real household identifier (NOT statistical matching).
# Measures (2026, ages 16–74):
#   HH_DIGER_INT   : another household member aged 16–74 used the internet in the last 3 months
#   HH_DIGER_MHRS  : ... another member booked a doctor appointment online
#   ERISIM_ILISKI  : the respondent OR another member uses the internet (relational access)
#   MHRS_ILISKI    : the respondent OR another member books appointments online (relational appointment)
#   YALITIK        : the respondent is offline AND no other member aged 16–74 is online (digital isolation)
# Compensation share = 1 − (relational gap / individual gap).

ILISKI_CIKTI = {
    "INTERNET": "Kendi internet kullanımı",
    "ERISIM_ILISKI": "İlişkisel erişim (kendisi veya hane)",
    "MHRS_TUM": "Kendi çevrimiçi randevu",
    "MHRS_ILISKI": "İlişkisel randevu (kendisi veya hane)",
    "KAYIT_TUM": "Kendi sağlık kaydı erişimi",
    "KAYIT_ILISKI": "İlişkisel kayıt erişimi (kendisi veya hane)",
    "YALITIK": "Dijital yalıtılmış (kendisi ve hane çevrimdışı)",
    "HH_DIGER_INT": "Hanede çevrimiçi başka fert var",
}
CIFTLER = [("INTERNET", "ERISIM_ILISKI"), ("MHRS_TUM", "MHRS_ILISKI"), ("KAYIT_TUM", "KAYIT_ILISKI")]


def _hane_degiskenleri(ctx, A):
    """Add household indicators to A (2026, 16–74); return the coverage-check table."""
    raw = load(ctx.cache("raw_2026.pkl"))
    yas = num(raw["YAS"])
    hid = raw["BULTEN_NO"].astype(str)
    roster = pd.DataFrame({"hid": hid, "yas": yas})
    g = roster.groupby("hid")
    kad = pd.DataFrame({
        "n_fert_ham": g.size(),
        "n_1674_ham": g["yas"].apply(lambda x: int(x.between(16, 74).sum())),
        "n_alti16": g["yas"].apply(lambda x: int((x < 16).sum())),
        "n_75ustu": g["yas"].apply(lambda x: int((x > 74).sum())),
    })
    A = A.copy()
    A["hid"] = A["BULTEN_NO"].astype(str)
    A["MHRS_TUM"] = all_denominator(A, "MHRS")
    A["KAYIT_TUM"] = all_denominator(A, "KAYIT")
    for k in ("INTERNET", "MHRS_TUM", "KAYIT_TUM"):
        x = A[k].fillna(0.0)
        tot = x.groupby(A["hid"]).transform("sum")
        A[f"_diger_{k}"] = ((tot - x) > 0).astype(float)
    A["n_1674_veri"] = A.groupby("hid")["hid"].transform("size")
    A = A.join(kad, on="hid")
    A["HH_DIGER_INT"] = A["_diger_INTERNET"]
    A["HH_DIGER_MHRS"] = A["_diger_MHRS_TUM"]
    A["ERISIM_ILISKI"] = np.maximum(A["INTERNET"], A["HH_DIGER_INT"]).where(A["INTERNET"].notna())
    A["MHRS_ILISKI"] = np.maximum(A["MHRS_TUM"], A["HH_DIGER_MHRS"]).where(A["MHRS_TUM"].notna())
    A["KAYIT_ILISKI"] = np.maximum(A["KAYIT_TUM"], A["_diger_KAYIT_TUM"]).where(A["KAYIT_TUM"].notna())
    A["YALITIK"] = ((A["INTERNET"] == 0) & (A["HH_DIGER_INT"] == 0)).astype(float).where(A["INTERNET"].notna())
    A["TEK_YETISKIN"] = (A["n_1674_veri"] == 1).astype(float)
    A["YAS55"] = np.where(A["AGE"] >= YAS_ESIK, "55-74", "16-54")
    kapsam = pd.DataFrame([dict(
        hane_sayisi=int(A["hid"].nunique()),
        fert_16_74_veride=int(len(A)),
        fert_16_74_kadroda=int(kad.loc[A["hid"].unique(), "n_1674_ham"].sum()),
        bireysel_veri_eksik_hane_orani=float((A.groupby("hid")["n_1674_veri"].first()
                                              < A.groupby("hid")["n_1674_ham"].first()).mean()),
        internet_eksik_oran=float(A["INTERNET"].isna().mean()),
        mhrs_tum_eksik_oran=float(A["MHRS_TUM"].isna().mean()),
    )])
    return A, kapsam


def asama_13(ctx: Ctx):
    ctx.require("gali_merge_levels")
    ctx.log_decisions("13_hane")
    cfg = ctx.cfg["models"]
    rep = Report(ctx, "13_hane", "Aşama 13 — İlişkisel (hane içi) dijital erişim")
    A_all = load(ctx.cache("analysis.pkl"))
    R_all = load(ctx.cache(f"rep_B{ctx.B}.pkl"))
    m = (A_all.YEAR == 2026).to_numpy()
    A = A_all[m].reset_index(drop=True)
    R = R_all[m]
    E, levels, ref = exposure(ctx)
    A, kapsam = _hane_degiskenleri(ctx, A)
    W = A["W"].to_numpy(float)
    rhs = cfg["total_rhs"]

    rep.h("Kapsam denetimi (yöntemin ön koşulu)")
    rep.table(kapsam, "kapsam", digits=3)
    rep.p("Yöntem, hanedeki TÜM 16–74 yaş fertlerin bireysel soru kâğıdını doldurmuş olmasına dayanır. "
          "`fert_16_74_veride` ≈ `fert_16_74_kadroda` ve `bireysel_veri_eksik_hane_orani` ≈ 0 olmalı. "
          "Değilse (ör. hanede tek fert seçiliyorsa) hane göstergeleri eksik sayılır ve bu aşama yorumlanmaz.")

    gm = A[E].notna().to_numpy()
    lv_all = [ref] + levels
    Eg = A[E].to_numpy()
    young = (A["YAS55"] == "16-54").to_numpy()

    # adjusted models (all 16–74; household size NOT adjusted for relational measures:
    # a larger household is the compensation mechanism itself)
    outs = list(ILISKI_CIKTI)
    models, Xcf = {}, {}
    atlanan = []
    for k in outs:
        try:
            md = Logit(A, f"{k} ~ C({E}) + {rhs}", subset=gm)
            md.fit(W)
        except Exception as e:  # e.g. prevalence ~100% → separation
            atlanan.append(f"{k}: {e}")
            continue
        models[k] = md
        Xcf[k] = {g: md.cf_matrix(md.rows, **{E: g}) for g in lv_all}

    def fn(w):
        out = {}
        for k in outs:
            x = A[k].to_numpy(float)
            ok = gm & ~np.isnan(x)
            pr = {}
            for g in lv_all:
                s = ok & (Eg == g)
                pr[g] = np.sum(w[s] * x[s]) / np.sum(w[s])
                out[f"ham_oran|{k}|{g}"] = pr[g]
            for sn, sm in (("16-54", young), ("55-74", ~young)):
                for g in lv_all:
                    s = ok & sm & (Eg == g)
                    out[f"oran_{sn}|{k}|{g}"] = np.sum(w[s] * x[s]) / np.sum(w[s])
            for l in levels:
                out[f"ham_acik|{k}|{l}"] = pr[l] - pr[ref]
            if k in models:
                b = models[k].fit(w)
                ww = w[models[k].rows]
                p = {g: np.sum(ww * expit(X @ b)) / np.sum(ww) for g, X in Xcf[k].items()}
                for l in levels:
                    out[f"ayarli_acik|{k}|{l}"] = p[l] - p[ref]
        for bir, ili in CIFTLER:
            for l in levels:
                for t in ("ham", "ayarli"):
                    if f"{t}_acik|{ili}|{l}" not in out or f"{t}_acik|{bir}|{l}" not in out:
                        continue
                    a_b, a_i = out[f"{t}_acik|{bir}|{l}"], out[f"{t}_acik|{ili}|{l}"]
                    out[f"{t}_telafi|{bir}|{l}"] = 1 - a_i / a_b if abs(a_b) > 1e-9 else np.nan
        # among internet non-users: share with an online household member
        nu = gm & (A["INTERNET"] == 0).to_numpy()
        for g in lv_all:
            for sn, sm in (("tum", np.ones_like(nu)), ("16-54", young), ("55-74", ~young)):
                s = nu & sm & (Eg == g)
                out[f"kullanmayan_vekil_{sn}|HH_DIGER_INT|{g}"] = (
                    np.sum(w[s] * A["HH_DIGER_INT"].to_numpy(float)[s]) / np.sum(w[s]) if np.sum(w[s]) > 0 else np.nan)
                out[f"kullanmayan_vekil_{sn}|HH_DIGER_MHRS|{g}"] = (
                    np.sum(w[s] * A["HH_DIGER_MHRS"].to_numpy(float)[s]) / np.sum(w[s]) if np.sum(w[s]) > 0 else np.nan)
        return pd.Series(out)

    point, reps = replicate(fn, W, R, progress="hane")
    s = summarize(point, reps)
    idx = s.index.str.split("|", expand=True)
    s["olcu"], s["cikti"], s["GALI"] = (idx.get_level_values(i) for i in range(3))
    s = s.reset_index(drop=True)
    s.to_csv(rep.dir / "hane_uzun.csv", index=False, encoding="utf-8-sig")

    def cell(olcu, k, g, pct=True, d=1):
        x = s[(s.olcu == olcu) & (s.cikti == k) & (s.GALI == str(g))]
        if not len(x):
            return ""
        x = x.iloc[0]
        f = 100 if pct else 1
        return f"{f * x.tahmin:.{d}f} [{f * x.ga_alt:.{d}f}; {f * x.ga_ust:.{d}f}]"

    # unweighted n, GALI × age
    nt = A[gm].groupby([E, "YAS55"]).size().unstack()
    rep.h("Örneklem: GALI × yaş (ağırlıksız n)")
    rep.table(nt, "n_gali_yas", digits=0, index=True)

    rep.h("Yaygınlıklar (%, ağırlıklı) — tüm 16–74")
    rows = []
    for k, lab in ILISKI_CIKTI.items():
        r = {"cikti": k, "etiket": lab}
        for g in lv_all:
            r[f"GALI{g}"] = cell("ham_oran", k, g)
        rows.append(r)
    rep.table(pd.DataFrame(rows), "yayginlik")

    rep.h("Yaygınlıklar yaş katmanında (%, ağırlıklı)")
    rows = []
    for k in ("INTERNET", "ERISIM_ILISKI", "MHRS_TUM", "MHRS_ILISKI", "YALITIK", "HH_DIGER_INT"):
        for sn in ("16-54", "55-74"):
            r = {"cikti": k, "yas": sn}
            for g in lv_all:
                r[f"GALI{g}"] = cell(f"oran_{sn}", k, g)
            rows.append(r)
    rep.table(pd.DataFrame(rows), "yayginlik_yas")

    rep.h("Bireysel ve ilişkisel açık (yüzde puan) ve hane telafi payı")
    rows = []
    for bir, ili in CIFTLER:
        for l in levels:
            rows.append(dict(
                olcu=bir, GALI=l,
                ham_bireysel=cell("ham_acik", bir, l), ham_iliskisel=cell("ham_acik", ili, l),
                ham_telafi_pay=cell("ham_telafi", bir, l, pct=True, d=0),
                ayarli_bireysel=cell("ayarli_acik", bir, l), ayarli_iliskisel=cell("ayarli_acik", ili, l),
                ayarli_telafi_pay=cell("ayarli_telafi", bir, l, pct=True, d=0)))
    rep.table(pd.DataFrame(rows), "telafi")
    rep.p("Telafi payı = 1 − ilişkisel açık / bireysel açık (%). Ayarlı modeller 05'teki 'toplam' belirtim "
          "(yaş spline, cinsiyet, eğitim, kır-kent, bölge); hane büyüklüğü bilerek ayarlanmadı, çünkü "
          "geniş hane telafinin mekanizmasıdır. Ayarlı bireysel açık küçükse (GA 0'ı içeriyorsa) ayarlı payı yorumlamayın.")

    rep.h("Dijital yalıtılmışlık ve hanede çevrimiçi fert: ayarlı açık (yüzde puan)")
    rows = []
    for k in ("YALITIK", "HH_DIGER_INT"):
        for l in levels:
            rows.append(dict(cikti=k, GALI=l, ham=cell("ham_acik", k, l), ayarli=cell("ayarli_acik", k, l)))
    rep.table(pd.DataFrame(rows), "yalitik")

    rep.h("İnternet kullanmayanlar içinde: hanede çevrimiçi / MHRS kullanan başka fert olanların payı (%)")
    rows = []
    for k in ("HH_DIGER_INT", "HH_DIGER_MHRS"):
        for sn in ("tum", "16-54", "55-74"):
            r = {"gosterge": k, "yas": sn}
            for g in lv_all:
                r[f"GALI{g}"] = cell(f"kullanmayan_vekil_{sn}", k, g)
            rows.append(r)
    rep.table(pd.DataFrame(rows), "kullanmayan_vekil")
    nu_n = A[gm & (A["INTERNET"] == 0)].groupby([E, "YAS55"]).size().unstack()
    rep.p("İnternet kullanmayan ağırlıksız n (GALI × yaş):")
    rep.table(nu_n, "n_kullanmayan", digits=0, index=True)

    # e-government proxy item (small subsample, descriptive)
    rows = []
    for g in lv_all:
        s_ = A[gm & (Eg == g) & A["VEKALET"].notna()]
        rows.append(dict(GALI=g, n_soruldu=len(s_),
                         baskasi_yapti_yuzde=100 * wmean(s_["VEKALET"], s_["W"]) if len(s_) else np.nan,
                         n_evet=int((s_["VEKALET"] == 1).sum())))
    rep.h("e-Devlet: 'formu başkası gönderdi' (yalnızca sorulan alt küme, betimsel)")
    rep.table(pd.DataFrame(rows), "vekalet", digits=1)

    if atlanan:
        rep.p("**Ayarlı modeli kurulamayan çıktılar (yalnızca ham açık raporlandı):** " + "; ".join(atlanan))
    bad = s[s.gecersiz_rep > 0]
    if len(bad):
        rep.p(f"**Uyarı:** {len(bad)} tahminde yakınsamayan replikasyon var.")
    rep.checkpoint([
        "Kapsam denetimi geçti mi? (hanedeki tüm 16–74 fertlerin verisi var mı)",
        "İnternet kullanmayan ağır kısıtlıların yüzde kaçının hanesinde çevrimiçi biri var? 55–74'te bu pay ne?",
        "İlişkisel ölçüde açık ne kadar küçülüyor (telafi payı)? Randevuda erişimden fazla mı?",
        "Dijital yalıtılmışlık ağır kısıtlı 55–74'te ne düzeyde ve ayarlı açık anlamlı mı?",
        "Hücre n'leri yeterli mi (ağır kısıtlı, internet kullanmayan, 16–54 hücresi küçük olabilir)?",
    ])
    rep.write()


# =============================================================================
# S14_HANE_ICI — within-household comparison (household fixed effects)
# =============================================================================
# Compares adults with and without limitation living in the same household.
# Household fixed effects remove ALL household-level confounders (observed and unobserved),
# such as connection, devices, income, region, urbanisation, culture and interviewer.
# Within-household gap: the INDIVIDUAL barrier associated with limitation. The difference from a
# pooled model (no household effects) in the same sample → size of household-level confounding / shared resources.
# Estimator: weighted linear probability model with within-household demeaning (within estimator).
# Covariates: age spline, sex, education (those that vary within households).

FE_CIKTI = {
    "INTERNET": ("tum", "Son 3 ayda internet"),
    "MHRS_TUM": ("tum", "Çevrimiçi randevu (tüm 16–74)"),
    "KAYIT_TUM": ("tum", "Sağlık kaydı erişimi (tüm 16–74)"),
    "BILGI_F_TUM": ("tum", "Fiziksel sağlık bilgisi (tüm 16–74)"),
    "YZ_TUM": ("tum", "Üretken YZ (tüm 16–74)"),
    "BILGI_F": ("kullanici", "Fiziksel sağlık bilgisi (kullanıcılar)"),
    "MHRS": ("kullanici", "Çevrimiçi randevu (kullanıcılar)"),
    "KAYIT": ("kullanici", "Sağlık kaydı erişimi (kullanıcılar)"),
    "YZ": ("kullanici", "Üretken YZ (kullanıcılar)"),
}


def _demean(M, g, w):
    """Weighted within-household demeaning: M − weighted household mean."""
    sw = np.bincount(g, weights=w)  # households with zero weight in a replicate: mean 0, contribution 0
    out = np.empty_like(M)
    for j in range(M.shape[1]):
        mu = np.divide(np.bincount(g, weights=w * M[:, j]), sw, out=np.zeros_like(sw), where=sw > 0)
        out[:, j] = M[:, j] - mu[g]
    return out


def asama_14(ctx: Ctx):
    ctx.require("gali_merge_levels")
    ctx.log_decisions("14_hane_ici")
    rep = Report(ctx, "14_hane_ici", "Aşama 14 — Hane içi karşılaştırma (hane sabit etkileri)")
    A_all = load(ctx.cache("analysis.pkl"))
    R_all = load(ctx.cache(f"rep_B{ctx.B}.pkl"))
    m = (A_all.YEAR == 2026).to_numpy()
    A = A_all[m].reset_index(drop=True)
    R = R_all[m]
    E, levels, ref = exposure(ctx)
    A["hid"] = A["BULTEN_NO"].astype(str)
    for k in ("MHRS", "KAYIT", "BILGI_F", "YZ"):
        A[f"{k}_TUM"] = all_denominator(A, k)
    A["OLD"] = (A["AGE"] >= YAS_ESIK).astype(float)
    # covariate design (variables that vary within households)
    base = A[E].notna() & A["EDU"].notna()
    Xc = patsy.dmatrix("cr(AGE, df=4, constraints='center') + FEMALE + C(EDU)", A[base],
                       return_type="dataframe")
    COV = np.full((len(A), Xc.shape[1]), np.nan)
    COV[base.to_numpy()] = Xc.to_numpy(float)
    COV = COV[:, 1:]  # intercept drops out after demeaning
    D = np.column_stack([(A[E] == l).astype(float).to_numpy() for l in levels])
    DO = D * A["OLD"].to_numpy()[:, None]            # age interaction (secondary model)
    REGURB = patsy.dmatrix("C(REGION) + C(URBAN)", A[base], return_type="dataframe").to_numpy(float)
    RU = np.full((len(A), REGURB.shape[1]), np.nan)
    RU[base.to_numpy()] = REGURB
    RU = RU[:, 1:]

    specs, nrows = {}, []
    for k, (payda, lab) in FE_CIKTI.items():
        y = A[k].to_numpy(float)
        ok = base.to_numpy() & ~np.isnan(y)
        if payda == "kullanici":
            ok &= (A["INTERNET"] == 1).to_numpy()
        # households with at least 2 eligible members
        cnt = pd.Series(ok.astype(int)).groupby(A["hid"]).transform("sum").to_numpy()
        ok &= cnt >= 2
        # households where GALI varies (informative households)
        gv = A.loc[ok].groupby("hid")[E].nunique()
        inf_h = set(gv[gv > 1].index)
        inf = ok & A["hid"].isin(inf_h).to_numpy()
        idx = np.where(ok)[0]
        g = pd.factorize(A["hid"].to_numpy()[idx])[0]
        specs[k] = dict(idx=idx, g=g, y=y[idx], D=D[idx], DO=DO[idx], C=COV[idx], RU=RU[idx], lab=lab)
        nrows.append(dict(cikti=k, etiket=lab, payda=payda, n_fert=int(ok.sum()),
                          n_hane=int(len(set(A["hid"].to_numpy()[idx]))),
                          n_bilgi_tasiyan_hane=len(inf_h), n_bilgi_tasiyan_fert=int(inf.sum()),
                          n_kisitli_bilgi_tasiyan=int((inf & (A[E] > 0).to_numpy()).sum())))

    def fn(w_full):
        out = {}
        for k, sp in specs.items():
            w = w_full[sp["idx"]]
            if np.sum(w) <= 0:
                continue
            y, g = sp["y"], sp["g"]
            # (1) pooled: same sample, no household effects, region + urbanisation included
            Xp = np.column_stack([np.ones(len(y)), sp["D"], sp["C"], sp["RU"]])
            bp = wls(Xp, y, w)
            # (2) household fixed effects
            Xf = _demean(np.column_stack([sp["D"], sp["C"]]), g, w)
            yf = _demean(y[:, None], g, w)[:, 0]
            bf = wls(Xf, yf, w)
            # (3) household fixed effects + GALI × 55+
            Xi = _demean(np.column_stack([sp["D"], sp["DO"], sp["C"]]), g, w)
            bi = wls(Xi, yf, w)
            L = len(levels)
            for j, l in enumerate(levels):
                out[f"havuz|{k}|{l}"] = bp[1 + j]
                out[f"hane_FE|{k}|{l}"] = bf[j]
                out[f"fark_FE_havuz|{k}|{l}"] = bf[j] - bp[1 + j]
                out[f"FE_16-54|{k}|{l}"] = bi[j]
                out[f"FE_55-74|{k}|{l}"] = bi[j] + bi[L + j]
                out[f"FE_yas_fark|{k}|{l}"] = bi[L + j]
        return pd.Series(out)

    W = A["W"].to_numpy(float)
    point, reps = replicate(fn, W, R, progress="hane içi")
    s = summarize(point, reps)
    idx = s.index.str.split("|", expand=True)
    s["olcu"], s["cikti"], s["GALI"] = (idx.get_level_values(i) for i in range(3))
    s = s.reset_index(drop=True)
    s.to_csv(rep.dir / "hane_ici_uzun.csv", index=False, encoding="utf-8-sig")

    def cell(o, k, l):
        x = s[(s.olcu == o) & (s.cikti == k) & (s.GALI == str(l))]
        if not len(x):
            return ""
        x = x.iloc[0]
        return f"{100 * x.tahmin:.1f} [{100 * x.ga_alt:.1f}; {100 * x.ga_ust:.1f}]"

    def pv(o, k, l):
        x = s[(s.olcu == o) & (s.cikti == k) & (s.GALI == str(l))]
        return round(float(x.iloc[0].p), 4) if len(x) else np.nan

    rep.p("Tahminci: ağırlıklı doğrusal olasılık modeli. 'havuz' = aynı örneklemde hane etkisi olmadan "
          "(yaş spline, cinsiyet, eğitim, bölge, kır-kent); 'hane_FE' = hane sabit etkileri (hane içi merkezleme; "
          "bölge ve kır-kent hane içinde sabit olduğundan düşer). Yalnızca ≥2 tanımlı yetişkini olan haneler; "
          "kısıtlılık katsayısı yalnızca GALI'si hane içinde değişen hanelerden öğrenilir. Varyans: Rao–Wu bootstrap.")
    rep.h("Örneklem")
    rep.table(pd.DataFrame(nrows), "orneklem", digits=0)
    rep.h("Havuzlanmış ve hane içi açık (yüzde puan)")
    rows = []
    for k, sp in specs.items():
        for l in levels:
            rows.append(dict(cikti=k, etiket=sp["lab"], GALI=l, havuz=cell("havuz", k, l),
                             hane_FE=cell("hane_FE", k, l), p_FE=pv("hane_FE", k, l),
                             FE_eksi_havuz=cell("fark_FE_havuz", k, l), p_fark=pv("fark_FE_havuz", k, l)))
    rep.table(pd.DataFrame(rows), "havuz_FE")
    rep.p("Yorum: hane_FE ≈ havuz → açık hane düzeyi etkenlerden (bağlantı, gelir, bölge) kaynaklanmıyor, "
          "bireysel. |hane_FE| < |havuz| → açığın bir kısmı haneler arası farklardan geliyor. "
          "Doğrusal model %100'e yakın yaygınlıkta (internet) açığı abartabilir (aşama 10); "
          "burada amaç aynı modelde havuz–FE karşılaştırmasıdır.")
    rep.h("Hane içi açık yaş katmanında (yüzde puan)")
    rows = []
    for k, sp in specs.items():
        for l in levels:
            rows.append(dict(cikti=k, GALI=l, FE_16_54=cell("FE_16-54", k, l), FE_55_74=cell("FE_55-74", k, l),
                             fark=cell("FE_yas_fark", k, l), p_fark=pv("FE_yas_fark", k, l)))
    rep.table(pd.DataFrame(rows), "FE_yas")
    bad = s[s.gecersiz_rep > 0]
    if len(bad):
        rep.p(f"**Uyarı:** {len(bad)} tahminde geçersiz replikasyon var.")
    rep.checkpoint([
        "Bilgi taşıyan (GALI'si hane içinde değişen) hane ve kısıtlı fert sayısı yeterli mi? (ağır kısıtlı için ≥300 hedef)",
        "İnternet erişim açığı hane içinde de sürüyor mu? (Sürüyorsa engel bireysel: aynı bağlantılı evde çevrimdışı.)",
        "Kullanıcılarda bilgi aramadaki pozitif ilişki hane içinde de görülüyor mu?",
        "55–74'teki büyük erişim açığı hane içinde de var mı?",
    ])
    rep.write()


# =============================================================================
# S15_YAYILIM — household need spillover
# =============================================================================
# Question: do internet users WITHOUT limitation who live with a limited adult
# seek more health information / book more appointments? (Use on behalf of family.)
# Exposure: highest GALI level among the OTHER household members aged 16–74 (0/1/2).
# Negative control: generative AI — weakly related to need; no association expected.
# Adjustment: 'toplam' covariates + household size (living with a limited adult relates to household size).
# Limitation: limitation is not measured for members aged 75+ or <16 → exposure misclassified (bias towards zero).

YAYILIM_CIKTI = ["BILGI_F", "BILGI_R", "MHRS", "KAYIT", "KONSULT", "YZ"]


def asama_15(ctx: Ctx):
    ctx.require("gali_merge_levels")
    ctx.log_decisions("15_yayilim")
    cfg = ctx.cfg["models"]
    rep = Report(ctx, "15_yayilim", "Aşama 15 — Hane içi ihtiyaç yayılımı (kısıtlı olmayan kullanıcılar)")
    A_all = load(ctx.cache("analysis.pkl"))
    R_all = load(ctx.cache(f"rep_B{ctx.B}.pkl"))
    m = (A_all.YEAR == 2026).to_numpy()
    A = A_all[m].reset_index(drop=True)
    R = R_all[m]
    W = A["W"].to_numpy(float)
    hid = A["BULTEN_NO"].astype(str)
    g = A["GALI3"].fillna(-1)
    # highest GALI of the other members: household maximum excluding the respondent
    def _diger_max(x):
        v = x.to_numpy()
        out = np.full(len(v), -1.0)
        for i in range(len(v)):
            o = np.delete(v, i)
            out[i] = o.max() if len(o) else -1.0
        return pd.Series(out, index=x.index)
    A["DIGER_GALI"] = g.groupby(hid).transform(_diger_max)
    A.loc[A["DIGER_GALI"] < 0, "DIGER_GALI"] = 0.0      # single-adult household or missing → no limited co-resident
    A["TEK_YETISKIN"] = (hid.map(hid.value_counts()) == 1).astype(float)
    drop = set(ctx.dec.get("drop_outcomes") or [])
    outs = [k for k in YAYILIM_CIKTI if k not in drop and A[k].notna().sum() > 0]
    sub = ((A["GALI3"] == 0) & (A["INTERNET"] == 1) & (A["TEK_YETISKIN"] == 0)).to_numpy()
    rhs = cfg["total_rhs"] + " + C(HHSIZE)"
    lv = [1.0, 2.0]
    fns, atlanan = {}, []
    for k in outs:
        try:
            md = Logit(A, f"{k} ~ C(DIGER_GALI) + {rhs}", subset=sub)
            md.fit(W)
            fns[k] = ame_fn(md, "DIGER_GALI", lv, 0.0)
        except Exception as e:
            atlanan.append(f"{k}: {e}")

    def fn(w):
        out = {}
        for k, f in fns.items():
            for l, v in f(w).items():
                out[f"{k}|{l}"] = v
        return pd.Series(out)

    point, reps = replicate(fn, W, R, progress="yayılım")
    s = summarize(point, reps)
    idx = s.index.str.split("|", expand=True)
    s["cikti"], s["DIGER_GALI"] = idx.get_level_values(0), idx.get_level_values(1)
    s = s.reset_index(drop=True)

    nt = A[sub].groupby("DIGER_GALI").size().rename("n").reset_index()
    rep.p("Örneklem: kısıtlılığı olmayan, son 3 ayda internet kullanan, en az iki 16–74 yaş yetişkinli hanede "
          "yaşayan fertler. Maruziyet: hanedeki diğer yetişkinlerin en yüksek GALI düzeyi. "
          "Ayar: yaş spline, cinsiyet, eğitim, kır-kent, bölge, hane büyüklüğü.")
    rep.h("Örneklem (ağırlıksız n)")
    rep.table(nt, "n", digits=0)
    rows = []
    for k in outs:
        for l in lv:
            x = s[(s.cikti == k) & (s.DIGER_GALI == str(l))]
            if len(x):
                x = x.iloc[0]
                rows.append(dict(cikti=k, etiket=OUTCOME_LABELS.get(k, k),
                                 hanede=("orta kısıtlı" if l == 1.0 else "ağır kısıtlı"),
                                 AME_pp=f"{100 * x.tahmin:.1f} [{100 * x.ga_alt:.1f}; {100 * x.ga_ust:.1f}]",
                                 p=round(float(x.p), 4)))
    rep.h("Hanede kısıtlı biriyle yaşamanın AME'si (yüzde puan)")
    rep.table(pd.DataFrame(rows), "yayilim")
    rep.p("Beklenti: sağlık bilgisi ve randevuda pozitif (aile adına kullanım); YZ'de ~0 (negatif kontrol). "
          "YZ'de de anlamlı pozitif ilişki varsa bulgu hane düzeyi karıştırıcılığa işaret eder ve "
          "yayılım yorumu zayıflar.")
    if atlanan:
        rep.p("**Atlanan:** " + "; ".join(atlanan))
    rep.checkpoint([
        "Sağlık bilgisi / randevu için pozitif yayılım var mı?",
        "Negatif kontrol (YZ) sıfıra yakın mı?",
    ])
    rep.write()


# =============================================================================
# RUN
# =============================================================================
STAGES = {"00": ("00_load", asama_00), "01": ("01_audit", asama_01), "02": ("02_derive", asama_02),
          "03": ("03_design", asama_03), "04": ("04_descriptives", asama_04),
          "05": ("05_models", asama_05), "06": ("06_hurdle", asama_06),
          "07": ("07_khb", asama_07), "08": ("08_trend", asama_08),
          "09": ("09_ham_ayarli", asama_09), "10": ("10_saglamlik", asama_10),
          "12": ("12_goreli_yas", asama_12),
          "13": ("13_hane", asama_13),
          "14": ("14_hane_ici", asama_14),
          "15": ("15_yayilim", asama_15), "11": ("11_manuscript", asama_11)}


def main():
    secilen = list(STAGES) if ASAMALAR == ["hepsi"] else list(ASAMALAR)
    ctx0 = Ctx(synthetic=SENTETIK, draft=TASLAK)
    out_dir = ctx0._abs(ctx0.paths["output_dir"])
    if "00" not in secilen and not ctx0.cache("raw_2026.pkl").exists():
        print("Önbellekte veri yok; önce 00 çalıştırılıyor.")
        secilen = ["00"] + secilen
    birlesik, durdu = [], None
    for st in secilen:
        if st not in STAGES:
            sys.exit(f"Bilinmeyen aşama: {st}. Geçerli: {', '.join(STAGES)} veya 'hepsi'")
        klasor, fn = STAGES[st]
        ctx = Ctx(synthetic=SENTETIK, draft=TASLAK)
        print(f"\n[{st}] {klasor} başlıyor (B={ctx.B})")
        t0 = time.time()
        try:
            fn(ctx)
        except DecisionMissing as e:
            durdu = f"**[{st}] DURDU** — {e}"
            print(f"\n[{st}] DURDU — {e}")
            break
        print(f"[{st}] bitti ({time.time() - t0:.0f} sn)")
        birlesik.append((out_dir / klasor / "report.md").read_text(encoding="utf-8"))
    son = out_dir / "SON_RAPOR.md"
    metin = (f"# SON RAPOR — {_dt.datetime.now():%Y-%m-%d %H:%M}\n\n"
             f"Çalışan aşamalar: {', '.join(secilen)} · B={ctx0.B}\n\n"
             + (durdu + "\n\n" if durdu else "") + "\n\n---\n\n".join(birlesik))
    son.write_text(metin, encoding="utf-8")
    print(f"\n==> Claude'a gönderilecek dosya: {son}")


if __name__ == "__main__":
    main()
