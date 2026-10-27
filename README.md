# Activity limitation and digital health in Türkiye — analysis code

Reproducible analysis code for the article:

> **Who Is Left Offline When Health Care Goes Digital? Individual and Household Digital Resources Among People with Activity Limitations in a Nationally Representative Survey from Türkiye.** Güvey, E. *et al.* (submitted to *Healthcare*, MDPI).

The code estimates crude and adjusted gaps in internet use and online health services by activity limitation (GALI), within-household (household fixed-effects) gaps, household-inclusive ("relational") access and digital isolation, using design-based Rao–Wu bootstrap inference. It produces all tables, figures and supplementary files of the article.

## Data

The analysis uses the anonymised microdata of the **TurkStat Survey on ICT Usage in Households and by Individuals, 2025 and 2026**. The microdata are **not included** in this repository and may not be redistributed. They can be requested from the Turkish Statistical Institute (TurkStat) microdata service. After obtaining access, place the four files in `data/`:

```
data/hbtka_2025_fert.xlsx   data/hbtka_2025_hh.xlsx
data/hbtka_2026_fert.xlsx   data/hbtka_2026_hh.xlsx
```

A **synthetic dataset** with the same variable names and coding can be generated to test the pipeline without the microdata (numbers are meaningless):

```bash
python tests/make_synthetic.py
```

## Running

Requires Python ≥ 3.10 (`pip install -r requirements.txt`). All settings are at the top of `hhbta_analiz.py`:

| Setting | Meaning |
|---|---|
| `ASAMALAR` | Stages to run, e.g. `["hepsi"]` (all) or `["05", "06"]` |
| `TASLAK` | `True` = 100 bootstrap replicates (quick test); `False` = 500 (final) |
| `VERI` | Paths to the four microdata files |
| `KARARLAR` | Analytic decisions taken at the documented checkpoints (see Table S13) |
| `SENTETIK` | `True` = run on the synthetic data |

```bash
python hhbta_analiz.py
```

Each stage writes a report (`outputs/<stage>/report.md`) and CSV files; a combined `outputs/SON_RAPOR.md` is written at the end. The full run with 500 replicates takes roughly 30–60 minutes on a laptop.

## Stages

| Stage | Content |
|---|---|
| 00 | Load Excel files |
| 01 | Data audit, codebook label check (Table S2) |
| 02 | Derive variables; check against TurkStat headline figures |
| 03 | Survey design, Rao–Wu replicate weights, minimum detectable effects |
| 04 | Descriptive cascade (Figure 2, Table S4) |
| 05 | Average marginal effects, H3 contrasts (Table 5, S5) |
| 06 | Two-part decomposition: access vs conditional use (Table S6, Figure S1) |
| 07 | Karlson–Holm–Breen decomposition (Table S7) |
| 08 | 2025–2026 change, DiNardo–Fortin–Lemieux reweighting, equivalence tests (Table S9, Figure S7) |
| 09 | Crude → adjusted gaps (Table 2, Figure 3) |
| 10 | Sensitivity analyses (Table S8, Figure S2) |
| 12 | Risk ratios and age strata (Tables S10–S11, Figures S3–S4) |
| 13 | Relational access and digital isolation (Table 4, S14; Figures 4–5) |
| 14 | Within-household comparison, household fixed effects (Table 3, S15; Figures 4, S5) |
| 15 | Household spillover with negative control (Table S16, Figure S6) |
| 11 | Manuscript outputs: Word tables, figures (PNG 600 dpi + PDF), supplementary files |

Code comments and docstrings are in English. Some identifiers, intermediate report texts (`report.md`) and CSV column names are in Turkish; all manuscript outputs (stage 11) are in English. Glossary of the main Turkish names:

| Turkish | English |
|---|---|
| `ASAMALAR`, `TASLAK`, `VERI`, `KARARLAR`, `SENTETIK` | stages, draft (quick) mode, data paths, decisions, synthetic data |
| `tahmin`, `se`, `ga_alt`, `ga_ust`, `p` | estimate, standard error, 95% CI lower, 95% CI upper, p value |
| `cikti`, `olcu`, `etiket`, `donem` | outcome, measure, label, period |
| `GALI3` (0/1/2), `GALI_ANY` | not / moderately / severely limited; any limitation |
| `INTERNET`, `BILGI_F`, `BILGI_R` | internet use (3 months); physical / mental health information seeking |
| `MHRS`, `KAYIT`, `KONSULT` | online appointment; online health records; online consultation / e-prescription |
| `YZ`, `AKILLI_SAAT`, `SAGLIK_CIHAZ` | generative AI; smartwatch / wearable; connected health device |
| `REPERTUAR`, `BECERI` | online activity repertoire; digital skills |
| `ham`, `ayarli`, `acik`, `telafi` | crude, adjusted, gap, compensation share |
| `hane`, `hane_ici`, `havuz`, `hane_FE` | household, within household, pooled, household fixed effects |
| `YALITIK`, `ERISIM_ILISKI`, `HH_DIGER_INT` | digitally isolated; relational access; online co-resident |
| `yayilim`, `saglamlik`, `yas` | spillover, sensitivity, age |
| `toplam`, `dogrudan` | total (confounders only) and direct (+ resource mediators) specifications |

## Outputs

`outputs/` contains only aggregated estimates (no record-level data). `cache/` contains record-level intermediate files derived from the microdata and is excluded from version control.

## License and citation

Code: MIT License (see `LICENSE`). Please cite the article and this repository (see `CITATION.cff`).
