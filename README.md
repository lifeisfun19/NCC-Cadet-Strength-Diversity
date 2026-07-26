# NCC Cadet Strength & Gender-Diversity Trend — Automated Reporting Pipeline

A Python + SQL pipeline that reconciles four **real** National Cadet Corps (NCC, Ministry
of Defence) data releases — each with a genuinely different structure — into one
consolidated, validated dataset and an auto-generated Excel report.

## Why this dataset
Unlike a single downloaded CSV, this mirrors an actual data-management problem: four
different government/press releases report NCC statistics in four different shapes
(a two-year wide trend table, a wing-level breakdown, a metric/value/unit table, and an
announcement with no year field at all). Reconciling them — and catching where they don't
quite agree — is the real work.

## Pipeline
1. **Normalize** — reads all 4 source files (`sources/`) and reshapes each into a common
   long format: metric_category, metric_name, year, value, unit, source_file.
2. **Cross-validate** — checks the girls'-count figure implied by (Total Cadets × Girls %)
   against the Senior + Junior Wing figures published separately, and **flags** the small
   discrepancy rather than silently reconciling it.
3. **Load** — writes the reconciled metrics into a local SQLite database.
4. **Analyze (SQL)** — pulls the cadet-strength trend, girls'-share trend, institutional
   coverage, and expansion-plan figures; computes CAGR for cadet-strength growth.
5. **Report** — auto-generates a 3-sheet Excel report: Summary (tables + charts), Sources
   (exact attribution for every figure), and a Data Quality Log of flagged discrepancies.

## Real figures used (see `output/.../Sources` sheet for full citations)
- Enrolled cadets: 14,64,727 (2021) → 15,58,199 (2025); girls' share 35.50% → 40.68%
- ~23,000 institutions currently covered, ~12,000 more in the pipeline
- Sanctioned-strength expansion announced: 17 lakh → 20 lakh, plus 3 lakh additional cadets planned
- 4 new NCC units raised in J&K and Ladakh in the last 5 years

## Run it
```bash
python3 etl_ncc_report.py
```

## Files
- `sources/` — the 4 real source CSVs (each independently shaped, as originally reported)
- `etl_ncc_report.py` — the full normalize + validate + SQL + report pipeline
- `output/` — generated SQLite DB and Excel report
