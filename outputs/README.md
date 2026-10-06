# ValueBench exploratory data analysis outputs

This folder contains the reproducible Python analysis, nine report-ready figures, and supporting tables for the ValueBench clinical-vignette corpus.

## Run the analysis

```bash
python src/valuebench_eda.py --input /path/to/valuebench_eda.xlsx --output-dir outputs
```

The script does not edit the Google Sheet. It reads an XLSX export, preserves the three-tab data lineage, and writes derived files locally.

## Primary analytic choices

- All 286 cases describe the curation pipeline and candidate corpus.
- The 51 approved cases are treated as the currently training-ready cohort.
- `Original Cases` is used only for revision auditing.
- The mech-interp tab is a repeated extract, not an independent dataset.
- `promotes`, `neutral`, and `violates` are ordered only to determine which choice each principle favors. The four principles are never summed into a single ethical score.
- Reviewer-comment themes and vignette context flags are transparent, non-exclusive, rule-based exploratory codes. Their row-level coding is exported for manual review.
- TF-IDF similarity flags possible duplicates or paraphrases; it does not establish semantic equivalence.

## Output structure

- `figures/`: high-resolution PNG and vector PDF versions of Figures 1–9, plus a contact sheet.
- `tables/`: CSVs supporting every figure and audit.
- `derived_case_level.csv`: one row per current case with derived variables and exploratory flags.
- `analysis_summary.json`: headline counts used to draft the report.
- `key_tables.md`: compact rendering of the most important tables.

## Source checks

- No structural warnings were raised.

Source: https://docs.google.com/spreadsheets/d/1-Zhf4Nr59UqoeJLqhFwLoGDRH8uAuMGKdNqnigVO01Y/edit?gid=220762358
