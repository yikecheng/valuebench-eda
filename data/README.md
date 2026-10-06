# Source data

The raw workbook is available from the [ValueBench Google Sheet](https://docs.google.com/spreadsheets/d/1-Zhf4Nr59UqoeJLqhFwLoGDRH8uAuMGKdNqnigVO01Y/edit?gid=220762358).

The analysis expects an XLSX workbook containing these tabs:

- `Cases`
- `Original Cases`
- `Downloaded for mech interp`

The script downloads the workbook export when `--input` is omitted. To use a local copy, download the Google Sheet as Microsoft Excel format and pass its path with `--input`.

Raw XLSX files are excluded from version control. Derived case level data and aggregate output tables are available in `outputs/`.

The public derived case level file omits reviewer account identifiers while retaining review decisions and comments used by the analysis.
