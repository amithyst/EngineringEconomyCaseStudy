# `src`

Repository scripts are organized by responsibility.

- `analysis/`
  Core model logic, assumptions, and reusable cash flow calculations.
- `excel/`
  Builders for the submission workbook and any spreadsheet-specific exports.
- `reporting/`
  Figure generation and report assembly helpers.

Compatibility entrypoints remain at the root of `src/` so existing commands still work:

- `calculate_case_study.py`
- `generate_report_figures.py`
- `build_submission.py`
