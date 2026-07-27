# Validation Analitika Web 1.11.1

## Automated checks

- `python -m pytest -q`: **225 passed, 3 skipped**.
- `python -m compileall -q streamlit_app.py src`: passed.
- Static version check: `1.11.1` in `version.json` and fallback metadata.
- Dependency pins preserved: Streamlit `1.60.0`, PyArrow `24.0.0`.
- Upload limit preserved: `150 MB` per file.

The three skipped tests are real Streamlit AppTest cases. Streamlit is not available in the local validation runtime; GitHub Actions installs the production dependencies and runs the complete suite.

## UX checks

- General sales analysis does not call the shared sidebar or mobile anchor navigation.
- Period comparison does not call the shared sidebar or mobile anchor navigation.
- Both modules retain inline actions for replacing uploaded reports.
- Metal/purity and FX controls are rendered once inside a collapsed `Настройки отчёта` expander.
- All five operational modules expose `Работа / Как с этим работать` tabs.
- Contextual help reads the same Markdown chapter used by the full user guide.
- Visible success banners repeating the active metal filter were removed.
- Sonu no longer repeats the FX rate in the persistent report caption.

## Documentation checks

- USER_GUIDE.md contains detailed chapters for sales analysis, comparison, supplier order, Sonu and Baserow.
- Release notes, deployment instructions, changelog and README are updated for 1.11.1.

## Business logic scope

No recommendation, parsing, classification, stock, order quantity or export formula was intentionally changed in this release.
