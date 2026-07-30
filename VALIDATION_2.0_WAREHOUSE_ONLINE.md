# Validation — Analitika 2.0 Princess Warehouse Online

## Source

Integrated into the uploaded `Analitika_Web_2.0_ORDER_LIBRARY_LAYOUT` source.
The public application version remains `2.0`.

## Automated validation

- Python compilation: passed;
- full repository tests: **261 passed, 3 skipped**;
- existing analytics, comparison, SONU and supplier-order tests: passed;
- warehouse service tests: passed;
- duplicate document and maximum quantity rules: covered;
- reverse operation logic: covered;
- empty-box Master parser regression: covered.

## Real Master regression

Tested with `Master_New_Souvenirs_2026-07-17.xlsx`:

- 125 SKU parsed;
- 1,341 units parsed;
- 125 embedded photographs extracted.

## Load protection

- only the active warehouse subsection renders;
- read cache remains 60 seconds;
- cache clears after writes;
- file upload limit remains 150 MB;
- uploaded Excel and extracted images are stored in a temporary runtime directory;
- parsed files are removed after a successful supply creation;
- existing application modes are not imported or recalculated by warehouse actions.

## Deployment limitation

The actual `[auth]` and `[order_storage]` secret values were not present in the
uploaded archive. They remain configured in the existing Streamlit Cloud app.
The Baserow warehouse access is embedded in the private repository.
