# Validation 2.5.12

## Automated checks

- Python compilation: passed.
- Pytest: **344 passed, 5 skipped, 0 failed**.
- Management report navigation and exact full-report path coexist in one build.
- Warehouse/Baserow source files were not changed by the merge.
- Supplier-order and Sonu business modules were not changed by the merge.

## Real 1C files

### fulljune(2).xlsx

- Management parser: quantity `16641.01`, revenue `26626829577 VND`, returns `34.0` / `231141000 VND`.
- Exact full parser: `5451` SKU facts, quantity `16641.0`, revenue `26626829577 VND`, returns `34.0` / `231141000 VND`.
- Dimensions: sellers `40`, categories `5`, subgroups `14`, nomenclature groups `102`, stones `18`, purities `5`.
- Detected as full sales report: `True`. Detected as supplier summary: `False`.

### fulljuly(1).xlsx

- Management parser: quantity `19801.451`, revenue `31918314500 VND`, returns `41.0` / `232748000 VND`.
- Exact full parser: `6542` SKU facts, quantity `19801.5`, revenue `31918314500 VND`, returns `41.0` / `232748000 VND`.
- Dimensions: sellers `43`, categories `5`, subgroups `14`, nomenclature groups `104`, stones `18`, purities `5`.
- Detected as full sales report: `True`. Detected as supplier summary: `False`.

## Store 63 split

- 63 Retail: period 1 `118103000 VND`; period 2 `425408000 VND`.
- 63 Timing: period 1 `301095000 VND`; period 2 `400984000 VND`.

## Runtime note

The Streamlit server executable is not installed in the packaging container, so an HTTP browser-start smoke test could not be run here. Import/syntax, unit, regression, real-file parser and navigation-source checks passed.
