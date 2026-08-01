# Validation — Analitika Web 2.5.13

## Automated tests

- **348 passed**
- **5 skipped**
- **0 failed**
- **127 Python files** compiled successfully with `py_compile`

The build environment did not include the Streamlit runtime package, so a live HTTP server launch was not performed. The repository contains the pinned Streamlit dependency in `requirements.txt`; source, parser, UI-contract and regression tests passed.

## Real 1C files used

- `ИЮНЬПРОД.xlsx`
- `ИЮНЬКОНС.xlsx`
- `ИЮНЬПОСТ(1).xlsx`
- `ИЮЛЬПРОД.xlsx`
- `ИЮЛЬКОНС.xlsx`
- `ИЮЛЬПОСТ(1).xlsx`

## Cross-block control

### June 2026

All three blocks produced the same explicit total:

- quantity: **16 641.01**
- revenue: **26 626 829 577 VND**
- returns: **34 units / 231 141 000 VND**

Primary dimensions:

- stores: 9 rows, revenue difference to total **0 VND**, grouped quantity difference **+0.99**
- consultants: 40 rows, revenue difference **0 VND**, grouped quantity difference **+0.99**
- suppliers: 12 rows, revenue difference **0 VND**, grouped quantity difference **+0.99**

### July 2026

All three blocks produced the same explicit total:

- quantity: **20 315.451**
- revenue: **32 872 187 500 VND**
- returns: **43 units / 235 554 000 VND**

Primary dimensions:

- stores: 10 rows, revenue difference to total **0 VND**, grouped quantity difference **+0.549**
- consultants: 43 rows, revenue difference **0 VND**, grouped quantity difference **-0.451**
- suppliers: 11 rows, revenue difference **0 VND**, grouped quantity difference **-0.451**

The sub-unit quantity differences are caused by grouped 1C rounding of weighted quantities. The exact KPI is always read from the `Итого` row.

## Supplier control

- June Own production: **475 units / 1 366 194 000 VND**
- July Own production: **839 units / 1 702 807 500 VND**
- supplier revenue sum matches the total row exactly in both periods
- blank supplier groups remain visible as `Не определен`
- `Own production service` aliases normalize to `Own production`

## Consultant control

- all level-0 manager rows are retained
- `Admin` is not removed from the source total
- blank manager is shown as `Менеджер не указан`
- consultant revenue and returns sum to the explicit period total

## Scope control

Compared with the user-provided 2.5.12 repository, functional changes are limited to:

- management-report parser and UI
- management-report tests
- version and release documentation
- updated PDF user guide

Warehouse, Baserow, supplier-order and Sonu source modules were not modified.
