# Analitika Web 2.6.0 — complete merged release

## What was repaired

- restored all files from the complete 2.5.8 package, including the online warehouse, silver workflow, weight-based receiving, migrations and warehouse regression tests;
- retained the newer private authentication and operator integration from 2.5.8;
- merged the 2.0 Analytics & Stock Workspaces instead of replacing the repository with the incomplete archive;
- added parsing for the current 1C hierarchy: Store → Stone/insert → Purity → Product group → Supplier;
- added end-of-period stock analysis and stock comparison without duplicate hierarchy subtotals;
- added compact workspaces for summary, stores, assortment, stock and suppliers;
- added 1C setup screenshots and user-facing “Model” terminology;
- retained all historical release notes and validation documents.

## Source packages

- `Analitika_Web_2.5.8_Weight_Based_Receiving.zip` — complete base;
- `Analitika_Web_2.0_ANALYTICS_STOCK_WORKSPACES.zip` — analytics changes selectively merged.
