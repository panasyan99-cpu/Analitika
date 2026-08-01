# Validation 2.5.14

## Automated checks

- `pytest -q`: **353 passed, 5 skipped, 0 failed**.
- Syntax compilation: **128 Python files**, 0 errors.
- No files from the 2.5.13 base repository were removed.

## Real 1C reports

The three blocks were parsed with deliberately arbitrary source filenames to confirm that the Excel filename is not used for classification. The selected block and the 1C grouping structure are the only inputs used.

### June 2026

All three reports matched:

- quantity: **16 641.01**;
- revenue: **26 626 829 577 VND**;
- returns: **34 units / 231 141 000 VND**;
- Own production: **475 units / 1 366 194 000 VND**.

### July 2026

All three reports matched:

- quantity: **20 315.451**;
- revenue: **32 872 187 500 VND**;
- returns: **43 units / 235 554 000 VND**;
- Own production: **839 units / 1 702 807 500 VND**.

## Presentation rules

- Block order: stores → consultants → suppliers → combined reconciliation.
- Positive deltas use green text/background; negative deltas use red text/background.
- The color follows the mathematical sign, including discount and return differences.
- Displayed USD amounts are rounded to the nearest $10. Raw VND values and calculation inputs remain unchanged.
- USD rounding is applied in the main analytics, management report and Sonu analytical presentation. Supplier-order and warehouse calculations are unchanged.

## Limitations of this validation environment

The Streamlit package is not installed in the build environment, so a live browser/HTTP render was not launched. UI contracts, source syntax, data parsing and the full automated regression suite were verified.
