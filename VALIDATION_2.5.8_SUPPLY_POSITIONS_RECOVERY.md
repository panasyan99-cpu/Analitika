# Validation — 2.5.8 Supply Positions Recovery

## Fixed

- The warehouse no longer depends solely on Baserow table-list metadata to find `Позиции поставок`.
- Known/configured table IDs are probed directly; production table 646 is used only after successful validation.
- If the table is absent and Builder/Admin credentials are available, it is created automatically.
- Existing tables are repaired idempotently once per authenticated session.
- Missing detail rows are migrated from both souvenir and component catalog links.
- Silver 925 supplies are included in migration and operation reconciliation.
- Supply cards and detail screens retain a legacy compatibility fallback.

## Verification

- Python compilation: passed.
- Targeted supply-position recovery tests: passed.
- Complete regression suite: **313 passed, 5 skipped, 0 failed**.
- Hardcoded production credentials: not present.
- Runtime caches and compiled Python files: excluded from delivery.

## Deployment note

Keep the existing Streamlit Secrets. The application will resolve or repair `Позиции поставок` automatically after the first authenticated warehouse opening.
