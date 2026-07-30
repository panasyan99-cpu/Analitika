# Validation 2.5.2

## Automated checks

```text
294 passed
3 skipped
0 failed
```

The skipped checks are optional Streamlit AppTest scenarios that are disabled by their existing environment guards.

## Regression checks

- Silver 925 product payloads do not send the legacy Baserow select field `Категория`.
- Silver classification is retained in `Серебряная категория`.
- Regular non-silver components still send their ordinary category unchanged.
- Numeric precision normalization from 2.5.1 remains covered.
- ISO date parsing no longer emits the pandas `dayfirst=True` warning.
- Warehouse segmented controls no longer combine an explicit default with a pre-populated Session State key.
