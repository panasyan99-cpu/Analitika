# Validation 2.4.4

## Automated checks

```text
287 passed
3 skipped
0 failed
```

The three skipped checks use Streamlit AppTest, which is intentionally unavailable in the lightweight validation environment.

## Verified changes

- transfer methods are rendered lazily and `st.tabs` is no longer used for operational modes;
- each editable quantity input receives a row-specific `max_value`;
- supply transfer no longer loads the full operations table when supply lines are active;
- Baserow rows are reused within one render;
- warehouse thumbnails are 320 px and table rows are 138 px;
- catalog and supply detail photos are paginated;
- receiving, manual transfer, supply transfer and Excel transfer use the highlighted quantity editor;
- version history contains 2.1.0 through 2.4.4.
