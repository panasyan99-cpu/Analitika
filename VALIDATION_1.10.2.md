# Validation 1.10.2

## Scope

Regression patch for completed-order pearl analytics.

## Verified behavior

- White freshwater pearl and round white freshwater pearl resolve to `White Freshwater`.
- Pink / Rose, Grey / Gray and Black freshwater pearl resolve to `Colored Freshwater`.
- Shape markers such as `Round` no longer override an explicit colour.
- The obsolete analytics family `Round White Freshwater` is not emitted.
- Supplier-order recommendations, drafts and Excel export remain unchanged.

## Automated validation

```text
pytest -q
182 passed, 1 skipped
```

Targeted analytics regression tests:

```text
pytest -q test_supplier_order_analytics_1100.py
5 passed
```
