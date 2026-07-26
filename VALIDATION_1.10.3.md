# Validation 1.10.3

## Scope

Supplier-order visibility regression patch.

## Verified behavior

- Every pearl material containing an explicit `Round` marker is excluded from both order modes.
- A `Round` marker present in the SKU also blocks a pearl row.
- Non-round white, pink/rose, grey/gray and black freshwater pearls remain available in the pearl order.
- MOP, MOR, AMA, unknown and empty material values use the `Other Stones` navigation bucket.
- Established top stones keep their individual navigation entries.
- The concrete stone value remains supplier-facing; `Other Stones` is navigation-only.

## Automated validation

```text
pytest -q
187 passed, 1 skipped
```

Targeted visibility regression tests:

```text
pytest -q test_supplier_order_visibility_1103.py
5 passed
```
