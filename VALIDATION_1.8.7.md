# Validation 1.8.7

- `python -m compileall -q src streamlit_app.py test_supplier_order_context_186.py` — passed.
- Supplier-order regression scenario: set card rendered with an explicitly supplied workbook `source_hash` — passed.
- `test_supplier_order_context_186.py` — 7 tests passed using a minimal Streamlit import stub because Streamlit is not installed in the build container.
- Full-suite collection was not claimed: the build container does not contain the pinned Streamlit package and has no package-index access.
