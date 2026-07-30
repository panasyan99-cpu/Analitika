# Analitika Web 2.5.8 — Warehouse visibility hotfix

- Restored the default Baserow table ID `646` for **Позиции поставок**.
- The supply registry no longer hides Baserow supply headers when detail rows are temporarily unavailable.
- Legacy supply links are now read from both **Сувенирка** and **Комплектующие** (including Silver 925).
- Receiving and supply detail screens fall back safely to legacy product links.
- No supplier-order (`order_storage`) logic is involved in this fix.
