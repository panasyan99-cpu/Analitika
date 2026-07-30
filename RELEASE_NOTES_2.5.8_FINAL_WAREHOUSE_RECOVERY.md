# Analitika Web 2.5.8 — Supply positions recovery

- Resolves `Позиции поставок` by direct table probe before metadata discovery, so restricted Baserow tokens work.
- Safely creates the table only when it is genuinely absent and operator credentials are available.
- Runs the schema repair once per authenticated session, including already existing tables.
- Restores missing supply-position rows from both `Сувенирка` and `Комплектующие`.
- Includes Silver 925 components in migration and operation reconciliation.
- Keeps legacy supplies visible while recovery is unavailable.
- Does not touch supplier-order/cloud-order storage.
