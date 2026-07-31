# Analitika Web 2.5.9 — stable warehouse recovery

This release intentionally rolls back the broad 2.6.0 changes.

Included:
- the complete working 2.5.8 site;
- analytics stock workspaces merged earlier;
- warehouse supply visibility fixes;
- direct discovery of the `Позиции поставок` table;
- recovery of missing supply-position rows for souvenirs, components/casts, and Silver 925;
- automatic Baserow sign-in using the private embedded deployment credentials from the last working build.

Not included:
- the 2.6.0 diagnostics workspace;
- changes to supplier-order persistence;
- Sonu storage changes;
- role/access redesign;
- global UI refactors.
