# Analitika Web 2.5.7

## Baserow reconciliation hotfix

- Removing a never-received position from a supply now also removes its stale supply link from the catalog.
- A zero-stock catalog card is deleted when it has no remaining supply line and no posted operation history.
- Historical zero-stock cards are preserved for audit, unlinked from deleted supplies and deactivated.
- Added the **Актуализировать Baserow** button. It recalculates receipt and transfer counters from posted operations, repairs supply and line statuses, rebuilds product-to-supply links and cleans safe orphan cards.
- Reconciliation never creates a receipt, never changes document quantities and never deletes cards with operation history.
- Baserow row caches are invalidated immediately after create, update and delete operations.
