# Validation 2.5.7

- Unit regression: removing an unreceived line deletes only a zero-stock/no-history orphan card.
- Unit regression: reconciliation ignores unfinished operations and respects posted corrections.
- Unit regression: receipt, transfer, line status and supply status are rebuilt consistently.
- Unit regression: stale product links are replaced by current supply-line links.
- Unit regression: historical orphan cards are deactivated rather than deleted.
- Full pytest suite and Python compilation are required before packaging.
