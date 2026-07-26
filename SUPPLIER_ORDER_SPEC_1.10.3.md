# Supplier Order Specification 1.10.3

## Hard exclusions

A pearl row is excluded when its material description or explicit SKU text contains both `PEARL` and `ROUND`. The exclusion applies before order sets, recommendations, summary, completed-order analytics and Excel export are built.

Bracelets, configured ready-item stone exclusions and confirmed duplicate SKUs continue to follow their existing rules.

## Other Stones navigation

The following remain orderable and are never discarded merely because classification is incomplete:

- MOP and other recognized rare materials;
- MOR and other unregistered abbreviations;
- AMA / unrecognized material;
- unknown or empty source values.

Only established top stones remain standalone filters. Green stones and Other Topaz keep their grouped filters. All remaining stone values are displayed under `Other Stones`, while cards and supplier Excel retain the concrete canonical value.
