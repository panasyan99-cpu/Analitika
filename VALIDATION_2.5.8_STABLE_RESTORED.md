# Validation — Analitika Web 2.5.8 stable restored

## Scope

The broad 2.6.0 changes were fully rolled back. This build is based on the complete working 2.5.8 warehouse-recovery release.

Only the following deployment-specific change was restored:

- the private embedded Baserow database token;
- the private embedded Baserow operator login and password used by the previously working site.

## Confirmed unchanged from the stable recovery build

- supplier-order workflow;
- Sonu workflow;
- main Streamlit navigation and workspaces;
- general analytics;
- comparison analytics;
- warehouse receiving and supply-position recovery.

## Removed

- the 2.6.0 diagnostics workspace;
- the 2.6.0 global Baserow permission gate;
- all 2.6.0 supplier-order persistence changes;
- all 2.6.0 Sonu storage changes;
- all 2.6.0 broad UI and configuration refactors.

## Automated validation

- Pytest: **313 passed, 5 skipped, 0 failed**.
- Python compile check: passed.
- Embedded Baserow operator credentials: present.
- Embedded Baserow database token: present.
- Diagnostics workspace strings: absent.
- ZIP integrity: checked after packaging.

## Live connection note

The current build environment could not resolve the private Baserow hostname, so a live network login was not possible here. The exact private connection values from the previously working 2.5.8 build were restored without displaying them in reports.
