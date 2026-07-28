# Private secret status

This repository is intended for the private Analitika deployment.

## Embedded

- Baserow URL;
- Baserow database token;
- database ID 148;
- table IDs 642–645.

The warehouse module can therefore connect without adding a new Baserow block
in Streamlit Cloud Secrets.

## Preserved in existing Streamlit Cloud App Settings

- `[auth] password`;
- `[order_storage] endpoint_url`;
- `[order_storage] access_key_id`;
- `[order_storage] secret_access_key`;
- existing order-storage bucket configuration.

The uploaded source archive contained only placeholders for those values, not
the actual credentials. The deployment code continues reading them exactly as
before, so updating the repository in the same Streamlit application does not
require replacing the current Cloud Secrets.

Do not create a new Streamlit application without first copying the existing
`[auth]` and `[order_storage]` sections from the current app settings.

## Baserow schema migration in 2.4.0

The existing database token is sufficient for row operations but cannot create tables or fields.
The maintenance page requests a Baserow Builder/Admin email and password once, obtains a short-lived JWT, creates/migrates the schema, and does not persist the password.
The resulting table ID is auto-discovered at runtime; it can then be copied into `supply_lines_table_id` in Streamlit Secrets.
