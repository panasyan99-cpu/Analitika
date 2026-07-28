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
