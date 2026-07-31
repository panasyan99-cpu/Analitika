# Secret configuration status

The final repository contains no real passwords, account credentials, or Baserow API tokens.

Configure deployment values through `.streamlit/secrets.toml` (excluded by `.gitignore`) or environment variables:

- shared site password;
- Baserow URL and API token;
- database and table identifiers;
- optional server-side Baserow email and password.

Use `.streamlit/secrets.toml.example` as the template. Never commit the real `secrets.toml`.
