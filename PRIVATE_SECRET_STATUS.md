# Private secret status

This repository is intended only for the private Analitika deployment.

## Private ZIP

The private release ZIP contains a ready `.streamlit/secrets.toml` with:

- the existing shared site password;
- Baserow URL and database token;
- Baserow database/table identifiers;
- server-side Baserow account credentials.

## Runtime behavior in 2.4.3

- the public login screen asks for one shared password only;
- no Baserow email/password field is rendered;
- opening the warehouse initializes the Baserow client server-side;
- a short-lived JWT is obtained and refreshed automatically;
- the safe `Позиции поставок` schema is discovered or prepared automatically.

## Git bundle

`.streamlit/secrets.toml` is excluded from Git history. The private fallback module remains in this owner-requested private release, so the repository must never be made public.
