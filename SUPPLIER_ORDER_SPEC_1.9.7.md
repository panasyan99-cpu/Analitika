# Supplier Order Specification 1.9.7

## SKU clipboard in ring-size stage

1. Every ordered ring card displays the textual SKU and a compact copy button in one horizontal row.
2. Clipboard payload is `str(value or "").strip()`.
3. No leading/trailing spaces, tabs or line breaks may be copied.
4. Visible SKU text is HTML-escaped. Clipboard payload is serialized through JSON before insertion into JavaScript.
5. Successful copy shows a temporary `✓`; failure shows `!`.
6. Clipboard API is preferred; a textarea/`execCommand("copy")` fallback is retained.
7. The change is UI-only and must not modify draft persistence, quantities, ring-size validation or export data.
