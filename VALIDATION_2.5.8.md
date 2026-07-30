# Validation 2.5.8

## Scope

Weight-based receiving for a supply that was physically delivered earlier and is already being used.

## Verified behavior

- Standard quantity receiving is unchanged.
- Weight mode is available only before the first receipt of a supply line.
- Clean measured weight is divided by the invoice average unit weight.
- Quantity uses decimal half-up rounding and cannot exceed the document quantity.
- The operator can override the calculated quantity before posting.
- A zero measured weight is valid and produces a zero current remainder.
- Selected lines are posted as a full document receipt.
- The difference between document quantity and current remainder is posted as a separate expense with the reason «Использовано до постановки на учёт».
- Lines without entered weight remain waiting and can be processed later.
- Receiving method, measured weight, calculated quantity and weight error are stored in Baserow supply-line fields.
- The weight operation is protected by an explicit confirmation checkbox and Command ID duplicate protection.

## Real invoice regression

The raw silver invoice dated 30.06.2026 was parsed directly:

- 14 positions recognized;
- all 14 positions contain a valid total weight and document quantity;
- average unit weight is available for every line, including the 1000-meter chain line;
- using the original full batch weight reconstructs the original document quantity for all 14 lines;
- automatic receipt remains disabled at import.

## Automated tests

- `307 passed`
- `3 skipped` Streamlit AppTest cases because AppTest is unavailable in the validation environment.
- Python compilation passed for the modified warehouse UI, service and schema modules.

## Limitation

No write was made to the live Baserow database during validation. The schema creation and posting workflow were verified with regression fakes and existing Baserow client contracts.
