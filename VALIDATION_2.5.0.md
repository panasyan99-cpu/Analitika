# Validation 2.5.0

## Automated checks

```text
290 passed
3 skipped
0 failed
```

The three skipped checks use Streamlit AppTest, which is unavailable in the lightweight validation environment.

## Actual invoice verification

The uploaded source file `Серебро INVOICE18.07.2026г.xlsx` was parsed directly without conversion.

- 18 product lines detected;
- 26 478 stock units detected;
- 5 173 puset pairs detected on line 2;
- all products classified as Silver 925 components;
- line 11 marked as a separately sellable chain;
- 18 product photo placements extracted and optimized;
- fixed purchase USD and imported VND sale values matched the workbook formulas;
- puset sale check at USD/VND 26 500 and coefficient 10: 642 000 VND;
- generated Master export opened successfully with 19 rows and 39 columns.

## Verified application behavior

- warehouse no longer receives the global purity filter or standard report FX block;
- warehouse shows independent USD/VND and coefficient controls with defaults 26 500 and 10;
- current sale VND is recalculated from fixed purchase USD and rounded up to 1 000 VND;
- historical supply-line prices are retained separately from the current display price;
- first silver import adds required additive Baserow fields automatically;
- an expected supply can be created with zero received quantity;
- receiving later is limited by the still-waiting quantity;
- non-silver warehouse imports retain their previous behavior.
