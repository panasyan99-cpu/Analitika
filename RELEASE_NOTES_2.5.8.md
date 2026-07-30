# Analitika Web 2.5.8

## Weight-based receiving for an already used supply

- Added **«По весу — товар уже в работе»** inside **Приёмка → По поставке**.
- The operator enters clean remaining weight without packaging; the site uses the average unit weight saved from the invoice and applies conventional half-up rounding.
- The calculated quantity is shown before posting and can be corrected manually.
- For every selected line the full invoice quantity is posted as receipt, while the difference is posted separately as **«Использовано до постановки на учёт»**. This preserves the real supplier delivery and restores the current live stock.
- Weight mode is allowed only before the first receipt of a supply line and only when the invoice contains a valid average unit weight.
- Zero weight is supported and means the current remainder is zero. Lines without a weight remain waiting and can be processed later.
- Added explicit confirmation before posting and saved the receiving method, entered weight, calculated quantity and rounding error in Baserow supply-line fields.
- Ordinary piece-by-piece receiving remains unchanged and is marked as **«По количеству»**.
