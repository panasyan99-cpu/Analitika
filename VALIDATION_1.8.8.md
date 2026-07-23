# Validation 1.8.8

- `python -m py_compile src/order_persistence.py src/order_workflow.py` — passed.
- Полный pytest-набор: `142 passed, 1 skipped`.
- Проверена разреженная сериализация черновика.
- Проверено сохранение workbook + draft + manifest в имитации S3.
- Проверено восстановление workbook в пустой новый runtime.
- Проверено преобразование cloud manifest в карточку «Продолжить сохранённый заказ».
- Проверено, что новый отчёт не удаляет предыдущую восстанавливаемую историю.
- Проверен аварийный ZIP с исходным Excel и JSON черновика.
- Один Streamlit AppTest пропущен, потому что в среде сборки нет установленного runtime Streamlit; остальные тесты используют test-only fallback.
