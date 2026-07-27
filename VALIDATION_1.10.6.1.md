# Validation 1.10.6.1

Проверяется:

- чистая установка `requirements.txt`;
- `pip check`;
- импорт `streamlit_app.py` и складского модуля;
- отсутствие конфликтующей пары Streamlit 1.60 / PyArrow 25;
- отсутствие `default=` у `supplier_order_mode`, значение которого управляется Session State;
- компиляция и полный набор тестов.
