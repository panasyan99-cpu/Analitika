# Установка Analitika Web 1.10.3

Версия устанавливается поверх 1.10.2 полной заменой файлов из архива. Она уже содержит исправление аналитики White/Colored из предыдущего патча.

Перед заменой файлов:

```bash
git add -A
git commit -m "Analitika Web 1.10.2 - Fix white and colored pearl order analytics"
```

Скопируйте содержимое папки `Analitika_Web_1.10.3_ORDER_VISIBILITY` в корень репозитория, затем выполните:

```bash
pip install -r requirements.txt
pytest -q
streamlit run streamlit_app.py
```

Коммит версии 1.10.3:

```bash
git add -A
git commit -m "Analitika Web 1.10.3 - Exclude round pearls and group unknown stones"
```
