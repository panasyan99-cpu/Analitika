# Установка Analitika Web 1.10.2

Версия устанавливается поверх текущего репозитория 1.10.1 полной заменой файлов из архива.

```bash
git add -A
git commit -m "Backup before Analitika Web 1.10.2"
```

Скопируйте содержимое папки `Analitika_Web_1.10.2_PEARL_ANALYTICS` в корень репозитория, затем выполните:

```bash
pip install -r requirements.txt
pytest -q
streamlit run streamlit_app.py
```

Рекомендуемое название коммита:

```bash
git add -A
git commit -m "Analitika Web 1.10.2 - Fix white and colored pearl order analytics"
```
