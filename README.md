# Analitika Web 0.1.0 TEST

Первая тестовая веб-версия аналитики Princess Jewelry.

## Возможности

- загрузка общей Excel-выгрузки;
- определение магазинов и периода по содержимому файла;
- сводка по сети прямо в браузере;
- страницы магазинов и OUTLET;
- диаграммы Top Stones / Pearls / Colored Stones;
- аналитические выводы;
- формирование и скачивание итогового Excel;
- работа через браузер на Windows и macOS.

## Локальный запуск

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Публикация в Streamlit Community Cloud

1. Скопируйте содержимое проекта в корень GitHub-репозитория.
2. Откройте https://share.streamlit.io/.
3. Нажмите **Create app**.
4. Выберите репозиторий `panasyan99-cpu/Analitika`.
5. Branch: `main`.
6. Main file path: `streamlit_app.py`.
7. Нажмите **Deploy**.

После каждого `Commit + Push` сайт будет автоматически обновляться.

## Google Sheets

В тестовой версии готовый `.xlsx` скачивается из браузера и открывается через Google Drive / Google Sheets. Прямое создание Google-таблицы будет добавлено после настройки Google OAuth или сервисного аккаунта.
