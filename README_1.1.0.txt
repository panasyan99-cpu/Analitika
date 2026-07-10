АНАЛИТИКА 1.1.4

ГОТОВЫЙ УСТАНОВЩИК БЕЗ PYTHON
Самый простой способ — GitHub Actions:
1. Создайте публичный репозиторий GitHub.
2. Загрузите в него содержимое этой папки.
3. Откройте Actions -> Build Windows installer -> Run workflow.
4. Скачайте artifact Analitika_Setup_1.1.4.
5. Передайте файл Analitika_Setup_1.1.4.exe руководителям.

Для автообновлений:
- замените OWNER/REPOSITORY в update_config.json на имя репозитория;
- создавайте релизы с тегами v1.1.4, v1.2.0 и прикладывайте установщик.

Руководителям Python и PyCharm не нужны.

Разработка: Vladimir Panasyan
