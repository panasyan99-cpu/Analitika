# Analitika Web 1.1.7 — deployment

1. Переключиться на ветку `staging` и распаковать архив в корень репозитория с заменой файлов.
2. Commit: `Analitika Web 1.1.7 - Stability and memory optimization`.
3. Push в `staging` и проверить тестовую ссылку.
4. После проверки переключиться на `main` и выполнить `Branch → Merge into current branch… → staging`.
5. Нажать `Push origin`; production-сайт обновится автоматически.

Критично: production deployment должен работать на Python 3.12. Если в логах указан Python 3.14, удалите приложение в Streamlit Community Cloud и разверните его заново с тем же URL, выбрав Python 3.12 в Advanced settings.
