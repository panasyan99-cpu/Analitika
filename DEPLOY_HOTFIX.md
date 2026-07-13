# Analitika Web 1.1.10 — тестовый deployment

1. Переключитесь на ветку `staging`.
2. Распакуйте архив в корень репозитория с заменой файлов.
3. Commit: `Analitika Web 1.1.10 - Fix table sorting and executive cards`.
4. Нажмите `Push origin` и проверьте тестовую ссылку Streamlit, привязанную к `staging`.
5. Production-ссылка на ветке `main` при этом не изменится.

После проверки:

1. Переключитесь на `main`.
2. Выполните `Branch → Merge into current branch… → staging`.
3. Нажмите `Push origin`.

Существующий production deployment обновится на той же ссылке. Удалять приложение в Streamlit не требуется.
