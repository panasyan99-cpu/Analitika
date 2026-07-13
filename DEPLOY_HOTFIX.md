# Analitika Web 1.1.1 — deployment hotfix

1. Распаковать содержимое архива в корень репозитория с заменой файлов.
2. Commit: `Analitika Web 1.1.1 - Cloud stability hotfix`.
3. Push в ветку `main`.
4. В Streamlit Community Cloud удалить текущий deployment приложения и развернуть его заново из того же репозитория.
5. В Advanced settings выбрать Python 3.12.
6. Main file path: `streamlit_app.py`.

Причина повторного deployment: версия Python для уже созданного приложения не меняется обычным commit/push.
