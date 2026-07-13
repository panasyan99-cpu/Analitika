# Analitika Web 1.1.10 — безопасный перенос в production

## Проверка

1. Распакуйте архив в ветку `staging` с заменой файлов.
2. Commit: `Analitika Web 1.1.10 - Fix table sorting and executive cards`.
3. Push и проверьте тестовую ссылку.
4. Убедитесь, что загружается реальная выгрузка, открывается «Оперативная сводка», таблицы сортируются, а диаграммы остаются заблокированными от масштабирования.

## Production без смены ссылки

1. Переключитесь на `main`.
2. Подтяните изменения через `Fetch origin` / `Pull origin`, если потребуется.
3. Выполните `Branch → Merge into current branch… → staging`.
4. Нажмите `Push origin`.

Streamlit обновит существующий сайт на прежнем URL. Deployment удалять не нужно.
