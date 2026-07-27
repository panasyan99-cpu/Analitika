# Установка Analitika Web 1.10.5

1. Полностью скопируйте содержимое репозитория поверх текущего проекта.
2. Не удаляйте существующие Streamlit Secrets. Убедитесь, что `[order_storage]` настроен и `required = true`.
3. Выполните коммит:

```bash
git add -A
git commit -m "Analitika Web 1.10.5 - Improve supplier order performance and storage reliability"
git push
```

4. После деплоя проверьте номер `1.10.5`, загрузку файла около 90 МБ, пакетное изменение количества/замка/размеров и восстановление черновика.

Лимит файла установлен на 150 МБ в `.streamlit/config.toml`.
