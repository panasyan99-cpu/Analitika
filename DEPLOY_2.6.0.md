# Deploy — Analitika Web 2.6.0

## 1. Обновление кода

Распакуйте релиз в корень репозитория, не копируя `.streamlit/secrets.toml` из других источников.

```bash
git add -A
git commit -m "release: Analitika Web 2.6.0 reliability and recovery"
git push
```

## 2. Streamlit Secrets

Существующие Secrets не заменять. Проверьте наличие `[auth]`, `[baserow]` и `[order_storage]` по примеру `.streamlit/secrets.toml.example`.

Для автоматического ремонта структуры Baserow нужны `email` и `password`. Для обычной работы таблиц используется database token.

## 3. Проверка после деплоя

1. Откройте **Склад → Диагностика**.
2. Нажмите **«Проверить всё снова»**.
3. Проверьте чтение и запись Baserow, загрузку файлов и R2.
4. Когда показан план изменения схемы, подтвердите и выполните ремонт один раз.
5. Проверьте две последние ожидаемые поставки и 32 SKU серебра.
6. Создайте тестовый черновик заказа поставщику и убедитесь, что статус сохранения показывает оба хранилища честно.

## 4. Browser smoke в GitHub Actions

Добавьте repository secrets:

- `ANALITIKA_BASE_URL`;
- `ANALITIKA_PASSWORD`.

Запустите workflow **Analitika browser smoke** вручную через Actions.
