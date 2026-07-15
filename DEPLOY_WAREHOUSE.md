# Deploy Analitika Web 1.2.0

Это полный проект, а не патч. Основной файл Streamlit: `streamlit_app.py`.

## 1. Заменить содержимое репозитория

Скопируйте все файлы этого архива в корень репозитория Analitika. Служебную папку `.git` существующего репозитория не удаляйте.

## 2. Создать отдельный read-only токен Baserow

Название: `Streamlit Warehouse Read Only`.

Выдать только право `Read` для таблиц:

- Сувенирка — 642;
- Комплектующие — 643;
- Операции — 644;
- Поставки — 645.

Права Create, Update и Delete не выдавать.

## 3. Добавить Streamlit Secrets

В Streamlit Community Cloud откройте приложение → Settings → Secrets и добавьте:

```toml
[baserow]
url = "https://storage.princess-jewelry.com"
token = "READ_ONLY_TOKEN"
souvenirs_table_id = 642
components_table_id = 643
operations_table_id = 644
supplies_table_id = 645
```

Секрет не добавляется в GitHub. Файл `.streamlit/secrets.toml` уже исключён через `.gitignore`.

## 4. Проверка локально

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Без локального секрета складской раздел покажет инструкцию подключения, а два существующих режима продаж продолжат работать.

## 5. Публикация

```bash
git add .
git commit -m "Release Analitika 1.2.0 warehouse analytics"
git push origin main
```

Streamlit Cloud автоматически пересоберёт существующий сайт по тому же URL.
