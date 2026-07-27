# Установка Analitika Web 1.10.6.1

1. Распаковать ZIP.
2. Скопировать **всё содержимое** папки репозитория поверх текущего репозитория.
3. Выполнить:

```bash
git add -A
git commit -m "Analitika Web 1.10.6.1 - Fix Streamlit deployment dependencies and startup"
git push
```

4. Дождаться зелёного GitHub Actions.
5. В Streamlit Community Cloud открыть **Manage app → Reboot app**.
6. В логах проверить:

```text
Found Streamlit version 1.60.0
pyarrow==24.0.0
```

Старое виртуальное окружение может временно показывать Streamlit 1.59.1 до полного Reboot.
