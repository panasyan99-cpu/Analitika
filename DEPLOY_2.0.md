# Развёртывание Analitika Web 2.0 — Clean Workspace

Это дополненная сборка того же релиза **2.0**. Номер версии не повышается.

1. Распакуйте архив.
2. Скопируйте **всё содержимое** папки репозитория поверх текущего репозитория.
3. Не удаляйте существующие Streamlit Secrets.
4. Выполните:

```bash
git add -A
git commit -m "Analitika Web 2.0 - Simplify upload workspaces and centralize guidance"
git push
```

5. Дождитесь зелёного GitHub Actions.
6. В Streamlit Cloud выполните **Manage app → Reboot app**.

После запуска в интерфейсе должна по-прежнему отображаться версия **2.0**.
