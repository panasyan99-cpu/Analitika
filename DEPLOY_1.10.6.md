# Установка Analitika Web 1.10.6

1. Полностью скопируйте содержимое репозитория поверх текущего проекта.
2. Не удаляйте Streamlit Secrets. Проверьте, что `[order_storage]` настроен и `required = true`.
3. Выполните:

```bash
git add -A
git commit -m "Analitika Web 1.10.6 - Add fast fragments and reliable cloud autosave"
git push
```

Streamlit Community Cloud увидит изменение `requirements.txt`, сам установит `streamlit==1.60.0` и пересоберёт приложение. Ручная установка на облачном сервере не требуется.

## Локальная установка на Windows

Из папки проекта:

```powershell
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
py -m streamlit version
py -m streamlit run streamlit_app.py
```

Ожидаемая версия: `Streamlit, version 1.60.0`.

## Smoke test после деплоя

1. Убедиться, что сайт показывает 1.10.6.
2. Загрузить рабочий файл около 90 МБ.
3. Принять рекомендацию, изменить количество и замок — карточка должна обновляться заметно быстрее.
4. Заполнить размеры кольца одной отправкой формы.
5. Подождать 12–20 секунд без действий, обновить страницу и проверить восстановление черновика.
6. Сформировать Excel и проверить `Change Lock To`.
7. Переключить «Камни / Жемчуг» и проверить сохранность обоих черновиков.
