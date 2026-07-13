# Analitika Web 1.1.7 — стабильный deployment

## Код

1. Распакуйте архив в ветку `staging` с заменой файлов.
2. Commit: `Analitika Web 1.1.7 - Stability and memory optimization`.
3. Push и проверьте тестовую ссылку.
4. Слейте `staging` в `main` и сделайте Push.

## Важно: Python 3.12

Файлы `.python-version` и `runtime.txt` фиксируют Python 3.12 для нового deployment, но не меняют Python уже созданного приложения Streamlit Cloud.

Если production в логах показывает Python 3.14, его необходимо удалить в Streamlit Community Cloud и развернуть заново с теми же параметрами и тем же URL, выбрав в Advanced settings Python 3.12.

Перед удалением запишите:
- repository;
- branch `main`;
- main file `streamlit_app.py`;
- текущий app URL;
- secrets, если они есть.
