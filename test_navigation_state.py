from pathlib import Path
import ast

APP = Path(__file__).with_name("streamlit_app.py")
source = APP.read_text(encoding="utf-8")
tree = ast.parse(source)

assert '"uploaded_payloads"' in source
assert 'key="upload_widget"' in source
assert 'files = saved_uploads()' in source
assert source.index('"📦 Поставщики"') < source.index('"📤 Экспорт"')

pages = [
    'elif page == "📊 Сводка":',
    'elif page == "🏪 Магазины":',
    'elif page == "🔎 Интерактивная аналитика":',
    'elif page == "📦 Поставщики":',
    'elif page == "📤 Экспорт":',
]
for page in pages:
    assert page in source, page

print("Navigation/state static checks passed")
