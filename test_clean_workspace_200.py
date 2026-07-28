from pathlib import Path

ROOT = Path(__file__).parent


def test_workspaces_do_not_repeat_upload_guidance():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    sonu = (ROOT / "src" / "sonu.py").read_text(encoding="utf-8")
    order = (ROOT / "src" / "order_workflow.py").read_text(encoding="utf-8")

    for phrase in (
        "Загрузите отчёт для анализа",
        "Загрузите отчёт Sonu",
        "Загрузите отчёт для нового заказа",
        'class="empty-state"',
    ):
        assert phrase not in app + sonu + order


def test_primary_uploaders_do_not_show_help_icons():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    sonu = (ROOT / "src" / "sonu.py").read_text(encoding="utf-8")
    order = (ROOT / "src" / "order_workflow.py").read_text(encoding="utf-8")

    assert "Используйте единый отчет с иерархией" not in app
    assert "Остаток в файле — общий по сети" not in sonu
    assert "Имя Excel-файла может быть любым" not in order


def test_version_is_current_hotfix():
    assert '"version": "2.4.3"' in (ROOT / "version.json").read_text(encoding="utf-8")
