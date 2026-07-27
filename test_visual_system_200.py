from pathlib import Path
import json

ROOT = Path(__file__).parent


def test_version_and_release_files():
    assert json.loads((ROOT / "version.json").read_text(encoding="utf-8"))["version"] == "2.0"
    assert (ROOT / "RELEASE_NOTES_2.0.md").exists()
    assert (ROOT / "DEPLOY_2.0.md").exists()


def test_all_management_headers_are_present():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    for phrase in (
        "Сводный обзор продаж по сети",
        "Сопоставление двух периодов",
        "Единое рабочее пространство для контроля остатков",
        "Анализ продаж ассортимента Sonu",
        "Расчёт потребности в товарах",
    ):
        assert phrase in app


def test_visual_tokens_cover_every_workspace():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    for token in (
        "Analitika 2.0 — unified visual system",
        ".report-context",
        ".empty-state",
        ".product-flow",
        ".warehouse-section-heading",
        ".sonu-data-card",
        "polish_chart_surface",
    ):
        assert token in app


def test_mobile_header_is_compact():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    mobile = app[app.index("@media (max-width:600px)", app.index("Analitika 2.0")):]
    assert ".luxury-badges { display:none; }" in mobile
    assert ".luxury-title { font-size:29px" in mobile
