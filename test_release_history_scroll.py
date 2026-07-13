from pathlib import Path


def source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def test_comparison_guidance_requires_same_report_structure_and_different_period():
    text = source()
    assert "выгрузите два одинаково настроенных отчета" in text
    assert "Между файлами должен отличаться только выбранный период" in text


def test_release_history_is_compact_and_scrollable():
    text = source()
    assert 'class="about-card updates-card"' in text
    assert 'class="updates-scroll"' in text
    assert "max-height:270px" in text
    assert "overflow-y:auto" in text
    assert "-webkit-overflow-scrolling:touch" in text


def test_release_history_contains_1_1_14():
    text = source()
    assert 'APP_VERSION = "1.1.14"' in text
    assert "Analitika Web 1.1.14 — Compact release history" in text
