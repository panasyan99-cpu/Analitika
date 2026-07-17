from pathlib import Path


def source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def test_report_setup_guidance_is_not_in_about_program():
    text = source()
    assert "выгрузите два одинаково настроенных отчета" not in text
    assert "Между файлами должен отличаться только выбранный период" not in text
    assert "Как настроить отчет" not in text
    assert "Как подготовить отчет" not in text


def test_release_history_is_compact_scrollable_and_dynamic():
    text = source()
    product_info = Path(__file__).with_name("src").joinpath("product_info.py").read_text(encoding="utf-8")
    assert 'class="about-card updates-card"' in text
    assert 'class="updates-scroll"' in text
    assert "max-height:270px" in text
    assert "overflow-y:auto" in text
    assert "-webkit-overflow-scrolling:touch" in text
    assert "release_history_html" in text
    assert "CHANGELOG.md" in product_info


def test_release_history_contains_restored_versions():
    changelog = Path(__file__).with_name("CHANGELOG.md").read_text(encoding="utf-8")
    version = Path(__file__).with_name("version.json").read_text(encoding="utf-8")
    assert '"version": "1.6.4"' in version
    assert "## 1.3.0 — Full-page unified UX" in changelog
    assert "## 1.2.6 — Complete Sonu report, sidebar navigation and full export" in changelog
    assert "## 1.2.1 — Fix warehouse chart sizing and label clipping" in changelog
    assert "## 1.2.0 — Baserow warehouse analytics" in changelog
    assert "## 1.1.14 — Compact release history" in changelog
