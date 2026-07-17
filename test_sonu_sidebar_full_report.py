from pathlib import Path


def test_sonu_navigation_is_sidebar_anchor_navigation():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def _sonu_sidebar_navigation" in text
    assert '"sonu-main-report"' in text
    assert '"sonu-extra"' in text
    assert '"sonu-export"' in text
    assert '"global-filter"' not in text[text.index("def sonu_navigation_items"):text.index("def _sonu_sidebar_navigation")]


def test_sonu_renders_five_groups_in_one_report():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    expected = [
        '_anchor("sonu-main-report")',
        '_render_sonu_group(title, table',
        '_anchor("sonu-extra")',
        '_render_sonu_extra(section_tables)',
        '_anchor("sonu-export")',
    ]
    positions = [text.index(token) for token in expected]
    assert positions == sorted(positions)
    for label in ["Серьги", "Кольца", "Подвески", "Браслеты не полный круг", "Браслеты полный круг"]:
        assert label in text


def test_sonu_full_excel_export_contains_new_business_blocks():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def build_full_sonu_export" in text
    assert '("Рекомендации", recommendations)' in text
    assert '("Категории", category_overview)' in text
    assert "К заказу на 30 дней" not in text[text.index("def build_full_sonu_export"):text.index("@dataclass", text.index("def build_full_sonu_export"))]
