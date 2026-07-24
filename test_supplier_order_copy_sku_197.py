from src.order_workflow import copy_sku_html, normalize_copy_sku


def test_normalize_copy_sku_removes_trailing_and_leading_whitespace():
    assert normalize_copy_sku("  SFR17N124B-FPB   \n\t") == "SFR17N124B-FPB"


def test_copy_markup_uses_trimmed_sku_for_text_and_clipboard_payload():
    markup = copy_sku_html("SFR17N124B-FPB   ")

    assert "SFR17N124B-FPB   " not in markup
    assert "const sku = \"SFR17N124B-FPB\";" in markup
    assert ">SFR17N124B-FPB</span>" in markup
    assert "navigator.clipboard.writeText(sku)" in markup
    assert "Копировать артикул" in markup


def test_copy_markup_escapes_visible_html_but_preserves_clipboard_value():
    markup = copy_sku_html("SKU<&>  ")

    assert "SKU&lt;&amp;&gt;" in markup
    assert 'const sku = "SKU<&>";' in markup
