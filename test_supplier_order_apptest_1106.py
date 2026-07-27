from __future__ import annotations

import io

from openpyxl import Workbook


def _minimal_supplier_report() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "TDSheet"
    sheet["A1"] = "Продажи товаров за период 01.03.2026 - 30.06.2026"
    sheet["A2"] = "Поставщик(и): Y&J"
    sheet["E7"] = "Продажи за период"
    sheet["G7"] = "Остатки"
    sheet["O7"] = "ТВП"
    sheet["G8"] = "TT"
    sheet["H8"] = "AB"
    sheet["I8"] = "NTR1"
    sheet["J8"] = "NTR2"
    sheet["K8"] = "SCR"
    sheet["L8"] = "63"
    sheet["M8"] = "20"
    sheet["N8"] = "Всего"
    sheet["A11"] = "Set# TEST"
    sheet["A12"] = "SKE24A001"
    sheet["B12"] = "Ruby"
    sheet["C12"] = "Earrings"
    sheet["E12"] = 8
    sheet["N12"] = 0
    sheet["O12"] = 0
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def test_supplier_order_upload_and_first_action_with_real_streamlit() -> None:
    from streamlit.testing.v1 import AppTest

    app = AppTest.from_string(
        "from src.order_workflow import render_supplier_order_dashboard\n"
        "render_supplier_order_dashboard()\n",
        default_timeout=60,
    ).run()
    assert not app.exception
    assert len(app.file_uploader) == 1

    app.file_uploader[0].upload(
        "supplier-report.xlsx",
        _minimal_supplier_report(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ).run(timeout=60)
    assert not app.exception
    assert any("Комплекты по камням" in str(item.value) for item in app.markdown)

    accept_buttons = [button for button in app.button if button.label == "Согласен с рекомендацией"]
    assert accept_buttons
    accept_buttons[0].click().run(timeout=60)
    assert not app.exception
