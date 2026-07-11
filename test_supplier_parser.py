"""Smoke test for supplier hierarchy parsing.
Run locally: python test_supplier_parser.py path/to/report.xlsx
"""
from pathlib import Path
import sys

from streamlit_app import is_supplier_report, parse_supplier_report, supplier_summary


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python test_supplier_parser.py report.xlsx")
        return 2
    path = Path(sys.argv[1])
    assert path.exists(), path
    assert is_supplier_report(path), "File is not recognized as supplier report"
    detail = parse_supplier_report(path)
    assert not detail.empty, "Supplier detail is empty"
    assert "Сеть" not in set(detail["Поставщик"].astype(str)), "Service supplier 'Сеть' was not normalized"
    summary = supplier_summary(detail)
    assert int(summary["Количество"].sum()) == int(detail["Количество"].sum())
    print(f"OK: {len(detail)} detail rows, {summary['Поставщик'].nunique()} suppliers")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
