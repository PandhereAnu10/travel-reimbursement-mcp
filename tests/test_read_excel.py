
from datetime import datetime
from openpyxl import Workbook

import tools.read_excel as read_excel_module


def create_test_workbook(output_dir):
    file_path = output_dir / "test.xlsx"

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Employee"

    worksheet["A1"] = "Employee Name"
    worksheet["B1"] = "Test User"
    worksheet["A2"] = "Employee ID"
    worksheet["B2"] = "E001"
    worksheet["A3"] = "Travel Date"
    worksheet["B3"] = datetime(2026, 11, 25)

    workbook.create_sheet("Expenses")
    workbook.save(file_path)
    workbook.close()

    return file_path


def test_read_all_worksheets(output_dir):
    create_test_workbook(output_dir)

    result = read_excel_module.read_excel("test.xlsx")

    assert result["status"] == "success"
    assert "Employee" in result["sheets"]
    assert "Expenses" in result["sheets"]
    assert result["sheets"]["Employee"][0] == [
        "Employee Name",
        "Test User",
    ]


def test_read_specific_worksheet(tmp_path, monkeypatch):
    create_test_workbook(tmp_path)
    monkeypatch.setattr(read_excel_module, "OUTPUT_DIR", tmp_path)

    result = read_excel_module.read_excel(
        "test.xlsx",
        sheet_name="Employee",
    )

    assert result["status"] == "success"
    assert list(result["sheets"].keys()) == ["Employee"]


def test_read_empty_worksheet(tmp_path, monkeypatch):
    create_test_workbook(tmp_path)
    monkeypatch.setattr(read_excel_module, "OUTPUT_DIR", tmp_path)

    result = read_excel_module.read_excel(
        "test.xlsx",
        sheet_name="Expenses",
    )

    assert result["status"] == "success"
    assert "Expenses" in result["sheets"]
    assert result["sheets"]["Expenses"] == []


def test_convert_dates_to_iso_format(tmp_path, monkeypatch):
    create_test_workbook(tmp_path)
    monkeypatch.setattr(read_excel_module, "OUTPUT_DIR", tmp_path)

    result = read_excel_module.read_excel(
        "test.xlsx",
        sheet_name="Employee",
    )

    date_value = result["sheets"]["Employee"][2][1]

    assert date_value == "2026-11-25T00:00:00"


def test_read_nonexistent_workbook(tmp_path, monkeypatch):
    monkeypatch.setattr(read_excel_module, "OUTPUT_DIR", tmp_path)

    try:
        read_excel_module.read_excel("missing.xlsx")
        assert False, "Expected FileNotFoundError"
    except FileNotFoundError:
        pass


def test_read_nonexistent_worksheet(tmp_path, monkeypatch):
    create_test_workbook(tmp_path)
    monkeypatch.setattr(read_excel_module, "OUTPUT_DIR", tmp_path)

    try:
        read_excel_module.read_excel(
            "test.xlsx",
            sheet_name="Unknown",
        )
        assert False, "Expected ValueError"
    except ValueError as error:
        assert "does not exist" in str(error)


def test_reject_invalid_filename(tmp_path, monkeypatch):
    monkeypatch.setattr(read_excel_module, "OUTPUT_DIR", tmp_path)

    try:
        read_excel_module.read_excel("../outside.xlsx")
        assert False, "Expected ValueError"
    except ValueError:
        pass


def test_read_keeps_formula_cells(output_dir):
    file_path = create_test_workbook(output_dir)
    workbook = __import__("openpyxl").load_workbook(file_path)
    workbook["Employee"]["B4"] = "=1+1"
    workbook.save(file_path)
    workbook.close()

    result = read_excel_module.read_excel("test.xlsx", sheet_name="Employee", include_formulas=True)

    assert [None, "=1+1"] in result["sheets"]["Employee"]
    assert "formula expressions" in result["formula_note"]


def test_read_omits_formula_only_rows_by_default(output_dir):
    file_path = create_test_workbook(output_dir)
    workbook = __import__("openpyxl").load_workbook(file_path)
    workbook["Employee"]["B4"] = "=1+1"
    workbook.save(file_path)
    workbook.close()

    result = read_excel_module.read_excel("test.xlsx", sheet_name="Employee")
    assert [None, "=1+1"] not in result["sheets"]["Employee"]
    assert "formula_note" not in result
