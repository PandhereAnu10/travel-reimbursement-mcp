
from openpyxl import Workbook

import tools.list_excel_files as list_module


def create_workbook(file_path, sheet_name="Sheet1"):
    workbook = Workbook()
    workbook.active.title = sheet_name
    workbook.save(file_path)
    workbook.close()


def test_list_excel_files(output_dir):
    create_workbook(output_dir / "first.xlsx")
    create_workbook(output_dir / "second.xlsx")

    result = list_module.list_excel_files()

    assert result["status"] == "success"

    filenames = [
        item["filename"]
        for item in result["files"]
    ]

    assert "first.xlsx" in filenames
    assert "second.xlsx" in filenames


def test_list_excel_file_details(output_dir):
    create_workbook(
        output_dir / "reimbursement.xlsx",
        sheet_name="Employee",
    )

    result = list_module.list_excel_files()

    file_info = next(
        item
        for item in result["files"]
        if item["filename"] == "reimbursement.xlsx"
    )

    assert file_info["size_bytes"] > 0
    assert file_info["last_modified"]
    assert "Employee" in file_info["worksheets"]


def test_ignore_non_excel_files(output_dir):
    create_workbook(output_dir / "valid.xlsx")

    (output_dir / "notes.txt").write_text(
        "This is not an Excel file."
    )

    result = list_module.list_excel_files()

    filenames = [
        item["filename"]
        for item in result["files"]
    ]

    assert "valid.xlsx" in filenames
    assert "notes.txt" not in filenames


def test_ignore_temporary_excel_files(output_dir):
    create_workbook(output_dir / "valid.xlsx")

    (output_dir / "~$temporary.xlsx").write_text("Temporary file")

    result = list_module.list_excel_files()

    filenames = [
        item["filename"]
        for item in result["files"]
    ]

    assert "valid.xlsx" in filenames
    assert "~$temporary.xlsx" not in filenames


def test_empty_output_directory(output_dir):
    result = list_module.list_excel_files()

    assert result["status"] == "success"
    assert result["files"] == []


def test_nonexistent_output_directory(tmp_path, monkeypatch):
    missing_directory = tmp_path / "missing"

    monkeypatch.setattr(
        list_module,
        "OUTPUT_DIR",
        missing_directory,
    )

    result = list_module.list_excel_files()

    assert result["status"] == "success"
    assert result["files"] == []
