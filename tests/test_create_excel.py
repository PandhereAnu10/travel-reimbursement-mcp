
import pytest
from openpyxl import load_workbook

import tools.create_excel as create_excel_module


def test_create_workbook(tmp_path, monkeypatch):
    # Using temporary directory instead of real output directory.
    monkeypatch.setattr(create_excel_module, "OUTPUT_DIR", tmp_path)

    result = create_excel_module.create_excel("test_reimbursement.xlsx")

    file_path = tmp_path / "test_reimbursement.xlsx"

    assert file_path.exists()
    assert result["status"] == "success"

    workbook = load_workbook(file_path)

    assert workbook.sheetnames == [
        "Employee & Travel Information",
        "Expense Details",
        "Expense Calculator",
        "Claim Summary & HR",
    ]

    assert workbook["Employee & Travel Information"]["A1"].value == (
        "Employee & Travel Information"
    )


def test_filename_without_extension():
    result = create_excel_module.validate_filename("reimbursement")

    assert result == "reimbursement.xlsx"


def test_reject_directory_path():
    with pytest.raises(ValueError):
        create_excel_module.validate_filename("../outside.xlsx")


def test_reject_wrong_extension():
    with pytest.raises(ValueError):
        create_excel_module.validate_filename("reimbursement.csv")


def test_prevent_overwriting(tmp_path, monkeypatch):
    monkeypatch.setattr(create_excel_module, "OUTPUT_DIR", tmp_path)

    create_excel_module.create_excel("existing.xlsx")

    with pytest.raises(FileExistsError):
        create_excel_module.create_excel("existing.xlsx")
