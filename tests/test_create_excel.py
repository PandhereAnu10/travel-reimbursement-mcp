
import pytest
from openpyxl import load_workbook

import tools.create_excel as create_excel_module


def test_create_workbook(output_dir):
    result = create_excel_module.create_excel("test_reimbursement.xlsx")

    file_path = output_dir / "test_reimbursement.xlsx"

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


def test_prevent_overwriting(output_dir):

    create_excel_module.create_excel("existing.xlsx")

    with pytest.raises(FileExistsError):
        create_excel_module.create_excel("existing.xlsx")


def test_new_workbook_starts_in_draft_status(output_dir):
    from openpyxl import load_workbook
    from templates.reimbursement import SUMMARY_SHEET, SUBMISSION_STATUS_CELL, DEFAULT_STATUS

    create_excel_module.create_excel("draft_status.xlsx")
    workbook = load_workbook(output_dir / "draft_status.xlsx")
    assert workbook[SUMMARY_SHEET][SUBMISSION_STATUS_CELL].value == DEFAULT_STATUS
    workbook.close()


def test_resolve_workbook_path_rejects_symlink(tmp_path):
    outside = tmp_path.parent / "outside.xlsx"
    outside.write_bytes(b"not a workbook")
    link = tmp_path / "linked.xlsx"
    try:
        link.symlink_to(outside)
    except (OSError, NotImplementedError):
        pytest.skip("Symlink creation is not available on this platform")
    with pytest.raises(ValueError, match="Symbolic-link"):
        create_excel_module.resolve_workbook_path("linked.xlsx", tmp_path)
