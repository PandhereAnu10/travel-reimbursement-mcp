
import pytest
from openpyxl import load_workbook

import tools.create_excel as create_excel_module
import tools.calculate_expense as calculate_module

from templates.reimbursement import CALCULATOR_SHEET


def test_mileage_calculation(reimbursement_file):
    result = calculate_module.calculate_expense(
        filename="test.xlsx",
        calculation_type="mileage",
        distance_km=100,
        rate_per_km=10,
    )

    assert result["status"] == "success"
    assert float(result["result"]) == 1000

    workbook = load_workbook(reimbursement_file)
    worksheet = workbook[CALCULATOR_SHEET]

    assert worksheet["B4"].value == 100
    assert worksheet["B5"].value == 10
    assert worksheet["C6"].value == 1000

    workbook.close()


def test_per_diem_calculation(reimbursement_file):
    result = calculate_module.calculate_expense(
        filename="test.xlsx",
        calculation_type="per_diem",
        travel_days=5,
        daily_allowance=1000,
    )

    assert result["status"] == "success"
    assert float(result["result"]) == 5000

    workbook = load_workbook(reimbursement_file)
    worksheet = workbook[CALCULATOR_SHEET]

    assert worksheet["B9"].value == 5
    assert worksheet["B10"].value == 1000
    assert worksheet["C11"].value == 5000

    workbook.close()


def test_currency_conversion(reimbursement_file):
    result = calculate_module.calculate_expense(
        filename="test.xlsx",
        calculation_type="currency_conversion",
        foreign_amount=100,
        exchange_rate=83,
        currency_code="USD",
    )

    assert result["status"] == "success"
    assert float(result["result"]) == 8300

    workbook = load_workbook(reimbursement_file)
    worksheet = workbook[CALCULATOR_SHEET]

    assert worksheet["B14"].value == 100
    assert worksheet["B15"].value == 83
    assert worksheet["B16"].value == "USD"
    assert worksheet["C17"].value == 8300

    workbook.close()


@pytest.mark.parametrize(
    ("number1", "number2", "operation", "expected"),
    [
        (10, 5, "+", 15),
        (10, 5, "-", 5),
        (10, 5, "*", 50),
        (10, 5, "/", 2),
    ],
)
def test_quick_calculation(
    reimbursement_file,
    number1,
    number2,
    operation,
    expected,
):
    result = calculate_module.calculate_expense(
        filename="test.xlsx",
        calculation_type="quick_calculation",
        number1=number1,
        number2=number2,
        operation=operation,
    )

    assert result["status"] == "success"
    assert float(result["result"]) == expected

    workbook = load_workbook(reimbursement_file)
    worksheet = workbook[CALCULATOR_SHEET]

    assert worksheet["B20"].value == number1
    assert worksheet["B21"].value == operation
    assert worksheet["B22"].value == number2
    assert worksheet["C23"].value == expected

    workbook.close()


def test_reject_negative_mileage(reimbursement_file):
    result = calculate_module.calculate_expense(
        filename="test.xlsx",
        calculation_type="mileage",
        distance_km=-10,
        rate_per_km=5,
    )

    assert result["status"] == "error"
    assert "negative" in result["message"].lower()


def test_reject_zero_per_diem_days(reimbursement_file):
    result = calculate_module.calculate_expense(
        filename="test.xlsx",
        calculation_type="per_diem",
        travel_days=0,
        daily_allowance=1000,
    )

    assert result["status"] == "error"


def test_reject_zero_exchange_rate(reimbursement_file):
    result = calculate_module.calculate_expense(
        filename="test.xlsx",
        calculation_type="currency_conversion",
        foreign_amount=100,
        exchange_rate=0,
        currency_code="USD",
    )

    assert result["status"] == "error"


def test_reject_missing_currency_code(reimbursement_file):
    result = calculate_module.calculate_expense(
        filename="test.xlsx",
        calculation_type="currency_conversion",
        foreign_amount=100,
        exchange_rate=83,
        currency_code="",
    )

    assert result["status"] == "error"


def test_reject_invalid_calculation_type(reimbursement_file):
    result = calculate_module.calculate_expense(
        filename="test.xlsx",
        calculation_type="invalid_type",
    )

    assert result["status"] == "error"


def test_reject_division_by_zero(reimbursement_file):
    result = calculate_module.calculate_expense(
        filename="test.xlsx",
        calculation_type="quick_calculation",
        number1=10,
        number2=0,
        operation="/",
    )

    assert result["status"] == "error"


def test_reject_negative_calculation_operand(reimbursement_file):
    result = calculate_module.calculate_expense(
        filename="test.xlsx",
        calculation_type="quick_calculation",
        number1=-10,
        number2=5,
        operation="+",
    )

    assert result["status"] == "error"


def test_reject_invalid_quick_calculation_operation(
    reimbursement_file,
):
    result = calculate_module.calculate_expense(
        filename="test.xlsx",
        calculation_type="quick_calculation",
        number1=10,
        number2=5,
        operation="%",
    )

    assert result["status"] == "error"


def test_calculate_for_missing_workbook(output_dir):
    result = calculate_module.calculate_expense(
        filename="missing.xlsx",
        calculation_type="mileage",
        distance_km=100,
        rate_per_km=10,
    )
    
    assert result["status"] == "error"
    assert "not found" in result["message"].lower()


def test_reject_invalid_filename(tmp_path, monkeypatch):
    monkeypatch.setattr(
        calculate_module,
        "OUTPUT_DIR",
        tmp_path,
    )

    result = calculate_module.calculate_expense(
        filename="../outside.xlsx",
        calculation_type="mileage",
        distance_km=100,
        rate_per_km=10,
    )

    assert result["status"] == "error"
