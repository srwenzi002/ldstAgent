from pathlib import Path

from openpyxl import load_workbook

from slack_excel_bot.excel_tools import ExcelToolService
from slack_excel_bot.config import Settings
from slack_excel_bot.excel_writer import ExcelWriteError, ExcelWriter


ROOT = Path(__file__).resolve().parents[1]
PACKAGE_DIR = ROOT / "src" / "slack_excel_bot"


def build_settings(tmp_path: Path) -> Settings:
    return Settings(
        slack_bot_token="xoxb-test",
        slack_app_token="xapp-test",
        openai_api_key="sk-test",
        openai_model="gpt-4.1-mini",
        port=3000,
        storage_dir=tmp_path,
        default_employee_name="山田太郎",
        default_employee_id="0001",
        default_department="開発本部",
        default_department_code="50",
        default_work_grade=1,
        default_clock_in="09:00",
        default_clock_out="18:00",
    )


def test_transport_writer_generates_workbook(tmp_path: Path) -> None:
    writer = ExcelWriter(package_dir=PACKAGE_DIR, draft_dir=tmp_path)
    payload = {
        "employee": {
            "department": "開発本部",
            "employee_id": "0001",
            "name": "山田太郎",
        },
        "items": [
            {
                "travel_date": "2026-03-10",
                "purpose": "客先作業",
                "visit_place": "渋谷オフィス",
                "transport_mode": "電車・バス",
                "route_from": "新宿",
                "route_to": "渋谷",
                "route_line": "JR山手線",
                "one_way_amount": 178,
                "is_round_trip": True,
                "receipt_no": "R-001",
            }
        ],
    }

    result = writer.write_draft("transport_jp_leadingsoft_v1", payload)
    wb = load_workbook(result.output_path)
    ws = wb["精算書（交通費）"]

    assert ws["J4"].value == "開発本部"
    assert ws["N4"].value == "0001"
    assert ws["S4"].value == "山田太郎"
    assert ws["A9"].value is not None
    assert ws["AD9"].value == "●"


def test_transport_tool_generates_workbook(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    service = ExcelToolService(settings)

    result = service.generate_transport_sheet(
        {
            "employee": {
                "department": "ソリューション開発部",
                "department_code": "51",
                "employee_id": "0001",
                "name": "山田太郎",
            },
            "items": [
                {
                    "travel_date": "2026-03-10",
                    "purpose": "客先作業",
                    "visit_place": "渋谷オフィス",
                    "transport_mode": "電車・バス",
                    "route_from": "新宿",
                    "route_to": "渋谷",
                    "route_line": "JR山手線",
                    "one_way_amount": 178,
                    "is_round_trip": True,
                    "receipt_no": "R-001",
                }
            ],
        }
    )
    wb = load_workbook(result["output_path"])
    ws = wb["精算書（交通費）"]

    assert ws["J4"].value == "ソリューション開発部"
    assert ws["N4"].value == "0001"
    assert ws["S4"].value == "山田太郎"
    assert ws["F9"].value == "客先作業"
    assert ws["K9"].value == "電車・バス"
    assert ws["AA9"].value == 178
    assert ws["AD9"].value == "●"


def test_transport_tool_applies_default_purpose_and_round_trip(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    service = ExcelToolService(settings)

    result = service.generate_transport_sheet(
        {
            "employee": {
                "department": "ソリューション開発部",
                "department_code": "51",
                "employee_id": "0001",
                "name": "山田太郎",
            },
            "items": [
                {
                    "travel_date": "2026-03-29",
                    "purpose": None,
                    "visit_place": None,
                    "transport_mode": "電車・バス",
                    "route_from": "青砥",
                    "route_to": "青物横丁",
                    "route_line": None,
                    "one_way_amount": 651,
                    "is_round_trip": None,
                    "receipt_no": None,
                }
            ],
        }
    )
    wb = load_workbook(result["output_path"])
    ws = wb["精算書（交通費）"]

    assert ws["F9"].value == "営業活動"
    assert ws["AD9"].value in (None, "")


def test_writer_raises_for_missing_required_field(tmp_path: Path) -> None:
    writer = ExcelWriter(package_dir=PACKAGE_DIR, draft_dir=tmp_path)
    payload = {
        "employee": {
            "department": "開発本部",
            "employee_id": "0001",
        },
        "items": [],
    }

    try:
        writer.write_draft("transport_jp_leadingsoft_v1", payload)
    except ExcelWriteError as exc:
        assert exc.code == "FIELD_MISSING"
        assert exc.details["field_path"] == "employee.name"
    else:
        raise AssertionError("Expected ExcelWriteError")


def test_attendance_tool_expands_weekdays(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    service = ExcelToolService(settings)

    result = service.generate_attendance_sheet(
        {
            "year": 2026,
            "month": 3,
            "employee": {"department_code": "51", "employee_id": "0001", "name": "山田太郎"},
            "days": [
                {"day": 2, "work_grade": 2, "clock_in": "09:00", "clock_out": "17:30"},
                {"day": 3, "work_grade": 2, "clock_in": "09:00", "clock_out": "17:30"},
            ],
        }
    )
    wb = load_workbook(result["output_path"])
    ws = wb["勤務状況表"]

    assert ws["AE4"].value == "0001"
    assert ws["AF4"].value == "山田太郎"
    assert ws["F12"].value is None
    assert ws["G12"].value is None
    assert ws["F13"].value is not None
    assert ws["G13"].value is not None
    assert ws["E13"].value == 2


def test_attendance_tool_writes_half_day_leave_directly(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    service = ExcelToolService(settings)

    result = service.generate_attendance_sheet(
        {
            "year": 2026,
            "month": 2,
            "employee": {"department_code": "51", "employee_id": "0001", "name": "山田太郎"},
            "days": [
                {
                    "day": 6,
                    "leave_item_no": 2,
                    "work_grade": 2,
                    "clock_in": "13:00",
                    "clock_out": "18:00",
                }
            ],
        }
    )
    wb = load_workbook(result["output_path"])
    ws = wb["勤務状況表"]

    assert ws["K17"].value == 2
    assert ws["E17"].value == 2
    assert ws["F17"].value is not None
    assert ws["G17"].value is not None


def test_personal_expense_tool_generates_workbook(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    service = ExcelToolService(settings)

    result = service.generate_personal_expense_sheet(
        {
            "employee": {
                "department": "ソリューション開発部",
                "department_code": "51",
                "employee_id": "0001",
                "name": "山田太郎",
            },
            "items": [
                {
                    "expense_date": "2026-03-15",
                    "purpose": "会議費",
                    "amount_jpy": 5500,
                    "payee_name": "渋谷カフェ",
                    "description": "顧客打合せの飲食代",
                    "burden_department": "ソリューション開発部",
                    "project_code_name": "SD0001：SD部門経費",
                    "counterparty_company": "株式会社サンプル",
                    "counterparty_attendees": "佐藤様",
                    "counterparty_count": 1,
                    "internal_attendees": "山田太郎",
                    "internal_count": 1,
                }
            ],
        }
    )
    wb = load_workbook(result["output_path"])
    ws = wb["経費（個人）精算書"]

    assert ws["J5"].value == "ソリューション開発部"
    assert ws["N5"].value == "0001"
    assert ws["R5"].value == "山田太郎"
    assert ws["J10"].value == "会議費"
    assert ws["Y10"].value == 5500
    assert ws["A12"].value == "渋谷カフェ"
    assert ws["I12"].value == "顧客打合せの飲食代"
