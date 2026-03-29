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
        slack_signing_secret="secret",
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

    result = service.generate_attendance_sheet({"year": 2026, "month": 3, "full_attendance": True})
    wb = load_workbook(result["output_path"])
    ws = wb["勤務状況表"]

    assert ws["AE4"].value == "0001"
    assert ws["AF4"].value == "山田太郎"
    assert ws["F12"].value is None
    assert ws["G12"].value is None
    assert ws["F13"].value is not None
    assert ws["G13"].value is not None
