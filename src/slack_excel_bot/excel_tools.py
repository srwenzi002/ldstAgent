from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from slack_excel_bot.config import Settings
from slack_excel_bot.excel_writer import DraftWriteResult, ExcelWriter
from slack_excel_bot.tool_schemas import (
    AttendanceDayOverride,
    AttendanceSheetInput,
    EmployeeInput,
    PersonalExpenseSheetInput,
    TransportSheetInput,
)


@dataclass
class GeneratedWorkbook:
    template_id: str
    output_path: str
    title: str
    payload: dict[str, Any]

    def as_tool_output(self) -> dict[str, Any]:
        return {
            "template_id": self.template_id,
            "output_path": self.output_path,
            "title": self.title,
            "payload": self.payload,
        }


class ExcelToolService:
    def __init__(self, settings: Settings):
        package_dir = Path(__file__).resolve().parent
        self.settings = settings
        self.writer = ExcelWriter(package_dir=package_dir, draft_dir=settings.storage_dir / "drafts")

    def generate_attendance_sheet(self, raw_args: dict[str, Any]) -> dict[str, Any]:
        args = AttendanceSheetInput.model_validate(raw_args)
        employee = self._merge_employee_defaults(args.employee)
        payload = {
            "year": args.year,
            "month": args.month,
            "employee": employee,
            "work_grade": None,
            "paid_leave_balance": args.paid_leave_balance,
            "items": self._build_attendance_items(args),
        }
        result = self.writer.write_draft("timesheet_jp_leadingsoft_v1", payload)
        return GeneratedWorkbook(
            template_id=result.template_id,
            output_path=result.output_path,
            title=f"{args.year}年{args.month}月_考勤表",
            payload=payload,
        ).as_tool_output()

    def generate_transport_sheet(self, raw_args: dict[str, Any]) -> dict[str, Any]:
        args = TransportSheetInput.model_validate(raw_args)
        payload = {
            "employee": self._merge_employee_defaults(args.employee),
            "items": [item.model_dump(mode="json", exclude_none=True) for item in args.items],
        }
        result = self.writer.write_draft("transport_jp_leadingsoft_v1", payload)
        return GeneratedWorkbook(
            template_id=result.template_id,
            output_path=result.output_path,
            title="交通费精算表",
            payload=payload,
        ).as_tool_output()

    def generate_personal_expense_sheet(self, raw_args: dict[str, Any]) -> dict[str, Any]:
        args = PersonalExpenseSheetInput.model_validate(raw_args)
        payload = {
            "employee": self._merge_employee_defaults(args.employee),
            "items": [item.model_dump(mode="json", exclude_none=True) for item in args.items],
        }
        result = self.writer.write_draft("personal_expense_jp_leadingsoft_v1", payload)
        return GeneratedWorkbook(
            template_id=result.template_id,
            output_path=result.output_path,
            title="个人报销计算表",
            payload=payload,
        ).as_tool_output()

    def _merge_employee_defaults(self, employee: EmployeeInput | None) -> dict[str, Any]:
        values = employee.model_dump(exclude_none=True) if employee else {}
        return {
            "employee_id": values.get("employee_id") or self.settings.default_employee_id,
            "name": values.get("name") or self.settings.default_employee_name,
            "department": values.get("department") or self.settings.default_department,
            "department_code": values.get("department_code") or self.settings.default_department_code,
        }

    def _build_attendance_items(self, args: AttendanceSheetInput) -> list[dict[str, Any]]:
        items = [item.model_dump(mode="json", exclude_none=True) for item in args.days]
        items.sort(key=lambda item: int(item["day"]))
        return items
