from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from slack_excel_bot.config import Settings
from slack_excel_bot.ekispert_client import EkispertError, EkispertMcpClient
from slack_excel_bot.excel_writer import DraftWriteResult, ExcelWriter
from slack_excel_bot.tool_schemas import (
    AttendanceDayOverride,
    AttendanceSheetInput,
    ExpenseEvidenceAnalysisInput,
    EmployeeInput,
    PersonalExpenseSheetInput,
    TransportRouteLookupInput,
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
        self.ekispert_client = (
            EkispertMcpClient(settings.ekispert_api_token) if settings.ekispert_api_token else None
        )

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
            "items": [self._normalize_transport_item(item) for item in args.items],
        }
        result = self.writer.write_draft("transport_jp_leadingsoft_v1", payload)
        return GeneratedWorkbook(
            template_id=result.template_id,
            output_path=result.output_path,
            title="交通费精算表",
            payload=payload,
        ).as_tool_output()

    def lookup_transport_route_options(self, raw_args: dict[str, Any]) -> dict[str, Any]:
        args = TransportRouteLookupInput.model_validate(raw_args)
        if self.ekispert_client is None:
            raise EkispertError("EXPENSES_EKISPERT_API_TOKEN is not configured.")

        options = self.ekispert_client.search_route_options(
            route_from=args.route_from,
            route_to=args.route_to,
            top_k=args.top_k or 3,
            travel_date=args.travel_date,
        )
        return {
            "ok": True,
            "title": "交通路线候选",
            "travel_date": args.travel_date,
            "route_from": args.route_from,
            "route_to": args.route_to,
            "options": [option.as_dict() for option in options],
        }

    def analyze_expense_evidence(self, raw_args: dict[str, Any]) -> dict[str, Any]:
        args = ExpenseEvidenceAnalysisInput.model_validate(raw_args)
        payload = args.model_dump(mode="json")
        payload["missing_fields"] = self._normalize_missing_expense_fields(payload)
        if payload.get("expense_type") == "transport" and payload.get("transport_mode") is None and (
            payload.get("route_from") or payload.get("route_to") or payload.get("route_line")
        ):
            payload["transport_mode"] = "電車・バス"
        for item in payload.get("transport_items", []):
            if item.get("transport_mode") is None:
                item["transport_mode"] = "電車・バス"
        return {
            "ok": True,
            "title": "票据分析结果",
            **payload,
        }

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

    @staticmethod
    def _normalize_transport_item(item) -> dict[str, Any]:
        values = item.model_dump(mode="json", exclude_none=True)
        if not values.get("purpose"):
            values["purpose"] = "営業活動"
        if "is_round_trip" not in values or values["is_round_trip"] is None:
            values["is_round_trip"] = False
        return values

    @staticmethod
    def _normalize_missing_expense_fields(values: dict[str, Any]) -> list[str]:
        expense_type = values.get("expense_type")
        if expense_type == "transport":
            transport_items = values.get("transport_items") or []
            if transport_items:
                return []
            required_fields = ("travel_date", "route_from", "route_to", "one_way_amount")
            missing = [field for field in required_fields if values.get(field) in (None, "", [])]
            route_hint_present = values.get("route_line") not in (None, "")
            if values.get("one_way_amount") is None and not route_hint_present and "route_line" not in missing:
                missing.append("route_line")
            return missing
        if expense_type == "personal_expense":
            required_fields = ("expense_date", "amount_jpy", "payee_name")
            return [field for field in required_fields if values.get(field) in (None, "", [])]
        return []
