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
    TransportRouteBatchLookupInput,
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
    FARE_TOLERANCE_JPY = 10.0

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
            title=self._build_attendance_title(payload),
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
            title=self._build_transport_title(payload),
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
            "title": "交通経路候補",
            "travel_date": args.travel_date,
            "route_from": args.route_from,
            "route_to": args.route_to,
            "options": [option.as_dict() for option in options],
        }

    def lookup_transport_route_batch(self, raw_args: dict[str, Any]) -> dict[str, Any]:
        args = TransportRouteBatchLookupInput.model_validate(raw_args)
        if self.ekispert_client is None:
            raise EkispertError("EXPENSES_EKISPERT_API_TOKEN is not configured.")

        results: list[dict[str, Any]] = []
        resolved_candidates: list[dict[str, Any]] = []
        needs_confirmation: list[dict[str, Any]] = []
        any_success = False
        for index, item in enumerate(args.items, start=1):
            base_result = {
                "item_id": str(index),
                "travel_date": item.travel_date,
                "route_from": item.route_from,
                "route_to": item.route_to,
                "image_one_way_amount": item.one_way_amount,
                "image_route_line": item.route_line,
            }
            try:
                options = self.ekispert_client.search_route_options(
                    route_from=item.route_from,
                    route_to=item.route_to,
                    top_k=args.top_k or 3,
                    travel_date=item.travel_date,
                )
                option_dicts = [option.as_dict() for option in options]
                match_summary = self._summarize_route_match(
                    image_one_way_amount=item.one_way_amount,
                    options=option_dicts,
                )
                result_item = {
                    **base_result,
                    "status": "ok",
                    **match_summary,
                    "options": option_dicts,
                    "error": None,
                }
                results.append(result_item)
                if result_item["matched_option"] is not None and not result_item["should_prompt_user"]:
                    resolved_candidates.append(
                        {
                            "item_id": str(index),
                            "travel_date": item.travel_date,
                            "purpose": None,
                            "visit_place": None,
                            "transport_mode": "電車・バス",
                            "route_from": item.route_from,
                            "route_to": item.route_to,
                            "route_line": result_item["matched_option"]["route_line"],
                            "one_way_amount": result_item["final_one_way_amount"],
                            "is_round_trip": False,
                            "receipt_no": None,
                        }
                    )
                else:
                    needs_confirmation.append(result_item)
                any_success = True
            except EkispertError as exc:
                failed_item = {
                    **base_result,
                    "status": "query_error",
                    "matched_option": None,
                    "recommended_option": None,
                    "final_one_way_amount": item.one_way_amount,
                    "match_type": "query_error",
                    "should_prompt_user": True,
                    "prompt_reason": "query_error",
                    "options": [],
                    "error": str(exc),
                }
                results.append(failed_item)
                needs_confirmation.append(failed_item)

        resolved_items, round_trip_suggestions = self._merge_round_trip_candidates(resolved_candidates)
        return {
            "ok": any_success,
            "title": "交通経路の一括確認結果",
            "has_partial_failures": any(item["status"] != "ok" for item in results),
            "resolved_items": resolved_items,
            "round_trip_suggestions": round_trip_suggestions,
            "needs_confirmation": needs_confirmation,
            "items": results,
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
            "title": "証憑の解析結果",
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
            title=self._build_personal_expense_title(payload),
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

    @staticmethod
    def _build_transport_title(payload: dict[str, Any]) -> str:
        employee = payload["employee"]
        items = payload.get("items") or []
        yyyymm = "YYYYMM"
        if items:
            travel_date = str(items[0].get("travel_date") or "")
            if len(travel_date) >= 7:
                yyyymm = travel_date[:7].replace("-", "")
        return f"精算書_集（交通費）_{employee['name']}_{yyyymm}"

    @staticmethod
    def _build_personal_expense_title(payload: dict[str, Any]) -> str:
        employee = payload["employee"]
        items = payload.get("items") or []
        yyyymm = "YYYYMM"
        if items:
            expense_date = str(items[0].get("expense_date") or "")
            if len(expense_date) >= 7:
                yyyymm = expense_date[:7].replace("-", "")
        return f"精算書_集（個人経費立替）_{employee['name']}_{yyyymm}"

    @staticmethod
    def _build_attendance_title(payload: dict[str, Any]) -> str:
        employee = payload["employee"]
        year = int(payload["year"])
        month = int(payload["month"])
        yymm = f"{year % 100:02d}{month:02d}"
        return f"Ldjpw668_{yymm}_{employee['department_code']}_{employee['employee_id']}_{employee['name']}"

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

    @classmethod
    def _summarize_route_match(
        cls,
        *,
        image_one_way_amount: float | None,
        options: list[dict[str, Any]],
    ) -> dict[str, Any]:
        ranked_options = sorted(
            options,
            key=lambda option: (
                int(option.get("transfer_count") or 99),
                int(option.get("total_minutes") or 999),
                float(option.get("one_way_amount") or 999999),
            ),
        )
        if image_one_way_amount is None:
            recommended_option = ranked_options[0] if ranked_options else None
            return {
                "match_type": "no_image_amount",
                "matched_option": None,
                "recommended_option": recommended_option,
                "final_one_way_amount": recommended_option["one_way_amount"] if recommended_option else None,
                "should_prompt_user": recommended_option is None,
                "prompt_reason": "missing_image_amount" if recommended_option else "no_route_candidates",
            }

        tolerance_matches = [
            option
            for option in ranked_options
            if abs(float(option["one_way_amount"]) - float(image_one_way_amount)) <= cls.FARE_TOLERANCE_JPY
        ]
        tolerance_matches.sort(
            key=lambda option: (
                abs(float(option["one_way_amount"]) - float(image_one_way_amount)),
                int(option.get("transfer_count") or 99),
                int(option.get("total_minutes") or 999),
                float(option.get("one_way_amount") or 999999),
            )
        )

        if not tolerance_matches:
            recommended_option = ranked_options[0] if ranked_options else None
            return {
                "match_type": "unmatched",
                "matched_option": None,
                "recommended_option": recommended_option,
                "final_one_way_amount": image_one_way_amount,
                "should_prompt_user": True,
                "prompt_reason": "fare_out_of_tolerance",
            }

        best_option = tolerance_matches[0]
        ambiguous = len(tolerance_matches) > 1 and cls._options_are_similar(best_option, tolerance_matches[1])
        return {
            "match_type": "exact"
            if float(best_option["one_way_amount"]) == float(image_one_way_amount)
            else "near_ic_fare",
            "matched_option": None if ambiguous else best_option,
            "recommended_option": best_option,
            "final_one_way_amount": image_one_way_amount,
            "should_prompt_user": ambiguous,
            "prompt_reason": "multiple_close_candidates" if ambiguous else None,
        }

    @staticmethod
    def _options_are_similar(first: dict[str, Any], second: dict[str, Any]) -> bool:
        return (
            abs(float(first["one_way_amount"]) - float(second["one_way_amount"])) <= 10
            and abs(int(first.get("transfer_count") or 99) - int(second.get("transfer_count") or 99)) <= 1
            and abs(int(first.get("total_minutes") or 999) - int(second.get("total_minutes") or 999)) <= 10
        )

    @staticmethod
    def _merge_round_trip_candidates(
        resolved_candidates: list[dict[str, Any]],
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        merged_items: list[dict[str, Any]] = []
        round_trip_suggestions: list[dict[str, Any]] = []
        used_indexes: set[int] = set()

        for index, item in enumerate(resolved_candidates):
            if index in used_indexes:
                continue

            pair_index = None
            for candidate_index in range(index + 1, len(resolved_candidates)):
                candidate = resolved_candidates[candidate_index]
                if candidate_index in used_indexes:
                    continue
                if (
                    candidate["travel_date"] == item["travel_date"]
                    and float(candidate["one_way_amount"]) == float(item["one_way_amount"])
                    and candidate["route_from"] == item["route_to"]
                    and candidate["route_to"] == item["route_from"]
                ):
                    pair_index = candidate_index
                    break

            if pair_index is None:
                merged_items.append({key: value for key, value in item.items() if key != "item_id"})
                continue

            pair_item = resolved_candidates[pair_index]
            used_indexes.add(index)
            used_indexes.add(pair_index)
            merged_items.append(
                {
                    "travel_date": item["travel_date"],
                    "purpose": item.get("purpose"),
                    "visit_place": item.get("visit_place"),
                    "transport_mode": item["transport_mode"],
                    "route_from": item["route_from"],
                    "route_to": item["route_to"],
                    "route_line": item.get("route_line"),
                    "one_way_amount": item["one_way_amount"],
                    "is_round_trip": True,
                    "receipt_no": item.get("receipt_no"),
                }
            )
            round_trip_suggestions.append(
                {
                    "travel_date": item["travel_date"],
                    "route_from": item["route_from"],
                    "route_to": item["route_to"],
                    "one_way_amount": item["one_way_amount"],
                    "merged_item_ids": [item["item_id"], pair_item["item_id"]],
                    "message": "同日・同額で往路と復路がそろっていたため、いったん往復としてまとめました✨ 別々の片道にしたい場合は教えてください。",
                }
            )

        return merged_items, round_trip_suggestions
