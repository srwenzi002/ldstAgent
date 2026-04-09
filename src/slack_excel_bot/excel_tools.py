from __future__ import annotations

import calendar
import contextvars
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import holidays

from slack_excel_bot.config import Settings
from slack_excel_bot.ekispert_client import EkispertError, EkispertMcpClient
from slack_excel_bot.excel_writer import DraftWriteResult, ExcelWriter
from slack_excel_bot.thread_draft_store import ThreadDraftStore
from slack_excel_bot.tool_schemas import (
    AttendanceDayOverride,
    AttendanceDraftUpsertInput,
    CalendarContextInput,
    AttendanceSheetInput,
    ExpenseEvidenceAnalysisInput,
    EmployeeInput,
    PersonalExpenseDraftUpsertInput,
    PersonalExpenseSheetInput,
    StationCandidateLookupInput,
    TransportDraftUpsertInput,
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


@dataclass
class DraftRunState:
    thread_ts: str
    store: ThreadDraftStore
    updated_templates: set[str]
    replacement_messages: list[str]


class ExcelToolService:
    FARE_TOLERANCE_JPY = 10.0
    TOKYO_STATION_PREFIX_RULES = {
        "京成": {"operator_keywords": ("京成",), "strip_prefix_for_query": True},
        "京急": {"operator_keywords": ("京急",), "strip_prefix_for_query": True},
        "JR": {"operator_keywords": ("JR", "ＪＲ"), "strip_prefix_for_query": True},
        "地": {"operator_keywords": ("東京メトロ", "都営"), "strip_prefix_for_query": True},
        "都": {"operator_keywords": ("都営",), "strip_prefix_for_query": True},
        "メトロ": {"operator_keywords": ("東京メトロ",), "strip_prefix_for_query": True},
        "東武": {"operator_keywords": ("東武",), "strip_prefix_for_query": True},
        "西武": {"operator_keywords": ("西武",), "strip_prefix_for_query": True},
        "京王": {"operator_keywords": ("京王",), "strip_prefix_for_query": True},
        "東急": {"operator_keywords": ("東急",), "strip_prefix_for_query": True},
        "小田急": {"operator_keywords": ("小田急",), "strip_prefix_for_query": True},
        "相鉄": {"operator_keywords": ("相鉄",), "strip_prefix_for_query": True},
    }
    TOKYO_STATION_ALIAS_QUERIES = {
        "京急八景": ("金沢八景", "金沢八景(京急線)"),
        "京成日暮": ("日暮里",),
        "京成日暮里": ("日暮里",),
        "京成上野": ("京成上野",),
    }
    WORK_GRADE_SCHEDULES = {
        1: {"clock_in": "09:30", "clock_out": "18:00", "half_day_cutoff": "12:30"},
        2: {"clock_in": "09:00", "clock_out": "17:30", "half_day_cutoff": "12:00"},
        3: {"clock_in": "10:00", "clock_out": "18:30", "half_day_cutoff": "13:00"},
        4: {"clock_in": "10:30", "clock_out": "19:00", "half_day_cutoff": "13:30"},
    }

    def __init__(self, settings: Settings):
        package_dir = Path(__file__).resolve().parent
        self.settings = settings
        self.writer = ExcelWriter(package_dir=package_dir, draft_dir=settings.storage_dir / "drafts")
        self.ekispert_client = (
            EkispertMcpClient(settings.ekispert_api_token) if settings.ekispert_api_token else None
        )
        self._draft_run_state: contextvars.ContextVar[DraftRunState | None] = contextvars.ContextVar(
            "draft_run_state",
            default=None,
        )

    def start_draft_run(self, thread_ts: str, store: ThreadDraftStore) -> contextvars.Token[DraftRunState | None]:
        return self._draft_run_state.set(
            DraftRunState(
                thread_ts=thread_ts,
                store=store,
                updated_templates=set(),
                replacement_messages=[],
            )
        )

    def finish_draft_run(self, token: contextvars.Token[DraftRunState | None]) -> DraftRunState | None:
        state = self._draft_run_state.get()
        self._draft_run_state.reset(token)
        return state

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

    def get_month_calendar_context(self, raw_args: dict[str, Any]) -> dict[str, Any]:
        args = CalendarContextInput.model_validate(raw_args)
        jp_holidays = holidays.Japan(years=[args.year], language="ja")
        weekday_labels = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        month_days = calendar.monthrange(args.year, args.month)[1]
        days: list[dict[str, Any]] = []

        for day in range(1, month_days + 1):
            current = date(args.year, args.month, day)
            weekday_index = current.weekday()
            holiday_name = jp_holidays.get(current)
            days.append(
                {
                    "date": current.isoformat(),
                    "day": day,
                    "weekday": weekday_labels[weekday_index],
                    "is_weekend": weekday_index >= 5,
                    "is_holiday": holiday_name is not None,
                    "holiday_name": holiday_name,
                }
            )

        return {
            "ok": True,
            "title": "月次カレンダー情報",
            "year": args.year,
            "month": args.month,
            "days": days,
        }

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

    def upsert_transport_draft(self, raw_args: dict[str, Any]) -> dict[str, Any]:
        args = TransportDraftUpsertInput.model_validate(raw_args)
        draft_patch = {
            "employee": self._merge_employee_defaults(args.employee),
            "items": [self._normalize_transport_item(item) for item in args.items],
        }
        return self._upsert_template_draft(
            template_type="transport",
            mode=args.mode,
            draft_patch=draft_patch,
            pending_questions=args.pending_questions,
        )

    def lookup_transport_route_options(self, raw_args: dict[str, Any]) -> dict[str, Any]:
        args = TransportRouteLookupInput.model_validate(raw_args)
        if self.ekispert_client is None:
            raise EkispertError("EXPENSES_EKISPERT_API_TOKEN is not configured.")

        try:
            options, normalization_info = self._search_route_options_with_station_resolution(
                route_from=args.route_from,
                route_to=args.route_to,
                top_k=args.top_k or 3,
                travel_date=args.travel_date,
            )
        except EkispertError as exc:
            return {
                "ok": False,
                "title": "交通経路候補",
                "travel_date": args.travel_date,
                "route_from": args.route_from,
                "route_to": args.route_to,
                "status": "query_error",
                "match_type": "query_error",
                "should_prompt_user": True,
                "prompt_reason": "query_error",
                "options": [],
                "error": str(exc),
            }
        response = {
            "ok": True,
            "title": "交通経路候補",
            "travel_date": args.travel_date,
            "route_from": args.route_from,
            "route_to": args.route_to,
            "options": [option.as_dict() for option in options],
        }
        if normalization_info:
            response["station_normalizations"] = normalization_info
        return response

    def lookup_station_candidates(self, raw_args: dict[str, Any]) -> dict[str, Any]:
        args = StationCandidateLookupInput.model_validate(raw_args)
        if self.ekispert_client is None:
            raise EkispertError("EXPENSES_EKISPERT_API_TOKEN is not configured.")

        try:
            normalized = self._lookup_station_candidates_with_variants(
                station_name=args.station_name,
                top_k=args.top_k or 5,
                prefecture_code=args.prefecture_code,
                match_type=args.match_type or "partial",
                station_type=args.station_type or "train",
            )
        except EkispertError as exc:
            return {
                "ok": False,
                "title": "駅名候補",
                "station_name": args.station_name,
                "candidates": [],
                "error": str(exc),
            }

        return {
            "ok": True,
            "title": "駅名候補",
            "station_name": args.station_name,
            "prefix_hint": normalized["prefix_hint"],
            "query_variants": normalized["query_variants"],
            "auto_selected_station_name": normalized["resolved_station_name"],
            "prefecture_code": args.prefecture_code,
            "match_type": args.match_type or "partial",
            "station_type": args.station_type or "train",
            "candidates": [candidate.as_dict() for candidate in normalized["candidates"]],
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
                options, normalization_info = self._search_route_options_with_station_resolution(
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
                if normalization_info:
                    result_item["station_normalizations"] = normalization_info
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

    def _search_route_options_with_station_resolution(
        self,
        *,
        route_from: str,
        route_to: str,
        top_k: int,
        travel_date: str | None,
    ) -> tuple[list[Any], list[dict[str, Any]]]:
        assert self.ekispert_client is not None
        original_error: EkispertError | None = None
        try:
            return (
                self.ekispert_client.search_route_options(
                    route_from=route_from,
                    route_to=route_to,
                    top_k=top_k,
                    travel_date=travel_date,
                ),
                [],
            )
        except EkispertError as exc:
            original_error = exc
            if "駅名が見つかりません" not in str(exc):
                raise

        normalization_info: list[dict[str, Any]] = []
        resolved_from = self._resolve_station_name(route_from)
        resolved_to = self._resolve_station_name(route_to)

        if resolved_from["resolved_station_name"] and resolved_from["resolved_station_name"] != route_from:
            normalization_info.append(
                {
                    "field": "route_from",
                    "original": route_from,
                    "resolved": resolved_from["resolved_station_name"],
                    "prefix_hint": resolved_from["prefix_hint"],
                    "query_variants": resolved_from["query_variants"],
                }
            )
        if resolved_to["resolved_station_name"] and resolved_to["resolved_station_name"] != route_to:
            normalization_info.append(
                {
                    "field": "route_to",
                    "original": route_to,
                    "resolved": resolved_to["resolved_station_name"],
                    "prefix_hint": resolved_to["prefix_hint"],
                    "query_variants": resolved_to["query_variants"],
                }
            )

        retry_from = resolved_from["resolved_station_name"] or route_from
        retry_to = resolved_to["resolved_station_name"] or route_to
        if retry_from == route_from and retry_to == route_to:
            raise original_error

        return (
            self.ekispert_client.search_route_options(
                route_from=retry_from,
                route_to=retry_to,
                top_k=top_k,
                travel_date=travel_date,
            ),
            normalization_info,
        )

    def _resolve_station_name(self, station_name: str) -> dict[str, Any]:
        normalized = self._lookup_station_candidates_with_variants(
            station_name=station_name,
            top_k=5,
            prefecture_code="13",
            match_type="partial",
            station_type="train",
        )
        return {
            "prefix_hint": normalized["prefix_hint"],
            "query_variants": normalized["query_variants"],
            "resolved_station_name": normalized["resolved_station_name"],
        }

    def _lookup_station_candidates_with_variants(
        self,
        *,
        station_name: str,
        top_k: int,
        prefecture_code: str | None,
        match_type: str,
        station_type: str,
    ) -> dict[str, Any]:
        assert self.ekispert_client is not None
        query_plan = self._build_station_query_plan(station_name)
        candidate_scores: dict[str, tuple[float, Any]] = {}
        search_station_candidates = getattr(self.ekispert_client, "search_station_candidates", None)
        if search_station_candidates is None:
            return {
                "prefix_hint": query_plan["prefix_hint"],
                "query_variants": query_plan["query_variants"],
                "resolved_station_name": None,
                "candidates": [],
            }

        for query in query_plan["query_variants"]:
            if not query:
                continue
            candidates = search_station_candidates(
                station_name=query,
                top_k=top_k,
                prefecture_code=prefecture_code,
                match_type=match_type,
                station_type=station_type,
            )
            for rank, candidate in enumerate(candidates):
                score = self._score_station_candidate(
                    raw_station_name=station_name,
                    query=query,
                    candidate=candidate,
                    prefix_hint=query_plan["prefix_hint"],
                    rank=rank,
                )
                current = candidate_scores.get(candidate.station_code)
                if current is None or score > current[0]:
                    candidate_scores[candidate.station_code] = (score, candidate)

        ordered_candidates = [
            item[1]
            for item in sorted(
                candidate_scores.values(),
                key=lambda item: (
                    -item[0],
                    item[1].station_name,
                ),
            )[:top_k]
        ]

        return {
            "prefix_hint": query_plan["prefix_hint"],
            "query_variants": query_plan["query_variants"],
            "resolved_station_name": self._select_resolved_station_name(
                raw_station_name=station_name,
                query_variants=query_plan["query_variants"],
                candidates=ordered_candidates,
            ),
            "candidates": ordered_candidates,
        }

    @classmethod
    def _build_station_query_plan(cls, station_name: str) -> dict[str, Any]:
        raw_name = cls._compact_station_name(station_name)
        prefix_hint = None
        query_variants: list[str] = []
        for prefix in sorted(cls.TOKYO_STATION_PREFIX_RULES, key=len, reverse=True):
            if raw_name.startswith(prefix):
                prefix_hint = prefix
                break

        def add_query(value: str | None) -> None:
            compacted = cls._compact_station_name(value)
            if compacted and compacted not in query_variants:
                query_variants.append(compacted)

        add_query(raw_name)

        if prefix_hint is not None:
            stripped = raw_name.removeprefix(prefix_hint)
            add_query(stripped)
            for alias in cls.TOKYO_STATION_ALIAS_QUERIES.get(raw_name, ()):
                add_query(alias)

        return {
            "prefix_hint": prefix_hint,
            "query_variants": query_variants,
        }

    @classmethod
    def _score_station_candidate(
        cls,
        *,
        raw_station_name: str,
        query: str,
        candidate: Any,
        prefix_hint: str | None,
        rank: int,
    ) -> float:
        raw = cls._compact_station_name(raw_station_name)
        query_compact = cls._compact_station_name(query)
        candidate_name = cls._compact_station_name(candidate.station_name)
        score = 0.0
        if candidate_name == raw:
            score += 120
        if candidate_name == query_compact:
            score += 90
        if query_compact and query_compact in candidate_name:
            score += 40
        if raw and raw in candidate_name:
            score += 20
        if prefix_hint is not None:
            rule = cls.TOKYO_STATION_PREFIX_RULES[prefix_hint]
            if any(keyword in candidate.station_name for keyword in rule["operator_keywords"]):
                score += 30
        if candidate.prefecture_code == "13":
            score += 10
        score -= rank
        return score

    @classmethod
    def _select_resolved_station_name(
        cls,
        *,
        raw_station_name: str,
        query_variants: list[str],
        candidates: list[Any],
    ) -> str | None:
        if not candidates:
            return None

        query_set = {cls._compact_station_name(query) for query in query_variants}
        for candidate in candidates:
            candidate_name = cls._compact_station_name(candidate.station_name)
            if candidate_name in query_set and candidate_name != cls._compact_station_name(raw_station_name):
                return candidate.station_name

        if len(candidates) == 1:
            return candidates[0].station_name
        return None

    @staticmethod
    def _compact_station_name(value: str | None) -> str:
        if value is None:
            return ""
        compacted = str(value).replace("\u3000", "").replace(" ", "").strip()
        compacted = re.sub(r"[()（）]", "", compacted)
        return compacted

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

    def upsert_personal_expense_draft(self, raw_args: dict[str, Any]) -> dict[str, Any]:
        args = PersonalExpenseDraftUpsertInput.model_validate(raw_args)
        draft_patch = {
            "employee": self._merge_employee_defaults(args.employee) if args.employee else None,
            "items": [item.model_dump(mode="json", exclude_none=True) for item in args.items],
        }
        return self._upsert_template_draft(
            template_type="personal_expense",
            mode=args.mode,
            draft_patch=draft_patch,
            pending_questions=args.pending_questions,
        )

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

    def upsert_attendance_draft(self, raw_args: dict[str, Any]) -> dict[str, Any]:
        args = AttendanceDraftUpsertInput.model_validate(raw_args)
        draft_patch: dict[str, Any] = {
            "days": [self._normalize_attendance_item(item) for item in args.days],
        }
        if args.year is not None:
            draft_patch["year"] = args.year
        if args.month is not None:
            draft_patch["month"] = args.month
        if args.employee is not None:
            draft_patch["employee"] = self._merge_employee_defaults(args.employee)
        if args.paid_leave_balance is not None:
            draft_patch["paid_leave_balance"] = args.paid_leave_balance
        return self._upsert_template_draft(
            template_type="attendance",
            mode=args.mode,
            draft_patch=draft_patch,
            pending_questions=args.pending_questions,
        )

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
        items = [self._normalize_attendance_item(item) for item in args.days]
        items.sort(key=lambda item: int(item["day"]))
        return items

    def _normalize_attendance_item(self, item: AttendanceDayOverride) -> dict[str, Any]:
        normalized = item.model_dump(mode="json", exclude_none=True)
        work_grade = normalized.get("work_grade")
        if work_grade is None:
            return normalized

        schedule = self.WORK_GRADE_SCHEDULES.get(int(work_grade))
        if schedule is None:
            return normalized

        leave_item_no = normalized.get("leave_item_no")
        if leave_item_no == 2:
            normalized["clock_in"] = schedule["half_day_cutoff"]
            normalized["clock_out"] = schedule["clock_out"]
            return normalized

        if leave_item_no == 3:
            normalized["clock_in"] = schedule["clock_in"]
            normalized["clock_out"] = schedule["half_day_cutoff"]
            return normalized

        normalized.setdefault("clock_in", schedule["clock_in"])
        normalized.setdefault("clock_out", schedule["clock_out"])
        return normalized

    @staticmethod
    def _normalize_transport_item(item) -> dict[str, Any]:
        values = item.model_dump(mode="json", exclude_none=True)
        if not values.get("purpose"):
            values["purpose"] = "営業活動"
        if "is_round_trip" not in values or values["is_round_trip"] is None:
            values["is_round_trip"] = False
        return values

    def _upsert_template_draft(
        self,
        *,
        template_type: str,
        mode: str,
        draft_patch: dict[str, Any],
        pending_questions: list[str],
    ) -> dict[str, Any]:
        run_state = self._draft_run_state.get()
        if run_state is None:
            raise RuntimeError("Draft run context is not set.")

        cleaned_patch = {key: value for key, value in draft_patch.items() if value not in (None, [], {})}
        result = run_state.store.upsert_draft(
            thread_ts=run_state.thread_ts,
            template_type=template_type,  # type: ignore[arg-type]
            mode=mode,  # type: ignore[arg-type]
            draft_patch=cleaned_patch,
            pending_questions=pending_questions,
        )
        run_state.updated_templates.add(template_type)
        if result.replaced_previous:
            run_state.replacement_messages.append(
                f"同じテンプレートの以前の草稿は最新版に置き換えました（{template_type}）。"
            )

        draft = result.snapshot["drafts_by_template"][template_type]
        return {
            "ok": True,
            "title": "草稿状態を更新しました",
            "template_type": template_type,
            "status": draft["status"],
            "pending_questions": draft["pending_questions"],
            "replaced_previous": result.replaced_previous,
            "canonical_state": draft["canonical_state"],
            "ready_to_generate": self.is_template_ready(template_type, draft["canonical_state"]),
        }

    def is_template_ready(self, template_type: str, canonical_state: dict[str, Any]) -> bool:
        try:
            if template_type == "transport":
                TransportSheetInput.model_validate(canonical_state)
                return True
            if template_type == "personal_expense":
                PersonalExpenseSheetInput.model_validate(canonical_state)
                return True
            if template_type == "attendance":
                AttendanceSheetInput.model_validate(canonical_state)
                return True
        except Exception:
            return False
        return False

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
