from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field


class EmployeeInput(BaseModel):
    employee_id: str | None = Field(default=None, description="社員番号")
    name: str | None = Field(default=None, description="氏名")
    department: str | None = Field(default=None, description="部署名")
    department_code: str | None = Field(default=None, description="部署コード")


class AttendanceDayOverride(BaseModel):
    day: int = Field(ge=1, le=31)
    work_grade: int | None = Field(default=None, ge=1, le=4)
    clock_in: str | None = Field(default=None, description="HH:MM")
    clock_out: str | None = Field(default=None, description="HH:MM")
    special_note: str | None = None
    leave_item_no: int | None = None


class AttendanceSheetInput(BaseModel):
    year: int = Field(ge=2000, le=2100)
    month: int = Field(ge=1, le=12)
    employee: EmployeeInput | None = None
    full_attendance: bool = Field(default=True, description="平日を全勤として自動展開するか")
    work_grade: int | None = Field(default=None, ge=1, le=4)
    clock_in: str | None = Field(default=None, description="HH:MM")
    clock_out: str | None = Field(default=None, description="HH:MM")
    paid_leave_balance: float | None = None
    day_overrides: list[AttendanceDayOverride] = Field(default_factory=list)


class TransportItemInput(BaseModel):
    travel_date: str = Field(description="YYYY-MM-DD")
    purpose: str
    visit_place: str | None = None
    transport_mode: str
    route_from: str
    route_to: str
    route_line: str | None = None
    one_way_amount: float
    is_round_trip: bool = False
    receipt_no: str | None = None


class TransportSheetInput(BaseModel):
    employee: EmployeeInput | None = None
    items: list[TransportItemInput] = Field(min_length=1, max_length=18)


class PersonalExpenseItemInput(BaseModel):
    expense_date: str = Field(description="YYYY-MM-DD")
    purpose: str
    amount_jpy: float
    payee_name: str
    description: str
    burden_department: str
    project_code_name: str
    counterparty_company: str
    counterparty_attendees: str
    counterparty_count: int
    internal_attendees: str
    internal_count: int


class PersonalExpenseSheetInput(BaseModel):
    employee: EmployeeInput | None = None
    items: list[PersonalExpenseItemInput] = Field(min_length=1, max_length=3)


def openai_function_tool(name: str, description: str, model: type[BaseModel]) -> dict[str, Any]:
    return {
        "type": "function",
        "name": name,
        "description": description,
        "parameters": model.model_json_schema(),
        "strict": True,
    }
