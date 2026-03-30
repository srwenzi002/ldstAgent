from __future__ import annotations

from copy import deepcopy
from typing import Any

from pydantic import BaseModel, ConfigDict, Field


class EmployeeInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    employee_id: str | None = Field(default=None, description="社員番号")
    name: str | None = Field(default=None, description="氏名")
    department: str | None = Field(default=None, description="部署名")
    department_code: str | None = Field(
        default=None,
        description=(
            "部署コード。"
            "10=取締役会, 20=営業部, 30=管理部, 50=開発本部, 51=ソリューション開発部, "
            "52=プラットフォーム開発部, 60=カスタマサポートエンジニアリング部, 70=その他指定部署。"
        ),
    )


class AttendanceDayOverride(BaseModel):
    model_config = ConfigDict(extra="forbid")

    day: int = Field(ge=1, le=31, description="日付の日部分。1-31。")
    work_grade: int | None = Field(
        default=None,
        ge=1,
        le=4,
        description="就業区分。1=09:30-18:00, 2=09:00-17:30, 3=10:00-18:30, 4=10:30-19:00。",
    )
    clock_in: str | None = Field(default=None, description="出勤時刻。HH:MM。")
    clock_out: str | None = Field(default=None, description="退勤時刻。HH:MM。")
    special_note: str | None = Field(default=None, description="特記事項。必要時のみ。")
    leave_item_no: int | None = Field(
        default=None,
        ge=1,
        le=15,
        description=(
            "休暇項目番号。"
            "1=有給休暇(全日), 2=有給休暇(午前), 3=有給休暇(午後), 4=欠勤, 5=健診BC, "
            "6=無給休暇, 7=振休, 8=代休, 9=特別代休, 10=結忌引配出産, 11=SP5(GW・夏季), "
            "12=その他特休, 13=積立休暇, 14=休業, 15=教育訓練。"
        ),
    )


class AttendanceSheetInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    year: int = Field(ge=2000, le=2100, description="対象年。")
    month: int = Field(ge=1, le=12, description="対象月。1-12。")
    employee: EmployeeInput
    paid_leave_balance: float | None = Field(default=None, description="有休残日数。")
    days: list[AttendanceDayOverride] = Field(
        default_factory=list,
        description="表に書き込む日ごとの勤怠データ。必要な日だけ出力する。",
    )


class TransportItemInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

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
    model_config = ConfigDict(extra="forbid")

    employee: EmployeeInput | None = None
    items: list[TransportItemInput] = Field(min_length=1, max_length=18)


class PersonalExpenseItemInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

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
    model_config = ConfigDict(extra="forbid")

    employee: EmployeeInput | None = None
    items: list[PersonalExpenseItemInput] = Field(min_length=1, max_length=3)


def _normalize_for_openai(schema: dict[str, Any]) -> dict[str, Any]:
    normalized = deepcopy(schema)

    def walk(node: Any) -> None:
        if isinstance(node, dict):
            if node.get("type") == "object" and "properties" in node:
                properties = node.get("properties", {})
                node["additionalProperties"] = False
                node["required"] = list(properties.keys())
            for value in node.values():
                walk(value)
        elif isinstance(node, list):
            for item in node:
                walk(item)

    walk(normalized)
    return normalized


def openai_function_tool(name: str, description: str, model: type[BaseModel]) -> dict[str, Any]:
    return {
        "type": "function",
        "name": name,
        "description": description,
        "parameters": _normalize_for_openai(model.model_json_schema()),
        "strict": True,
    }
