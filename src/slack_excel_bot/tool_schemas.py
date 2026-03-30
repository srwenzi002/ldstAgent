from __future__ import annotations

from copy import deepcopy
from typing import Any, Literal

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

    expense_date: str = Field(description="精算対象日。YYYY-MM-DD。")
    purpose: Literal[
        "交際費",
        "会議費",
        "旅費交通費",
        "通信費",
        "消耗品費",
        "図書費",
        "福利厚生費＿レクレーション補助",
        "福利厚生費＿社内福利厚生行事",
        "福利厚生費＿健康診断",
        "福利厚生費",
        "他支払手数料",
        "印紙税",
        "他租税公課",
        "水道光熱費",
        "荷造運賃",
        "諸会費",
        "保険料",
        "立替金＿LDNS",
        "立替金",
        "その他",
    ] = Field(
        description=(
            "精算科目。必须使用模板中的精确值。"
            "可选值包括 交際費, 会議費, 旅費交通費, 通信費, 消耗品費, 図書費, "
            "福利厚生費＿レクレーション補助, 福利厚生費＿社内福利厚生行事, 福利厚生費＿健康診断, "
            "福利厚生費, 他支払手数料, 印紙税, 他租税公課, 水道光熱費, 荷造運賃, 諸会費, 保険料, "
            "立替金＿LDNS, 立替金, その他。"
        )
    )
    amount_jpy: float = Field(description="税込金額。単位は円。")
    payee_name: str = Field(description="支払先名称。例: 店舗名、会社名、サービス名。")
    description: str = Field(description="内容説明。購入内容、会食内容、用途などを簡潔に記入。")
    burden_department: Literal[
        "取締役会",
        "営業部",
        "管理部",
        "開発本部",
        "ソリューション開発部",
        "プラットフォーム開発部",
        "カスタマサポートエンジニアリング部",
    ] = Field(
        description=(
            "負担部署。必须使用模板中的精确值。"
            "可选值: 取締役会, 営業部, 管理部, 開発本部, ソリューション開発部, "
            "プラットフォーム開発部, カスタマサポートエンジニアリング部。"
        )
    )
    project_code_name: str = Field(
        description=(
            "案件コード名称。必须填写模板中的精确值，例如 "
            "SD0001：SD部門経費, PF0001：PF部門経費, CSE001：CSE部門経費, "
            "KH0001：開発本部部門経費, BL0010：先端技術開発室, BL0021：第一開発部。"
            "如果用户没有明确给出且无法从上下文确定，应先追问，不要猜测。"
        )
    )
    counterparty_company: str = Field(description="会食或交际对象的公司名。没有公司名时也要填可识别的对象名称。")
    counterparty_attendees: str = Field(description="对方参加者姓名列表。多人时用顿号、逗号或中点分隔。")
    counterparty_count: int = Field(ge=0, description="对方参加人数。")
    internal_attendees: str = Field(description="我方参加者姓名列表。多人时用顿号、逗号或中点分隔。")
    internal_count: int = Field(ge=0, description="我方参加人数。")


class PersonalExpenseSheetInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    employee: EmployeeInput | None = None
    items: list[PersonalExpenseItemInput] = Field(
        min_length=1,
        max_length=3,
        description="個人立替精算の明細。1回の帳票で最大3件まで。",
    )


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
