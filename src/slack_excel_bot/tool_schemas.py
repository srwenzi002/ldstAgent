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

    travel_date: str = Field(description="乘车或出行日期。YYYY-MM-DD。")
    purpose: Literal[
        "営業活動",
        "客先作業",
        "研修・セミナー参加",
        "深夜帰宅",
        "接待関連",
        "その他会社業務",
    ] | None = Field(
        default=None,
        description=(
            "出行目的。必须使用模板中的精确值。"
            "可选值: 営業活動, 客先作業, 研修・セミナー参加, 深夜帰宅, 接待関連, その他会社業務。"
            "如果用户没有说明，允许留空，程序会默认补成 営業活動。"
        )
    )
    visit_place: str | None = Field(default=None, description="访问地点、客户地点或目的地名称。没有时可省略。")
    transport_mode: Literal["電車・バス", "タクシー"] = Field(
        description="交通手段。只能是 電車・バス 或 タクシー。"
    )
    route_from: str = Field(description="出发地。")
    route_to: str = Field(description="到达地。")
    route_line: str | None = Field(default=None, description="线路名。比如 JR山手線、東京メトロ銀座線。没有时可省略。")
    one_way_amount: float = Field(description="单程金额，单位日元。")
    is_round_trip: bool | None = Field(
        default=None,
        description="是否往返。未说明时允许留空，程序会默认补成 true。",
    )
    receipt_no: str | None = Field(default=None, description="领收书编号或票据编号。没有时可省略。")


class TransportSheetInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    employee: EmployeeInput | None = None
    items: list[TransportItemInput] = Field(
        min_length=1,
        max_length=18,
        description="交通费明细。1张表最多 18 条。",
    )


class TransportRouteLookupInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    travel_date: str = Field(description="出行日期。YYYY-MM-DD。")
    route_from: str = Field(description="出发站名。")
    route_to: str = Field(description="到达站名。")
    top_k: int | None = Field(default=3, ge=1, le=5, description="返回候选路线数量，默认 3。")


class TransportEvidenceItemInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    travel_date: str = Field(description="识别出的出行日期。YYYY-MM-DD。")
    route_from: str = Field(description="识别出的出发站或地点。")
    route_to: str = Field(description="识别出的到达站或地点。")
    one_way_amount: float = Field(description="识别出的单程金额。")
    route_line: str | None = Field(default=None, description="识别出的线路描述。无法确定时填 null。")
    transport_mode: Literal["電車・バス", "タクシー"] | None = Field(
        default=None,
        description="识别出的交通方式。无法确定时填 null。",
    )
    is_round_trip: bool | None = Field(default=None, description="是否往返。无法确定时填 null。")
    purpose: Literal[
        "営業活動",
        "客先作業",
        "研修・セミナー参加",
        "深夜帰宅",
        "接待関連",
        "その他会社業務",
    ] | None = Field(default=None, description="识别出的出行目的。无法确定时填 null。")
    confidence: Literal["high", "medium", "low"] = Field(description="该条交通明细的置信度。")
    notes: str | None = Field(default=None, description="该条明细的说明，例如是否由入/出记录配对推断。")


class TransportLedgerEventInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    travel_date: str = Field(description="履历事件日期。YYYY-MM-DD。")
    event_kind: Literal["入", "出", "窓出", "物販", "定", "charge", "other"] = Field(
        description="交通卡履历中的原始事件类型。注意：定 不能直接丢弃，它可能代表定期区间相关的进站或出站事件。"
    )
    event_role: Literal["entry", "exit", "pass_entry_or_exit", "shopping", "adjustment", "charge", "unknown"] | None = (
        Field(
            default=None,
            description=(
                "对该原始事件的业务解释。"
                "入 通常对应 entry，出 通常对应 exit，"
                "定 可以填写 pass_entry_or_exit，表示它可能是定期区间相关的进/出站事件，不能轻易舍弃。"
            ),
        )
    )
    station_or_merchant: str | None = Field(default=None, description="站名或商户名。无法确定时填 null。")
    amount_jpy: float | None = Field(default=None, description="该事件对应的金额；无法确定时填 null。")
    balance_jpy: float | None = Field(default=None, description="如果能看出余额则填写；无法确定时填 null。")
    paired_group: str | None = Field(
        default=None,
        description="如果模型认为若干事件属于同一次行程，可用同一个任意字符串分组；无法判断时填 null。",
    )
    confidence: Literal["high", "medium", "low"] = Field(description="该条原始事件识别的置信度。")
    notes: str | None = Field(default=None, description="该条原始事件的补充说明。")


class ExpenseEvidenceAnalysisInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    expense_type: Literal["transport", "personal_expense", "unknown"] = Field(
        description="图片或图文内容对应的精算类型。"
    )
    document_kind: Literal["transport_screenshot", "receipt", "invoice", "unknown"] = Field(
        description="识别到的票据或截图类型。"
    )
    travel_date: str | None = Field(default=None, description="交通出行日期。YYYY-MM-DD；无法确定时填 null。")
    route_from: str | None = Field(default=None, description="交通场景中的出发站或地点。无法确定时填 null。")
    route_to: str | None = Field(default=None, description="交通场景中的到达站或地点。无法确定时填 null。")
    route_line: str | None = Field(default=None, description="交通场景中的线路描述。无法确定时填 null。")
    one_way_amount: float | None = Field(default=None, description="交通场景中的单程金额。无法确定时填 null。")
    transport_events: list[TransportLedgerEventInput] = Field(
        default_factory=list,
        description="如果图片像交通卡履历截图，请尽量先逐条抽出原始事件，包含 入/出/物販/定 等。",
    )
    transport_items: list[TransportEvidenceItemInput] = Field(
        default_factory=list,
        description="如果一张交通截图里能可靠识别出多条乘车记录，就逐条输出在这里。",
    )
    transport_mode: Literal["電車・バス", "タクシー"] | None = Field(
        default=None,
        description="交通场景中的交通方式。公共交通统一为 電車・バス；无法确定时填 null。",
    )
    is_round_trip: bool | None = Field(default=None, description="交通场景中是否往返。无法确定时填 null。")
    purpose: Literal[
        "営業活動",
        "客先作業",
        "研修・セミナー参加",
        "深夜帰宅",
        "接待関連",
        "その他会社業務",
    ] | None = Field(default=None, description="交通场景中的出行目的。无法确定时填 null。")
    expense_date: str | None = Field(default=None, description="个人报销场景中的发生日期。YYYY-MM-DD；无法确定时填 null。")
    amount_jpy: float | None = Field(default=None, description="个人报销场景中的金额。无法确定时填 null。")
    payee_name: str | None = Field(default=None, description="个人报销场景中的支付对象或商户名称。无法确定时填 null。")
    description: str | None = Field(default=None, description="个人报销场景中的费用说明。无法确定时填 null。")
    confidence: Literal["high", "medium", "low"] = Field(description="本次识别的整体置信度。")
    evidence_sources: list[Literal["text", "image"]] = Field(
        default_factory=list,
        description="本次识别依赖的证据来源，可包含 text、image。",
    )
    missing_fields: list[str] = Field(
        default_factory=list,
        description="当前仍缺失、会阻碍后续生成对应精算表的字段名列表。",
    )
    notes: str | None = Field(default=None, description="补充说明，例如冲突、歧义或不确定点。没有时填 null。")


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
