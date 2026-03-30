from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable
from zoneinfo import ZoneInfo

from openai import OpenAI

from slack_excel_bot.config import Settings
from slack_excel_bot.debug_trace import DebugTrace
from slack_excel_bot.excel_tools import ExcelToolService
from slack_excel_bot.tool_schemas import (
    AttendanceSheetInput,
    ExpenseEvidenceAnalysisInput,
    PersonalExpenseSheetInput,
    TransportRouteLookupInput,
    TransportSheetInput,
    openai_function_tool,
)


@dataclass
class AgentResult:
    text: str
    generated_files: list[dict[str, Any]] = field(default_factory=list)


class OpenAIExcelAgent:
    def __init__(self, settings: Settings, tool_service: ExcelToolService):
        self.settings = settings
        self.tool_service = tool_service
        self.client = OpenAI(api_key=settings.openai_api_key)
        self.tools = [
            openai_function_tool(
                "generate_attendance_sheet",
                (
                    "生成日本月度考勤表。"
                    "当你调用此工具时，必须通过 function call arguments 提供接近最终 Excel 的 JSON 数据。"
                    "不要把 JSON 直接回复给用户。"
                    "employee.department_code 只能是 10/20/30/50/51/52/60/70。"
                    "days[].work_grade 只能是 1/2/3/4，其中 1=09:30-18:00, 2=09:00-17:30, 3=10:00-18:30, 4=10:30-19:00。"
                    "days[].leave_item_no 只能是 1..15，其中 1=有給休暇(全日), 2=有給休暇(午前), 3=有給休暇(午後), 4=欠勤, 5=健診BC, 6=無給休暇, 7=振休, 8=代休, 9=特別代休, 10=結忌引配出産, 11=SP5(GW・夏季), 12=その他特休, 13=積立休暇, 14=休業, 15=教育訓練。"
                    "不要提供 schema 之外的字段。"
                ),
                AttendanceSheetInput,
            ),
            openai_function_tool(
                "analyze_expense_evidence",
                (
                    "从用户的文字和图片中分析这是一张什么票据或截图，并抽取对应的精算字段。"
                    "这是一个通用分析工具，不生成 Excel，也不查询外部路线。"
                    "图片可能是交通截图，也可能是个人报销相关的发票或小票。"
                    "请先判断 expense_type 是 transport、personal_expense 还是 unknown，再只填写有证据支持的字段。"
                    "没有把握的金额、日期、线路、商户名都必须填 null，不要猜。"
                    "如果图片是 Suica/PASMO 等交通卡的履历截图，请优先逐条填写 transport_events，保留原始事件类型，例如 入、出、窓出、物販、定。"
                    "只有当某些原始事件能够被可靠地解释成可报销的乘车记录时，才把它们进一步汇总到 transport_items。"
                    "不要把 物販 误当成普通乘车明细。"
                    "定 不能直接舍弃；它可能表示定期区间相关的进站或出站事件，应先保留在 transport_events，并在需要时参与前后记录的推断。"
                    "如果记录跨多张截图连续出现，可以结合多张图片一起判断。"
                    "当 transport_items 已经足够完整时，不要再保守地只填一个 top-level 的单条交通字段。"
                    "如果图片和文字都参与了判断，evidence_sources 要同时包含 text 和 image。"
                    "missing_fields 中请列出还缺哪些关键字段。"
                ),
                ExpenseEvidenceAnalysisInput,
            ),
            openai_function_tool(
                "lookup_transport_route_options",
                (
                    "查询交通路线与金额候选。"
                    "当用户想做交通费精算，但没有提供准确路线、线路名或金额时，优先调用此工具。"
                    "参数中的日期必须是绝对日期 YYYY-MM-DD。"
                    "route_from 和 route_to 应尽量只填写站名，不要加入多余说明。"
                    "拿到候选后，先用自然语言列出 2-3 个候选给用户确认，不要立刻生成 Excel。"
                ),
                TransportRouteLookupInput,
            ),
            openai_function_tool(
                "generate_transport_sheet",
                (
                    "生成交通费精算表。"
                    "当你调用此工具时，必须通过 function call arguments 提供接近最终 Excel 的 JSON 数据。"
                    "不要把 JSON 直接回复给用户。"
                    "items 最多 18 条。"
                    "items[].purpose 必须使用模板中的精确值：営業活動, 客先作業, 研修・セミナー参加, 深夜帰宅, 接待関連, その他会社業務。"
                    "items[].transport_mode 只能是 電車・バス 或 タクシー。"
                    "items[].one_way_amount 必须填写单程金额，is_round_trip=true 时表示模板会标记往返。"
                    "如果用户没有说明 purpose，可以留空，程序会默认补成 営業活動。"
                    "如果用户没有说明是否往返，可以留空，程序会默认补成 false。"
                    "visit_place、route_line、receipt_no 在用户没有提供时可以留空。"
                    "不要输出 schema 之外的字段。"
                ),
                TransportSheetInput,
            ),
            openai_function_tool(
                "generate_personal_expense_sheet",
                (
                    "生成个人立替经费精算表。"
                    "当你调用此工具时，必须通过 function call arguments 提供接近最终 Excel 的 JSON 数据。"
                    "不要把 JSON 直接回复给用户。"
                    "items 最多 3 条。"
                    "items[].purpose 必须使用模板中的精确值，例如 交際費, 会議費, 旅費交通費, 通信費, "
                    "消耗品費, 図書費, 福利厚生費＿レクレーション補助, 福利厚生費＿社内福利厚生行事, "
                    "福利厚生費＿健康診断, 福利厚生費, 他支払手数料, 印紙税, 他租税公課, 水道光熱費, "
                    "荷造運賃, 諸会費, 保険料, 立替金＿LDNS, 立替金, その他。"
                    "items[].burden_department 必须使用精确部署名。"
                    "items[].project_code_name 必须填写模板允许的完整案件コード名称；如果用户没有明确提供且无法可靠判断，应先追问，不要猜测。"
                    "不要输出 schema 之外的字段。"
                ),
                PersonalExpenseSheetInput,
            ),
        ]
        self.handlers: dict[str, Callable[[dict[str, Any]], dict[str, Any]]] = {
            "generate_attendance_sheet": self.tool_service.generate_attendance_sheet,
            "analyze_expense_evidence": self.tool_service.analyze_expense_evidence,
            "lookup_transport_route_options": self.tool_service.lookup_transport_route_options,
            "generate_transport_sheet": self.tool_service.generate_transport_sheet,
            "generate_personal_expense_sheet": self.tool_service.generate_personal_expense_sheet,
        }

    def run(self, conversation_input: list[dict[str, Any]], trace: DebugTrace | None = None) -> AgentResult:
        today_jst = datetime.now(ZoneInfo("Asia/Tokyo")).date()
        instructions = (
            "你是一个 Slack 助手。"
            "你和用户之间永远使用自然语言交流。"
            "你和工具之间使用结构化 JSON。"
            "你需要根据用户在私聊中的文字和图片，判断是否需要生成 Excel。"
            "如果只是寒暄或普通问题，直接自然回复。"
            "如果用户要做考勤表、交通费精算表、个人报销计算表，请优先调用工具。"
            "当你决定调用工具时，JSON 只能出现在 function call arguments 中，绝不能出现在你给用户的聊天回复中。"
            "用户在 Slack 中看到的最终回复必须是自然语言，不得是 JSON、代码块、函数参数或内部对象。"
            "你可以从图片中读取信息。"
            "如果信息明显不足以填写 Excel，也可以先提出一条简短的追问。"
            f"当前日本时间日期是 {today_jst.isoformat()}。如果用户说 今天/昨天/前天 等相对日期，请先换算成绝对日期再调用工具。"
            "对于交通费精算表：如果用户没有说明 purpose，默认用 営業活動。"
            "如果用户描述了一次移动但没有明确说往返或片道，默认按片道处理。"
            "如果用户说 电车、地铁、公交 等公共交通，都归类到 電車・バス。"
            "当用户发送图片、截图、发票、小票时，先调用 analyze_expense_evidence 来判断这是交通费还是个人报销，或者是否无法判断。"
            "对于交通卡履历截图，先识别 transport_events，再判断哪些事件属于可报销乘车，哪些只是 物販、定、窓出 或无法报销的原始记录。"
            "其中 定 不是自动排除项；它可能是定期区间中的进站或出站线索，需要结合相邻事件和跨图上下文判断。"
            "如果 analyze_expense_evidence 识别出了 transport_items，说明一张图里已经抽出了多条交通明细；这时优先直接生成包含多条 items 的交通费精算表。"
            "如果 transport_events 很多，但只有一部分能可靠形成 transport_items，就只生成那一部分，并在回复里明确说明哪些事件未纳入。"
            "如果 analyze_expense_evidence 判断为 transport，且已经抽到 travel_date、route_from、route_to、one_way_amount，就可以直接调用 generate_transport_sheet。"
            "如果 analyze_expense_evidence 判断为 transport，但缺少明确金额、路线或线路名，再调用 lookup_transport_route_options。"
            "如果 analyze_expense_evidence 判断为 personal_expense，请继续收集个人报销所需字段，再调用 generate_personal_expense_sheet。"
            "如果 analyze_expense_evidence 判断为 unknown，不要贸然生成表格，先向用户确认这是交通费还是个人报销。"
            "拿到交通候选后，优先用中文列出候选编号、路线、单程金额、时长和换乘次数，请用户回复编号确认。"
            "只有当用户已经明确确认某一个交通候选，或者自己明确给出金额和路线时，才调用 generate_transport_sheet。"
            "如果图片里金额清楚、站名清楚，优先信任图片识别出的金额，不必再查 Ekispert。"
            "如果图片和文字冲突，请先向用户确认，不要擅自决定。"
            "对于只有月日没有年份的交通截图，如果没有相反证据，优先按当前日本日期所在年份推断。"
            "当工具已经成功生成文件后，用简洁中文告诉用户文件已准备好，不要伪造下载链接。"
            "不要向用户暴露任何内部 JSON、函数参数、工具返回对象、文件路径或系统字段。"
        )

        if trace is not None:
            trace.write_section(
                "openai_request_1",
                {
                    "model": self.settings.openai_model,
                    "instructions": instructions,
                    "input": conversation_input,
                    "tools": self.tools,
                },
            )

        response = self.client.responses.create(
            model=self.settings.openai_model,
            instructions=instructions,
            input=conversation_input,
            tools=self.tools,
        )
        if trace is not None:
            trace.write_section("openai_response_1", response)
        generated_files: list[dict[str, Any]] = []

        for round_index in range(5):
            function_calls = [item for item in response.output if item.type == "function_call"]
            if not function_calls:
                return AgentResult(text=response.output_text.strip() or "好的，我来处理。", generated_files=generated_files)

            tool_outputs = []
            for call in function_calls:
                handler = self.handlers[call.name]
                arguments = json.loads(call.arguments)
                result = handler(arguments)
                if result.get("output_path"):
                    generated_files.append(result)
                if trace is not None:
                    trace.write_section(
                        f"tool_call_{round_index + 1}_{call.name}",
                        {
                            "call_id": call.call_id,
                            "arguments": arguments,
                            "result": result,
                        },
                    )
                tool_outputs.append(
                    {
                        "type": "function_call_output",
                        "call_id": call.call_id,
                        "output": json.dumps(self._tool_result_summary(result), ensure_ascii=False),
                    }
                )

            if trace is not None:
                trace.write_section(
                    f"openai_followup_request_{round_index + 2}",
                    {
                        "model": self.settings.openai_model,
                        "previous_response_id": response.id,
                        "input": tool_outputs,
                        "tools": self.tools,
                    },
                )

            response = self.client.responses.create(
                model=self.settings.openai_model,
                instructions=instructions,
                previous_response_id=response.id,
                input=tool_outputs,
                tools=self.tools,
            )
            if trace is not None:
                trace.write_section(f"openai_followup_response_{round_index + 2}", response)

        return AgentResult(
            text="我已经处理了请求，但本轮工具调用次数达到上限，请检查输入后重试。",
            generated_files=generated_files,
        )

    @staticmethod
    def _tool_result_summary(result: dict[str, Any]) -> dict[str, Any]:
        if result.get("output_path"):
            return {
                "ok": True,
                "title": result.get("title"),
                "message": "Excel 文件已生成，系统会自动上传到当前 Slack 会话。",
            }
        return result
