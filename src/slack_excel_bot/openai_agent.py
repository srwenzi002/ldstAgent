from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any, Callable

from openai import OpenAI

from slack_excel_bot.config import Settings
from slack_excel_bot.debug_trace import DebugTrace
from slack_excel_bot.excel_tools import ExcelToolService
from slack_excel_bot.tool_schemas import (
    AttendanceSheetInput,
    PersonalExpenseSheetInput,
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
                "生成日本月度考勤表。你必须直接输出接近最终Excel的数据。employee.department_code 只能是 10/20/30/50/51/52/60/70。days[].work_grade 只能是 1/2/3/4，其中 1=09:30-18:00, 2=09:00-17:30, 3=10:00-18:30, 4=10:30-19:00。days[].leave_item_no 只能是 1..15，其中 1=有給休暇(全日), 2=有給休暇(午前), 3=有給休暇(午後), 4=欠勤, 5=健診BC, 6=無給休暇, 7=振休, 8=代休, 9=特別代休, 10=結忌引配出産, 11=SP5(GW・夏季), 12=その他特休, 13=積立休暇, 14=休業, 15=教育訓練。不要输出 schema 之外的字段。",
                AttendanceSheetInput,
            ),
            openai_function_tool(
                "generate_transport_sheet",
                "生成交通费精算表。适用于用户提供交通记录、车费截图、通勤报销信息。",
                TransportSheetInput,
            ),
            openai_function_tool(
                "generate_personal_expense_sheet",
                "生成个人报销计算表。适用于餐费、招待费、采购等个人代垫报销。",
                PersonalExpenseSheetInput,
            ),
        ]
        self.handlers: dict[str, Callable[[dict[str, Any]], dict[str, Any]]] = {
            "generate_attendance_sheet": self.tool_service.generate_attendance_sheet,
            "generate_transport_sheet": self.tool_service.generate_transport_sheet,
            "generate_personal_expense_sheet": self.tool_service.generate_personal_expense_sheet,
        }

    def run(self, conversation_input: list[dict[str, Any]], trace: DebugTrace | None = None) -> AgentResult:
        instructions = (
            "你是一个 Slack 助手。"
            "你需要根据用户在私聊中的文字和图片，判断是否需要生成 Excel。"
            "如果只是寒暄或普通问题，直接自然回复。"
            "如果用户要做考勤表、交通费精算表、个人报销计算表，请优先调用工具。"
            "你可以从图片中读取信息。"
            "如果信息明显不足以填写 Excel，也可以先提出一条简短的追问。"
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
                return AgentResult(
                    text=self._final_user_text(response.output_text, generated_files),
                    generated_files=generated_files,
                )

            tool_outputs = []
            for call in function_calls:
                handler = self.handlers[call.name]
                arguments = json.loads(call.arguments)
                result = handler(arguments)
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
        return {
            "ok": True,
            "title": result.get("title"),
            "message": "Excel 文件已生成，系统会自动上传到当前 Slack 会话。",
        }

    @staticmethod
    def _looks_like_json(text: str) -> bool:
        stripped = text.strip()
        if not stripped:
            return False
        if stripped.startswith("```json") or stripped.startswith("```"):
            return True
        if stripped.startswith("{") or stripped.startswith("["):
            return True
        suspicious_tokens = ['"output_path"', '"generated_files"', '"leave_item_no"', '"department_code"']
        return any(token in stripped for token in suspicious_tokens)

    def _final_user_text(self, output_text: str, generated_files: list[dict[str, Any]]) -> str:
        text = (output_text or "").strip()
        if generated_files:
            if not text or self._looks_like_json(text):
                return "好的，文件已经准备好了，我这就发到当前会话里。"
            return text
        return text or "好的，我来处理。"
