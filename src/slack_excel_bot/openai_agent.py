from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any, Callable

from openai import OpenAI

from slack_excel_bot.config import Settings
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
                "生成考勤表。适用于用户要求全勤表、出勤表、勤务表或月度考勤。",
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

    def run(self, conversation_input: list[dict[str, Any]]) -> AgentResult:
        instructions = (
            "你是一个 Slack 助手。"
            "你需要根据用户在私聊中的文字和图片，判断是否需要生成 Excel。"
            "如果只是寒暄或普通问题，直接自然回复。"
            "如果用户要做考勤表、交通费精算表、个人报销计算表，请优先调用工具。"
            "你可以从图片中读取信息。"
            "如果信息明显不足以填写 Excel，也可以先提出一条简短的追问。"
            "当工具已经成功生成文件后，用简洁中文告诉用户文件已准备好，不要伪造下载链接。"
        )

        response = self.client.responses.create(
            model=self.settings.openai_model,
            instructions=instructions,
            input=conversation_input,
            tools=self.tools,
        )
        generated_files: list[dict[str, Any]] = []

        for _ in range(5):
            function_calls = [item for item in response.output if item.type == "function_call"]
            if not function_calls:
                return AgentResult(text=response.output_text.strip() or "好的，我来处理。", generated_files=generated_files)

            tool_outputs = []
            for call in function_calls:
                handler = self.handlers[call.name]
                arguments = json.loads(call.arguments)
                result = handler(arguments)
                generated_files.append(result)
                tool_outputs.append(
                    {
                        "type": "function_call_output",
                        "call_id": call.call_id,
                        "output": json.dumps(result, ensure_ascii=False),
                    }
                )

            response = self.client.responses.create(
                model=self.settings.openai_model,
                previous_response_id=response.id,
                input=tool_outputs,
                tools=self.tools,
            )

        return AgentResult(
            text="我已经处理了请求，但本轮工具调用次数达到上限，请检查输入后重试。",
            generated_files=generated_files,
        )
