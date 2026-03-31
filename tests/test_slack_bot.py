import asyncio
from pathlib import Path

from slack_excel_bot.config import Settings
from slack_excel_bot.excel_tools import ExcelToolService
from slack_excel_bot.openai_agent import OpenAIExcelAgent
from slack_excel_bot.slack_bot import SlackExcelBot


def build_settings(tmp_path: Path) -> Settings:
    return Settings(
        slack_bot_token="xoxb-test",
        slack_app_token="xapp-test",
        openai_api_key="sk-test",
        openai_model="gpt-4.1-mini",
        ekispert_api_token="ek-test",
        port=3000,
        storage_dir=tmp_path,
        default_employee_name="山田太郎",
        default_employee_id="0001",
        default_department="開発本部",
        default_department_code="50",
        default_work_grade=1,
        default_clock_in="09:00",
        default_clock_out="18:00",
        max_concurrent_requests=50,
    )


def build_bot(tmp_path: Path) -> SlackExcelBot:
    settings = build_settings(tmp_path)
    tool_service = ExcelToolService(settings)
    agent = OpenAIExcelAgent(settings, tool_service)
    return SlackExcelBot(slack_client=None, agent=agent, bot_user_id="UBOT", bot_id="BBOT")  # type: ignore[arg-type]


def test_thread_lock_is_scoped_by_thread_ts(tmp_path: Path) -> None:
    bot = build_bot(tmp_path)

    first = bot._get_thread_lock("thread-1")
    second = bot._get_thread_lock("thread-1")
    third = bot._get_thread_lock("thread-2")

    assert first is second
    assert first is not third


def test_handle_socket_event_schedules_message_in_background(tmp_path: Path) -> None:
    async def scenario() -> None:
        bot = build_bot(tmp_path)
        scheduled: list[str] = []

        async def fake_handle(event, payload):
            scheduled.append(f"{event['ts']}:{payload['event']['type']}")

        bot._handle_message_event_with_limits = fake_handle  # type: ignore[method-assign]

        payload = {"event": {"type": "message"}}
        event = {
            "type": "message",
            "channel_type": "im",
            "user": "U123",
            "channel": "D123",
            "ts": "123.456",
            "text": "hello",
        }

        await bot.handle_socket_event(payload, event)
        await asyncio.sleep(0)

        assert scheduled == ["123.456:message"]

    asyncio.run(scenario())


def test_should_not_skip_image_file_share_message(tmp_path: Path) -> None:
    bot = build_bot(tmp_path)
    event = {
        "type": "message",
        "channel_type": "im",
        "subtype": "file_share",
        "user": "U123",
        "channel": "D123",
        "ts": "123.456",
        "files": [{"mimetype": "image/png", "url_private": "https://example.com/image.png"}],
        "text": "",
    }

    assert bot._should_skip_message_event(event) is False


def test_should_skip_non_image_empty_message(tmp_path: Path) -> None:
    bot = build_bot(tmp_path)
    event = {
        "type": "message",
        "channel_type": "im",
        "user": "U123",
        "channel": "D123",
        "ts": "123.456",
        "text": "",
        "files": [],
    }

    assert bot._should_skip_message_event(event) is True


def test_filter_context_keeps_older_user_file_messages(tmp_path: Path) -> None:
    bot = build_bot(tmp_path)
    current_event = {
        "type": "message",
        "channel_type": "im",
        "user": "U123",
        "channel": "D123",
        "ts": "200.000",
        "text": "请继续处理",
        "files": [],
    }
    messages = [
        {
            "type": "message",
            "channel_type": "im",
            "subtype": "file_share",
            "user": "U123",
            "channel": "D123",
            "ts": "100.000",
            "text": "",
            "files": [{"mimetype": "image/png", "url_private": "https://example.com/1.png"}],
        },
        current_event,
    ]

    filtered = bot._filter_context_messages(messages=messages, current_event=current_event)

    assert len(filtered) == 2
    assert filtered[0]["files"][0]["mimetype"] == "image/png"


def test_agent_status_mapping_for_route_lookup(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    tool_service = ExcelToolService(settings)
    agent = OpenAIExcelAgent(settings, tool_service)

    status, messages = agent._status_for_tool_name("lookup_transport_route_batch")

    assert status == "is checking routes..."
    assert messages == ["🚃 経路と運賃を確認中です", "💴 金額を照合しています"]


def test_agent_status_mapping_for_attendance_generation(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    tool_service = ExcelToolService(settings)
    agent = OpenAIExcelAgent(settings, tool_service)

    status, messages = agent._status_for_tool_name("generate_attendance_sheet")

    assert status == "is generating Excel..."
    assert messages == ["📅 勤務データを整理中です", "📎 Excel を仕上げています"]
