import asyncio
from pathlib import Path

from slack_excel_bot.config import Settings
from slack_excel_bot.excel_tools import ExcelToolService
from slack_excel_bot.openai_agent import OpenAIExcelAgent
from slack_excel_bot.slack_bot import SlackExcelBot


class FakeSlackClient:
    def __init__(self) -> None:
        self.published_views: list[dict[str, object]] = []

    async def views_publish(self, *, user_id: str, view: dict[str, object]) -> None:
        self.published_views.append({"user_id": user_id, "view": view})


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


def test_handle_socket_event_publishes_home_view(tmp_path: Path) -> None:
    async def scenario() -> None:
        settings = build_settings(tmp_path)
        tool_service = ExcelToolService(settings)
        agent = OpenAIExcelAgent(settings, tool_service)
        slack_client = FakeSlackClient()
        bot = SlackExcelBot(slack_client=slack_client, agent=agent, bot_user_id="UBOT", bot_id="BBOT")

        await bot.handle_socket_event(payload={"event": {"type": "app_home_opened"}}, event={"type": "app_home_opened", "user": "U123"})

        assert len(slack_client.published_views) == 1
        published = slack_client.published_views[0]
        assert published["user_id"] == "U123"
        view = published["view"]
        assert isinstance(view, dict)
        assert view["type"] == "home"
        blocks = view["blocks"]
        assert isinstance(blocks, list)
        page_text = "\n".join(
            block["text"]["text"]
            for block in blocks
            if isinstance(block, dict) and isinstance(block.get("text"), dict) and "text" in block["text"]
        )
        assert "技術スタック" in page_text
        assert "実装済み機能" in page_text
        assert "精算くん" in page_text
        assert "v0.4.0" in page_text
        assert "v0.3.0" in page_text

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

    assert status == "is working..."
    assert messages == ["STEP 2/4 経路と運賃を照会しています"]


def test_agent_status_mapping_for_attendance_generation(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    tool_service = ExcelToolService(settings)
    agent = OpenAIExcelAgent(settings, tool_service)

    status, messages = agent._status_for_tool_name("generate_attendance_sheet")

    assert status == "is working..."
    assert messages == ["STEP 4/4 勤務表を作成しています"]


def test_auto_generate_ready_transport_draft_uses_full_state(tmp_path: Path) -> None:
    bot = build_bot(tmp_path)
    token = bot.agent.tool_service.start_draft_run("thread-1", bot.thread_store)
    try:
        bot.agent.tool_service.upsert_transport_draft(
            {
                "mode": "merge",
                "employee": {
                    "department": "開発本部",
                    "department_code": "50",
                    "employee_id": "0001",
                    "name": "山田太郎",
                },
                "items": [
                    {
                        "travel_date": "2026-03-10",
                        "purpose": "営業活動",
                        "transport_mode": "電車・バス",
                        "route_from": "新宿",
                        "route_to": "渋谷",
                        "route_line": "JR山手線",
                        "one_way_amount": 178,
                        "is_round_trip": False,
                    }
                ],
                "pending_questions": [],
            }
        )
    finally:
        run_state = bot.agent.tool_service.finish_draft_run(token)

    assert run_state is not None
    generated = bot._auto_generate_ready_drafts(
        thread_ts="thread-1",
        already_generated_files=[],
        updated_templates=run_state.updated_templates,
    )

    assert len(generated) == 1
    assert generated[0]["template_id"] == "transport_jp_leadingsoft_v1"
    draft = bot.thread_store.get_draft("thread-1", "transport")
    assert draft["latest_generated_file"]["title"] == generated[0]["title"]


def test_agent_limits_route_retry_chain_after_second_query_error(tmp_path: Path, monkeypatch) -> None:
    settings = build_settings(tmp_path)
    tool_service = ExcelToolService(settings)
    agent = OpenAIExcelAgent(settings, tool_service)

    class StubResponse:
        def __init__(self, response_id, output, output_text=""):
            self.id = response_id
            self.output = output
            self.output_text = output_text

    class FakeResponses:
        def __init__(self):
            self.calls = 0

        def create(self, **kwargs):
            self.calls += 1
            if self.calls == 1:
                return StubResponse(
                    "resp-1",
                    [
                        type(
                            "FunctionCall",
                            (),
                            {
                                "type": "function_call",
                                "name": "lookup_transport_route_batch",
                                "arguments": '{"items":[{"travel_date":"2026-04-06","route_from":"早稲田","route_to":"日本橋","one_way_amount":178,"route_line":null}],"top_k":3}',
                                "call_id": "call-1",
                            },
                        )()
                    ],
                )
            if self.calls == 2:
                return StubResponse(
                    "resp-2",
                    [
                        type(
                            "FunctionCall",
                            (),
                            {
                                "type": "function_call",
                                "name": "lookup_station_candidates",
                                "arguments": '{"station_name":"日本橋","top_k":5,"prefecture_code":"13","match_type":"partial","station_type":"train"}',
                                "call_id": "call-2",
                            },
                        )()
                    ],
                )
            if self.calls == 3:
                return StubResponse(
                    "resp-3",
                    [
                        type(
                            "FunctionCall",
                            (),
                            {
                                "type": "function_call",
                                "name": "lookup_transport_route_batch",
                                "arguments": '{"items":[{"travel_date":"2026-04-06","route_from":"早稲田(東京メトロ)","route_to":"日本橋","one_way_amount":178,"route_line":null}],"top_k":3}',
                                "call_id": "call-3",
                            },
                        )()
                    ],
                )
            return StubResponse("resp-4", [], "候補が絞れないため、駅名確認に切り替えます。")

    class FakeClient:
        def __init__(self):
            self.responses = FakeResponses()

    monkeypatch.setattr("slack_excel_bot.openai_agent.OpenAI", lambda api_key: FakeClient())

    def fake_lookup_transport_route_batch(args):
        route_from = args["items"][0]["route_from"]
        return {
            "ok": False,
            "title": "交通経路の一括確認結果",
            "has_partial_failures": True,
            "resolved_items": [],
            "round_trip_suggestions": [],
            "needs_confirmation": [],
            "items": [],
            "prompt_reason": "query_error",
            "error": f'{{"status":400,"message":"駅名が見つかりません。({route_from})"}}',
        }

    def fake_lookup_station_candidates(args):
        return {
            "ok": True,
            "title": "駅名候補",
            "station_name": args["station_name"],
            "candidates": [{"station_name": "日本橋(東京都)"}],
        }

    agent.handlers["lookup_transport_route_batch"] = fake_lookup_transport_route_batch
    agent.handlers["lookup_station_candidates"] = fake_lookup_station_candidates

    result = agent.run(
        conversation_input=[{"role": "user", "content": [{"type": "input_text", "text": "交通履歴を確認して"}]}]
    )

    assert "駅名確認" in result.text
