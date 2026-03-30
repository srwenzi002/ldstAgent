from __future__ import annotations

import base64
import logging
import mimetypes
import traceback
from typing import Any

import httpx
from slack_sdk.web.async_client import AsyncWebClient

from slack_excel_bot.debug_trace import DebugTrace
from slack_excel_bot.openai_agent import OpenAIExcelAgent


logger = logging.getLogger(__name__)


class SlackExcelBot:
    def __init__(
        self,
        slack_client: AsyncWebClient,
        agent: OpenAIExcelAgent,
        *,
        bot_user_id: str | None = None,
        bot_id: str | None = None,
    ):
        self.slack_client = slack_client
        self.agent = agent
        self.bot_user_id = bot_user_id
        self.bot_id = bot_id
        self.thread_contexts: dict[str, dict[str, Any]] = {}

    async def handle_socket_event(self, payload: dict[str, Any], event: dict[str, Any]) -> None:
        event_type = event.get("type")
        if event_type == "assistant_thread_started":
            await self.handle_assistant_thread_started(event, payload)
            return
        if event_type == "assistant_thread_context_changed":
            await self.handle_assistant_thread_context_changed(event, payload)
            return
        if event_type == "message":
            await self.handle_message_event(event, payload)

    async def handle_assistant_thread_started(self, event: dict[str, Any], payload: dict[str, Any]) -> None:
        thread_info = self._extract_thread_info(event)
        if not thread_info:
            logger.info("assistant_thread_started received but thread info could not be extracted")
            return
        channel, thread_ts = thread_info
        self.thread_contexts[thread_ts] = {"started_event": event, "payload": payload}
        logger.info("Assistant thread started channel=%s thread_ts=%s", channel, thread_ts)
        await self._safe_set_suggested_prompts(channel, thread_ts)

    async def handle_assistant_thread_context_changed(self, event: dict[str, Any], payload: dict[str, Any]) -> None:
        thread_info = self._extract_thread_info(event)
        if not thread_info:
            logger.info("assistant_thread_context_changed received but thread info could not be extracted")
            return
        channel, thread_ts = thread_info
        self.thread_contexts.setdefault(thread_ts, {}).update({"context_changed_event": event, "payload": payload})
        logger.info("Assistant thread context changed channel=%s thread_ts=%s", channel, thread_ts)

    async def handle_message_event(self, event: dict[str, Any], payload: dict[str, Any]) -> None:
        trace: DebugTrace | None = None
        try:
            if self._should_skip_message_event(event):
                return

            channel = event["channel"]
            ts = event["ts"]
            thread_ts = event.get("thread_ts") or ts
            assistant_action_token = (event.get("assistant_thread") or {}).get("action_token")
            session_id = assistant_action_token or thread_ts
            session_key = f"{event.get('user', 'unknown_user')}__{channel}__{session_id}"
            trace = DebugTrace(self.agent.settings.storage_dir, session_key=session_key, timestamp_key=ts)
            trace.write_section("socket_payload", payload)
            trace.write_section("slack_event", event)
            logger.info(
                "Handling Slack DM event channel=%s ts=%s thread_ts=%s subtype=%s",
                channel,
                ts,
                thread_ts,
                event.get("subtype"),
            )
            await self._safe_set_status(channel, thread_ts, "is thinking...", ["正在理解你的请求", "正在整理上下文"])
            conversation_input = await self._build_openai_input(event=event, trace=trace)
            trace.write_section("conversation_input", conversation_input)
            result = self.agent.run(conversation_input, trace=trace)
            trace.write_section("agent_result", {"text": result.text, "generated_files": result.generated_files})

            for item in result.generated_files:
                await self.slack_client.files_upload_v2(
                    channel=channel,
                    file=item["output_path"],
                    filename=item["title"] + ".xlsx",
                    title=item["title"],
                    thread_ts=thread_ts,
                )
                trace.write_section("slack_file_upload", item)

            await self.slack_client.chat_postMessage(
                channel=channel,
                text=result.text,
                thread_ts=thread_ts,
            )
            await self._safe_set_status(channel, thread_ts, "")
            await self._safe_set_title(channel, thread_ts, event.get("text") or "申請アシスト")
            trace.write_section(
                "slack_reply",
                {"channel": channel, "thread_ts": thread_ts, "text": result.text},
            )
        except Exception:
            if trace is not None:
                trace.write_section(
                    "error",
                    {
                        "type": "exception",
                        "traceback": traceback.format_exc(),
                    },
                )
            logger.exception("Failed to handle Slack event")
            raise

    async def _build_openai_input(self, event: dict[str, Any], trace: DebugTrace | None = None) -> list[dict[str, Any]]:
        channel = event["channel"]
        latest_ts = event["ts"]
        thread_ts = event.get("thread_ts")

        messages: list[dict[str, Any]] = []
        if thread_ts and thread_ts != latest_ts:
            replies = await self.slack_client.conversations_replies(
                channel=channel,
                ts=thread_ts,
                latest=latest_ts,
                inclusive=True,
                limit=20,
            )
            messages = sorted(replies.get("messages", []), key=lambda item: float(item["ts"]))
            if trace is not None:
                trace.write_section("slack_thread_replies", messages)
        else:
            messages = [event]
            if trace is not None:
                trace.write_section("slack_thread_replies", messages)

        messages = self._filter_context_messages(messages=messages, current_event=event)
        openai_messages: list[dict[str, Any]] = []

        for index, message in enumerate(messages):
            role = "assistant" if message.get("bot_id") or message.get("subtype") == "bot_message" else "user"
            content: list[dict[str, Any]] = []
            text = (message.get("text") or "").strip()
            if text:
                if role == "assistant":
                    content.append({"type": "output_text", "text": text})
                else:
                    content.append({"type": "input_text", "text": text})
            if role == "user":
                content.extend(await self._load_image_inputs(message.get("files", [])))
            if content:
                openai_messages.append({"role": role, "content": content})

        if not openai_messages:
            openai_messages.append(
                {
                    "role": "user",
                    "content": [{"type": "input_text", "text": (event.get("text") or "").strip()}],
                }
            )
        return openai_messages

    async def _load_image_inputs(self, files: list[dict[str, Any]]) -> list[dict[str, Any]]:
        image_inputs: list[dict[str, Any]] = []
        for file_info in files:
            mimetype = file_info.get("mimetype") or ""
            if not mimetype.startswith("image/"):
                continue
            private_url = file_info.get("url_private_download") or file_info.get("url_private")
            if not private_url:
                continue
            data_url = await self._download_as_data_url(private_url, mimetype)
            image_inputs.append({"type": "input_image", "image_url": data_url})
        return image_inputs

    async def _download_as_data_url(self, url: str, mimetype: str) -> str:
        headers = {"Authorization": f"Bearer {self.slack_client.token}"}
        async with httpx.AsyncClient(timeout=60.0) as client:
            response = await client.get(url, headers=headers)
            response.raise_for_status()
        media_type = mimetype or mimetypes.guess_type(url)[0] or "application/octet-stream"
        encoded = base64.b64encode(response.content).decode("ascii")
        return f"data:{media_type};base64,{encoded}"

    async def _safe_set_status(
        self,
        channel: str,
        thread_ts: str,
        status: str,
        loading_messages: list[str] | None = None,
    ) -> None:
        try:
            await self.slack_client.assistant_threads_setStatus(
                channel_id=channel,
                thread_ts=thread_ts,
                status=status,
                loading_messages=loading_messages,
            )
        except Exception:
            logger.exception("Failed to set assistant thread status channel=%s thread_ts=%s", channel, thread_ts)

    async def _safe_set_title(self, channel: str, thread_ts: str, raw_title: str) -> None:
        title = (raw_title or "").strip()[:100]
        if not title:
            return
        try:
            await self.slack_client.assistant_threads_setTitle(
                channel_id=channel,
                thread_ts=thread_ts,
                title=title,
            )
        except Exception:
            logger.exception("Failed to set assistant thread title channel=%s thread_ts=%s", channel, thread_ts)

    async def _safe_set_suggested_prompts(self, channel: str, thread_ts: str) -> None:
        prompts = [
            {"title": "做三月全勤表", "message": "帮我做一个2026年3月全勤的考勤表"},
            {"title": "做交通费表", "message": "我会发交通记录截图，帮我生成交通费精算表"},
            {"title": "做个人报销表", "message": "帮我做个人立替经费精算表"},
        ]
        try:
            await self.slack_client.assistant_threads_setSuggestedPrompts(
                channel_id=channel,
                thread_ts=thread_ts,
                prompts=prompts,
            )
        except Exception:
            logger.exception("Failed to set suggested prompts channel=%s thread_ts=%s", channel, thread_ts)

    @staticmethod
    def _filter_context_messages(messages: list[dict[str, Any]], current_event: dict[str, Any]) -> list[dict[str, Any]]:
        filtered: list[dict[str, Any]] = []
        current_ts = current_event.get("ts")

        for message in messages:
            text = (message.get("text") or "").strip()
            subtype = message.get("subtype")

            if subtype in {"assistant_app_thread", "message_deleted", "message_changed"}:
                continue
            if text == "新しいアシスタントスレッド":
                continue
            if not text and not message.get("files"):
                continue
            filtered.append(message)

        # If Slack replies API didn't include the current user message for some reason,
        # append the event explicitly so the model always sees the fresh input.
        if current_ts and all(item.get("ts") != current_ts for item in filtered):
            filtered.append(current_event)

        return sorted(filtered, key=lambda item: float(item["ts"]))

    @staticmethod
    def _extract_thread_info(event: dict[str, Any]) -> tuple[str, str] | None:
        channel = (
            event.get("channel")
            or event.get("channel_id")
            or (event.get("assistant_thread") or {}).get("channel_id")
            or (event.get("assistant_thread") or {}).get("channel")
        )
        thread_ts = (
            event.get("thread_ts")
            or (event.get("assistant_thread") or {}).get("thread_ts")
            or (event.get("assistant_thread") or {}).get("ts")
        )
        if channel and thread_ts:
            return channel, thread_ts
        return None

    def _should_skip_message_event(self, event: dict[str, Any]) -> bool:
        if event.get("channel_type") != "im":
            return True
        if event.get("subtype") in {
            "bot_message",
            "message_changed",
            "message_deleted",
            "assistant_app_thread",
            "channel_join",
        }:
            return True
        if event.get("bot_id") == self.bot_id:
            logger.info("Skipping self bot_id event channel=%s ts=%s", event.get("channel"), event.get("ts"))
            return True
        if event.get("user") == self.bot_user_id:
            logger.info("Skipping self user event channel=%s ts=%s", event.get("channel"), event.get("ts"))
            return True
        if not event.get("user"):
            return True
        if not (event.get("text") or event.get("files")):
            return True
        return False
