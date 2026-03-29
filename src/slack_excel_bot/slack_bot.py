from __future__ import annotations

import base64
import logging
import mimetypes
from typing import Any

import httpx
from slack_sdk.web.async_client import AsyncWebClient

from slack_excel_bot.openai_agent import OpenAIExcelAgent


logger = logging.getLogger(__name__)


class SlackExcelBot:
    def __init__(self, slack_client: AsyncWebClient, agent: OpenAIExcelAgent):
        self.slack_client = slack_client
        self.agent = agent

    async def handle_event(self, event: dict[str, Any]) -> None:
        try:
            if event.get("channel_type") != "im":
                return
            if event.get("subtype") in {"bot_message", "message_changed", "message_deleted"}:
                return
            if not event.get("user"):
                return

            channel = event["channel"]
            ts = event["ts"]
            logger.info("Handling Slack DM event channel=%s ts=%s", channel, ts)
            conversation_input = await self._build_openai_input(channel=channel, latest_ts=ts)
            result = self.agent.run(conversation_input)

            for item in result.generated_files:
                await self.slack_client.files_upload_v2(
                    channel=channel,
                    file=item["output_path"],
                    filename=item["title"] + ".xlsx",
                    title=item["title"],
                )

            await self.slack_client.chat_postMessage(
                channel=channel,
                text=result.text,
            )
        except Exception:
            logger.exception("Failed to handle Slack event")
            raise

    async def _build_openai_input(self, channel: str, latest_ts: str) -> list[dict[str, Any]]:
        history = await self.slack_client.conversations_history(
            channel=channel,
            inclusive=True,
            latest=latest_ts,
            limit=6,
        )
        messages = sorted(history.get("messages", []), key=lambda item: float(item["ts"]))
        openai_messages: list[dict[str, Any]] = []

        for index, message in enumerate(messages):
            role = "assistant" if message.get("bot_id") or message.get("subtype") == "bot_message" else "user"
            content: list[dict[str, Any]] = []
            text = (message.get("text") or "").strip()
            if text:
                content.append({"type": "input_text", "text": text})
            if index == len(messages) - 1:
                content.extend(await self._load_image_inputs(message.get("files", [])))
            if content:
                openai_messages.append({"role": role, "content": content})
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
