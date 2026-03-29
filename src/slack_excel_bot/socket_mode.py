from __future__ import annotations

from typing import Any

from slack_sdk.socket_mode.aiohttp import SocketModeClient
from slack_sdk.socket_mode.request import SocketModeRequest
from slack_sdk.socket_mode.response import SocketModeResponse

from slack_excel_bot.slack_bot import SlackExcelBot


class SlackSocketModeRunner:
    def __init__(self, app_token: str, bot: SlackExcelBot):
        self.bot = bot
        self.client = SocketModeClient(app_token=app_token, web_client=bot.slack_client)
        self.client.socket_mode_request_listeners.append(self.process)

    async def connect(self) -> None:
        await self.client.connect()

    async def close(self) -> None:
        await self.client.close()

    async def process(self, client: SocketModeClient, request: SocketModeRequest) -> None:
        if request.type != "events_api":
            return

        await client.send_socket_mode_response(SocketModeResponse(envelope_id=request.envelope_id))
        payload: dict[str, Any] = request.payload or {}
        event = payload.get("event", {})
        await self.bot.handle_socket_event(payload=payload, event=event)
