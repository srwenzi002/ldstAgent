from __future__ import annotations

import asyncio
import logging

from slack_sdk.web.async_client import AsyncWebClient

from slack_excel_bot.config import Settings
from slack_excel_bot.excel_tools import ExcelToolService
from slack_excel_bot.openai_agent import OpenAIExcelAgent
from slack_excel_bot.slack_bot import SlackExcelBot
from slack_excel_bot.socket_mode import SlackSocketModeRunner


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


async def run() -> None:
    settings = Settings.from_env()
    settings.validate_runtime()

    slack_client = AsyncWebClient(token=settings.slack_bot_token or None)
    tool_service = ExcelToolService(settings)
    agent = OpenAIExcelAgent(settings, tool_service)
    bot = SlackExcelBot(slack_client, agent)
    socket_mode_runner = SlackSocketModeRunner(settings.slack_app_token, bot)

    logger.info("Starting Slack Excel Bot in Socket Mode")
    await socket_mode_runner.connect()
    logger.info("Socket Mode connected")

    try:
        await asyncio.Event().wait()
    finally:
        await socket_mode_runner.close()


def main() -> None:
    asyncio.run(run())


if __name__ == "__main__":
    main()
