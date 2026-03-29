from __future__ import annotations

import logging

from fastapi import BackgroundTasks, FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from slack_sdk.signature import SignatureVerifier
from slack_sdk.web.async_client import AsyncWebClient

from slack_excel_bot.config import Settings
from slack_excel_bot.excel_tools import ExcelToolService
from slack_excel_bot.openai_agent import OpenAIExcelAgent
from slack_excel_bot.socket_mode import SlackSocketModeRunner
from slack_excel_bot.slack_bot import SlackExcelBot


logging.basicConfig(level=logging.INFO)

settings = Settings.from_env()
slack_client = AsyncWebClient(token=settings.slack_bot_token or None)
tool_service = ExcelToolService(settings)
agent = OpenAIExcelAgent(settings, tool_service)
bot = SlackExcelBot(slack_client, agent)
signature_verifier = SignatureVerifier(signing_secret=settings.slack_signing_secret or "")
socket_mode_runner = SlackSocketModeRunner(settings.slack_app_token, bot) if settings.slack_app_token else None
app = FastAPI(title="Slack Excel Bot")


@app.on_event("startup")
async def startup() -> None:
    settings.validate_runtime()
    if socket_mode_runner is not None:
        await socket_mode_runner.connect()


@app.on_event("shutdown")
async def shutdown() -> None:
    if socket_mode_runner is not None:
        await socket_mode_runner.close()


@app.get("/healthz")
async def healthz() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/slack/events")
async def slack_events(request: Request, background_tasks: BackgroundTasks) -> JSONResponse:
    if not settings.slack_signing_secret:
        raise HTTPException(status_code=400, detail="HTTP events mode is disabled because SLACK_SIGNING_SECRET is not set")

    body = await request.body()
    if not signature_verifier.is_valid_request(body, request.headers):
        raise HTTPException(status_code=401, detail="invalid slack signature")

    payload = await request.json()
    if payload.get("type") == "url_verification":
        return JSONResponse({"challenge": payload["challenge"]})

    if payload.get("type") == "event_callback":
        event = payload.get("event", {})
        if event.get("type") == "message":
            background_tasks.add_task(bot.handle_event, event)
    return JSONResponse({"ok": True})
