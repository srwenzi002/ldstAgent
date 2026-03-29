from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


@dataclass(frozen=True)
class Settings:
    slack_bot_token: str
    slack_app_token: str
    openai_api_key: str
    openai_model: str
    port: int
    storage_dir: Path
    default_employee_name: str
    default_employee_id: str
    default_department: str
    default_department_code: str
    default_work_grade: int
    default_clock_in: str
    default_clock_out: str

    @classmethod
    def from_env(cls) -> "Settings":
        load_dotenv()
        storage_dir = Path(os.getenv("STORAGE_DIR", ".data")).resolve()
        storage_dir.mkdir(parents=True, exist_ok=True)
        (storage_dir / "drafts").mkdir(parents=True, exist_ok=True)

        return cls(
            slack_bot_token=os.getenv("SLACK_BOT_TOKEN", "").strip(),
            slack_app_token=os.getenv("SLACK_APP_TOKEN", "").strip(),
            openai_api_key=(os.getenv("OPENAI_API_KEY") or os.getenv("EXPENSES_LLM_API_KEY") or "").strip(),
            openai_model=(os.getenv("OPENAI_MODEL") or os.getenv("EXPENSES_LLM_MODEL") or "gpt-4.1-mini").strip(),
            port=int(os.getenv("PORT", "3000")),
            storage_dir=storage_dir,
            default_employee_name=os.getenv("DEFAULT_EMPLOYEE_NAME", "氏名未設定").strip(),
            default_employee_id=os.getenv("DEFAULT_EMPLOYEE_ID", "社員番号未設定").strip(),
            default_department=os.getenv("DEFAULT_DEPARTMENT", "開発本部").strip(),
            default_department_code=os.getenv("DEFAULT_DEPARTMENT_CODE", "50").strip(),
            default_work_grade=int(os.getenv("DEFAULT_WORK_GRADE", "1")),
            default_clock_in=os.getenv("DEFAULT_CLOCK_IN", "09:00").strip(),
            default_clock_out=os.getenv("DEFAULT_CLOCK_OUT", "18:00").strip(),
        )

    def validate_runtime(self) -> None:
        missing = []
        if not self.slack_bot_token:
            missing.append("SLACK_BOT_TOKEN")
        if not self.openai_api_key:
            missing.append("OPENAI_API_KEY or EXPENSES_LLM_API_KEY")
        if not self.slack_app_token:
            missing.append("SLACK_APP_TOKEN")
        if missing:
            raise RuntimeError(f"Missing required environment variables: {', '.join(missing)}")
