from __future__ import annotations

import json
from pathlib import Path

import yaml

from slack_excel_bot.template_schema import TemplateMapping, TemplateRegistryEntry


class TemplateLoaderError(Exception):
    def __init__(self, code: str, message: str, details: dict | None = None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.details = details or {}


def load_registry(registry_path: Path) -> list[TemplateRegistryEntry]:
    if not registry_path.exists():
        raise TemplateLoaderError(
            code="REGISTRY_NOT_FOUND",
            message="Template registry file not found.",
            details={"registry_path": str(registry_path)},
        )
    raw = json.loads(registry_path.read_text(encoding="utf-8"))
    return [TemplateRegistryEntry.model_validate(item) for item in raw]


def load_registry_entry(template_id: str, registry_path: Path) -> TemplateRegistryEntry:
    for entry in load_registry(registry_path):
        if entry.template_id == template_id:
            return entry
    raise TemplateLoaderError(
        code="TEMPLATE_NOT_FOUND",
        message="Template ID is not registered.",
        details={"template_id": template_id},
    )


def load_mapping(mapping_path: Path) -> TemplateMapping:
    if not mapping_path.exists():
        raise TemplateLoaderError(
            code="MAPPING_NOT_FOUND",
            message="Template mapping file not found.",
            details={"mapping_path": str(mapping_path)},
        )
    raw = yaml.safe_load(mapping_path.read_text(encoding="utf-8"))
    return TemplateMapping.model_validate(raw)
