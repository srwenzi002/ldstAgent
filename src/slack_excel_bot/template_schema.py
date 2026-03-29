from __future__ import annotations

from enum import Enum
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field


class MissingBehavior(str, Enum):
    BLANK = "blank"
    ERROR = "error"


class SingleFieldMapping(BaseModel):
    model_config = ConfigDict(extra="forbid")

    field_path: str = Field(min_length=1)
    cell: str = Field(min_length=2)
    missing: MissingBehavior | None = None
    allowed_values: list[Any] = Field(default_factory=list)
    value_type: Literal["auto", "string", "int", "float", "date", "time", "datetime"] = "auto"


class ConstantMapping(BaseModel):
    model_config = ConfigDict(extra="forbid")

    cell: str = Field(min_length=2)
    value: Any = None


class ItemFieldMapping(BaseModel):
    model_config = ConfigDict(extra="forbid")

    column: str = Field(min_length=1)
    row_offset: int = 0
    missing: MissingBehavior | None = None
    true_value: Any = None
    false_value: Any = None
    allowed_values: list[Any] = Field(default_factory=list)
    value_type: Literal["auto", "string", "int", "float", "date", "time", "datetime"] = "auto"


class ItemsMapping(BaseModel):
    model_config = ConfigDict(extra="forbid")

    path: str = Field(default="items", min_length=1)
    start_row: int = Field(ge=1)
    max_rows: int = Field(ge=1)
    row_stride: int = Field(default=1, ge=1)
    columns: dict[str, str] = Field(default_factory=dict)
    fields: dict[str, ItemFieldMapping] = Field(default_factory=dict)
    row_index_mode: Literal["sequence", "day_of_month"] = "sequence"
    row_index_field: str = Field(default="day", min_length=1)
    missing: MissingBehavior | None = None


class MissingStrategy(BaseModel):
    model_config = ConfigDict(extra="forbid")

    default: MissingBehavior = MissingBehavior.ERROR
    fields: dict[str, MissingBehavior] = Field(default_factory=dict)


class TemplateMapping(BaseModel):
    model_config = ConfigDict(extra="forbid")

    template_id: str = Field(min_length=1)
    sheet: str = Field(min_length=1)
    single_fields: list[SingleFieldMapping] = Field(default_factory=list)
    constants: list[ConstantMapping] = Field(default_factory=list)
    items: ItemsMapping | None = None
    missing_strategy: MissingStrategy = Field(default_factory=MissingStrategy)


class TemplateRegistryEntry(BaseModel):
    model_config = ConfigDict(extra="forbid")

    template_id: str = Field(min_length=1)
    name: str = Field(min_length=1)
    claim_type: str = Field(min_length=1)
    company: str = Field(min_length=1)
    version: str = Field(min_length=1)
    enabled: bool = True
    file_path: str = Field(min_length=1)
    mapping_path: str = Field(min_length=1)
