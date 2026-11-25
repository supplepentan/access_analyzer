"""Database, table, field, index, and relationship data models."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field


class VbaReferenceInfo(BaseModel):
    name: str
    description: str = ""
    full_path: str = ""
    guid: str = ""
    major: int = 0
    minor: int = 0
    is_broken: bool = False
    built_in: bool = False


class DatabaseMetadata(BaseModel):
    file_path: str
    file_name: str
    file_size_bytes: int = 0
    access_version: str = ""
    dao_version: str = ""
    startup_form: str | None = None
    app_title: str | None = None
    allow_bypass_key: bool | None = None
    properties: dict[str, Any] = Field(default_factory=dict)
    vba_references: list[VbaReferenceInfo] = Field(default_factory=list)
    linked_databases: list[str] = Field(default_factory=list)
    extracted_at: str = ""


class LookupInfo(BaseModel):
    display_control: str = ""
    row_source_type: str = ""
    row_source: str = ""
    bound_column: int | None = None
    column_count: int | None = None
    column_widths: str = ""
    limit_to_list: bool | None = None
    allow_multiple_values: bool = False


class FieldDataProfile(BaseModel):
    non_null_count: int = 0
    null_count: int = 0
    null_ratio: float = 0.0
    distinct_count: int = 0
    min_value: str | None = None
    max_value: str | None = None
    sample_values: list[str] = Field(default_factory=list)


class FieldInfo(BaseModel):
    name: str
    ordinal_position: int = 0
    dao_type: int
    type_name: str
    python_type: str = "Any"
    sqlalchemy_type: str = "Text"
    size: int = 0
    required: bool = False
    allow_zero_length: bool = False
    default_value: str = ""
    validation_rule: str = ""
    validation_text: str = ""
    input_mask: str = ""
    format: str = ""
    caption: str = ""
    description: str = ""
    is_autonumber: bool = False
    is_primary_key: bool = False
    is_hyperlink: bool = False
    is_multivalue: bool = False
    is_attachment: bool = False
    attributes: int = 0
    lookup: LookupInfo | None = None
    profile: FieldDataProfile | None = None


class IndexFieldInfo(BaseModel):
    name: str
    descending: bool = False


class IndexInfo(BaseModel):
    name: str
    primary: bool = False
    unique: bool = False
    foreign: bool = False
    ignore_nulls: bool = False
    required: bool = False
    fields: list[IndexFieldInfo] = Field(default_factory=list)


class TableInfo(BaseModel):
    name: str
    kind: Literal["local", "linked_access", "linked_odbc", "system"] = "local"
    is_hidden: bool = False
    connect: str = ""
    source_table_name: str = ""
    description: str = ""
    validation_rule: str = ""
    validation_text: str = ""
    record_count: int | None = None
    fields: list[FieldInfo] = Field(default_factory=list)
    indexes: list[IndexInfo] = Field(default_factory=list)
    primary_key_fields: list[str] = Field(default_factory=list)
    date_created: str = ""
    last_updated: str = ""


class RelationFieldMapping(BaseModel):
    source_field: str
    target_field: str


class RelationInfo(BaseModel):
    name: str
    table: str
    foreign_table: str
    enforce_referential_integrity: bool = False
    cascade_update: bool = False
    cascade_delete: bool = False
    is_one_to_one: bool = False
    join_type: Literal["inner", "left", "right"] = "inner"
    attributes: int = 0
    fields: list[RelationFieldMapping] = Field(default_factory=list)
