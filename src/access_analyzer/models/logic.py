"""Macro and VBA module/procedure data models."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


class MacroActionInfo(BaseModel):
    action_name: str
    arguments: dict[str, str] = Field(default_factory=dict)
    condition: str | None = None
    submacro_name: str | None = None
    comment: str | None = None


class MacroInfo(BaseModel):
    name: str
    is_autoexec: bool = False
    is_autokeys: bool = False
    format: Literal["xml", "legacy", "unknown"] = "unknown"
    submacros: list[str] = Field(default_factory=list)
    actions: list[MacroActionInfo] = Field(default_factory=list)
    referenced_forms: list[str] = Field(default_factory=list)
    referenced_reports: list[str] = Field(default_factory=list)
    referenced_queries: list[str] = Field(default_factory=list)
    referenced_tables: list[str] = Field(default_factory=list)
    referenced_macros: list[str] = Field(default_factory=list)
    referenced_functions: list[str] = Field(default_factory=list)
    raw_text: str = ""


class VbaParameterInfo(BaseModel):
    name: str
    passing: Literal["ByRef", "ByVal", "Implicit"] = "Implicit"
    param_type: str = "Variant"
    is_optional: bool = False
    default_value: str | None = None


class VbaDeclarationInfo(BaseModel):
    name: str
    kind: Literal["Declare", "Type", "Enum", "Const", "Variable"]
    visibility: Literal["Public", "Private", "Global", "Dim"] = "Private"
    signature: str = ""
    lib_name: str | None = None


class VbaProcedureInfo(BaseModel):
    name: str
    kind: Literal["Sub", "Function", "Property Get", "Property Let", "Property Set"]
    visibility: Literal["Public", "Private", "Friend"] = "Public"
    parameters: list[VbaParameterInfo] = Field(default_factory=list)
    return_type: str | None = None
    start_line: int = 0
    end_line: int = 0
    line_count: int = 0
    calls: list[str] = Field(default_factory=list)
    docmd_calls: list[dict[str, str]] = Field(default_factory=list)
    sql_snippets: list[str] = Field(default_factory=list)
    domain_aggregates: list[dict[str, str]] = Field(default_factory=list)
    form_references: list[str] = Field(default_factory=list)
    report_references: list[str] = Field(default_factory=list)
    tempvar_references: list[str] = Field(default_factory=list)
    external_objects: list[str] = Field(default_factory=list)
    recordset_sources: list[str] = Field(default_factory=list)


class VbaModuleInfo(BaseModel):
    name: str
    module_type: Literal["standard", "class", "form", "report"] = "standard"
    owner_object: str | None = None
    option_explicit: bool = False
    option_compare: str | None = None
    line_count: int = 0
    code_line_count: int = 0
    declarations: list[VbaDeclarationInfo] = Field(default_factory=list)
    procedures: list[VbaProcedureInfo] = Field(default_factory=list)
    raw_code: str = ""
