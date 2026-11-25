"""Query and SQL analysis data models."""

from __future__ import annotations

from pydantic import BaseModel, Field


class QueryParameterInfo(BaseModel):
    name: str
    dao_type: int = 0
    type_name: str = ""
    python_type: str = "Any"


class ParsedSqlInfo(BaseModel):
    statement_type: str = "UNKNOWN"
    referenced_sources: list[str] = Field(default_factory=list)
    target_table: str | None = None
    joins: list[str] = Field(default_factory=list)
    access_functions: list[str] = Field(default_factory=list)
    form_references: list[str] = Field(default_factory=list)
    report_references: list[str] = Field(default_factory=list)
    tempvar_references: list[str] = Field(default_factory=list)
    has_parameters: bool = False
    has_subquery: bool = False
    has_wildcard_like: bool = False
    is_crosstab: bool = False
    is_action_query: bool = False
    is_union: bool = False
    is_passthrough: bool = False


class QueryInfo(BaseModel):
    name: str
    query_type_code: int = 0
    query_type: str = "Select"
    sql: str = ""
    description: str = ""
    connect: str = ""
    returns_records: bool = True
    max_records: int = 0
    parameters: list[QueryParameterInfo] = Field(default_factory=list)
    date_created: str = ""
    last_updated: str = ""
    parsed: ParsedSqlInfo = Field(default_factory=ParsedSqlInfo)
