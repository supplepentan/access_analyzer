"""Unified analysis, dependency graph, and Python migration inventory models."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field

from access_analyzer.models.database import DatabaseMetadata, RelationInfo, TableInfo
from access_analyzer.models.logic import MacroInfo, VbaModuleInfo
from access_analyzer.models.query import QueryInfo
from access_analyzer.models.ui import FormInfo, ReportInfo


ObjectType = Literal["table", "query", "form", "report", "macro", "vba_module", "external"]


class ObjectNode(BaseModel):
    id: str
    object_type: ObjectType
    name: str
    is_entry_point: bool = False
    is_orphan: bool = False
    in_degree: int = 0
    out_degree: int = 0


class DependencyEdge(BaseModel):
    source_id: str
    target_id: str
    edge_type: str
    detail: str = ""


class DependencyGraph(BaseModel):
    nodes: list[ObjectNode] = Field(default_factory=list)
    edges: list[DependencyEdge] = Field(default_factory=list)
    entry_points: list[str] = Field(default_factory=list)
    orphan_objects: list[str] = Field(default_factory=list)
    topological_order: list[str] = Field(default_factory=list)
    cycles: list[list[str]] = Field(default_factory=list)


class MigrationHotspot(BaseModel):
    severity: Literal["high", "medium", "low"]
    category: str
    object_id: str
    summary: str
    recommendation: str


class ObjectComplexity(BaseModel):
    object_id: str
    object_type: str
    name: str
    score: int
    metrics: dict[str, int | float] = Field(default_factory=dict)


class ScreenBlueprint(BaseModel):
    form_name: str
    caption: str = ""
    default_view: str = "SingleForm"
    primary_source: str = ""
    subforms: list[str] = Field(default_factory=list)
    bound_tables: list[str] = Field(default_factory=list)
    bound_queries: list[str] = Field(default_factory=list)
    event_handlers: list[str] = Field(default_factory=list)
    opens_forms: list[str] = Field(default_factory=list)
    opens_reports: list[str] = Field(default_factory=list)
    complexity_score: int = 0


class MigrationInventory(BaseModel):
    total_tables: int = 0
    total_linked_tables: int = 0
    total_queries: int = 0
    total_forms: int = 0
    total_reports: int = 0
    total_macros: int = 0
    total_vba_modules: int = 0
    total_vba_lines: int = 0
    total_vba_procedures: int = 0
    hotspots: list[MigrationHotspot] = Field(default_factory=list)
    complexity_ranking: list[ObjectComplexity] = Field(default_factory=list)
    screen_blueprints: list[ScreenBlueprint] = Field(default_factory=list)
    recommended_python_layers: dict[str, list[str]] = Field(default_factory=dict)


class DatabaseAnalysis(BaseModel):
    analyzer_version: str = "0.1.0"
    metadata: DatabaseMetadata
    tables: list[TableInfo] = Field(default_factory=list)
    relations: list[RelationInfo] = Field(default_factory=list)
    queries: list[QueryInfo] = Field(default_factory=list)
    forms: list[FormInfo] = Field(default_factory=list)
    reports: list[ReportInfo] = Field(default_factory=list)
    macros: list[MacroInfo] = Field(default_factory=list)
    vba_modules: list[VbaModuleInfo] = Field(default_factory=list)
    dependency_graph: DependencyGraph = Field(default_factory=DependencyGraph)
    migration_inventory: MigrationInventory = Field(default_factory=MigrationInventory)
    warnings: list[str] = Field(default_factory=list)
