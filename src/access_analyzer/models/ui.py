"""Form, report, section, group level, and control UI data models."""

from __future__ import annotations

from pydantic import BaseModel, Field

from access_analyzer.models.logic import MacroActionInfo, VbaModuleInfo


class ControlBounds(BaseModel):
    left: int = 0
    top: int = 0
    width: int = 0
    height: int = 0


class ControlInfo(BaseModel):
    name: str
    control_type: str
    section: str = ""
    control_source: str = ""
    row_source: str = ""
    row_source_type: str = ""
    bound_column: str = ""
    column_count: str = ""
    column_widths: str = ""
    caption: str = ""
    format: str = ""
    input_mask: str = ""
    default_value: str = ""
    validation_rule: str = ""
    validation_text: str = ""
    visible: bool = True
    enabled: bool = True
    locked: bool = False
    tab_index: int | None = None
    bounds: ControlBounds | None = None
    subform_source: str = ""
    link_master_fields: str = ""
    link_child_fields: str = ""
    events: dict[str, str] = Field(default_factory=dict)
    embedded_macros: dict[str, list[MacroActionInfo]] = Field(default_factory=dict)
    properties: dict[str, str] = Field(default_factory=dict)
    children: list[ControlInfo] = Field(default_factory=list)


class SectionInfo(BaseModel):
    name: str
    section_type: str
    height: int | None = None
    visible: bool = True
    events: dict[str, str] = Field(default_factory=dict)
    control_names: list[str] = Field(default_factory=list)


class GroupLevelInfo(BaseModel):
    control_source: str = ""
    group_on: str = ""
    group_interval: str = ""
    sort_order: str = "Ascending"
    group_header: bool = False
    group_footer: bool = False
    keep_together: str = ""


class FormInfo(BaseModel):
    name: str
    caption: str = ""
    record_source: str = ""
    filter: str = ""
    filter_on: bool = False
    order_by: str = ""
    default_view: str = "SingleForm"
    allow_additions: bool = True
    allow_edits: bool = True
    allow_deletions: bool = True
    allow_filters: bool = True
    data_entry: bool = False
    modal: bool = False
    popup: bool = False
    navigation_buttons: bool = True
    record_selectors: bool = True
    has_module: bool = False
    is_subform_target: bool = False
    sections: list[SectionInfo] = Field(default_factory=list)
    controls: list[ControlInfo] = Field(default_factory=list)
    subforms: list[str] = Field(default_factory=list)
    events: dict[str, str] = Field(default_factory=dict)
    embedded_macros: dict[str, list[MacroActionInfo]] = Field(default_factory=dict)
    vba_module: VbaModuleInfo | None = None
    raw_properties: dict[str, str] = Field(default_factory=dict)


class ReportInfo(BaseModel):
    name: str
    caption: str = ""
    record_source: str = ""
    filter: str = ""
    order_by: str = ""
    has_module: bool = False
    is_subreport_target: bool = False
    group_levels: list[GroupLevelInfo] = Field(default_factory=list)
    sections: list[SectionInfo] = Field(default_factory=list)
    controls: list[ControlInfo] = Field(default_factory=list)
    subreports: list[str] = Field(default_factory=list)
    events: dict[str, str] = Field(default_factory=dict)
    embedded_macros: dict[str, list[MacroActionInfo]] = Field(default_factory=dict)
    vba_module: VbaModuleInfo | None = None
    raw_properties: dict[str, str] = Field(default_factory=dict)
