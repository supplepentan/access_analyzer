"""Converter from SaveAsText AST into FormInfo and ReportInfo models."""

from __future__ import annotations

from access_analyzer.extractors.dao_constants import DEFAULT_VIEW_MAP
from access_analyzer.models.logic import MacroActionInfo
from access_analyzer.models.ui import (
    ControlBounds,
    ControlInfo,
    FormInfo,
    GroupLevelInfo,
    ReportInfo,
    SectionInfo,
)
from access_analyzer.parsers.macro_parser import parse_xml_macro_actions
from access_analyzer.parsers.saveastext_parser import SaveAsTextNode, parse_saveastext
from access_analyzer.parsers.vba_parser import parse_vba_module

_EVENT_PREFIXES = ("On", "Before", "After")
_NON_EVENT_KEYS = {"Online", "Orientation", "OrderBy", "OrderByOn", "OrderByOnLoad"}

_CONTROL_TYPES = {
    "Label",
    "TextBox",
    "ComboBox",
    "ListBox",
    "CommandButton",
    "CheckBox",
    "OptionButton",
    "OptionGroup",
    "ToggleButton",
    "Subform",
    "Tab",
    "Page",
    "Image",
    "Line",
    "Rectangle",
    "PageBreak",
    "BoundObjectFrame",
    "ObjectFrame",
    "CustomControl",
    "Attachment",
    "NavigationControl",
    "NavigationButton",
    "WebBrowserControl",
    "EmptyCell",
}


def _to_bool(val: str | None, default: bool) -> bool:
    if val is None or val == "":
        return default
    low = val.strip().lower()
    if low in ("notdefault", "true", "-1", "1", "yes"):
        return True if low != "notdefault" else (not default)
    if low in ("false", "0", "no"):
        return False
    return default


def _to_int(val: str | None, default: int | None = None) -> int | None:
    if val is None or val == "":
        return default
    try:
        return int(val.strip())
    except ValueError:
        return default


def _normalize_subobject_name(raw: str) -> str:
    s = raw.strip()
    for prefix in ("Form.", "Report.", "Table.", "Query."):
        if s.startswith(prefix):
            return s[len(prefix) :]
    return s


def _extract_events_and_macros(
    props: dict[str, str],
) -> tuple[dict[str, str], dict[str, list[MacroActionInfo]]]:
    events: dict[str, str] = {}
    embedded: dict[str, list[MacroActionInfo]] = {}

    for k, v in props.items():
        if k.endswith("EmMacro") and v:
            ev_name = k[: -len("EmMacro")]
            actions, _ = parse_xml_macro_actions(v)
            if actions:
                embedded[ev_name] = actions
            events.setdefault(ev_name, "[Embedded Macro]")
        elif k.startswith(_EVENT_PREFIXES) and k not in _NON_EVENT_KEYS and v:
            events[k] = v

    return events, embedded


def _build_control(node: SaveAsTextNode, section_name: str, subobjects: list[str]) -> ControlInfo:
    props = node.properties
    name = props.get("Name", f"<{node.node_type}>")
    events, embedded = _extract_events_and_macros(props)

    bounds: ControlBounds | None = None
    if any(k in props for k in ("Left", "Top", "Width", "Height")):
        bounds = ControlBounds(
            left=_to_int(props.get("Left"), 0) or 0,
            top=_to_int(props.get("Top"), 0) or 0,
            width=_to_int(props.get("Width"), 0) or 0,
            height=_to_int(props.get("Height"), 0) or 0,
        )

    subform_src = ""
    if node.node_type == "Subform":
        raw_src = props.get("SourceObject", "")
        subform_src = _normalize_subobject_name(raw_src)
        if subform_src and subform_src not in subobjects:
            subobjects.append(subform_src)

    children: list[ControlInfo] = []
    for child in node.children:
        if child.node_type == "Anonymous":
            for grand in child.children:
                if grand.node_type in _CONTROL_TYPES:
                    children.append(_build_control(grand, section_name, subobjects))
        elif child.node_type in _CONTROL_TYPES:
            children.append(_build_control(child, section_name, subobjects))

    keep_props = {
        k: v
        for k, v in props.items()
        if not v.startswith("<binary:") and not k.endswith("EmMacro")
    }

    return ControlInfo(
        name=name,
        control_type=node.node_type,
        section=section_name,
        control_source=props.get("ControlSource", ""),
        row_source=props.get("RowSource", ""),
        row_source_type=props.get("RowSourceType", ""),
        bound_column=props.get("BoundColumn", ""),
        column_count=props.get("ColumnCount", ""),
        column_widths=props.get("ColumnWidths", ""),
        caption=props.get("Caption", ""),
        format=props.get("Format", ""),
        input_mask=props.get("InputMask", ""),
        default_value=props.get("DefaultValue", ""),
        validation_rule=props.get("ValidationRule", ""),
        validation_text=props.get("ValidationText", ""),
        visible=_to_bool(props.get("Visible"), True),
        enabled=_to_bool(props.get("Enabled"), True),
        locked=_to_bool(props.get("Locked"), False),
        tab_index=_to_int(props.get("TabIndex")),
        bounds=bounds,
        subform_source=subform_src,
        link_master_fields=props.get("LinkMasterFields", ""),
        link_child_fields=props.get("LinkChildFields", ""),
        events=events,
        embedded_macros=embedded,
        properties=keep_props,
        children=children,
    )


def _walk_ui_root(
    root: SaveAsTextNode,
) -> tuple[list[SectionInfo], list[ControlInfo], list[str], list[GroupLevelInfo]]:
    sections: list[SectionInfo] = []
    controls: list[ControlInfo] = []
    subobjects: list[str] = []
    group_levels: list[GroupLevelInfo] = []

    def visit(nodes: list[SaveAsTextNode], current_section: str) -> None:
        for n in nodes:
            if n.node_type == "Anonymous":
                visit(n.children, current_section)
            elif n.node_type == "Section":
                sec_name = n.properties.get("Name", "Section")
                sec_events, _ = _extract_events_and_macros(n.properties)
                sec_ctrl_names: list[str] = []
                # Collect direct controls inside this section
                sec_children: list[ControlInfo] = []

                def collect_sec_ctrls(c_nodes: list[SaveAsTextNode]) -> None:
                    for cn in c_nodes:
                        if cn.node_type == "Anonymous":
                            collect_sec_ctrls(cn.children)
                        elif cn.node_type in _CONTROL_TYPES:
                            ctrl = _build_control(cn, sec_name, subobjects)
                            sec_children.append(ctrl)
                            sec_ctrl_names.append(ctrl.name)

                collect_sec_ctrls(n.children)
                controls.extend(sec_children)
                sections.append(
                    SectionInfo(
                        name=sec_name,
                        section_type=n.properties.get("DisplayWhen", "Section"),
                        height=_to_int(n.properties.get("Height")),
                        visible=_to_bool(n.properties.get("Visible"), True),
                        events=sec_events,
                        control_names=sec_ctrl_names,
                    )
                )
            elif n.node_type == "GroupLevel":
                p = n.properties
                sort_raw = p.get("SortOrder", "0")
                sort_order = "Descending" if sort_raw in ("1", "-1", "True", "NotDefault") else "Ascending"
                group_levels.append(
                    GroupLevelInfo(
                        control_source=p.get("ControlSource", ""),
                        group_on=p.get("GroupOn", ""),
                        group_interval=p.get("GroupInterval", ""),
                        sort_order=sort_order,
                        group_header=_to_bool(p.get("GroupHeader"), False),
                        group_footer=_to_bool(p.get("GroupFooter"), False),
                        keep_together=p.get("KeepTogether", ""),
                    )
                )
            elif n.node_type in _CONTROL_TYPES:
                controls.append(_build_control(n, current_section, subobjects))
            else:
                visit(n.children, current_section)

    visit(root.children, "")
    return sections, controls, subobjects, group_levels


def parse_form(name: str, raw_text: str) -> FormInfo:
    """Parse an Access Form SaveAsText file into FormInfo."""
    doc = parse_saveastext(raw_text)
    form_node = next((n for n in doc.root_nodes if n.node_type == "Form"), None)
    if form_node is None and doc.root_nodes:
        form_node = doc.root_nodes[0]

    props = form_node.properties if form_node else {}
    sections, controls, subforms, _ = _walk_ui_root(form_node) if form_node else ([], [], [], [])
    events, embedded = _extract_events_and_macros(props)

    vba_mod = None
    if doc.code_behind.strip():
        vba_mod = parse_vba_module(
            name=f"Form_{name}",
            raw_code=doc.code_behind,
            module_type="form",
            owner_object=name,
        )

    default_view_code = props.get("DefaultView", "0")
    default_view = DEFAULT_VIEW_MAP.get(default_view_code, default_view_code)

    raw_props = {k: v for k, v in props.items() if not v.startswith("<binary:") and not k.endswith("EmMacro")}

    return FormInfo(
        name=name,
        caption=props.get("Caption", ""),
        record_source=props.get("RecordSource", ""),
        filter=props.get("Filter", ""),
        filter_on=_to_bool(props.get("FilterOn"), False),
        order_by=props.get("OrderBy", ""),
        default_view=default_view,
        allow_additions=_to_bool(props.get("AllowAdditions"), True),
        allow_edits=_to_bool(props.get("AllowEdits"), True),
        allow_deletions=_to_bool(props.get("AllowDeletions"), True),
        allow_filters=_to_bool(props.get("AllowFilters"), True),
        data_entry=_to_bool(props.get("DataEntry"), False),
        modal=_to_bool(props.get("Modal"), False),
        popup=_to_bool(props.get("PopUp"), False),
        navigation_buttons=_to_bool(props.get("NavigationButtons"), True),
        record_selectors=_to_bool(props.get("RecordSelectors"), True),
        has_module=vba_mod is not None or _to_bool(props.get("HasModule"), False),
        sections=sections,
        controls=controls,
        subforms=subforms,
        events=events,
        embedded_macros=embedded,
        vba_module=vba_mod,
        raw_properties=raw_props,
    )


def parse_report(name: str, raw_text: str) -> ReportInfo:
    """Parse an Access Report SaveAsText file into ReportInfo."""
    doc = parse_saveastext(raw_text)
    rep_node = next((n for n in doc.root_nodes if n.node_type == "Report"), None)
    if rep_node is None and doc.root_nodes:
        rep_node = doc.root_nodes[0]

    props = rep_node.properties if rep_node else {}
    sections, controls, subreports, group_levels = _walk_ui_root(rep_node) if rep_node else ([], [], [], [])
    events, embedded = _extract_events_and_macros(props)

    vba_mod = None
    if doc.code_behind.strip():
        vba_mod = parse_vba_module(
            name=f"Report_{name}",
            raw_code=doc.code_behind,
            module_type="report",
            owner_object=name,
        )

    raw_props = {k: v for k, v in props.items() if not v.startswith("<binary:") and not k.endswith("EmMacro")}

    return ReportInfo(
        name=name,
        caption=props.get("Caption", ""),
        record_source=props.get("RecordSource", ""),
        filter=props.get("Filter", ""),
        order_by=props.get("OrderBy", ""),
        has_module=vba_mod is not None or _to_bool(props.get("HasModule"), False),
        group_levels=group_levels,
        sections=sections,
        controls=controls,
        subreports=subreports,
        events=events,
        embedded_macros=embedded,
        vba_module=vba_mod,
        raw_properties=raw_props,
    )
