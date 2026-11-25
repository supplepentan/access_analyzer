"""Parser for Microsoft Access standalone and embedded macros (XML and legacy Begin-End)."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET

from access_analyzer.models.logic import MacroActionInfo, MacroInfo
from access_analyzer.parsers.saveastext_parser import SaveAsTextNode, parse_saveastext


def _strip_ns(tag: str) -> str:
    """Remove XML namespace prefix from a tag name."""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _parse_xml_statements(
    elem: ET.Element,
    actions: list[MacroActionInfo],
    submacro_name: str | None = None,
    condition: str | None = None,
) -> None:
    """Recursively extract MacroActionInfo items from Access XML macro statements."""
    for child in elem:
        tag = _strip_ns(child.tag)
        if tag == "Action":
            action_name = child.attrib.get("Name", "")
            args: dict[str, str] = {}
            for arg in child:
                if _strip_ns(arg.tag) == "Argument":
                    arg_name = arg.attrib.get("Name", f"arg{len(args)}")
                    args[arg_name] = (arg.text or "").strip()
            actions.append(
                MacroActionInfo(
                    action_name=action_name,
                    arguments=args,
                    condition=condition,
                    submacro_name=submacro_name,
                )
            )
        elif tag == "ConditionalBlock":
            for branch in child:
                b_tag = _strip_ns(branch.tag)
                if b_tag in ("If", "ElseIf", "Else"):
                    branch_cond = "ELSE" if b_tag == "Else" else None
                    stmts_elem = None
                    for sub in branch:
                        s_tag = _strip_ns(sub.tag)
                        if s_tag == "Condition":
                            branch_cond = (sub.text or "").strip()
                        elif s_tag == "Statements":
                            stmts_elem = sub
                    if stmts_elem is not None:
                        combined_cond = (
                            f"{condition} AND {branch_cond}"
                            if condition and branch_cond
                            else (branch_cond or condition)
                        )
                        _parse_xml_statements(
                            stmts_elem, actions, submacro_name=submacro_name, condition=combined_cond
                        )
        elif tag in ("Statements", "Group", "ForEachRecord"):
            _parse_xml_statements(child, actions, submacro_name=submacro_name, condition=condition)
        elif tag == "Comment":
            comment_text = (child.text or "").strip()
            if comment_text:
                actions.append(
                    MacroActionInfo(
                        action_name="Comment",
                        comment=comment_text,
                        condition=condition,
                        submacro_name=submacro_name,
                    )
                )


def parse_xml_macro_actions(xml_text: str) -> tuple[list[MacroActionInfo], list[str]]:
    """Parse Access XML macro text (<UserInterfaceMacros> / <DataMacros>)."""
    xml_start = xml_text.find("<")
    if xml_start == -1:
        return [], []
    cleaned = xml_text[xml_start:].strip()
    try:
        root = ET.fromstring(cleaned)
    except ET.ParseError:
        return [], []

    actions: list[MacroActionInfo] = []
    submacros: list[str] = []

    def walk(node: ET.Element, current_sub: str | None = None) -> None:
        tag = _strip_ns(node.tag)
        if tag == "Sub":
            sub_name = node.attrib.get("Name", "")
            if sub_name:
                submacros.append(sub_name)
            for c in node:
                if _strip_ns(c.tag) == "Statements":
                    _parse_xml_statements(c, actions, submacro_name=sub_name or current_sub)
                else:
                    walk(c, sub_name or current_sub)
            return
        if tag == "Statements":
            _parse_xml_statements(node, actions, submacro_name=current_sub)
            return
        for c in node:
            walk(c, current_sub)

    walk(root)
    return actions, submacros


def parse_legacy_macro_node(node: SaveAsTextNode) -> tuple[list[MacroActionInfo], list[str]]:
    """Parse a legacy Begin ... End macro node (or embedded macro property block)."""
    actions: list[MacroActionInfo] = []
    submacros: list[str] = []
    current_sub: str | None = None

    def visit(n: SaveAsTextNode) -> None:
        nonlocal current_sub
        props = n.properties
        if "MacroName" in props:
            current_sub = props["MacroName"]
            if current_sub and current_sub not in submacros:
                submacros.append(current_sub)

        if "Action" in props:
            args: dict[str, str] = {}
            if "Argument" in props:
                args["Argument0"] = props["Argument"]
            # Also collect additional Argument lines from raw_lines or repeated patterns
            for k, v in props.items():
                if k not in ("MacroName", "Condition", "Action", "Comment", "Argument"):
                    args[k] = v
            actions.append(
                MacroActionInfo(
                    action_name=props["Action"],
                    arguments=args,
                    condition=props.get("Condition"),
                    submacro_name=current_sub,
                    comment=props.get("Comment"),
                )
            )

        for child in n.children:
            visit(child)

    visit(node)
    return actions, submacros


def _parse_legacy_macro_text_lines(text: str) -> tuple[list[MacroActionInfo], list[str]]:
    """Parse legacy SaveAsText macro format preserving multiple Argument = lines per Begin block."""
    actions: list[MacroActionInfo] = []
    submacros: list[str] = []
    current_sub: str | None = None

    lines = text.replace("\r\n", "\n").split("\n")
    in_block = False
    block_action: str | None = None
    block_condition: str | None = None
    block_comment: str | None = None
    block_args: list[str] = []

    prop_re = re.compile(r"^([A-Za-z0-9_]+)\s*=(.*)$")

    def flush_block() -> None:
        nonlocal block_action, block_condition, block_comment, block_args
        if block_action:
            args = {f"Argument{i}": val for i, val in enumerate(block_args)}
            actions.append(
                MacroActionInfo(
                    action_name=block_action,
                    arguments=args,
                    condition=block_condition,
                    submacro_name=current_sub,
                    comment=block_comment,
                )
            )
        block_action = None
        block_condition = None
        block_comment = None
        block_args = []

    for line in lines:
        s = line.strip()
        if s == "Begin":
            in_block = True
            block_action = None
            block_condition = None
            block_comment = None
            block_args = []
            continue
        if s == "End" and in_block:
            flush_block()
            in_block = False
            continue
        if not in_block:
            continue

        m = prop_re.match(s)
        if not m:
            continue
        key, raw_val = m.group(1), m.group(2).strip()
        val = raw_val[1:-1].replace('""', '"') if len(raw_val) >= 2 and raw_val[0] == '"' and raw_val[-1] == '"' else raw_val
        if key == "MacroName":
            current_sub = val
            if val and val not in submacros:
                submacros.append(val)
        elif key == "Action":
            block_action = val
        elif key == "Condition":
            block_condition = val
        elif key == "Comment":
            block_comment = val
        elif key == "Argument":
            block_args.append(val)

    return actions, submacros


def _extract_macro_references(actions: list[MacroActionInfo]) -> dict[str, list[str]]:
    refs: dict[str, set[str]] = {
        "forms": set(),
        "reports": set(),
        "queries": set(),
        "tables": set(),
        "macros": set(),
        "functions": set(),
    }
    for act in actions:
        name_lower = act.action_name.lower()
        args = act.arguments
        first_arg = next(iter(args.values()), "") if args else ""

        if name_lower == "openform":
            target = args.get("FormName") or args.get("Argument0") or first_arg
            if target:
                refs["forms"].add(target)
        elif name_lower == "openreport":
            target = args.get("ReportName") or args.get("Argument0") or first_arg
            if target:
                refs["reports"].add(target)
        elif name_lower == "openquery":
            target = args.get("QueryName") or args.get("Argument0") or first_arg
            if target:
                refs["queries"].add(target)
        elif name_lower == "opentable":
            target = args.get("TableName") or args.get("Argument0") or first_arg
            if target:
                refs["tables"].add(target)
        elif name_lower == "runmacro":
            target = args.get("MacroName") or args.get("Argument0") or first_arg
            if target:
                refs["macros"].add(target.split(".")[0])
        elif name_lower == "runcode":
            fn_expr = args.get("FunctionName") or args.get("Argument0") or first_arg
            fn_match = re.match(r"^\s*([A-Za-z0-9_\u3000-\u9fff]+)", fn_expr)
            if fn_match:
                refs["functions"].add(fn_match.group(1))
        elif name_lower in ("closewindow", "close"):
            obj_type = (args.get("ObjectType") or args.get("Argument0") or "").lower()
            obj_name = args.get("ObjectName") or args.get("Argument1") or ""
            if obj_name:
                if "form" in obj_type:
                    refs["forms"].add(obj_name)
                elif "report" in obj_type:
                    refs["reports"].add(obj_name)

    return {k: sorted(v) for k, v in refs.items()}


def parse_embedded_macro(text_or_node: str | SaveAsTextNode) -> list[MacroActionInfo]:
    """Parse an embedded macro from either an XML/legacy string or a SaveAsTextNode."""
    if isinstance(text_or_node, SaveAsTextNode):
        actions, _ = parse_legacy_macro_node(text_or_node)
        return actions

    text = text_or_node.strip()
    if "<UserInterfaceMacros" in text or "<DataMacros" in text or text.startswith("<?xml"):
        actions, _ = parse_xml_macro_actions(text)
        return actions
    if "Action" in text and "Begin" in text:
        actions, _ = _parse_legacy_macro_text_lines(text)
        return actions
    return []


def parse_macro(name: str, raw_text: str) -> MacroInfo:
    """Parse a standalone Access macro (from SaveAsText acMacro)."""
    stripped = raw_text.strip()
    fmt: str = "unknown"
    actions: list[MacroActionInfo] = []
    submacros: list[str] = []

    if "<UserInterfaceMacros" in stripped or "<DataMacros" in stripped or stripped.startswith("<?xml"):
        fmt = "xml"
        actions, submacros = parse_xml_macro_actions(stripped)
    elif "Action" in stripped and "Begin" in stripped:
        fmt = "legacy"
        actions, submacros = _parse_legacy_macro_text_lines(stripped)
        if not actions:
            doc = parse_saveastext(stripped)
            for root in doc.root_nodes:
                a, s = parse_legacy_macro_node(root)
                actions.extend(a)
                submacros.extend(s)

    refs = _extract_macro_references(actions)
    lower_name = name.lower()

    return MacroInfo(
        name=name,
        is_autoexec=(lower_name == "autoexec"),
        is_autokeys=(lower_name == "autokeys"),
        format=fmt,  # type: ignore[arg-type]
        submacros=sorted(set(submacros)),
        actions=actions,
        referenced_forms=refs["forms"],
        referenced_reports=refs["reports"],
        referenced_queries=refs["queries"],
        referenced_tables=refs["tables"],
        referenced_macros=refs["macros"],
        referenced_functions=refs["functions"],
        raw_text=raw_text,
    )
