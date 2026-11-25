"""Static analyzer for Microsoft Access VBA modules (Standard, Class, Form, Report)."""

from __future__ import annotations

import re
from typing import Literal

from access_analyzer.models.logic import (
    VbaDeclarationInfo,
    VbaModuleInfo,
    VbaParameterInfo,
    VbaProcedureInfo,
)

_PROC_START_RE = re.compile(
    r"^\s*(?:(?P<vis>Public|Private|Friend)\s+)?"
    r"(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+"
    r"(?P<name>[A-Za-z_\u3000-\u9fff][A-Za-z0-9_\u3000-\u9fff]*)\s*"
    r"\((?P<params>[^)]*)\)"
    r"(?:\s+As\s+(?P<ret>[A-Za-z0-9_.\(\)\s]+))?",
    re.IGNORECASE,
)

_PROC_END_RE = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)

_DECLARE_RE = re.compile(
    r"^\s*(?:(?P<vis>Public|Private)\s+)?Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+"
    r"(?P<name>[A-Za-z0-9_]+)\s+Lib\s+\"(?P<lib>[^\"]+)\"",
    re.IGNORECASE,
)
_TYPE_ENUM_RE = re.compile(
    r"^\s*(?:(?P<vis>Public|Private)\s+)?(?P<kind>Type|Enum)\s+(?P<name>[A-Za-z0-9_\u3000-\u9fff]+)",
    re.IGNORECASE,
)
_CONST_RE = re.compile(
    r"^\s*(?:(?P<vis>Public|Private|Global)\s+)?Const\s+(?P<name>[A-Za-z0-9_\u3000-\u9fff]+)",
    re.IGNORECASE,
)
_VAR_RE = re.compile(
    r"^\s*(?P<vis>Public|Private|Global|Dim)\s+(?!WithEvents\s+)(?P<name>[A-Za-z0-9_\u3000-\u9fff]+)",
    re.IGNORECASE,
)

_DOCMD_RE = re.compile(
    r"""\bDoCmd\.(?P<method>[A-Za-z0-9_]+)\s*(?:\(\s*|\s+)(?:"(?P<arg1>[^"]+)"|(?P<arg2>[A-Za-z0-9_]+))?""",
    re.IGNORECASE,
)
_DOMAIN_AGG_RE = re.compile(
    r"""\b(?P<func>DLookup|DCount|DSum|DMax|DMin|DAvg|DFirst|DLast)\s*\(\s*"(?P<expr>[^"]*)"\s*,\s*"(?P<domain>[^"]*)"(?:\s*,\s*(?P<crit>[^)]*))?\)""",
    re.IGNORECASE,
)
_CREATE_OBJ_RE = re.compile(
    r"""\b(?:CreateObject|GetObject)\s*\(\s*(?:[^,]*,\s*)?"(?P<progid>[^"]+)"\s*\)""",
    re.IGNORECASE,
)
_SHELL_RE = re.compile(r"\bShell\s*\(", re.IGNORECASE)
_OPEN_RS_RE = re.compile(r"""\bOpenRecordset\s*\(\s*"(?P<src>[^"]+)""", re.IGNORECASE)
_SQL_LITERAL_RE = re.compile(
    r'"((?:""|[^"])*\b(?:SELECT\s+.+?\bFROM\b|INSERT\s+INTO\s+|UPDATE\s+.+?\bSET\b|DELETE\s+.*?\bFROM\b)(?:""|[^"])*)"',
    re.IGNORECASE,
)
_FORM_REF_RE = re.compile(
    r"""\bForms\s*(?:!\s*(?:\[([^\]]+)\]|([A-Za-z0-9_\u3000-\u9fff]+))|\(\s*"([^"]+)"\s*\))""",
    re.IGNORECASE,
)
_REPORT_REF_RE = re.compile(
    r"""\bReports\s*(?:!\s*(?:\[([^\]]+)\]|([A-Za-z0-9_\u3000-\u9fff]+))|\(\s*"([^"]+)"\s*\))""",
    re.IGNORECASE,
)
_TEMPVAR_REF_RE = re.compile(
    r"""\bTempVars\s*(?:!\s*(?:\[([^\]]+)\]|([A-Za-z0-9_\u3000-\u9fff]+))|\(\s*"([^"]+)"\s*\))""",
    re.IGNORECASE,
)
_CALL_STMT_RE = re.compile(r"\bCall\s+([A-Za-z_\u3000-\u9fff][A-Za-z0-9_\u3000-\u9fff]*)", re.IGNORECASE)
_FUNC_CALL_RE = re.compile(r"(?<!\.)\b([A-Za-z_\u3000-\u9fff][A-Za-z0-9_\u3000-\u9fff]*)\s*\(")

_VBA_BUILTINS = {
    "if", "elseif", "while", "until", "for", "select", "case", "with", "set", "dim", "redim",
    "msgbox", "inputbox", "nz", "iif", "isnull", "isempty", "isnumeric", "isdate", "isobject",
    "isarray", "iserror", "ismissing", "cint", "clng", "cdbl", "csng", "ccur", "cstr", "cdate",
    "cbool", "cbyte", "cvar", "val", "str", "len", "left", "right", "mid", "trim", "ltrim",
    "rtrim", "ucase", "lcase", "instr", "instrrev", "replace", "split", "join", "format",
    "date", "now", "time", "year", "month", "day", "hour", "minute", "second", "dateadd",
    "datediff", "datepart", "dateserial", "datevalue", "array", "ubound", "lbound", "err",
    "createobject", "getobject", "shell", "dir", "filelen", "freefile", "chr", "asc", "abs",
    "int", "fix", "round", "sgn", "sqr", "rnd", "nz", "dlookup", "dcount", "dsum", "dmax",
    "dmin", "davg", "dfirst", "dlast", "forms", "reports", "tempvars", "me", "currentdb",
    "docmd", "screen", "application", "eval", "nz", "openrecordset", "execute", "typeof",
    "switch", "choose", "space", "string", "strconv", "vbcrlf", "debug", "print",
}


def _strip_vba_comment(line: str) -> str:
    """Remove trailing VBA comment starting with `'` outside double-quoted strings."""
    in_str = False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            if in_str and i + 1 < len(line) and line[i + 1] == '"':
                i += 2
                continue
            in_str = not in_str
        elif ch == "'" and not in_str:
            return line[:i]
        i += 1
    return line


def _join_continuations(raw_lines: list[str]) -> list[tuple[int, str]]:
    """Join VBA line-continuation sequences (` _`) while recording original 1-based start line."""
    logical: list[tuple[int, str]] = []
    idx = 0
    n = len(raw_lines)
    while idx < n:
        start_line = idx + 1
        buf = raw_lines[idx]
        while buf.rstrip().endswith(" _") and idx + 1 < n:
            buf = buf.rstrip()[:-2] + " " + raw_lines[idx + 1].lstrip()
            idx += 1
        logical.append((start_line, buf))
        idx += 1
    return logical


def _parse_params(param_str: str) -> list[VbaParameterInfo]:
    params: list[VbaParameterInfo] = []
    if not param_str.strip():
        return params
    for raw_p in param_str.split(","):
        p = raw_p.strip()
        if not p:
            continue
        is_opt = False
        if re.match(r"^Optional\b", p, re.IGNORECASE):
            is_opt = True
            p = re.sub(r"^Optional\s+", "", p, flags=re.IGNORECASE).strip()

        passing: Literal["ByRef", "ByVal", "Implicit"] = "Implicit"
        if re.match(r"^ByVal\b", p, re.IGNORECASE):
            passing = "ByVal"
            p = re.sub(r"^ByVal\s+", "", p, flags=re.IGNORECASE).strip()
        elif re.match(r"^ByRef\b", p, re.IGNORECASE):
            passing = "ByRef"
            p = re.sub(r"^ByRef\s+", "", p, flags=re.IGNORECASE).strip()

        default_val: str | None = None
        if "=" in p:
            p, default_val = [x.strip() for x in p.split("=", 1)]

        param_type = "Variant"
        m_as = re.search(r"\bAs\s+(.+)$", p, re.IGNORECASE)
        if m_as:
            param_type = m_as.group(1).strip()
            p = p[: m_as.start()].strip()

        p_name = p.replace("()", "").strip()
        if p_name:
            params.append(
                VbaParameterInfo(
                    name=p_name,
                    passing=passing,
                    param_type=param_type,
                    is_optional=is_opt,
                    default_value=default_val,
                )
            )
    return params


def _analyze_procedure_body(proc: VbaProcedureInfo, body_lines: list[str]) -> None:
    calls: set[str] = set()
    docmd_calls: list[dict[str, str]] = []
    sql_snippets: list[str] = []
    domain_aggs: list[dict[str, str]] = []
    form_refs: set[str] = set()
    report_refs: set[str] = set()
    tempvars: set[str] = set()
    ext_objs: set[str] = set()
    rs_sources: list[str] = []

    for line in body_lines:
        code = _strip_vba_comment(line).strip()
        if not code:
            continue

        for dm in _DOCMD_RE.finditer(code):
            method = dm.group("method")
            target = (dm.group("arg1") or dm.group("arg2") or "").strip()
            docmd_calls.append({"method": method, "target": target})

        for dam in _DOMAIN_AGG_RE.finditer(code):
            domain_aggs.append(
                {
                    "function": dam.group("func"),
                    "expr": dam.group("expr"),
                    "domain": dam.group("domain"),
                    "criteria": (dam.group("crit") or "").strip(),
                }
            )

        for cm in _CREATE_OBJ_RE.finditer(code):
            ext_objs.add(cm.group("progid"))

        if _SHELL_RE.search(code):
            ext_objs.add("Shell")

        for rsm in _OPEN_RS_RE.finditer(code):
            src = rsm.group("src").replace('""', '"').strip()
            if src and src not in rs_sources:
                rs_sources.append(src)

        for sq in _SQL_LITERAL_RE.finditer(code):
            snip = sq.group(1).replace('""', '"').strip()
            if snip and snip not in sql_snippets:
                sql_snippets.append(snip)

        for fm in _FORM_REF_RE.finditer(code):
            val = (fm.group(1) or fm.group(2) or fm.group(3) or "").strip()
            if val:
                form_refs.add(val)

        for rm in _REPORT_REF_RE.finditer(code):
            val = (rm.group(1) or rm.group(2) or rm.group(3) or "").strip()
            if val:
                report_refs.add(val)

        for tm in _TEMPVAR_REF_RE.finditer(code):
            val = (tm.group(1) or tm.group(2) or tm.group(3) or "").strip()
            if val:
                tempvars.add(val)

        for call_m in _CALL_STMT_RE.finditer(code):
            cname = call_m.group(1)
            if cname.lower() not in _VBA_BUILTINS and cname.lower() != proc.name.lower():
                calls.add(cname)

        # Strip strings before scanning bare function call patterns
        code_no_str = re.sub(r'"(?:""|[^"])*"', '""', code)
        for fn_m in _FUNC_CALL_RE.finditer(code_no_str):
            cname = fn_m.group(1)
            if cname.lower() not in _VBA_BUILTINS and cname.lower() != proc.name.lower():
                calls.add(cname)

    proc.calls = sorted(calls)
    proc.docmd_calls = docmd_calls
    proc.sql_snippets = sql_snippets
    proc.domain_aggregates = domain_aggs
    proc.form_references = sorted(form_refs)
    proc.report_references = sorted(report_refs)
    proc.tempvar_references = sorted(tempvars)
    proc.external_objects = sorted(ext_objs)
    proc.recordset_sources = rs_sources


def parse_vba_module(
    name: str,
    raw_code: str,
    module_type: Literal["standard", "class", "form", "report"] = "standard",
    owner_object: str | None = None,
) -> VbaModuleInfo:
    """Parse a VBA module (standard, class, or form/report code-behind)."""
    text = raw_code.replace("\r\n", "\n").replace("\r", "\n")
    raw_lines = text.split("\n")

    # Detect class module from SaveAsText Attribute headers if caller passed "standard"
    detected_type = module_type
    filtered_lines: list[str] = []
    for line in raw_lines:
        stripped = line.strip()
        if stripped.startswith("Attribute VB_"):
            if "VB_Creatable" in stripped and "= True" in stripped and detected_type == "standard":
                detected_type = "class"
            continue
        if stripped.startswith("VERSION ") and "CLASS" in stripped.upper():
            if detected_type == "standard":
                detected_type = "class"
            continue
        if stripped in ("BEGIN", "END") or stripped.startswith("MultiUse ="):
            continue
        filtered_lines.append(line)

    clean_code = "\n".join(filtered_lines).strip("\n")
    code_lines_only = [
        ln for ln in filtered_lines if ln.strip() and not ln.strip().startswith("'") and not ln.strip().lower().startswith("rem ")
    ]

    logical_lines = _join_continuations(filtered_lines)

    option_explicit = False
    option_compare: str | None = None
    declarations: list[VbaDeclarationInfo] = []
    procedures: list[VbaProcedureInfo] = []

    current_proc: VbaProcedureInfo | None = None
    current_body: list[str] = []
    in_type_or_enum = False

    for line_no, raw_logical in logical_lines:
        code = _strip_vba_comment(raw_logical).strip()
        if not code:
            if current_proc is not None:
                current_body.append(raw_logical)
            continue

        if current_proc is None:
            if re.match(r"^Option\s+Explicit\b", code, re.IGNORECASE):
                option_explicit = True
                continue
            m_cmp = re.match(r"^Option\s+Compare\s+(\w+)", code, re.IGNORECASE)
            if m_cmp:
                option_compare = m_cmp.group(1)
                continue

            if in_type_or_enum:
                if re.match(r"^End\s+(?:Type|Enum)\b", code, re.IGNORECASE):
                    in_type_or_enum = False
                continue

            m_proc = _PROC_START_RE.match(code)
            if m_proc:
                vis_raw = (m_proc.group("vis") or "Public").capitalize()
                kind_raw = re.sub(r"\s+", " ", m_proc.group("kind").strip()).title()
                ret_raw = m_proc.group("ret").strip() if m_proc.group("ret") else None
                current_proc = VbaProcedureInfo(
                    name=m_proc.group("name"),
                    kind=kind_raw,  # type: ignore[arg-type]
                    visibility=vis_raw,  # type: ignore[arg-type]
                    parameters=_parse_params(m_proc.group("params") or ""),
                    return_type=ret_raw,
                    start_line=line_no,
                )
                current_body = []
                continue

            m_dec = _DECLARE_RE.match(code)
            if m_dec:
                vis = (m_dec.group("vis") or "Public").capitalize()
                declarations.append(
                    VbaDeclarationInfo(
                        name=m_dec.group("name"),
                        kind="Declare",
                        visibility=vis,  # type: ignore[arg-type]
                        signature=code,
                        lib_name=m_dec.group("lib"),
                    )
                )
                continue

            m_te = _TYPE_ENUM_RE.match(code)
            if m_te:
                in_type_or_enum = True
                vis = (m_te.group("vis") or "Public").capitalize()
                declarations.append(
                    VbaDeclarationInfo(
                        name=m_te.group("name"),
                        kind=m_te.group("kind").capitalize(),  # type: ignore[arg-type]
                        visibility=vis,  # type: ignore[arg-type]
                        signature=code,
                    )
                )
                continue

            m_const = _CONST_RE.match(code)
            if m_const:
                vis = (m_const.group("vis") or "Private").capitalize()
                declarations.append(
                    VbaDeclarationInfo(
                        name=m_const.group("name"),
                        kind="Const",
                        visibility=vis,  # type: ignore[arg-type]
                        signature=code,
                    )
                )
                continue

            m_var = _VAR_RE.match(code)
            if m_var:
                vis = m_var.group("vis").capitalize()
                declarations.append(
                    VbaDeclarationInfo(
                        name=m_var.group("name"),
                        kind="Variable",
                        visibility=vis,  # type: ignore[arg-type]
                        signature=code,
                    )
                )
                continue
        else:
            if _PROC_END_RE.match(code):
                current_proc.end_line = line_no
                current_proc.line_count = max(1, line_no - current_proc.start_line + 1)
                _analyze_procedure_body(current_proc, current_body)
                procedures.append(current_proc)
                current_proc = None
                current_body = []
            else:
                current_body.append(raw_logical)

    # Flush unclosed procedure if truncated
    if current_proc is not None:
        current_proc.end_line = len(filtered_lines)
        current_proc.line_count = max(1, current_proc.end_line - current_proc.start_line + 1)
        _analyze_procedure_body(current_proc, current_body)
        procedures.append(current_proc)

    return VbaModuleInfo(
        name=name,
        module_type=detected_type,
        owner_object=owner_object,
        option_explicit=option_explicit,
        option_compare=option_compare,
        line_count=len(filtered_lines),
        code_line_count=len(code_lines_only),
        declarations=declarations,
        procedures=procedures,
        raw_code=clean_code,
    )
