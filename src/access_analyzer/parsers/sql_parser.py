"""Static analyzer for Microsoft Access SQL dialect."""

from __future__ import annotations

import re

from access_analyzer.models.query import ParsedSqlInfo

_ACCESS_FUNCTIONS = [
    "Nz",
    "IIf",
    "Switch",
    "Choose",
    "DLookup",
    "DCount",
    "DSum",
    "DMax",
    "DMin",
    "DAvg",
    "DFirst",
    "DLast",
    "Format",
    "DateAdd",
    "DateDiff",
    "DatePart",
    "DateSerial",
    "DateValue",
    "TimeSerial",
    "Date",
    "Now",
    "Year",
    "Month",
    "Day",
    "Left",
    "Right",
    "Mid",
    "InStr",
    "InStrRev",
    "Trim",
    "LTrim",
    "RTrim",
    "Len",
    "Replace",
    "StrConv",
    "Val",
    "CInt",
    "CLng",
    "CDbl",
    "CCur",
    "CStr",
    "CDate",
    "CBool",
    "IsNull",
    "IsNumeric",
    "IsDate",
    "Eval",
    "First",
    "Last",
    "Partition",
]

_SQL_KEYWORDS = {
    "SELECT",
    "FROM",
    "WHERE",
    "GROUP",
    "ORDER",
    "HAVING",
    "INNER",
    "LEFT",
    "RIGHT",
    "OUTER",
    "JOIN",
    "ON",
    "AS",
    "AND",
    "OR",
    "NOT",
    "IN",
    "SET",
    "VALUES",
    "INTO",
    "UNION",
    "ALL",
    "DISTINCT",
    "TOP",
    "PIVOT",
    "TRANSFORM",
    "PARAMETERS",
    "WITH",
    "OWNERACCESS",
    "OPTION",
    "BY",
    "ASC",
    "DESC",
    "BETWEEN",
    "LIKE",
    "IS",
    "NULL",
    "EXISTS",
    "CASE",
    "WHEN",
    "THEN",
    "ELSE",
    "END",
    "UPDATE",
    "DELETE",
    "INSERT",
    "CREATE",
    "ALTER",
    "DROP",
    "TABLE",
    "INDEX",
    "VIEW",
}

_IDENT_PATTERN = r"(?:\[([^\]]+)\]|([A-Za-z_\u3000-\u9fff\uff00-\uffef][A-Za-z0-9_\u3000-\u9fff\uff00-\uffef$]*))"

_FORM_REF_RE = re.compile(
    r"(?:\[?Forms\]?)\s*!\s*(?:\[([^\]]+)\]|([A-Za-z0-9_\u3000-\u9fff]+))"
    r"(?:\s*!\s*(?:\[([^\]]+)\]|([A-Za-z0-9_\u3000-\u9fff]+)))?",
    re.IGNORECASE,
)
_REPORT_REF_RE = re.compile(
    r"(?:\[?Reports\]?)\s*!\s*(?:\[([^\]]+)\]|([A-Za-z0-9_\u3000-\u9fff]+))"
    r"(?:\s*!\s*(?:\[([^\]]+)\]|([A-Za-z0-9_\u3000-\u9fff]+)))?",
    re.IGNORECASE,
)
_TEMPVAR_REF_RE = re.compile(
    r"(?:\[?TempVars\]?)\s*!\s*(?:\[([^\]]+)\]|([A-Za-z0-9_\u3000-\u9fff]+))",
    re.IGNORECASE,
)
_JOIN_RE = re.compile(r"\b(INNER\s+JOIN|LEFT\s+(?:OUTER\s+)?JOIN|RIGHT\s+(?:OUTER\s+)?JOIN|CROSS\s+JOIN|JOIN)\b", re.IGNORECASE)
_WILDCARD_LIKE_RE = re.compile(r"""\bLIKE\s+(?:"[^"*?]*[*?][^"]*"|'[^'*?]*[*?][^']*')""", re.IGNORECASE)


def _strip_sql_strings(sql: str) -> str:
    """Replace string literals with placeholders so keywords inside strings do not falsely match."""
    sql = re.sub(r'"(?:""|[^"])*"', '""', sql)
    sql = re.sub(r"'(?:''|[^'])*'", "''", sql)
    return sql


def _extract_ident(match: re.Match[str], g1: int = 1, g2: int = 2) -> str:
    return (match.group(g1) or match.group(g2) or "").strip()


def _extract_from_and_join_sources(sql_no_strings: str) -> list[str]:
    sources: list[str] = []
    seen: set[str] = set()

    def add(name: str) -> None:
        clean = name.strip().strip("[]")
        if not clean or clean.upper() in _SQL_KEYWORDS:
            return
        if clean.lower() not in seen:
            seen.add(clean.lower())
            sources.append(clean)

    # FROM <ident> (handling Access nested parentheses like FROM ((T1 INNER JOIN T2 ON ...) LEFT JOIN T3))
    for m in re.finditer(rf"\bFROM\s+\(*\s*{_IDENT_PATTERN}", sql_no_strings, re.IGNORECASE):
        add(_extract_ident(m))

    # Comma-separated tables after FROM: FROM T1, T2 WHERE ...
    for m in re.finditer(
        r"\bFROM\s+([^;]+?)(?:\bWHERE\b|\bGROUP\s+BY\b|\bORDER\s+BY\b|\bHAVING\b|\bUNION\b|\bPIVOT\b|$)",
        sql_no_strings,
        re.IGNORECASE | re.DOTALL,
    ):
        clause = m.group(1)
        # Skip subqueries for simple comma split
        if "SELECT" not in clause.upper() and "JOIN" not in clause.upper():
            for part in clause.split(","):
                pm = re.match(rf"^\s*\(*\s*{_IDENT_PATTERN}", part)
                if pm:
                    add(_extract_ident(pm))

    # JOIN <ident>
    for m in re.finditer(rf"\bJOIN\s+\(*\s*{_IDENT_PATTERN}", sql_no_strings, re.IGNORECASE):
        add(_extract_ident(m))

    # INSERT INTO <ident>
    for m in re.finditer(rf"\bINSERT\s+INTO\s+{_IDENT_PATTERN}", sql_no_strings, re.IGNORECASE):
        add(_extract_ident(m))

    # UPDATE <ident>
    for m in re.finditer(rf"\bUPDATE\s+\(*\s*{_IDENT_PATTERN}", sql_no_strings, re.IGNORECASE):
        add(_extract_ident(m))

    return sources


def parse_access_sql(
    sql: str,
    query_type: str = "",
    is_passthrough: bool = False,
) -> ParsedSqlInfo:
    """Statically analyze an Access SQL statement."""
    if not sql or not sql.strip():
        return ParsedSqlInfo(is_passthrough=is_passthrough)

    raw = sql.strip()
    has_parameters = bool(re.match(r"^\s*PARAMETERS\b", raw, re.IGNORECASE))

    # Strip PARAMETERS ...; preamble to determine main statement type
    body = re.sub(r"^\s*PARAMETERS\b.*?;", "", raw, count=1, flags=re.IGNORECASE | re.DOTALL).strip()
    body_no_str = _strip_sql_strings(body)
    upper_body = body_no_str.upper()

    is_crosstab = query_type == "Crosstab" or upper_body.startswith("TRANSFORM")
    is_union = query_type == "Union" or bool(re.search(r"\bUNION\b", upper_body))

    target_table: str | None = None
    statement_type = "SELECT"
    is_action_query = False

    if is_passthrough or query_type == "PassThrough":
        statement_type = "PASSTHROUGH"
        is_passthrough = True
    elif is_crosstab:
        statement_type = "CROSSTAB"
    elif upper_body.startswith("INSERT"):
        statement_type = "INSERT"
        is_action_query = True
        m = re.search(rf"\bINSERT\s+INTO\s+{_IDENT_PATTERN}", body_no_str, re.IGNORECASE)
        if m:
            target_table = _extract_ident(m)
    elif upper_body.startswith("UPDATE"):
        statement_type = "UPDATE"
        is_action_query = True
        m = re.search(rf"\bUPDATE\s+\(*\s*{_IDENT_PATTERN}", body_no_str, re.IGNORECASE)
        if m:
            target_table = _extract_ident(m)
    elif upper_body.startswith("DELETE"):
        statement_type = "DELETE"
        is_action_query = True
        m = re.search(rf"\bFROM\s+\(*\s*{_IDENT_PATTERN}", body_no_str, re.IGNORECASE)
        if m:
            target_table = _extract_ident(m)
    elif re.match(r"^(CREATE|ALTER|DROP)\b", upper_body):
        statement_type = "DDL"
        is_action_query = True
    elif re.search(r"\bSELECT\b.+\bINTO\s+", upper_body, re.DOTALL) and not upper_body.startswith("INSERT"):
        statement_type = "MAKE_TABLE"
        is_action_query = True
        m = re.search(rf"\bINTO\s+{_IDENT_PATTERN}", body_no_str, re.IGNORECASE)
        if m:
            target_table = _extract_ident(m)
    elif is_union:
        statement_type = "UNION"
    elif upper_body.startswith("SELECT"):
        statement_type = "SELECT"
    else:
        statement_type = query_type.upper() if query_type else "UNKNOWN"

    sources = _extract_from_and_join_sources(body_no_str)

    joins: list[str] = []
    for jm in _JOIN_RE.finditer(body_no_str):
        jnorm = re.sub(r"\s+", " ", jm.group(1).upper())
        if jnorm not in joins:
            joins.append(jnorm)

    access_funcs: list[str] = []
    for fn in _ACCESS_FUNCTIONS:
        if re.search(rf"\b{re.escape(fn)}\s*\(", body_no_str, re.IGNORECASE):
            access_funcs.append(fn)

    form_refs: list[str] = []
    for fm in _FORM_REF_RE.finditer(raw):
        fname = (fm.group(1) or fm.group(2) or "").strip()
        cname = (fm.group(3) or fm.group(4) or "").strip()
        ref = f"{fname}!{cname}" if cname else fname
        if ref and ref not in form_refs:
            form_refs.append(ref)

    report_refs: list[str] = []
    for rm in _REPORT_REF_RE.finditer(raw):
        rname = (rm.group(1) or rm.group(2) or "").strip()
        cname = (rm.group(3) or rm.group(4) or "").strip()
        ref = f"{rname}!{cname}" if cname else rname
        if ref and ref not in report_refs:
            report_refs.append(ref)

    tempvars: list[str] = []
    for tm in _TEMPVAR_REF_RE.finditer(raw):
        tv = (tm.group(1) or tm.group(2) or "").strip()
        if tv and tv not in tempvars:
            tempvars.append(tv)

    # Subquery detection: more than one SELECT keyword outside strings (excluding UNION branches)
    select_count = len(re.findall(r"\bSELECT\b", upper_body))
    union_count = len(re.findall(r"\bUNION\b", upper_body))
    has_subquery = select_count > (1 + union_count)

    has_wildcard_like = bool(_WILDCARD_LIKE_RE.search(raw))

    return ParsedSqlInfo(
        statement_type=statement_type,
        referenced_sources=sources,
        target_table=target_table,
        joins=joins,
        access_functions=access_funcs,
        form_references=form_refs,
        report_references=report_refs,
        tempvar_references=tempvars,
        has_parameters=has_parameters,
        has_subquery=has_subquery,
        has_wildcard_like=has_wildcard_like,
        is_crosstab=is_crosstab,
        is_action_query=is_action_query,
        is_union=is_union,
        is_passthrough=is_passthrough,
    )
