"""Cross-object dependency graph builder, entry-point reachability, and topological sorter."""

from __future__ import annotations

from collections import defaultdict, deque

from access_analyzer.models.analysis import DependencyEdge, DependencyGraph, ObjectNode, ObjectType
from access_analyzer.models.database import DatabaseMetadata, RelationInfo, TableInfo
from access_analyzer.models.logic import MacroActionInfo, MacroInfo, VbaModuleInfo
from access_analyzer.models.query import QueryInfo
from access_analyzer.models.ui import ControlInfo, FormInfo, ReportInfo
from access_analyzer.parsers.sql_parser import parse_access_sql


def _node_id(obj_type: ObjectType, name: str) -> str:
    return f"{obj_type}:{name}"


def _iter_controls_recursive(controls: list[ControlInfo]):
    for ctrl in controls:
        yield ctrl
        if ctrl.children:
            yield from _iter_controls_recursive(ctrl.children)


def _looks_like_sql(text: str) -> bool:
    up = text.strip().upper()
    return up.startswith(("SELECT ", "TRANSFORM ", "PARAMETERS ", "INSERT ", "UPDATE ", "DELETE ", "WITH "))


def build_dependency_graph(
    metadata: DatabaseMetadata,
    tables: list[TableInfo],
    relations: list[RelationInfo],
    queries: list[QueryInfo],
    forms: list[FormInfo],
    reports: list[ReportInfo],
    macros: list[MacroInfo],
    vba_modules: list[VbaModuleInfo],
) -> DependencyGraph:
    """Build a unified dependency graph across all Access objects."""
    # Case-insensitive lookup maps
    table_map = {t.name.lower(): t.name for t in tables}
    query_map = {q.name.lower(): q.name for q in queries}
    form_map = {f.name.lower(): f.name for f in forms}
    report_map = {r.name.lower(): r.name for r in reports}
    macro_map = {m.name.lower(): m.name for m in macros}
    module_map = {m.name.lower(): m.name for m in vba_modules}

    # Map VBA procedure names (case-insensitive) -> owning node id
    proc_to_node: dict[str, str] = {}
    for mod in vba_modules:
        mod_node = _node_id("vba_module", mod.name)
        for proc in mod.procedures:
            if proc.visibility != "Private":
                proc_to_node.setdefault(proc.name.lower(), mod_node)

    nodes: dict[str, ObjectNode] = {}
    for t in tables:
        nid = _node_id("table", t.name)
        nodes[nid] = ObjectNode(id=nid, object_type="table", name=t.name)
    for q in queries:
        nid = _node_id("query", q.name)
        nodes[nid] = ObjectNode(id=nid, object_type="query", name=q.name)
    for f in forms:
        nid = _node_id("form", f.name)
        nodes[nid] = ObjectNode(id=nid, object_type="form", name=f.name)
    for r in reports:
        nid = _node_id("report", r.name)
        nodes[nid] = ObjectNode(id=nid, object_type="report", name=r.name)
    for m in macros:
        nid = _node_id("macro", m.name)
        nodes[nid] = ObjectNode(id=nid, object_type="macro", name=m.name)
    for mod in vba_modules:
        nid = _node_id("vba_module", mod.name)
        nodes[nid] = ObjectNode(id=nid, object_type="vba_module", name=mod.name)

    edges_set: set[tuple[str, str, str, str]] = set()

    def ensure_external(ext_name: str) -> str:
        nid = _node_id("external", ext_name)
        if nid not in nodes:
            nodes[nid] = ObjectNode(id=nid, object_type="external", name=ext_name)
        return nid

    def add_edge(src: str, dst: str, edge_type: str, detail: str = "") -> None:
        if src == dst or src not in nodes or dst not in nodes:
            return
        edges_set.add((src, dst, edge_type, detail))

    def resolve_data_source(name: str) -> str | None:
        clean = name.strip().strip("[]")
        low = clean.lower()
        if low in table_map:
            return _node_id("table", table_map[low])
        if low in query_map:
            return _node_id("query", query_map[low])
        return None

    def link_sql_or_source(src_node: str, expr: str, edge_type: str, detail: str = "") -> None:
        s = expr.strip()
        if not s:
            return
        if _looks_like_sql(s):
            parsed = parse_access_sql(s)
            for ref in parsed.referenced_sources:
                dst = resolve_data_source(ref)
                if dst:
                    add_edge(src_node, dst, edge_type, detail)
            for fref in parsed.form_references:
                fname = fref.split("!", 1)[0].lower()
                if fname in form_map:
                    add_edge(src_node, _node_id("form", form_map[fname]), "sql_form_param", fref)
        else:
            dst = resolve_data_source(s)
            if dst:
                add_edge(src_node, dst, edge_type, detail)

    def link_macro_actions(src_node: str, actions: list[MacroActionInfo], prefix: str) -> None:
        for act in actions:
            low = act.action_name.lower()
            args = act.arguments
            first_arg = next(iter(args.values()), "") if args else ""
            if low == "openform":
                t = (args.get("FormName") or args.get("Argument0") or first_arg).lower()
                if t in form_map:
                    add_edge(src_node, _node_id("form", form_map[t]), f"{prefix}_open_form")
            elif low == "openreport":
                t = (args.get("ReportName") or args.get("Argument0") or first_arg).lower()
                if t in report_map:
                    add_edge(src_node, _node_id("report", report_map[t]), f"{prefix}_open_report")
            elif low == "openquery":
                t = (args.get("QueryName") or args.get("Argument0") or first_arg).lower()
                if t in query_map:
                    add_edge(src_node, _node_id("query", query_map[t]), f"{prefix}_open_query")
            elif low == "opentable":
                t = (args.get("TableName") or args.get("Argument0") or first_arg).lower()
                if t in table_map:
                    add_edge(src_node, _node_id("table", table_map[t]), f"{prefix}_open_table")
            elif low == "runmacro":
                t = (args.get("MacroName") or args.get("Argument0") or first_arg).split(".")[0].lower()
                if t in macro_map:
                    add_edge(src_node, _node_id("macro", macro_map[t]), f"{prefix}_run_macro")
            elif low == "runcode":
                fn = (args.get("FunctionName") or args.get("Argument0") or first_arg).split("(")[0].strip().lower()
                if fn in proc_to_node:
                    add_edge(src_node, proc_to_node[fn], f"{prefix}_run_code", fn)

    def link_vba_module(src_node: str, mod: VbaModuleInfo) -> None:
        for dec in mod.declarations:
            if dec.kind == "Declare" and dec.lib_name:
                ext = ensure_external(f"Win32:{dec.lib_name}")
                add_edge(src_node, ext, "vba_win32_api", dec.name)

        for proc in mod.procedures:
            for dc in proc.docmd_calls:
                m_low = dc.get("method", "").lower()
                tgt = dc.get("target", "").strip()
                tgt_low = tgt.lower()
                if m_low == "openform" and tgt_low in form_map:
                    add_edge(src_node, _node_id("form", form_map[tgt_low]), "vba_open_form", proc.name)
                elif m_low == "openreport" and tgt_low in report_map:
                    add_edge(src_node, _node_id("report", report_map[tgt_low]), "vba_open_report", proc.name)
                elif m_low == "openquery" and tgt_low in query_map:
                    add_edge(src_node, _node_id("query", query_map[tgt_low]), "vba_open_query", proc.name)
                elif m_low == "opentable" and tgt_low in table_map:
                    add_edge(src_node, _node_id("table", table_map[tgt_low]), "vba_open_table", proc.name)
                elif m_low == "runmacro" and tgt_low.split(".")[0] in macro_map:
                    add_edge(
                        src_node,
                        _node_id("macro", macro_map[tgt_low.split(".")[0]]),
                        "vba_run_macro",
                        proc.name,
                    )
                elif m_low in ("transferspreadsheet", "transfertext", "outputto"):
                    dst = resolve_data_source(tgt)
                    if dst:
                        add_edge(src_node, dst, f"vba_{m_low}", proc.name)

            for sql_snip in proc.sql_snippets:
                link_sql_or_source(src_node, sql_snip, "vba_sql", proc.name)

            for rs_src in proc.recordset_sources:
                link_sql_or_source(src_node, rs_src, "vba_recordset", proc.name)

            for da in proc.domain_aggregates:
                dst = resolve_data_source(da.get("domain", ""))
                if dst:
                    add_edge(src_node, dst, "vba_domain_agg", f"{proc.name}:{da.get('function', '')}")

            for fref in proc.form_references:
                if fref.lower() in form_map:
                    add_edge(src_node, _node_id("form", form_map[fref.lower()]), "vba_form_ref", proc.name)

            for rref in proc.report_references:
                if rref.lower() in report_map:
                    add_edge(src_node, _node_id("report", report_map[rref.lower()]), "vba_report_ref", proc.name)

            for ext_name in proc.external_objects:
                ext = ensure_external(ext_name)
                add_edge(src_node, ext, "vba_external_com", proc.name)

            for callee in proc.calls:
                target_mod = proc_to_node.get(callee.lower())
                if target_mod and target_mod != src_node:
                    add_edge(src_node, target_mod, "vba_call", f"{proc.name}->{callee}")

    # 1. Relations (foreign_table depends on table)
    for rel in relations:
        parent = table_map.get(rel.table.lower())
        child = table_map.get(rel.foreign_table.lower())
        if parent and child:
            add_edge(_node_id("table", child), _node_id("table", parent), "relation", rel.name)

    # 2. Table lookup fields
    for tbl in tables:
        src = _node_id("table", tbl.name)
        for fld in tbl.fields:
            if fld.lookup and fld.lookup.row_source:
                link_sql_or_source(src, fld.lookup.row_source, "table_lookup", fld.name)

    # 3. Queries
    for q in queries:
        src = _node_id("query", q.name)
        for ref in q.parsed.referenced_sources:
            dst = resolve_data_source(ref)
            if dst:
                add_edge(src, dst, "sql_source")
        for fref in q.parsed.form_references:
            fname = fref.split("!", 1)[0].lower()
            if fname in form_map:
                add_edge(src, _node_id("form", form_map[fname]), "sql_form_param", fref)
        for rref in q.parsed.report_references:
            rname = rref.split("!", 1)[0].lower()
            if rname in report_map:
                add_edge(src, _node_id("report", report_map[rname]), "sql_report_param", rref)

    # 4. Forms
    subform_targets: set[str] = set()
    for frm in forms:
        src = _node_id("form", frm.name)
        if frm.record_source:
            link_sql_or_source(src, frm.record_source, "record_source")

        for sub in frm.subforms:
            sub_low = sub.lower()
            if sub_low in form_map:
                sub_actual = form_map[sub_low]
                subform_targets.add(sub_actual.lower())
                add_edge(src, _node_id("form", sub_actual), "subform")
            else:
                dst = resolve_data_source(sub)
                if dst:
                    add_edge(src, dst, "subform_source")

        for ctrl in _iter_controls_recursive(frm.controls):
            if ctrl.row_source and ctrl.row_source_type.lower() != "value list":
                link_sql_or_source(src, ctrl.row_source, "control_row_source", ctrl.name)
            for ev_val in ctrl.events.values():
                if ev_val and ev_val not in ("[Event Procedure]", "[Embedded Macro]"):
                    if ev_val.startswith("="):
                        fn = ev_val[1:].split("(")[0].strip().lower()
                        if fn in proc_to_node:
                            add_edge(src, proc_to_node[fn], "event_function", ctrl.name)
                    elif ev_val.split(".")[0].lower() in macro_map:
                        add_edge(src, _node_id("macro", macro_map[ev_val.split(".")[0].lower()]), "event_macro", ctrl.name)
            for actions in ctrl.embedded_macros.values():
                link_macro_actions(src, actions, "embedded_macro")

        for ev_val in frm.events.values():
            if ev_val and ev_val not in ("[Event Procedure]", "[Embedded Macro]"):
                if ev_val.startswith("="):
                    fn = ev_val[1:].split("(")[0].strip().lower()
                    if fn in proc_to_node:
                        add_edge(src, proc_to_node[fn], "event_function")
                elif ev_val.split(".")[0].lower() in macro_map:
                    add_edge(src, _node_id("macro", macro_map[ev_val.split(".")[0].lower()]), "event_macro")

        for actions in frm.embedded_macros.values():
            link_macro_actions(src, actions, "embedded_macro")

        if frm.vba_module is not None:
            link_vba_module(src, frm.vba_module)

    for frm in forms:
        if frm.name.lower() in subform_targets:
            frm.is_subform_target = True

    # 5. Reports
    subreport_targets: set[str] = set()
    for rep in reports:
        src = _node_id("report", rep.name)
        if rep.record_source:
            link_sql_or_source(src, rep.record_source, "record_source")
        for sub in rep.subreports:
            sub_low = sub.lower()
            if sub_low in report_map:
                sub_actual = report_map[sub_low]
                subreport_targets.add(sub_actual.lower())
                add_edge(src, _node_id("report", sub_actual), "subreport")
        for ctrl in _iter_controls_recursive(rep.controls):
            if ctrl.row_source and ctrl.row_source_type.lower() != "value list":
                link_sql_or_source(src, ctrl.row_source, "control_row_source", ctrl.name)
            for actions in ctrl.embedded_macros.values():
                link_macro_actions(src, actions, "embedded_macro")
        for actions in rep.embedded_macros.values():
            link_macro_actions(src, actions, "embedded_macro")
        if rep.vba_module is not None:
            link_vba_module(src, rep.vba_module)

    for rep in reports:
        if rep.name.lower() in subreport_targets:
            rep.is_subreport_target = True

    # 6. Macros
    for mac in macros:
        src = _node_id("macro", mac.name)
        link_macro_actions(src, mac.actions, "macro")

    # 7. Standalone VBA Modules
    for mod in vba_modules:
        src = _node_id("vba_module", mod.name)
        link_vba_module(src, mod)

    edges = [
        DependencyEdge(source_id=s, target_id=t, edge_type=et, detail=d)
        for s, t, et, d in sorted(edges_set)
    ]

    # Degree counts & adjacency
    out_adj: dict[str, set[str]] = defaultdict(set)
    in_adj: dict[str, set[str]] = defaultdict(set)
    for e in edges:
        out_adj[e.source_id].add(e.target_id)
        in_adj[e.target_id].add(e.source_id)

    for nid, node in nodes.items():
        node.out_degree = len(out_adj[nid])
        node.in_degree = len(in_adj[nid])

    # Entry points
    entry_points: list[str] = []
    if metadata.startup_form and metadata.startup_form.lower() in form_map:
        ep = _node_id("form", form_map[metadata.startup_form.lower()])
        entry_points.append(ep)
    for mac in macros:
        if mac.is_autoexec or mac.is_autokeys:
            ep = _node_id("macro", mac.name)
            if ep not in entry_points:
                entry_points.append(ep)
    for ep in entry_points:
        if ep in nodes:
            nodes[ep].is_entry_point = True

    # Orphan detection
    orphan_objects: list[str] = []
    if entry_points:
        visited: set[str] = set()
        queue = deque(entry_points)
        while queue:
            cur = queue.popleft()
            if cur in visited:
                continue
            visited.add(cur)
            for nxt in out_adj[cur]:
                if nxt not in visited:
                    queue.append(nxt)
        for nid, node in nodes.items():
            if node.object_type == "external":
                continue
            if nid not in visited:
                node.is_orphan = True
                orphan_objects.append(nid)
    else:
        for nid, node in nodes.items():
            if node.object_type == "external":
                continue
            if node.in_degree == 0 and node.out_degree == 0:
                node.is_orphan = True
                orphan_objects.append(nid)

    # Topological build order (dependencies first: if A uses B, B comes before A)
    # Build reversed graph: edge B -> A when A depends on B
    internal_ids = [nid for nid, n in nodes.items() if n.object_type != "external"]
    dep_in_degree: dict[str, int] = {nid: 0 for nid in internal_ids}
    rev_adj: dict[str, list[str]] = defaultdict(list)

    for src, targets in out_adj.items():
        if src not in dep_in_degree:
            continue
        for tgt in targets:
            if tgt in dep_in_degree:
                rev_adj[tgt].append(src)
                dep_in_degree[src] += 1

    # Prioritize tables -> queries -> vba_module -> form -> report -> macro when tying
    type_rank = {"table": 0, "query": 1, "vba_module": 2, "form": 3, "report": 4, "macro": 5}
    ready = sorted(
        [nid for nid, deg in dep_in_degree.items() if deg == 0],
        key=lambda x: (type_rank.get(nodes[x].object_type, 9), x),
    )
    topo_order: list[str] = []
    while ready:
        cur = ready.pop(0)
        topo_order.append(cur)
        for dependent in rev_adj[cur]:
            dep_in_degree[dependent] -= 1
            if dep_in_degree[dependent] == 0:
                ready.append(dependent)
                ready.sort(key=lambda x: (type_rank.get(nodes[x].object_type, 9), x))

    # Detect cycles among remaining nodes
    cycles: list[list[str]] = []
    remaining = {nid for nid, deg in dep_in_degree.items() if deg > 0}
    if remaining:
        visited_cyc: set[str] = set()
        for start in sorted(remaining):
            if start in visited_cyc:
                continue
            path: list[str] = []
            path_idx: dict[str, int] = {}

            def dfs(u: str) -> None:
                visited_cyc.add(u)
                path_idx[u] = len(path)
                path.append(u)
                for v in sorted(out_adj[u]):
                    if v not in remaining:
                        continue
                    if v in path_idx:
                        cyc = path[path_idx[v] :]
                        if cyc and cyc not in cycles:
                            cycles.append(cyc)
                    elif v not in visited_cyc:
                        dfs(v)
                path.pop()
                del path_idx[u]

            dfs(start)

    return DependencyGraph(
        nodes=sorted(nodes.values(), key=lambda n: (n.object_type, n.name)),
        edges=edges,
        entry_points=entry_points,
        orphan_objects=sorted(orphan_objects),
        topological_order=topo_order,
        cycles=cycles,
    )
