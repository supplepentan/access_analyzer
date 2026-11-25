"""Python reconstruction migration inventory, hotspot detector, and screen blueprint generator."""

from __future__ import annotations

from access_analyzer.models.analysis import (
    DependencyGraph,
    MigrationHotspot,
    MigrationInventory,
    ObjectComplexity,
    ScreenBlueprint,
)
from access_analyzer.models.database import DatabaseMetadata, TableInfo
from access_analyzer.models.logic import MacroInfo, VbaModuleInfo
from access_analyzer.models.query import QueryInfo
from access_analyzer.models.ui import ControlInfo, FormInfo, ReportInfo


def _count_controls(controls: list[ControlInfo]) -> int:
    total = 0
    for c in controls:
        total += 1 + _count_controls(c.children)
    return total


def _count_embedded_macros(controls: list[ControlInfo]) -> int:
    total = 0
    for c in controls:
        total += len(c.embedded_macros) + _count_embedded_macros(c.children)
    return total


def _collect_control_events(controls: list[ControlInfo]) -> list[str]:
    items: list[str] = []
    for c in controls:
        for ev, val in c.events.items():
            items.append(f"{c.name}.{ev} ({val})")
        items.extend(_collect_control_events(c.children))
    return items


def build_migration_inventory(
    metadata: DatabaseMetadata,
    tables: list[TableInfo],
    queries: list[QueryInfo],
    forms: list[FormInfo],
    reports: list[ReportInfo],
    macros: list[MacroInfo],
    vba_modules: list[VbaModuleInfo],
    dep_graph: DependencyGraph,
) -> MigrationInventory:
    """Analyze hotspots, complexity, and Python reconstruction blueprints."""
    hotspots: list[MigrationHotspot] = []
    complexity_list: list[ObjectComplexity] = []

    # 0. Broken VBA references
    for ref in metadata.vba_references:
        if ref.is_broken:
            hotspots.append(
                MigrationHotspot(
                    severity="high",
                    category="BrokenReference",
                    object_id=f"reference:{ref.name}",
                    summary=f"参照設定 '{ref.name}' が破損/未解決です ({ref.full_path})",
                    recommendation="元の外部ライブラリ仕様を確認し、Python標準/サードパーティライブラリで代替してください。",
                )
            )

    # 1. Tables
    linked_count = 0
    for tbl in tables:
        oid = f"table:{tbl.name}"
        if tbl.kind in ("linked_access", "linked_odbc"):
            linked_count += 1
            sev = "high" if tbl.kind == "linked_odbc" else "medium"
            hotspots.append(
                MigrationHotspot(
                    severity=sev,  # type: ignore[arg-type]
                    category="LinkedTable",
                    object_id=oid,
                    summary=f"リンクテーブル ({tbl.kind}): 接続先 `{tbl.connect}` / 元テーブル `{tbl.source_table_name}`",
                    recommendation="移行先の統合DBスキーマに含めるか、外部DB接続(SQLAlchemy Engine)として分離設計してください。",
                )
            )

        if tbl.kind == "local" and not tbl.primary_key_fields:
            hotspots.append(
                MigrationHotspot(
                    severity="medium",
                    category="MissingPrimaryKey",
                    object_id=oid,
                    summary=f"ローカルテーブル '{tbl.name}' に主キー(Primary Key)が定義されていません。",
                    recommendation="ORM(SQLAlchemy等)で扱うため、サロゲートキー(id)または複合主キーを設計してください。",
                )
            )

        lookup_cnt = 0
        for fld in tbl.fields:
            if fld.is_attachment:
                hotspots.append(
                    MigrationHotspot(
                        severity="high",
                        category="AttachmentField",
                        object_id=oid,
                        summary=f"フィールド '{tbl.name}.{fld.name}' は Access 固有の添付ファイル型です。",
                        recommendation="ファイル保存先(ローカル/S3等)と添付メタデータ子テーブルに正規化してください。",
                    )
                )
            if fld.is_multivalue:
                hotspots.append(
                    MigrationHotspot(
                        severity="high",
                        category="MultiValueField",
                        object_id=oid,
                        summary=f"フィールド '{tbl.name}.{fld.name}' は複数値フィールド(MVF)です。",
                        recommendation="多対多の中間テーブルまたは JSON カラムへ正規化してください。",
                    )
                )
            if fld.lookup is not None:
                lookup_cnt += 1
                hotspots.append(
                    MigrationHotspot(
                        severity="medium",
                        category="TableLookupField",
                        object_id=oid,
                        summary=f"フィールド '{tbl.name}.{fld.name}' にテーブルレベルのルックアップ定義があります (RowSource: {fld.lookup.row_source[:60]})。",
                        recommendation="DB層では純粋な外部キーとし、選択肢表示はUI層(Select/ComboBox)へ分離してください。",
                    )
                )

        score = len(tbl.fields) + lookup_cnt * 3 + (5 if tbl.kind != "local" else 0)
        complexity_list.append(
            ObjectComplexity(
                object_id=oid,
                object_type="table",
                name=tbl.name,
                score=score,
                metrics={"fields": len(tbl.fields), "indexes": len(tbl.indexes), "lookups": lookup_cnt},
            )
        )

    # 2. Queries
    for q in queries:
        oid = f"query:{q.name}"
        p = q.parsed
        if p.is_crosstab:
            hotspots.append(
                MigrationHotspot(
                    severity="high",
                    category="CrosstabQuery",
                    object_id=oid,
                    summary=f"クロス集計クエリ (TRANSFORM ... PIVOT): '{q.name}'",
                    recommendation="Python の pandas.pivot_table または SQL の条件付き集計 (CASE WHEN + GROUP BY) で再実装してください。",
                )
            )
        if p.is_passthrough:
            hotspots.append(
                MigrationHotspot(
                    severity="high",
                    category="PassThroughQuery",
                    object_id=oid,
                    summary=f"パススルークエリ: '{q.name}' (接続: {q.connect[:60]})",
                    recommendation="ターゲットDB方言のSQLをそのままSQLAlchemy `text()` / Repository層へ移植してください。",
                )
            )
        if p.is_action_query:
            hotspots.append(
                MigrationHotspot(
                    severity="high",
                    category="ActionQuery",
                    object_id=oid,
                    summary=f"アクションクエリ ({p.statement_type}): '{q.name}' -> 対象 `{p.target_table or '不明'}`",
                    recommendation="トランザクション管理付きの Python サービス関数として明示的に実装してください。",
                )
            )
        if p.form_references or p.tempvar_references:
            refs = ", ".join(p.form_references + [f"TempVars!{t}" for t in p.tempvar_references])
            hotspots.append(
                MigrationHotspot(
                    severity="medium",
                    category="UiBoundSqlParameter",
                    object_id=oid,
                    summary=f"クエリ '{q.name}' がUIコントロール/TempVarsを直接参照しています: {refs}",
                    recommendation="UI依存を排除し、明示的な引数を受け取るパラメータ化クエリ関数へ変換してください。",
                )
            )
        if p.has_wildcard_like:
            hotspots.append(
                MigrationHotspot(
                    severity="medium",
                    category="AccessWildcard",
                    object_id=oid,
                    summary=f"クエリ '{q.name}' で Access 固有ワイルドカード (`*` / `?`) が LIKE に使用されています。",
                    recommendation="標準SQLの `%` / `_` へ変換してください。",
                )
            )
        if p.access_functions:
            domain_in_sql = [f for f in p.access_functions if f.lower().startswith("d") and f.lower() in ("dlookup", "dcount", "dsum", "dmax", "dmin", "davg")]
            hotspots.append(
                MigrationHotspot(
                    severity="medium" if domain_in_sql else "low",
                    category="AccessSqlFunctions",
                    object_id=oid,
                    summary=f"クエリ '{q.name}' で Access 固有関数を使用: {', '.join(p.access_functions)}",
                    recommendation="Nz→COALESCE, IIf→CASE WHEN, DLookup→JOIN/サブクエリ, 日付関数→SQL標準/Python側処理へ置換してください。",
                )
            )

        score = (
            len(p.referenced_sources) * 2
            + len(p.joins) * 3
            + len(p.access_functions) * 2
            + len(p.form_references) * 4
            + (10 if p.is_crosstab else 0)
            + (6 if p.is_action_query else 0)
            + (4 if p.has_subquery else 0)
        )
        complexity_list.append(
            ObjectComplexity(
                object_id=oid,
                object_type="query",
                name=q.name,
                score=score,
                metrics={
                    "sources": len(p.referenced_sources),
                    "joins": len(p.joins),
                    "access_functions": len(p.access_functions),
                    "form_refs": len(p.form_references),
                },
            )
        )

    # Helper to scan VBA module hotspots (used for standard/class AND form/report code-behind)
    def scan_vba_hotspots(owner_id: str, mod: VbaModuleInfo) -> tuple[int, int, int]:
        if not mod.option_explicit and mod.code_line_count > 0:
            hotspots.append(
                MigrationHotspot(
                    severity="low",
                    category="MissingOptionExplicit",
                    object_id=owner_id,
                    summary=f"VBAモジュール '{mod.name}' に `Option Explicit` がありません。",
                    recommendation="暗黙の Variant 変数やタイポに注意して型付けしながら Python へ移植してください。",
                )
            )
        for dec in mod.declarations:
            if dec.kind == "Declare":
                hotspots.append(
                    MigrationHotspot(
                        severity="high",
                        category="Win32ApiDeclare",
                        object_id=owner_id,
                        summary=f"Win32 API 宣言を検出: `{dec.name}` (Lib: {dec.lib_name})",
                        recommendation="OS依存処理を Python 標準ライブラリ (os, pathlib, ctypes, tkinter/PySide 等) で置換してください。",
                    )
                )

        ext_objs: set[str] = set()
        dom_agg_count = 0
        sql_count = 0
        for proc in mod.procedures:
            ext_objs.update(proc.external_objects)
            dom_agg_count += len(proc.domain_aggregates)
            sql_count += len(proc.sql_snippets)

        if ext_objs:
            hotspots.append(
                MigrationHotspot(
                    severity="high",
                    category="ExternalComAutomation",
                    object_id=owner_id,
                    summary=f"外部COM/シェル連携を検出 ({mod.name}): {', '.join(sorted(ext_objs))}",
                    recommendation="Excel操作は openpyxl/pandas、メールは smtplib、ファイル操作は pathlib 等の Pure Python ライブラリへ置換してください。",
                )
            )
        if dom_agg_count >= 3:
            hotspots.append(
                MigrationHotspot(
                    severity="medium",
                    category="DomainAggregatesInVba",
                    object_id=owner_id,
                    summary=f"VBA '{mod.name}' 内で DLookup/DSum 等のドメイン集計関数が {dom_agg_count} 箇所使われています。",
                    recommendation="N+1クエリを避けるため、Repository層の集約SQLまたはJOINクエリへ統合してください。",
                )
            )
        return len(mod.procedures), mod.code_line_count, dom_agg_count + sql_count

    # 3. Forms & Screen Blueprints
    edge_by_src: dict[str, list] = {}
    for e in dep_graph.edges:
        edge_by_src.setdefault(e.source_id, []).append(e)

    total_vba_modules = len(vba_modules)
    total_vba_lines = sum(m.code_line_count for m in vba_modules)
    total_vba_procs = sum(len(m.procedures) for m in vba_modules)

    screen_blueprints: list[ScreenBlueprint] = []
    for frm in forms:
        oid = f"form:{frm.name}"
        ctrl_count = _count_controls(frm.controls)
        em_count = len(frm.embedded_macros) + _count_embedded_macros(frm.controls)
        if em_count > 0:
            hotspots.append(
                MigrationHotspot(
                    severity="medium",
                    category="EmbeddedMacro",
                    object_id=oid,
                    summary=f"フォーム '{frm.name}' に {em_count} 件の埋め込みマクロが含まれています。",
                    recommendation="イベントハンドラとして Python のコントローラ/ビュー層へコード化してください。",
                )
            )

        vba_lines = 0
        vba_procs = 0
        if frm.vba_module is not None:
            total_vba_modules += 1
            total_vba_lines += frm.vba_module.code_line_count
            total_vba_procs += len(frm.vba_module.procedures)
            vba_procs, vba_lines, _ = scan_vba_hotspots(oid, frm.vba_module)

        score = ctrl_count + len(frm.subforms) * 8 + em_count * 4 + vba_procs * 4 + (vba_lines // 20)
        complexity_list.append(
            ObjectComplexity(
                object_id=oid,
                object_type="form",
                name=frm.name,
                score=score,
                metrics={
                    "controls": ctrl_count,
                    "subforms": len(frm.subforms),
                    "embedded_macros": em_count,
                    "vba_procs": vba_procs,
                    "vba_lines": vba_lines,
                },
            )
        )

        # Build screen blueprint
        out_edges = edge_by_src.get(oid, [])
        bound_tables = sorted(
            {e.target_id.split(":", 1)[1] for e in out_edges if e.target_id.startswith("table:")}
        )
        bound_queries = sorted(
            {e.target_id.split(":", 1)[1] for e in out_edges if e.target_id.startswith("query:")}
        )
        opens_forms = sorted(
            {
                e.target_id.split(":", 1)[1]
                for e in out_edges
                if e.target_id.startswith("form:") and "open_form" in e.edge_type
            }
        )
        opens_reports = sorted(
            {
                e.target_id.split(":", 1)[1]
                for e in out_edges
                if e.target_id.startswith("report:") and "open_report" in e.edge_type
            }
        )
        ev_handlers = [f"Form.{k} ({v})" for k, v in frm.events.items()] + _collect_control_events(frm.controls)

        screen_blueprints.append(
            ScreenBlueprint(
                form_name=frm.name,
                caption=frm.caption,
                default_view=frm.default_view,
                primary_source=frm.record_source,
                subforms=frm.subforms,
                bound_tables=bound_tables,
                bound_queries=bound_queries,
                event_handlers=ev_handlers,
                opens_forms=opens_forms,
                opens_reports=opens_reports,
                complexity_score=score,
            )
        )

    # 4. Reports
    for rep in reports:
        oid = f"report:{rep.name}"
        ctrl_count = _count_controls(rep.controls)
        vba_lines = 0
        vba_procs = 0
        if rep.vba_module is not None:
            total_vba_modules += 1
            total_vba_lines += rep.vba_module.code_line_count
            total_vba_procs += len(rep.vba_module.procedures)
            vba_procs, vba_lines, _ = scan_vba_hotspots(oid, rep.vba_module)

        score = ctrl_count + len(rep.group_levels) * 5 + len(rep.subreports) * 8 + vba_procs * 3 + (vba_lines // 20)
        complexity_list.append(
            ObjectComplexity(
                object_id=oid,
                object_type="report",
                name=rep.name,
                score=score,
                metrics={
                    "controls": ctrl_count,
                    "group_levels": len(rep.group_levels),
                    "subreports": len(rep.subreports),
                    "vba_lines": vba_lines,
                },
            )
        )

    # 5. Standalone VBA Modules
    for mod in vba_modules:
        oid = f"vba_module:{mod.name}"
        vba_procs, vba_lines, data_ops = scan_vba_hotspots(oid, mod)
        score = vba_procs * 4 + (vba_lines // 15) + data_ops * 3 + len(mod.declarations) * 2
        complexity_list.append(
            ObjectComplexity(
                object_id=oid,
                object_type="vba_module",
                name=mod.name,
                score=score,
                metrics={
                    "procedures": vba_procs,
                    "code_lines": vba_lines,
                    "declarations": len(mod.declarations),
                    "data_ops": data_ops,
                },
            )
        )

    # 6. Macros
    for mac in macros:
        oid = f"macro:{mac.name}"
        score = len(mac.actions) * 3 + len(mac.submacros) * 2
        complexity_list.append(
            ObjectComplexity(
                object_id=oid,
                object_type="macro",
                name=mac.name,
                score=score,
                metrics={"actions": len(mac.actions), "submacros": len(mac.submacros)},
            )
        )

    # 7. Orphan objects notice
    if dep_graph.orphan_objects:
        hotspots.append(
            MigrationHotspot(
                severity="low",
                category="OrphanObjects",
                object_id="graph:orphans",
                summary=f"エントリーポイントから到達不能または未使用と推定されるオブジェクトが {len(dep_graph.orphan_objects)} 件あります。",
                recommendation="不要資産の移植を避けるため、運用上本当に使用されているか確認してください（動的文字列で開かれるフォーム等は要目視確認）。",
            )
        )

    sev_rank = {"high": 0, "medium": 1, "low": 2}
    hotspots.sort(key=lambda h: (sev_rank[h.severity], h.category, h.object_id))
    complexity_list.sort(key=lambda c: (-c.score, c.object_id))
    screen_blueprints.sort(key=lambda s: (-s.complexity_score, s.form_name))

    recommended_layers = {
        "models (ORM / Pydantic)": [t.name for t in tables if t.kind != "system"],
        "repositories (Queries / Data Access)": [q.name for q in queries],
        "services (Business Logic / Batch)": [m.name for m in vba_modules] + [m.name for m in macros],
        "ui_screens (Top-Level Forms)": [f.name for f in forms if not f.is_subform_target],
        "ui_components (Subforms)": [f.name for f in forms if f.is_subform_target],
        "reports (PDF / Excel / HTML)": [r.name for r in reports],
    }

    return MigrationInventory(
        total_tables=len(tables),
        total_linked_tables=linked_count,
        total_queries=len(queries),
        total_forms=len(forms),
        total_reports=len(reports),
        total_macros=len(macros),
        total_vba_modules=total_vba_modules,
        total_vba_lines=total_vba_lines,
        total_vba_procedures=total_vba_procs,
        hotspots=hotspots,
        complexity_ranking=complexity_list,
        screen_blueprints=screen_blueprints,
        recommended_python_layers=recommended_layers,
    )
