"""Export DatabaseAnalysis to comprehensive Markdown design specifications with Mermaid diagrams."""

from __future__ import annotations

import hashlib
import re
from collections import defaultdict
from pathlib import Path

from access_analyzer.models.analysis import DatabaseAnalysis
from access_analyzer.models.database import TableInfo
from access_analyzer.models.logic import VbaModuleInfo
from access_analyzer.models.ui import ControlInfo


def _md_escape(text: str | None) -> str:
    if not text:
        return ""
    return text.replace("|", "\\|").replace("\n", " ").strip()


def _mermaid_label(text: str) -> str:
    return text.replace('"', "'").replace("\n", " ").replace("[", "(").replace("]", ")")


def _mermaid_id(prefix: str, raw: str) -> str:
    ascii_part = re.sub(r"[^A-Za-z0-9_]", "", raw)
    digest = hashlib.md5(raw.encode("utf-8")).hexdigest()[:6]
    if ascii_part:
        return f"{prefix}_{ascii_part[:20]}_{digest}"
    return f"{prefix}_{digest}"


def _er_entity_id(name: str, index: int) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", name)
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    if not cleaned or not cleaned[0].isalpha():
        return f"TBL_{index:02d}"
    return cleaned


def _er_field_name(name: str, index: int) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", name)
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    if not cleaned or not cleaned[0].isalpha():
        return f"col_{index:02d}"
    return cleaned


def _iter_controls_flat(controls: list[ControlInfo], prefix: str = "") -> list[tuple[str, ControlInfo]]:
    rows: list[tuple[str, ControlInfo]] = []
    for c in controls:
        full_name = f"{prefix}{c.name}" if prefix else c.name
        rows.append((full_name, c))
        if c.children:
            rows.extend(_iter_controls_flat(c.children, prefix=f"{full_name} > "))
    return rows


def _python_class_name(table_name: str, index: int) -> str:
    parts = re.findall(r"[A-Za-z0-9]+", table_name)
    if parts:
        return "".join(p.capitalize() for p in parts)
    return f"TableModel{index:02d}"


def _python_attr_name(field_name: str, index: int) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", field_name).strip("_").lower()
    if not cleaned or not cleaned[0].isalpha():
        return f"field_{index:02d}"
    return cleaned


def _render_readme(analysis: DatabaseAnalysis) -> str:
    meta = analysis.metadata
    inv = analysis.migration_inventory
    high_cnt = sum(1 for h in inv.hotspots if h.severity == "high")
    med_cnt = sum(1 for h in inv.hotspots if h.severity == "medium")
    low_cnt = sum(1 for h in inv.hotspots if h.severity == "low")

    lines = [
        f"# Access 構造解析レポート: `{meta.file_name}`",
        "",
        "## 1. データベース基本情報",
        "",
        "| 項目 | 値 |",
        "| :--- | :--- |",
        f"| ファイル名 | `{meta.file_name}` |",
        f"| ファイルサイズ | {meta.file_size_bytes:,} bytes |",
        f"| Access バージョン | {meta.access_version or '-'} |",
        f"| DAO バージョン | {meta.dao_version or '-'} |",
        f"| アプリタイトル | {_md_escape(meta.app_title) or '-'} |",
        f"| スタートアップフォーム | `{meta.startup_form}` |" if meta.startup_form else "| スタートアップフォーム | - |",
        f"| 抽出日時 (UTC) | {meta.extracted_at or '-'} |",
        "",
    ]

    if meta.linked_databases:
        lines.append("### リンク先データベース")
        lines.append("")
        for ldb in meta.linked_databases:
            lines.append(f"- `{ldb}`")
        lines.append("")

    lines.extend(
        [
            "## 2. オブジェクト規模サマリー",
            "",
            "| カテゴリ | 件数 |",
            "| :--- | ---: |",
            f"| テーブル (うちリンクテーブル) | {inv.total_tables} ({inv.total_linked_tables}) |",
            f"| リレーションシップ | {len(analysis.relations)} |",
            f"| クエリ | {inv.total_queries} |",
            f"| フォーム | {inv.total_forms} |",
            f"| レポート | {inv.total_reports} |",
            f"| マクロ | {inv.total_macros} |",
            f"| VBA モジュール (フォーム/レポート付随含む) | {inv.total_vba_modules} |",
            f"| VBA プロシージャ総数 | {inv.total_vba_procedures} |",
            f"| VBA 実コード行数 | {inv.total_vba_lines:,} |",
            f"| Python 移行ホットスポット (高 / 中 / 低) | {high_cnt} / {med_cnt} / {low_cnt} |",
            "",
            "## 3. エントリーポイント",
            "",
        ]
    )
    if analysis.dependency_graph.entry_points:
        for ep in analysis.dependency_graph.entry_points:
            lines.append(f"- `{ep}`")
    else:
        lines.append("- （明示的なスタートアップフォーム / `AutoExec` マクロは検出されませんでした）")

    lines.extend(
        [
            "",
            "## 4. 設計書ドキュメント一覧",
            "",
            "1. [01_er_diagram.md](./01_er_diagram.md) — ER図（Mermaid）・全テーブル/フィールド/インデックス/Python型定義",
            "2. [02_queries.md](./02_queries.md) — クエリ一覧・SQL解析・Access固有関数/フォーム参照チェック",
            "3. [03_forms_ui.md](./03_forms_ui.md) — 画面遷移図（Mermaid）・フォーム/サブフォーム・コントロール・イベント定義",
            "4. [04_reports.md](./04_reports.md) — レポート定義・グループ化/並べ替え・コントロール構成",
            "5. [05_macros_and_vba.md](./05_macros_and_vba.md) — マクロ・VBA解析・コールグラフ（Mermaid）・外部依存",
            "6. [06_dependency_graph.md](./06_dependency_graph.md) — 全オブジェクト横断依存関係図・逆引き影響表・未使用候補",
            "7. [07_python_migration_guide.md](./07_python_migration_guide.md) — Python 再構築ガイド・ホットスポット・画面設計ブループリント",
            "",
        ]
    )

    if analysis.warnings:
        lines.append("## 5. 解析時の警告")
        lines.append("")
        for w in analysis.warnings:
            lines.append(f"- {w}")
        lines.append("")

    return "\n".join(lines)


def _render_er_diagram(analysis: DatabaseAnalysis) -> str:
    user_tables = [t for t in analysis.tables if t.kind != "system"]
    lines = ["# 01. データベース構造・ER図", ""]

    # Build Mermaid erDiagram
    entity_ids: dict[str, str] = {}
    used_eids: set[str] = set()
    for idx, tbl in enumerate(user_tables, start=1):
        eid = _er_entity_id(tbl.name, idx)
        if eid in used_eids:
            eid = f"{eid}_{idx}"
        used_eids.add(eid)
        entity_ids[tbl.name.lower()] = eid

    fk_fields_by_table: dict[str, set[str]] = defaultdict(set)
    for rel in analysis.relations:
        for fm in rel.fields:
            fk_fields_by_table[rel.foreign_table.lower()].add(fm.target_field.lower())

    if user_tables:
        lines.append("## Mermaid ER図")
        lines.append("")
        lines.append("```mermaid")
        lines.append("erDiagram")
        for tbl in user_tables:
            eid = entity_ids[tbl.name.lower()]
            lines.append(f"    {eid} {{")
            fks = fk_fields_by_table.get(tbl.name.lower(), set())
            for f_idx, fld in enumerate(tbl.fields, start=1):
                ftype = re.sub(r"[^A-Za-z0-9_]", "_", fld.python_type) or "Any"
                fname = _er_field_name(fld.name, f_idx)
                key_tags = []
                if fld.is_primary_key:
                    key_tags.append("PK")
                if fld.name.lower() in fks:
                    key_tags.append("FK")
                key_str = f" {','.join(key_tags)}" if key_tags else ""
                comment = _mermaid_label(fld.name)
                lines.append(f'        {ftype} {fname}{key_str} "{comment}"')
            lines.append("    }")

        for rel in analysis.relations:
            parent_eid = entity_ids.get(rel.table.lower())
            child_eid = entity_ids.get(rel.foreign_table.lower())
            if parent_eid and child_eid:
                card = "||--||" if rel.is_one_to_one else "||--o{"
                label = _mermaid_label(
                    ", ".join(f"{m.source_field}->{m.target_field}" for m in rel.fields) or rel.name
                )
                lines.append(f'    {parent_eid} {card} {child_eid} : "{label}"')
        lines.append("```")
        lines.append("")

        # Entity ID legend if any Japanese/renamed entity ids exist
        lines.append("### エンティティ対応表")
        lines.append("")
        lines.append("| ER図ID | Access テーブル名 | 種別 | レコード数 | 主キー |")
        lines.append("| :--- | :--- | :--- | ---: | :--- |")
        for tbl in user_tables:
            eid = entity_ids[tbl.name.lower()]
            rc = f"{tbl.record_count:,}" if tbl.record_count is not None else "-"
            pks = ", ".join(f"`{k}`" for k in tbl.primary_key_fields) or "（なし）"
            lines.append(f"| `{eid}` | `{tbl.name}` | {tbl.kind} | {rc} | {pks} |")
        lines.append("")

    # Relationships table
    lines.append("## リレーションシップ一覧")
    lines.append("")
    if analysis.relations:
        lines.append("| 名前 | 親テーブル | 子テーブル | 結合項目 | 参照整合性 | 連鎖更新 | 連鎖削除 | 1:1 |")
        lines.append("| :--- | :--- | :--- | :--- | :---: | :---: | :---: | :---: |")
        for rel in analysis.relations:
            fmap = ", ".join(f"`{m.source_field}` → `{m.target_field}`" for m in rel.fields)
            lines.append(
                f"| `{rel.name}` | `{rel.table}` | `{rel.foreign_table}` | {fmap} "
                f"| {'Yes' if rel.enforce_referential_integrity else 'No'} "
                f"| {'Yes' if rel.cascade_update else 'No'} "
                f"| {'Yes' if rel.cascade_delete else 'No'} "
                f"| {'Yes' if rel.is_one_to_one else 'No'} |"
            )
    else:
        lines.append("- リレーションシップ定義はありません。")
    lines.append("")

    # Per-table detail + SQLAlchemy blueprint
    lines.append("## テーブル詳細定義")
    lines.append("")
    for idx, tbl in enumerate(analysis.tables, start=1):
        lines.append(f"### `{tbl.name}` ({tbl.kind})")
        lines.append("")
        if tbl.description:
            lines.append(f"- **説明**: {_md_escape(tbl.description)}")
        if tbl.connect:
            lines.append(f"- **リンク接続文字列**: `{tbl.connect}` (元テーブル: `{tbl.source_table_name}`)")
        if tbl.validation_rule:
            lines.append(f"- **テーブル入力規則**: `{tbl.validation_rule}` ({_md_escape(tbl.validation_text)})")
        if tbl.record_count is not None:
            lines.append(f"- **レコード数**: {tbl.record_count:,}")
        lines.append("")

        lines.append("| # | フィールド名 | DAO型 | Python型 | SQLAlchemy型 | PK | 必須 | 既定値 | ルックアップ / 備考 |")
        lines.append("| ---: | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- |")
        for f_i, fld in enumerate(tbl.fields, start=1):
            notes = []
            if fld.is_autonumber:
                notes.append("AutoNumber")
            if fld.validation_rule:
                notes.append(f"Rule: `{_md_escape(fld.validation_rule)}`")
            if fld.lookup:
                notes.append(f"Lookup({fld.lookup.display_control}): `{_md_escape(fld.lookup.row_source[:50])}`")
            if fld.description:
                notes.append(_md_escape(fld.description))
            lines.append(
                f"| {f_i} | `{fld.name}` | {fld.type_name} | `{fld.python_type}` | `{fld.sqlalchemy_type}` "
                f"| {'PK' if fld.is_primary_key else ''} | {'Yes' if fld.required else ''} "
                f"| `{_md_escape(fld.default_value)}` | {' / '.join(notes)} |"
            )
        lines.append("")

        if tbl.indexes:
            lines.append("#### インデックス")
            lines.append("")
            lines.append("| インデックス名 | Primary | Unique | Foreign | 対象フィールド |")
            lines.append("| :--- | :---: | :---: | :---: | :--- |")
            for ix in tbl.indexes:
                flds = ", ".join(f"`{f.name}`{' DESC' if f.descending else ''}" for f in ix.fields)
                lines.append(
                    f"| `{ix.name}` | {'Yes' if ix.primary else ''} | {'Yes' if ix.unique else ''} "
                    f"| {'Yes' if ix.foreign else ''} | {flds} |"
                )
            lines.append("")

        # Optional data profile
        profiled = [f for f in tbl.fields if f.profile is not None]
        if profiled:
            lines.append("#### データプロファイリング")
            lines.append("")
            lines.append("| フィールド | 非NULL | NULL率 | ユニーク数 | 最小値 | 最大値 | サンプル |")
            lines.append("| :--- | ---: | ---: | ---: | :--- | :--- | :--- |")
            for fld in profiled:
                pr = fld.profile
                assert pr is not None
                samples = ", ".join(pr.sample_values[:5])
                lines.append(
                    f"| `{fld.name}` | {pr.non_null_count:,} | {pr.null_ratio:.1%} | {pr.distinct_count:,} "
                    f"| `{_md_escape(pr.min_value)}` | `{_md_escape(pr.max_value)}` | {_md_escape(samples)} |"
                )
            lines.append("")

        # Python SQLAlchemy model sketch
        lines.extend(_render_sqlalchemy_sketch(tbl, idx))

    return "\n".join(lines)


def _render_sqlalchemy_sketch(tbl: TableInfo, index: int) -> list[str]:
    cls_name = _python_class_name(tbl.name, index)
    lines = [
        "<details>",
        f"<summary>Python モデル案 (SQLAlchemy 2.0): <code>{cls_name}</code></summary>",
        "",
        "```python",
        f"class {cls_name}(Base):",
        f'    __tablename__ = "{tbl.name}"',
        "",
    ]
    for f_idx, fld in enumerate(tbl.fields, start=1):
        attr = _python_attr_name(fld.name, f_idx)
        py_t = fld.python_type
        if not fld.required and not fld.is_primary_key:
            py_t = f"{py_t} | None"
        col_args = [f'"{fld.name}"', fld.sqlalchemy_type.split(" / ")[0]]
        if fld.is_primary_key:
            col_args.append("primary_key=True")
        if fld.is_autonumber:
            col_args.append("autoincrement=True")
        if not fld.required and not fld.is_primary_key:
            col_args.append("nullable=True")
        lines.append(f"    {attr}: Mapped[{py_t}] = mapped_column({', '.join(col_args)})")
    lines.extend(["```", "", "</details>", ""])
    return lines


def _render_queries(analysis: DatabaseAnalysis) -> str:
    lines = ["# 02. クエリ・SQL解析", ""]
    if not analysis.queries:
        lines.append("- クエリ定義はありません。")
        return "\n".join(lines)

    lines.append("## クエリ一覧")
    lines.append("")
    lines.append("| クエリ名 | 種別 | SQL種別 | 参照ソース | Access固有関数 | フォーム参照 | アクション |")
    lines.append("| :--- | :--- | :--- | :--- | :--- | :--- | :---: |")
    for q in analysis.queries:
        p = q.parsed
        srcs = ", ".join(f"`{s}`" for s in p.referenced_sources)
        fns = ", ".join(f"`{f}`" for f in p.access_functions)
        frefs = ", ".join(f"`{r}`" for r in p.form_references)
        lines.append(
            f"| `{q.name}` | {q.query_type} | {p.statement_type} | {srcs} | {fns} | {frefs} "
            f"| {'Yes' if p.is_action_query else ''} |"
        )
    lines.append("")

    lines.append("## クエリ詳細")
    lines.append("")
    for q in analysis.queries:
        p = q.parsed
        lines.append(f"### `{q.name}` ({q.query_type} / {p.statement_type})")
        lines.append("")
        if q.description:
            lines.append(f"- **説明**: {_md_escape(q.description)}")
        if p.target_table:
            lines.append(f"- **更新対象テーブル**: `{p.target_table}`")
        if p.referenced_sources:
            lines.append(f"- **参照ソース**: {', '.join(f'`{s}`' for s in p.referenced_sources)}")
        if p.joins:
            lines.append(f"- **JOIN**: {', '.join(p.joins)}")
        if q.parameters:
            param_str = ", ".join(f"`{pr.name}` ({pr.type_name} / `{pr.python_type}`)" for pr in q.parameters)
            lines.append(f"- **パラメータ**: {param_str}")
        if p.form_references:
            lines.append(f"- **フォーム参照（要引数化）**: {', '.join(f'`{r}`' for r in p.form_references)}")
        if p.tempvar_references:
            lines.append(f"- **TempVars 参照**: {', '.join(f'`{t}`' for t in p.tempvar_references)}")
        if p.access_functions:
            lines.append(f"- **Access固有関数**: {', '.join(f'`{f}`' for f in p.access_functions)}")
        if p.has_wildcard_like:
            lines.append("- **注意**: `LIKE` で Access ワイルドカード（`*` / `?`）を使用")
        lines.append("")
        lines.append("```sql")
        lines.append(q.sql.strip())
        lines.append("```")
        lines.append("")

    return "\n".join(lines)


def _render_forms_ui(analysis: DatabaseAnalysis) -> str:
    lines = ["# 03. フォーム・画面遷移・UI定義", ""]
    if not analysis.forms:
        lines.append("- フォーム定義はありません。")
        return "\n".join(lines)

    # Mermaid Screen Transition & Subform Diagram
    lines.append("## 画面遷移・サブフォーム構成図")
    lines.append("")
    lines.append("```mermaid")
    lines.append("flowchart LR")
    for frm in analysis.forms:
        fid = _mermaid_id("F", frm.name)
        label = _mermaid_label(f"{frm.name}\\n({frm.caption})" if frm.caption else frm.name)
        if frm.is_subform_target:
            lines.append(f'    {fid}["[Sub] {label}"]')
        else:
            lines.append(f'    {fid}["{label}"]')
    for rep in analysis.reports:
        rid = _mermaid_id("R", rep.name)
        lines.append(f'    {rid}(["Report: {_mermaid_label(rep.name)}"])')

    nav_edges_seen: set[tuple[str, str, str]] = set()
    for e in analysis.dependency_graph.edges:
        if e.source_id.startswith("form:") and (
            e.edge_type in ("subform", "vba_open_form", "embedded_macro_open_form", "vba_open_report", "embedded_macro_open_report")
        ):
            s_name = e.source_id.split(":", 1)[1]
            t_type, t_name = e.target_id.split(":", 1)
            s_id = _mermaid_id("F", s_name)
            t_id = _mermaid_id("F" if t_type == "form" else "R", t_name)
            edge_label = "subform" if e.edge_type == "subform" else ("open report" if "report" in e.edge_type else "open form")
            key = (s_id, t_id, edge_label)
            if key not in nav_edges_seen:
                nav_edges_seen.add(key)
                arrow = "-.->" if edge_label == "subform" else "-->"
                lines.append(f'    {s_id} {arrow}|"{edge_label}"| {t_id}')
    lines.append("```")
    lines.append("")

    # Summary table
    lines.append("## フォーム一覧")
    lines.append("")
    lines.append("| フォーム名 | 標題 | ビュー | レコードソース | サブフォーム | コントロール数 | VBA行数 | 備考 |")
    lines.append("| :--- | :--- | :--- | :--- | :--- | ---: | ---: | :--- |")
    for frm in analysis.forms:
        flat_ctrls = _iter_controls_flat(frm.controls)
        vba_lines = frm.vba_module.code_line_count if frm.vba_module else 0
        notes = []
        if analysis.metadata.startup_form and analysis.metadata.startup_form.lower() == frm.name.lower():
            notes.append("StartUp")
        if frm.is_subform_target:
            notes.append("Subform")
        if frm.modal:
            notes.append("Modal")
        if frm.popup:
            notes.append("PopUp")
        lines.append(
            f"| `{frm.name}` | {_md_escape(frm.caption)} | {frm.default_view} "
            f"| `{_md_escape(frm.record_source[:45])}` | {', '.join(f'`{s}`' for s in frm.subforms)} "
            f"| {len(flat_ctrls)} | {vba_lines} | {', '.join(notes)} |"
        )
    lines.append("")

    # Per-form detail
    lines.append("## フォーム詳細仕様")
    lines.append("")
    for frm in analysis.forms:
        lines.append(f"### `{frm.name}`" + (f" — {frm.caption}" if frm.caption else ""))
        lines.append("")
        lines.append(f"- **レコードソース**: `{frm.record_source or '（非連結）'}`")
        lines.append(
            f"- **ビュー / 操作制御**: `{frm.default_view}` / "
            f"追加={frm.allow_additions}, 更新={frm.allow_edits}, 削除={frm.allow_deletions}, "
            f"データ入力専用={frm.data_entry}, Modal={frm.modal}, PopUp={frm.popup}"
        )
        if frm.filter:
            lines.append(f"- **フィルタ**: `{frm.filter}` (FilterOn={frm.filter_on})")
        if frm.order_by:
            lines.append(f"- **並び順**: `{frm.order_by}`")
        if frm.events:
            ev_s = ", ".join(f"`{k}={v}`" for k, v in frm.events.items())
            lines.append(f"- **フォームイベント**: {ev_s}")
        lines.append("")

        flat_ctrls = _iter_controls_flat(frm.controls)
        subform_ctrls = [c for _, c in flat_ctrls if c.subform_source]
        if subform_ctrls:
            lines.append("#### サブフォーム連携")
            lines.append("")
            lines.append("| コントロール名 | ソースオブジェクト | LinkMasterFields | LinkChildFields |")
            lines.append("| :--- | :--- | :--- | :--- |")
            for sc in subform_ctrls:
                lines.append(
                    f"| `{sc.name}` | `{sc.subform_source}` | `{sc.link_master_fields}` | `{sc.link_child_fields}` |"
                )
            lines.append("")

        if flat_ctrls:
            lines.append("#### コントロール一覧")
            lines.append("")
            lines.append("| コントロール名 | 種別 | セクション | ControlSource | RowSource | イベント |")
            lines.append("| :--- | :--- | :--- | :--- | :--- | :--- |")
            for fname, c in flat_ctrls:
                evs = ", ".join(f"{k}:{v}" for k, v in c.events.items())
                rs = _md_escape(c.row_source[:50])
                lines.append(
                    f"| `{_md_escape(fname)}` | {c.control_type} | {_md_escape(c.section)} "
                    f"| `{_md_escape(c.control_source)}` | `{rs}` | {_md_escape(evs)} |"
                )
            lines.append("")

        # Embedded macros on form or controls
        em_rows: list[tuple[str, str, str, str]] = []
        for ev, acts in frm.embedded_macros.items():
            for a in acts:
                em_rows.append(("Form", ev, a.action_name, str(a.arguments)))
        for fname, c in flat_ctrls:
            for ev, acts in c.embedded_macros.items():
                for a in acts:
                    em_rows.append((fname, ev, a.action_name, str(a.arguments)))
        if em_rows:
            lines.append("#### 埋め込みマクロ")
            lines.append("")
            lines.append("| 対象 | イベント | アクション | 引数 |")
            lines.append("| :--- | :--- | :--- | :--- |")
            for target, ev, act_name, args_s in em_rows:
                lines.append(f"| `{_md_escape(target)}` | `{ev}` | `{act_name}` | `{_md_escape(args_s)}` |")
            lines.append("")

        if frm.vba_module:
            lines.append(f"#### フォーム VBA ({len(frm.vba_module.procedures)} プロシージャ / {frm.vba_module.code_line_count} 行)")
            lines.append("")
            lines.extend(_render_vba_proc_table(frm.vba_module))

    return "\n".join(lines)


def _render_reports(analysis: DatabaseAnalysis) -> str:
    lines = ["# 04. レポート（帳票）定義", ""]
    if not analysis.reports:
        lines.append("- レポート定義はありません。")
        return "\n".join(lines)

    lines.append("## レポート一覧")
    lines.append("")
    lines.append("| レポート名 | 標題 | レコードソース | グループ階層数 | サブレポート | コントロール数 | VBA行数 |")
    lines.append("| :--- | :--- | :--- | ---: | :--- | ---: | ---: |")
    for rep in analysis.reports:
        flat_ctrls = _iter_controls_flat(rep.controls)
        vba_lines = rep.vba_module.code_line_count if rep.vba_module else 0
        lines.append(
            f"| `{rep.name}` | {_md_escape(rep.caption)} | `{_md_escape(rep.record_source[:45])}` "
            f"| {len(rep.group_levels)} | {', '.join(f'`{s}`' for s in rep.subreports)} "
            f"| {len(flat_ctrls)} | {vba_lines} |"
        )
    lines.append("")

    for rep in analysis.reports:
        lines.append(f"### `{rep.name}`" + (f" — {rep.caption}" if rep.caption else ""))
        lines.append("")
        lines.append(f"- **レコードソース**: `{rep.record_source or '（非連結）'}`")
        if rep.filter:
            lines.append(f"- **フィルタ**: `{rep.filter}`")
        if rep.order_by:
            lines.append(f"- **並び順**: `{rep.order_by}`")
        lines.append("")

        if rep.group_levels:
            lines.append("#### グループ化・並べ替え")
            lines.append("")
            lines.append("| ControlSource | 並び順 | Header | Footer | KeepTogether |")
            lines.append("| :--- | :--- | :---: | :---: | :--- |")
            for gl in rep.group_levels:
                lines.append(
                    f"| `{gl.control_source}` | {gl.sort_order} "
                    f"| {'Yes' if gl.group_header else 'No'} | {'Yes' if gl.group_footer else 'No'} "
                    f"| {gl.keep_together} |"
                )
            lines.append("")

        flat_ctrls = _iter_controls_flat(rep.controls)
        if flat_ctrls:
            lines.append("#### コントロール・集計式")
            lines.append("")
            lines.append("| コントロール名 | 種別 | セクション | ControlSource / Caption | 書式 |")
            lines.append("| :--- | :--- | :--- | :--- | :--- |")
            for cname, c in flat_ctrls:
                expr = c.control_source or c.caption
                lines.append(
                    f"| `{_md_escape(cname)}` | {c.control_type} | {_md_escape(c.section)} "
                    f"| `{_md_escape(expr)}` | `{_md_escape(c.format)}` |"
                )
            lines.append("")

        if rep.vba_module:
            lines.append(f"#### レポート VBA ({len(rep.vba_module.procedures)} プロシージャ)")
            lines.append("")
            lines.extend(_render_vba_proc_table(rep.vba_module))

    return "\n".join(lines)


def _render_vba_proc_table(mod: VbaModuleInfo) -> list[str]:
    if not mod.procedures:
        return ["- （プロシージャなし）", ""]
    lines = [
        "| プロシージャ | 種別 | 可視性 | 行数 | DoCmd / SQL / ドメイン集計 / 外部連携 |",
        "| :--- | :--- | :--- | ---: | :--- |",
    ]
    for p in mod.procedures:
        highlights = []
        if p.docmd_calls:
            dc_s = ", ".join(f"DoCmd.{d['method']}({d['target']})" for d in p.docmd_calls[:4])
            highlights.append(dc_s)
        if p.sql_snippets:
            highlights.append(f"SQL x{len(p.sql_snippets)}")
        if p.domain_aggregates:
            da_s = ", ".join(f"{d['function']}({d['domain']})" for d in p.domain_aggregates[:3])
            highlights.append(da_s)
        if p.external_objects:
            highlights.append(f"COM: {', '.join(p.external_objects)}")
        lines.append(
            f"| `{p.name}` | {p.kind} | {p.visibility} | {p.line_count} | {_md_escape(' / '.join(highlights))} |"
        )
    lines.append("")
    return lines


def _render_macros_and_vba(analysis: DatabaseAnalysis) -> str:
    lines = ["# 05. マクロ・VBA ロジック解析", ""]

    # VBA References
    lines.append("## VBA 参照設定")
    lines.append("")
    if analysis.metadata.vba_references:
        lines.append("| 名前 | 説明 | バージョン | 破損 | パス |")
        lines.append("| :--- | :--- | :--- | :---: | :--- |")
        for ref in analysis.metadata.vba_references:
            lines.append(
                f"| `{ref.name}` | {_md_escape(ref.description)} | {ref.major}.{ref.minor} "
                f"| {'BROKEN' if ref.is_broken else 'OK'} | `{_md_escape(ref.full_path)}` |"
            )
    else:
        lines.append("- 参照設定情報なし。")
    lines.append("")

    # Macros
    lines.append("## スタンドアロンマクロ")
    lines.append("")
    if analysis.macros:
        for mac in analysis.macros:
            badges = []
            if mac.is_autoexec:
                badges.append("AutoExec")
            if mac.is_autokeys:
                badges.append("AutoKeys")
            badge_str = f" [{', '.join(badges)}]" if badges else ""
            lines.append(f"### `{mac.name}`{badge_str} ({mac.format})")
            lines.append("")
            if mac.actions:
                lines.append("| サブマクロ | 条件 | アクション | 引数 | コメント |")
                lines.append("| :--- | :--- | :--- | :--- | :--- |")
                for act in mac.actions:
                    args_s = ", ".join(f"{k}={v}" for k, v in act.arguments.items())
                    lines.append(
                        f"| `{_md_escape(act.submacro_name)}` | `{_md_escape(act.condition)}` "
                        f"| `{act.action_name}` | `{_md_escape(args_s)}` | {_md_escape(act.comment)} |"
                    )
            else:
                lines.append("- （アクションなし）")
            lines.append("")
    else:
        lines.append("- スタンドアロンマクロはありません。")
        lines.append("")

    # VBA Call Graph (Mermaid)
    vba_call_edges = [
        e for e in analysis.dependency_graph.edges if e.edge_type in ("vba_call", "vba_external_com", "vba_win32_api")
    ]
    if vba_call_edges:
        lines.append("## VBA 呼び出し・外部連携グラフ")
        lines.append("")
        lines.append("```mermaid")
        lines.append("flowchart LR")
        seen_nodes: set[str] = set()
        for e in vba_call_edges:
            for nid in (e.source_id, e.target_id):
                if nid not in seen_nodes:
                    seen_nodes.add(nid)
                    mid = _mermaid_id("V", nid)
                    lines.append(f'    {mid}["{_mermaid_label(nid)}"]')
            s_id = _mermaid_id("V", e.source_id)
            t_id = _mermaid_id("V", e.target_id)
            lbl = _mermaid_label(e.detail or e.edge_type)
            lines.append(f'    {s_id} -->|"{lbl}"| {t_id}')
        lines.append("```")
        lines.append("")

    # All VBA modules (standard/class + form/report code-behind)
    all_mods: list[VbaModuleInfo] = list(analysis.vba_modules)
    all_mods.extend(f.vba_module for f in analysis.forms if f.vba_module is not None)
    all_mods.extend(r.vba_module for r in analysis.reports if r.vba_module is not None)

    lines.append("## VBA モジュール詳細")
    lines.append("")
    if not all_mods:
        lines.append("- VBA モジュールはありません。")
        return "\n".join(lines)

    for mod in all_mods:
        lines.append(f"### `{mod.name}` ({mod.module_type})")
        lines.append("")
        lines.append(
            f"- **行数**: {mod.line_count} 行 (コード行: {mod.code_line_count}) / "
            f"**Option Explicit**: {'Yes' if mod.option_explicit else 'No'} / "
            f"**プロシージャ数**: {len(mod.procedures)}"
        )
        if mod.declarations:
            lines.append("- **モジュールレベル宣言**:")
            for d in mod.declarations:
                lib_s = f" [Lib: {d.lib_name}]" if d.lib_name else ""
                lines.append(f"  - `{d.visibility} {d.kind} {d.name}`{lib_s}")
        lines.append("")
        lines.extend(_render_vba_proc_table(mod))

    return "\n".join(lines)


def _render_dependency_graph(analysis: DatabaseAnalysis) -> str:
    dg = analysis.dependency_graph
    lines = ["# 06. 全オブジェクト横断 依存関係グラフ", ""]

    # Mermaid flowchart grouped by object type
    lines.append("## 横断依存関係図")
    lines.append("")
    if dg.nodes:
        lines.append("```mermaid")
        lines.append("flowchart LR")
        by_type: dict[str, list] = defaultdict(list)
        for n in dg.nodes:
            by_type[n.object_type].append(n)

        subgraph_titles = [
            ("macro", "Macros"),
            ("form", "Forms"),
            ("report", "Reports"),
            ("vba_module", "VBA Modules"),
            ("query", "Queries"),
            ("table", "Tables"),
            ("external", "External Dependencies"),
        ]
        for obj_t, title in subgraph_titles:
            group = by_type.get(obj_t, [])
            if not group:
                continue
            lines.append(f'    subgraph SG_{obj_t}["{title}"]')
            for n in group:
                mid = _mermaid_id("N", n.id)
                tag = " [ENTRY]" if n.is_entry_point else (" [ORPHAN]" if n.is_orphan else "")
                lines.append(f'        {mid}["{_mermaid_label(n.name + tag)}"]')
            lines.append("    end")

        seen_pairs: set[tuple[str, str, str]] = set()
        for e in dg.edges:
            key = (e.source_id, e.target_id, e.edge_type)
            if key in seen_pairs:
                continue
            seen_pairs.add(key)
            s_id = _mermaid_id("N", e.source_id)
            t_id = _mermaid_id("N", e.target_id)
            lines.append(f'    {s_id} -->|"{_mermaid_label(e.edge_type)}"| {t_id}')
        lines.append("```")
        lines.append("")

    # Build order
    lines.append("## 推奨実装順序（トポロジカル順：依存元から順に構築）")
    lines.append("")
    if dg.topological_order:
        for i, nid in enumerate(dg.topological_order, start=1):
            lines.append(f"{i}. `{nid}`")
    else:
        lines.append("- 対象オブジェクトなし。")
    lines.append("")

    if dg.cycles:
        lines.append("## 循環参照の検出")
        lines.append("")
        for cyc in dg.cycles:
            lines.append(f"- {' → '.join(f'`{c}`' for c in cyc)} → `{cyc[0]}`")
        lines.append("")

    # Reverse impact matrix
    uses_map: dict[str, list[str]] = defaultdict(list)
    used_by_map: dict[str, list[str]] = defaultdict(list)
    for e in dg.edges:
        uses_map[e.source_id].append(f"`{e.target_id}` ({e.edge_type})")
        used_by_map[e.target_id].append(f"`{e.source_id}` ({e.edge_type})")

    lines.append("## 逆引き影響分析マトリクス")
    lines.append("")
    lines.append("| オブジェクト | 種別 | 利用先 (Uses) | 被参照元 (Used By) | 状態 |")
    lines.append("| :--- | :--- | :--- | :--- | :--- |")
    for n in dg.nodes:
        status = "EntryPoint" if n.is_entry_point else ("Orphan?" if n.is_orphan else "Active")
        uses_s = ", ".join(sorted(set(uses_map.get(n.id, [])))[:8]) or "-"
        used_s = ", ".join(sorted(set(used_by_map.get(n.id, [])))[:8]) or "-"
        lines.append(f"| `{n.name}` | {n.object_type} | {uses_s} | {used_s} | {status} |")
    lines.append("")

    if dg.orphan_objects:
        lines.append("## 未使用・孤立候補オブジェクト")
        lines.append("")
        for oid in dg.orphan_objects:
            lines.append(f"- `{oid}`")
        lines.append("")

    return "\n".join(lines)


def _render_python_migration_guide(analysis: DatabaseAnalysis) -> str:
    inv = analysis.migration_inventory
    lines = [
        "# 07. Python 再構築ガイド",
        "",
        "Access アプリケーションを Python で再構築するための移行ホットスポット、推奨レイヤー構成、画面設計ブループリント、および複雑度ランキングです。",
        "",
        "## 1. Python 移行ホットスポット（要注意箇所）",
        "",
    ]

    if inv.hotspots:
        lines.append("| 重要度 | カテゴリ | 対象 | 内容 | Python 置換方針 |")
        lines.append("| :---: | :--- | :--- | :--- | :--- |")
        for h in inv.hotspots:
            lines.append(
                f"| **{h.severity.upper()}** | `{h.category}` | `{h.object_id}` "
                f"| {_md_escape(h.summary)} | {_md_escape(h.recommendation)} |"
            )
    else:
        lines.append("- 特筆すべき高リスクな移行ホットスポットは検出されませんでした。")
    lines.append("")

    lines.append("## 2. 推奨 Python パッケージ構成マッピング")
    lines.append("")
    for layer, items in inv.recommended_python_layers.items():
        lines.append(f"### `{layer}` ({len(items)} 件)")
        lines.append("")
        if items:
            lines.append(", ".join(f"`{x}`" for x in items))
        else:
            lines.append("- なし")
        lines.append("")

    lines.append("## 3. 画面設計ブループリント（フォーム単位）")
    lines.append("")
    if inv.screen_blueprints:
        for sb in inv.screen_blueprints:
            lines.append(f"### `{sb.form_name}`" + (f" ({sb.caption})" if sb.caption else "") + f" — 複雑度: {sb.complexity_score}")
            lines.append("")
            lines.append(f"- **ビュー形式**: `{sb.default_view}` / **主データソース**: `{sb.primary_source or '非連結'}`")
            if sb.subforms:
                lines.append(f"- **埋め込みサブフォーム**: {', '.join(f'`{s}`' for s in sb.subforms)}")
            if sb.bound_tables:
                lines.append(f"- **関連テーブル**: {', '.join(f'`{t}`' for t in sb.bound_tables)}")
            if sb.bound_queries:
                lines.append(f"- **関連クエリ**: {', '.join(f'`{q}`' for q in sb.bound_queries)}")
            if sb.opens_forms:
                lines.append(f"- **遷移先フォーム**: {', '.join(f'`{f}`' for f in sb.opens_forms)}")
            if sb.opens_reports:
                lines.append(f"- **出力レポート**: {', '.join(f'`{r}`' for r in sb.opens_reports)}")
            if sb.event_handlers:
                lines.append(f"- **イベントハンドラ ({len(sb.event_handlers)}件)**: {', '.join(f'`{e}`' for e in sb.event_handlers[:10])}")
            lines.append("")
    else:
        lines.append("- フォームはありません。")
        lines.append("")

    lines.append("## 4. オブジェクト複雑度ランキング（上位20件）")
    lines.append("")
    lines.append("| 順位 | オブジェクト | 種別 | スコア | 内訳メトリクス |")
    lines.append("| ---: | :--- | :--- | ---: | :--- |")
    for rank, item in enumerate(inv.complexity_ranking[:20], start=1):
        m_str = ", ".join(f"{k}={v}" for k, v in item.metrics.items())
        lines.append(f"| {rank} | `{item.name}` | {item.object_type} | {item.score} | {m_str} |")
    lines.append("")

    return "\n".join(lines)


def export_markdown(analysis: DatabaseAnalysis, docs_dir: str | Path) -> dict[str, Path]:
    """Generate the full Markdown + Mermaid specification pack inside docs_dir."""
    out = Path(docs_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)

    files = {
        "README.md": _render_readme(analysis),
        "01_er_diagram.md": _render_er_diagram(analysis),
        "02_queries.md": _render_queries(analysis),
        "03_forms_ui.md": _render_forms_ui(analysis),
        "04_reports.md": _render_reports(analysis),
        "05_macros_and_vba.md": _render_macros_and_vba(analysis),
        "06_dependency_graph.md": _render_dependency_graph(analysis),
        "07_python_migration_guide.md": _render_python_migration_guide(analysis),
    }

    written: dict[str, Path] = {}
    for fname, content in files.items():
        path = out / fname
        path.write_text(content, encoding="utf-8")
        written[fname] = path
    return written
