"""End-to-end tests: raw dump -> Stage 2 analysis -> JSON & Markdown/Mermaid export -> CLI."""

from __future__ import annotations

import json
from pathlib import Path

from access_analyzer.cli import main as cli_main
from access_analyzer.exporters.json_exporter import export_json
from access_analyzer.exporters.markdown_exporter import export_markdown
from access_analyzer.models.analysis import DatabaseAnalysis
from access_analyzer.pipeline import build_analysis_from_raw


def _create_sample_raw_dump(raw_dir: Path) -> Path:
    """Create a realistic Japanese Access raw dump directory for end-to-end testing."""
    for sub in ("forms", "reports", "macros", "modules"):
        (raw_dir / sub).mkdir(parents=True, exist_ok=True)

    dao_schema = {
        "metadata": {
            "file_path": "D:/sample/受注管理.accdb",
            "file_name": "受注管理.accdb",
            "file_size_bytes": 1048576,
            "access_version": "16.0",
            "dao_version": "12.0",
            "startup_form": "F_メニュー",
            "app_title": "受注管理システム",
            "allow_bypass_key": False,
            "properties": {"AppTitle": "受注管理システム", "StartUpForm": "F_メニュー"},
            "vba_references": [
                {
                    "name": "VBA",
                    "description": "Visual Basic For Applications",
                    "full_path": "C:/Program Files/Common Files/Microsoft Shared/VBA/VBA7.1/VBE7.DLL",
                    "guid": "{000204EF-0000-0000-C000-000000000046}",
                    "major": 4,
                    "minor": 2,
                    "is_broken": False,
                    "built_in": True,
                },
                {
                    "name": "LegacyOCX",
                    "description": "Missing Legacy Control",
                    "full_path": "C:/Windows/System32/OLDCTRL.OCX",
                    "guid": "{11111111-2222-3333-4444-555555555555}",
                    "major": 1,
                    "minor": 0,
                    "is_broken": True,
                    "built_in": False,
                },
            ],
            "linked_databases": ["ODBC;DSN=ERP_PROD;DATABASE=inventory"],
            "extracted_at": "2026-11-25T13:00:00+00:00",
        },
        "tables": [
            {
                "name": "T_顧客",
                "kind": "local",
                "record_count": 250,
                "primary_key_fields": ["顧客ID"],
                "fields": [
                    {
                        "name": "顧客ID",
                        "ordinal_position": 0,
                        "dao_type": 4,
                        "type_name": "Long",
                        "python_type": "int",
                        "sqlalchemy_type": "Integer",
                        "required": True,
                        "is_autonumber": True,
                        "is_primary_key": True,
                    },
                    {
                        "name": "顧客名",
                        "ordinal_position": 1,
                        "dao_type": 10,
                        "type_name": "Text",
                        "python_type": "str",
                        "sqlalchemy_type": "String(100)",
                        "size": 100,
                        "required": True,
                    },
                    {
                        "name": "ランク",
                        "ordinal_position": 2,
                        "dao_type": 10,
                        "type_name": "Text",
                        "python_type": "str",
                        "sqlalchemy_type": "String(10)",
                        "size": 10,
                        "lookup": {
                            "display_control": "ComboBox",
                            "row_source_type": "Value List",
                            "row_source": '"A";"B";"C"',
                            "bound_column": 1,
                            "column_count": 1,
                        },
                    },
                ],
                "indexes": [
                    {
                        "name": "PrimaryKey",
                        "primary": True,
                        "unique": True,
                        "fields": [{"name": "顧客ID", "descending": False}],
                    }
                ],
            },
            {
                "name": "T_受注",
                "kind": "local",
                "record_count": 1200,
                "primary_key_fields": ["受注ID"],
                "fields": [
                    {
                        "name": "受注ID",
                        "ordinal_position": 0,
                        "dao_type": 4,
                        "type_name": "Long",
                        "python_type": "int",
                        "sqlalchemy_type": "Integer",
                        "required": True,
                        "is_autonumber": True,
                        "is_primary_key": True,
                    },
                    {
                        "name": "顧客ID",
                        "ordinal_position": 1,
                        "dao_type": 4,
                        "type_name": "Long",
                        "python_type": "int",
                        "sqlalchemy_type": "Integer",
                        "required": True,
                    },
                    {
                        "name": "受注日",
                        "ordinal_position": 2,
                        "dao_type": 8,
                        "type_name": "Date/Time",
                        "python_type": "datetime",
                        "sqlalchemy_type": "DateTime",
                        "required": True,
                    },
                    {
                        "name": "添付資料",
                        "ordinal_position": 3,
                        "dao_type": 101,
                        "type_name": "Attachment",
                        "python_type": "list[bytes]",
                        "sqlalchemy_type": "JSON / 添付ファイル子テーブル",
                        "is_attachment": True,
                    },
                ],
                "indexes": [
                    {
                        "name": "PrimaryKey",
                        "primary": True,
                        "unique": True,
                        "fields": [{"name": "受注ID", "descending": False}],
                    }
                ],
            },
            {
                "name": "T_受注明細",
                "kind": "local",
                "record_count": 4800,
                "primary_key_fields": ["明細ID"],
                "fields": [
                    {
                        "name": "明細ID",
                        "ordinal_position": 0,
                        "dao_type": 4,
                        "type_name": "Long",
                        "python_type": "int",
                        "sqlalchemy_type": "Integer",
                        "required": True,
                        "is_autonumber": True,
                        "is_primary_key": True,
                    },
                    {
                        "name": "受注ID",
                        "ordinal_position": 1,
                        "dao_type": 4,
                        "type_name": "Long",
                        "python_type": "int",
                        "sqlalchemy_type": "Integer",
                        "required": True,
                    },
                    {
                        "name": "金額",
                        "ordinal_position": 2,
                        "dao_type": 5,
                        "type_name": "Currency",
                        "python_type": "Decimal",
                        "sqlalchemy_type": "Numeric(19, 4)",
                        "required": True,
                    },
                ],
                "indexes": [
                    {
                        "name": "PrimaryKey",
                        "primary": True,
                        "unique": True,
                        "fields": [{"name": "明細ID", "descending": False}],
                    }
                ],
            },
            {
                "name": "T_外部在庫",
                "kind": "linked_odbc",
                "connect": "ODBC;DSN=ERP_PROD;DATABASE=inventory",
                "source_table_name": "dbo.stock",
                "fields": [
                    {
                        "name": "商品CD",
                        "dao_type": 10,
                        "type_name": "Text",
                        "python_type": "str",
                        "sqlalchemy_type": "String(20)",
                        "size": 20,
                    }
                ],
                "indexes": [],
                "primary_key_fields": [],
            },
            {
                "name": "T_未使用アーカイブ",
                "kind": "local",
                "record_count": 0,
                "primary_key_fields": [],
                "fields": [
                    {
                        "name": "メモ",
                        "dao_type": 12,
                        "type_name": "Memo (Long Text)",
                        "python_type": "str",
                        "sqlalchemy_type": "Text",
                    }
                ],
                "indexes": [],
            },
        ],
        "relations": [
            {
                "name": "Rel_顧客_受注",
                "table": "T_顧客",
                "foreign_table": "T_受注",
                "enforce_referential_integrity": True,
                "cascade_update": False,
                "cascade_delete": True,
                "is_one_to_one": False,
                "join_type": "inner",
                "fields": [{"source_field": "顧客ID", "target_field": "顧客ID"}],
            },
            {
                "name": "Rel_受注_明細",
                "table": "T_受注",
                "foreign_table": "T_受注明細",
                "enforce_referential_integrity": True,
                "cascade_update": False,
                "cascade_delete": True,
                "is_one_to_one": False,
                "join_type": "inner",
                "fields": [{"source_field": "受注ID", "target_field": "受注ID"}],
            },
        ],
        "queries": [
            {
                "name": "Q_受注一覧",
                "query_type_code": 0,
                "query_type": "Select",
                "sql": (
                    'SELECT T_受注.受注ID, Nz(T_顧客.顧客名, "") AS 顧客名, T_受注.受注日 '
                    "FROM T_受注 INNER JOIN T_顧客 ON T_受注.顧客ID = T_顧客.顧客ID "
                    'WHERE T_顧客.顧客名 LIKE "*株式会社*" '
                    "AND T_受注.受注日 >= [Forms]![F_受注管理]![txt開始日];"
                ),
                "description": "受注一覧検索用クエリ",
            },
            {
                "name": "Q_月別売上クロス",
                "query_type_code": 16,
                "query_type": "Crosstab",
                "sql": (
                    "TRANSFORM Sum(T_受注明細.金額) AS 合計 "
                    "SELECT T_受注.顧客ID FROM T_受注 INNER JOIN T_受注明細 ON T_受注.受注ID = T_受注明細.受注ID "
                    'GROUP BY T_受注.顧客ID PIVOT Format(T_受注.受注日, "yyyy/mm");'
                ),
            },
            {
                "name": "Q_ステータス更新",
                "query_type_code": 48,
                "query_type": "Update",
                "sql": "UPDATE T_受注 SET 受注日 = Date() WHERE 受注ID = 1;",
                "returns_records": False,
            },
        ],
    }

    (raw_dir / "dao_schema.json").write_text(
        json.dumps(dao_schema, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # Forms
    (raw_dir / "forms" / "0001_F_メニュー.txt").write_text(
        """Version =21
Begin Form
    Caption ="メインメニュー"
    Begin Section
        Name ="詳細"
        Begin
            Begin CommandButton
                Name ="btn受注管理"
                Caption ="受注管理を開く"
                OnClick ="[Embedded Macro]"
                OnClickEmMacro ="<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?><UserInterfaceMacros><Statements><Action Name=""OpenForm""><Argument Name=""FormName"">F_受注管理</Argument></Action></Statements></UserInterfaceMacros>"
            End
            Begin CommandButton
                Name ="btn請求書"
                Caption ="請求書印刷"
                OnClick ="[Event Procedure]"
            End
        End
    End
End
CodeBehindForm
Option Compare Database
Option Explicit
Private Sub btn請求書_Click()
    DoCmd.OpenReport "R_請求書", acViewPreview
End Sub
""",
        encoding="utf-16",
    )

    (raw_dir / "forms" / "0002_F_受注管理.txt").write_text(
        """Version =21
Begin Form
    Caption ="受注管理"
    RecordSource ="Q_受注一覧"
    Begin Section
        Name ="詳細"
        Begin
            Begin TextBox
                Name ="txt開始日"
            End
            Begin Subform
                Name ="sub明細"
                SourceObject ="Form.F_受注明細"
                LinkMasterFields ="受注ID"
                LinkChildFields ="受注ID"
            End
            Begin CommandButton
                Name ="btnExcel出力"
                OnClick ="[Event Procedure]"
            End
        End
    End
End
CodeBehindForm
Option Compare Database
Option Explicit
Private Sub btnExcel出力_Click()
    Dim cName As String
    cName = DLookup("顧客名", "T_顧客", "顧客ID=1")
    Call ExportToExcel("Q_月別売上クロス")
    DoCmd.OpenQuery "Q_ステータス更新"
End Sub
""",
        encoding="utf-8",
    )

    (raw_dir / "forms" / "0003_F_受注明細.txt").write_text(
        """Version =21
Begin Form
    Caption ="受注明細"
    RecordSource ="T_受注明細"
    DefaultView =1
    Begin Section
        Name ="詳細"
        Begin
            Begin TextBox
                Name ="txt金額"
                ControlSource ="金額"
            End
        End
    End
End
""",
        encoding="utf-8",
    )

    # Report
    (raw_dir / "reports" / "0001_R_請求書.txt").write_text(
        """Version =21
Begin Report
    Caption ="請求書"
    RecordSource ="Q_受注一覧"
    Begin GroupLevel
        ControlSource ="顧客名"
        GroupHeader =-1
        GroupFooter =-1
    End
    Begin Section
        Name ="詳細"
        Begin
            Begin TextBox
                Name ="txt顧客名"
                ControlSource ="顧客名"
            End
        End
    End
End
""",
        encoding="utf-8",
    )

    # Macro
    (raw_dir / "macros" / "0001_AutoExec.txt").write_text(
        """<?xml version="1.0" encoding="UTF-16" standalone="no"?>
<UserInterfaceMacros>
  <Macro>
    <Statements>
      <Action Name="RunCode"><Argument Name="FunctionName">InitApp()</Argument></Action>
    </Statements>
  </Macro>
</UserInterfaceMacros>
""",
        encoding="utf-8",
    )

    # VBA Module
    (raw_dir / "modules" / "0001_Mod_共通.txt").write_text(
        """Option Compare Database
Public Declare PtrSafe Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function InitApp() As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("T_外部在庫")
    InitApp = True
End Function

Public Sub ExportToExcel(ByVal queryName As String)
    Dim xl As Object
    Set xl = CreateObject("Excel.Application")
    CurrentDb.Execute "UPDATE T_受注 SET 受注日 = Date()"
End Sub
""",
        encoding="utf-8",
    )

    manifest = {
        "format_version": "1.0",
        "extracted_at": "2026-11-25T13:00:00+00:00",
        "source_file": "D:/sample/受注管理.accdb",
        "dao_schema_file": "dao_schema.json",
        "object_files": {
            "forms": {
                "F_メニュー": "forms/0001_F_メニュー.txt",
                "F_受注管理": "forms/0002_F_受注管理.txt",
                "F_受注明細": "forms/0003_F_受注明細.txt",
            },
            "reports": {"R_請求書": "reports/0001_R_請求書.txt"},
            "macros": {"AutoExec": "macros/0001_AutoExec.txt"},
            "modules": {"Mod_共通": "modules/0001_Mod_共通.txt"},
        },
        "warnings": [],
    }
    (raw_dir / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return raw_dir


def test_end_to_end_pipeline_and_exporters(tmp_path: Path) -> None:
    raw_dir = _create_sample_raw_dump(tmp_path / "raw")
    analysis = build_analysis_from_raw(raw_dir)

    # Basic counts
    assert len(analysis.tables) == 5
    assert len(analysis.relations) == 2
    assert len(analysis.queries) == 3
    assert len(analysis.forms) == 3
    assert len(analysis.reports) == 1
    assert len(analysis.macros) == 1
    assert len(analysis.vba_modules) == 1

    # Subform flag propagated
    sub_form = next(f for f in analysis.forms if f.name == "F_受注明細")
    assert sub_form.is_subform_target is True
    assert sub_form.default_view == "ContinuousForms"

    # Dependency graph verification
    dg = analysis.dependency_graph
    assert "form:F_メニュー" in dg.entry_points
    assert "macro:AutoExec" in dg.entry_points
    assert "table:T_未使用アーカイブ" in dg.orphan_objects
    assert "table:T_顧客" not in dg.orphan_objects

    edge_triples = {(e.source_id, e.target_id, e.edge_type) for e in dg.edges}
    assert ("form:F_メニュー", "form:F_受注管理", "embedded_macro_open_form") in edge_triples
    assert ("form:F_メニュー", "report:R_請求書", "vba_open_report") in edge_triples
    assert ("form:F_受注管理", "form:F_受注明細", "subform") in edge_triples
    assert ("form:F_受注管理", "vba_module:Mod_共通", "vba_call") in edge_triples
    assert ("macro:AutoExec", "vba_module:Mod_共通", "macro_run_code") in edge_triples
    assert ("vba_module:Mod_共通", "external:Excel.Application", "vba_external_com") in edge_triples

    # Migration hotspots verification
    hotspot_cats = {h.category for h in analysis.migration_inventory.hotspots}
    expected_cats = {
        "BrokenReference",
        "LinkedTable",
        "MissingPrimaryKey",
        "AttachmentField",
        "TableLookupField",
        "CrosstabQuery",
        "ActionQuery",
        "UiBoundSqlParameter",
        "AccessWildcard",
        "Win32ApiDeclare",
        "ExternalComAutomation",
        "EmbeddedMacro",
        "OrphanObjects",
    }
    assert expected_cats.issubset(hotspot_cats)

    # JSON export + roundtrip validation
    out_dir = tmp_path / "output"
    json_files = export_json(analysis, out_dir / "json")
    assert len(json_files) == 6
    reloaded = DatabaseAnalysis.model_validate_json(json_files["analysis"].read_text(encoding="utf-8"))
    assert reloaded.metadata.file_name == "受注管理.accdb"
    assert len(reloaded.tables) == 5

    # Markdown export verification
    md_files = export_markdown(analysis, out_dir / "docs")
    assert len(md_files) == 8
    er_md = md_files["01_er_diagram.md"].read_text(encoding="utf-8")
    assert "```mermaid\nerDiagram" in er_md
    assert "T_顧客" in er_md
    assert "class " in er_md  # SQLAlchemy sketch

    ui_md = md_files["03_forms_ui.md"].read_text(encoding="utf-8")
    assert "```mermaid\nflowchart LR" in ui_md
    assert "F_受注明細" in ui_md

    guide_md = md_files["07_python_migration_guide.md"].read_text(encoding="utf-8")
    assert "ExternalComAutomation" in guide_md
    assert "CrosstabQuery" in guide_md


def test_cli_build_command(tmp_path: Path) -> None:
    raw_dir = _create_sample_raw_dump(tmp_path / "raw_cli")
    out_dir = tmp_path / "out_cli"
    exit_code = cli_main(["build", str(raw_dir), "-o", str(out_dir)])
    assert exit_code == 0
    assert (out_dir / "json" / "analysis.json").exists()
    assert (out_dir / "docs" / "README.md").exists()
    assert (out_dir / "docs" / "07_python_migration_guide.md").exists()
