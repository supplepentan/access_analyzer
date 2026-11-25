"""Unit tests for Stage 2 parsers (encoding, SaveAsText, UI, Macro, SQL, VBA)."""

from __future__ import annotations

from access_analyzer.parsers.encoding import decode_access_bytes
from access_analyzer.parsers.macro_parser import parse_macro
from access_analyzer.parsers.saveastext_parser import parse_saveastext
from access_analyzer.parsers.sql_parser import parse_access_sql
from access_analyzer.parsers.ui_parser import parse_form, parse_report
from access_analyzer.parsers.vba_parser import parse_vba_module


def test_decode_access_bytes_encodings() -> None:
    sample = "フォーム名=\"F_顧客管理\"\r\nCaption=\"顧客マスタ\""
    assert decode_access_bytes(sample.encode("utf-16")) == sample.replace("\r\n", "\n")
    assert decode_access_bytes(sample.encode("utf-8-sig")) == sample.replace("\r\n", "\n")
    assert decode_access_bytes(sample.encode("cp932")) == sample.replace("\r\n", "\n")


def test_parse_saveastext_multiline_binary_and_codebehind() -> None:
    raw = '''Version =21
VersionRequired =20
Begin Form
    Caption ="顧客管理"
    RecordSource ="SELECT * "
        "FROM [T_顧客] "
        "WHERE 区分=""法人"""
    PrtDevMode = Begin
        0x001122334455
        0x66778899AABB
    End
    Begin Section
        Name ="詳細"
        Height =3000
        Begin
            Begin TextBox
                Name ="txt顧客名"
                ControlSource ="顧客名"
            End
        End
    End
End
CodeBehindForm
Option Compare Database
Option Explicit
Private Sub Form_Load()
    MsgBox "Loaded"
End Sub
'''
    doc = parse_saveastext(raw)
    assert doc.header_properties["Version"] == "21"
    assert len(doc.root_nodes) == 1
    form = doc.root_nodes[0]
    assert form.node_type == "Form"
    assert form.properties["Caption"] == "顧客管理"
    assert form.properties["RecordSource"] == 'SELECT * FROM [T_顧客] WHERE 区分="法人"'
    assert form.properties["PrtDevMode"].startswith("<binary:")
    assert "Private Sub Form_Load()" in doc.code_behind


def test_parse_form_with_subform_embedded_macro_and_vba() -> None:
    raw = """Version =21
Begin Form
    Caption ="受注入力画面"
    RecordSource ="Q_受注一覧"
    DefaultView =0
    PopUp = NotDefault
    AllowDeletions =0
    OnLoad ="[Event Procedure]"
    Begin
        Begin Section
            Name ="詳細"
            Height =4500
            Begin
                Begin ComboBox
                    Name ="cbo顧客"
                    ControlSource ="顧客ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT 顧客ID, 顧客名 "
                        "FROM T_顧客;"
                    BoundColumn =1
                    ColumnCount =2
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Subform
                    Name ="sub明細"
                    SourceObject ="Form.F_受注明細"
                    LinkMasterFields ="受注ID"
                    LinkChildFields ="受注ID"
                End
                Begin CommandButton
                    Name ="btn請求書"
                    Caption ="請求書印刷"
                    OnClick ="[Embedded Macro]"
                    OnClickEmMacro ="<?xml version=\\"1.0\\" encoding=\\"UTF-16\\" standalone=\\"no\\"?>"
                        "<UserInterfaceMacros xmlns=\\"http://schemas.microsoft.com/office/accessservices/2009/11/application\\">"
                        "<Statements><Action Name=\\"OpenReport\\"><Argument Name=\\"ReportName\\">R_請求書</Argument></Action></Statements>"
                        "</UserInterfaceMacros>"
                End
            End
        End
    End
End
CodeBehindForm
Option Compare Database
Option Explicit

Private Sub cbo顧客_AfterUpdate()
    Me.Requery
End Sub
"""
    # Use Access-style doubled quotes for the embedded XML macro property
    raw_fixed = raw.replace('\\"', '""')
    frm = parse_form("F_受注入力", raw_fixed)
    assert frm.name == "F_受注入力"
    assert frm.caption == "受注入力画面"
    assert frm.record_source == "Q_受注一覧"
    assert frm.default_view == "SingleForm"
    assert frm.popup is True
    assert frm.allow_deletions is False
    assert frm.subforms == ["F_受注明細"]
    assert len(frm.controls) == 3

    btn = next(c for c in frm.controls if c.name == "btn請求書")
    assert "OnClick" in btn.embedded_macros
    assert btn.embedded_macros["OnClick"][0].action_name == "OpenReport"
    assert btn.embedded_macros["OnClick"][0].arguments["ReportName"] == "R_請求書"

    assert frm.vba_module is not None
    assert frm.vba_module.option_explicit is True
    assert [p.name for p in frm.vba_module.procedures] == ["cbo顧客_AfterUpdate"]


def test_parse_report_with_group_levels() -> None:
    raw = """Version =21
Begin Report
    Caption ="月次請求書"
    RecordSource ="Q_受注一覧"
    Begin GroupLevel
        ControlSource ="顧客ID"
        GroupHeader =-1
        GroupFooter =-1
        SortOrder =1
        KeepTogether =1
    End
    Begin Section
        Name ="詳細"
        Begin
            Begin TextBox
                Name ="txt金額合計"
                ControlSource ="=Sum([金額])"
                Format ="#,##0"
            End
        End
    End
End
"""
    rep = parse_report("R_請求書", raw)
    assert rep.caption == "月次請求書"
    assert len(rep.group_levels) == 1
    assert rep.group_levels[0].control_source == "顧客ID"
    assert rep.group_levels[0].sort_order == "Descending"
    assert rep.group_levels[0].group_header is True
    assert rep.controls[0].control_source == "=Sum([金額])"


def test_parse_xml_and_legacy_macros() -> None:
    xml_macro = """<?xml version="1.0" encoding="UTF-16" standalone="no"?>
<UserInterfaceMacros xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application">
  <Macro>
    <Statements>
      <ConditionalBlock>
        <If>
          <Condition>True</Condition>
          <Statements>
            <Action Name="RunCode"><Argument Name="FunctionName">InitApp()</Argument></Action>
            <Action Name="OpenForm"><Argument Name="FormName">F_メニュー</Argument></Action>
          </Statements>
        </If>
      </ConditionalBlock>
    </Statements>
  </Macro>
</UserInterfaceMacros>
"""
    m_xml = parse_macro("AutoExec", xml_macro)
    assert m_xml.is_autoexec is True
    assert m_xml.format == "xml"
    assert m_xml.referenced_forms == ["F_メニュー"]
    assert m_xml.referenced_functions == ["InitApp"]
    assert len(m_xml.actions) == 2

    legacy_macro = """Version =196611
ColumnsShown =8
Begin
    MacroName ="バッチ1"
    Condition ="[Forms]![F_メニュー]![chk有効]=True"
    Action ="OpenQuery"
    Argument ="Q_ステータス更新"
    Comment ="ステータス一括更新"
End
"""
    m_leg = parse_macro("M_バッチ", legacy_macro)
    assert m_leg.format == "legacy"
    assert m_leg.submacros == ["バッチ1"]
    assert m_leg.referenced_queries == ["Q_ステータス更新"]
    assert m_leg.actions[0].comment == "ステータス一括更新"


def test_parse_access_sql_variants() -> None:
    sql_select = """
    PARAMETERS [Forms]![F_受注管理]![txt開始日] DateTime;
    SELECT T_受注.受注ID, Nz(T_顧客.顧客名, "未設定") AS 顧客名, IIf(T_受注.金額 > 10000, "大口", "通常") AS 区分
    FROM T_受注 INNER JOIN T_顧客 ON T_受注.顧客ID = T_顧客.顧客ID
    WHERE T_顧客.顧客名 LIKE "*株式会社*" AND T_受注.受注日 >= [Forms]![F_受注管理]![txt開始日]
      AND T_受注.担当者 = TempVars!LoginUser;
    """
    p = parse_access_sql(sql_select)
    assert p.statement_type == "SELECT"
    assert p.has_parameters is True
    assert p.has_wildcard_like is True
    assert set(p.referenced_sources) == {"T_受注", "T_顧客"}
    assert "INNER JOIN" in p.joins
    assert {"Nz", "IIf"}.issubset(set(p.access_functions))
    assert "F_受注管理!txt開始日" in p.form_references
    assert "LoginUser" in p.tempvar_references

    sql_crosstab = """
    TRANSFORM Sum(T_受注明細.金額) AS 金額の合計
    SELECT T_受注明細.商品CD
    FROM T_受注明細
    GROUP BY T_受注明細.商品CD
    PIVOT Format([受注日],"yyyy/mm");
    """
    pc = parse_access_sql(sql_crosstab, query_type="Crosstab")
    assert pc.is_crosstab is True
    assert pc.statement_type == "CROSSTAB"
    assert "Format" in pc.access_functions

    sql_update = "UPDATE T_受注 SET ステータス = '確定' WHERE 受注ID = 10;"
    pu = parse_access_sql(sql_update, query_type="Update")
    assert pu.is_action_query is True
    assert pu.statement_type == "UPDATE"
    assert pu.target_table == "T_受注"


def test_parse_vba_module_comprehensive() -> None:
    code = """Option Compare Database
Option Explicit

Public Declare PtrSafe Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Public Const APP_ENV As String = "PROD"

Public Function ExportOrders(ByVal targetYear As Long, Optional ByVal includeClosed As Boolean = False) As Boolean
    Dim xl As Object
    Dim rs As DAO.Recordset
    Dim cnt As Long
    Set xl = CreateObject("Excel.Application")
    cnt = DCount("*", "T_受注", "Year(受注日)=" & targetYear)
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Q_受注一覧 WHERE 金額 > 0")
    If Forms!F_受注管理!chk印刷 = True Then
        DoCmd.OpenReport "R_請求書", _
            acViewPreview
    End If
    Call HelperLog("Done")
    ExportOrders = True
End Function

Private Sub HelperLog(ByVal msg As String)
    CurrentDb.Execute "INSERT INTO T_ログ (内容) VALUES ('" & msg & "')"
End Sub
"""
    mod = parse_vba_module("Mod_共通", code)
    assert mod.option_explicit is True
    assert mod.option_compare == "Database"
    assert any(d.kind == "Declare" and d.lib_name == "advapi32.dll" for d in mod.declarations)
    assert len(mod.procedures) == 2

    fn = mod.procedures[0]
    assert fn.name == "ExportOrders"
    assert fn.kind == "Function"
    assert fn.return_type == "Boolean"
    assert len(fn.parameters) == 2
    assert fn.parameters[1].is_optional is True
    assert "Excel.Application" in fn.external_objects
    assert fn.domain_aggregates[0]["function"] == "DCount"
    assert fn.domain_aggregates[0]["domain"] == "T_受注"
    assert any(d["method"] == "OpenReport" and d["target"] == "R_請求書" for d in fn.docmd_calls)
    assert "F_受注管理" in fn.form_references
    assert "HelperLog" in fn.calls
    assert any("Q_受注一覧" in s for s in fn.recordset_sources)

    sub = mod.procedures[1]
    assert any("INSERT INTO T_ログ" in s for s in sub.sql_snippets)
