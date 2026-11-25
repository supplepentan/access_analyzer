"""DAO / Access COM constants and Python / SQLAlchemy type mappings."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class FieldTypeMapping:
    dao_code: int
    dao_name: str
    python_type: str
    sqlalchemy_type: str
    is_complex: bool = False
    migration_note: str = ""


DAO_FIELD_TYPES: dict[int, FieldTypeMapping] = {
    1: FieldTypeMapping(1, "Boolean", "bool", "Boolean"),
    2: FieldTypeMapping(2, "Byte", "int", "SmallInteger"),
    3: FieldTypeMapping(3, "Integer", "int", "SmallInteger"),
    4: FieldTypeMapping(4, "Long", "int", "Integer"),
    5: FieldTypeMapping(5, "Currency", "Decimal", "Numeric(19, 4)"),
    6: FieldTypeMapping(6, "Single", "float", "Float"),
    7: FieldTypeMapping(7, "Double", "float", "Float"),
    8: FieldTypeMapping(8, "Date/Time", "datetime", "DateTime"),
    9: FieldTypeMapping(9, "Binary", "bytes", "LargeBinary"),
    10: FieldTypeMapping(10, "Text", "str", "String"),
    11: FieldTypeMapping(11, "LongBinary (OLE Object)", "bytes", "LargeBinary", migration_note="OLEバイナリはファイルストレージ分離を検討"),
    12: FieldTypeMapping(12, "Memo (Long Text)", "str", "Text"),
    15: FieldTypeMapping(15, "GUID", "UUID", "String(36)"),
    16: FieldTypeMapping(16, "BigInt", "int", "BigInteger"),
    17: FieldTypeMapping(17, "VarBinary", "bytes", "LargeBinary"),
    18: FieldTypeMapping(18, "Char", "str", "String"),
    19: FieldTypeMapping(19, "Numeric", "Decimal", "Numeric"),
    20: FieldTypeMapping(20, "Decimal", "Decimal", "Numeric"),
    21: FieldTypeMapping(21, "Float", "float", "Float"),
    22: FieldTypeMapping(22, "Time", "time", "Time"),
    23: FieldTypeMapping(23, "TimeStamp", "datetime", "DateTime"),
    101: FieldTypeMapping(
        101,
        "Attachment",
        "list[bytes]",
        "JSON / 添付ファイル子テーブル",
        is_complex=True,
        migration_note="Access固有の添付ファイル型。子テーブル+オブジェクトストレージへ正規化が必要",
    ),
    102: FieldTypeMapping(102, "ComplexByte", "list[int]", "JSON / 中間テーブル", is_complex=True, migration_note="複数値フィールド(MVF)。中間テーブルへの正規化が必要"),
    103: FieldTypeMapping(103, "ComplexInteger", "list[int]", "JSON / 中間テーブル", is_complex=True, migration_note="複数値フィールド(MVF)。中間テーブルへの正規化が必要"),
    104: FieldTypeMapping(104, "ComplexLong", "list[int]", "JSON / 中間テーブル", is_complex=True, migration_note="複数値フィールド(MVF)。中間テーブルへの正規化が必要"),
    105: FieldTypeMapping(105, "ComplexSingle", "list[float]", "JSON / 中間テーブル", is_complex=True, migration_note="複数値フィールド(MVF)。中間テーブルへの正規化が必要"),
    106: FieldTypeMapping(106, "ComplexDouble", "list[float]", "JSON / 中間テーブル", is_complex=True, migration_note="複数値フィールド(MVF)。中間テーブルへの正規化が必要"),
    107: FieldTypeMapping(107, "ComplexGUID", "list[UUID]", "JSON / 中間テーブル", is_complex=True, migration_note="複数値フィールド(MVF)。中間テーブルへの正規化が必要"),
    108: FieldTypeMapping(108, "ComplexDecimal", "list[Decimal]", "JSON / 中間テーブル", is_complex=True, migration_note="複数値フィールド(MVF)。中間テーブルへの正規化が必要"),
    109: FieldTypeMapping(109, "ComplexText", "list[str]", "JSON / 中間テーブル", is_complex=True, migration_note="複数値フィールド(MVF)。中間テーブルへの正規化が必要"),
}

DAO_QUERY_TYPES: dict[int, str] = {
    0: "Select",
    16: "Crosstab",
    32: "Delete",
    48: "Update",
    64: "Append",
    80: "MakeTable",
    96: "DataDefinition",
    112: "PassThrough",
    128: "Union",
    144: "SPTBulk",
    160: "Compound",
    224: "Procedure",
    240: "Action",
}

# DAO TableDef Attributes
DB_SYSTEM_OBJECT = -2147483648  # 0x80000000
DB_HIDDEN_OBJECT = 1
DB_ATTACHED_TABLE = 1073741824  # 0x40000000
DB_ATTACHED_ODBC = 536870912    # 0x20000000

# DAO Field Attributes
DB_FIXED_FIELD = 1
DB_VARIABLE_FIELD = 2
DB_AUTO_INCR_FIELD = 16
DB_UPDATABLE_FIELD = 32
DB_SYSTEM_FIELD = 8192
DB_HYPERLINK_FIELD = 32768

# DAO Relation Attributes
DB_RELATION_UNIQUE = 1
DB_RELATION_DONT_ENFORCE = 2
DB_RELATION_INHERITED = 4
DB_RELATION_UPDATE_CASCADE = 256
DB_RELATION_DELETE_CASCADE = 4096
DB_RELATION_LEFT = 16777216
DB_RELATION_RIGHT = 33554432

# Access SaveAsText Object Types (AcObjectType)
AC_TABLE = 0
AC_QUERY = 1
AC_FORM = 2
AC_REPORT = 3
AC_MACRO = 4
AC_MODULE = 5

# Access Quit Options
AC_QUIT_SAVE_NONE = 2

# Office AutomationSecurity
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# Lookup DisplayControl codes
DISPLAY_CONTROL_MAP: dict[int, str] = {
    106: "CheckBox",
    109: "TextBox",
    110: "ListBox",
    111: "ComboBox",
}

# Form DefaultView codes
DEFAULT_VIEW_MAP: dict[str, str] = {
    "0": "SingleForm",
    "1": "ContinuousForms",
    "2": "Datasheet",
    "3": "PivotTable",
    "4": "PivotChart",
    "5": "SplitForm",
}


def map_dao_field_type(dao_code: int, size: int = 0) -> FieldTypeMapping:
    """Return FieldTypeMapping for a DAO field type code, formatting String(size) when applicable."""
    mapping = DAO_FIELD_TYPES.get(
        dao_code,
        FieldTypeMapping(dao_code, f"Unknown({dao_code})", "Any", "Text"),
    )
    if dao_code in (10, 18) and size > 0:
        return FieldTypeMapping(
            dao_code=mapping.dao_code,
            dao_name=mapping.dao_name,
            python_type=mapping.python_type,
            sqlalchemy_type=f"String({size})",
            is_complex=mapping.is_complex,
            migration_note=mapping.migration_note,
        )
    return mapping
