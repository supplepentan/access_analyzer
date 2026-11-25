"""Parsers for Microsoft Access SaveAsText artifacts, SQL, Macros, and VBA."""

from access_analyzer.parsers.encoding import decode_access_bytes, read_access_text_file
from access_analyzer.parsers.macro_parser import parse_macro, parse_xml_macro_actions
from access_analyzer.parsers.saveastext_parser import SaveAsTextDocument, SaveAsTextNode, parse_saveastext
from access_analyzer.parsers.sql_parser import parse_access_sql
from access_analyzer.parsers.ui_parser import parse_form, parse_report
from access_analyzer.parsers.vba_parser import parse_vba_module

__all__ = [
    "SaveAsTextDocument",
    "SaveAsTextNode",
    "decode_access_bytes",
    "parse_access_sql",
    "parse_form",
    "parse_macro",
    "parse_report",
    "parse_saveastext",
    "parse_vba_module",
    "parse_xml_macro_actions",
    "read_access_text_file",
]
