"""Stage 1: Extract full metadata and SaveAsText definitions from Microsoft Access via COM/DAO."""

from __future__ import annotations

import datetime
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from access_analyzer.extractors.dao_constants import (
    AC_FORM,
    AC_MACRO,
    AC_MODULE,
    AC_QUIT_SAVE_NONE,
    AC_REPORT,
    DAO_QUERY_TYPES,
    DB_ATTACHED_ODBC,
    DB_ATTACHED_TABLE,
    DB_AUTO_INCR_FIELD,
    DB_HIDDEN_OBJECT,
    DB_HYPERLINK_FIELD,
    DB_RELATION_DELETE_CASCADE,
    DB_RELATION_DONT_ENFORCE,
    DB_RELATION_LEFT,
    DB_RELATION_RIGHT,
    DB_RELATION_UNIQUE,
    DB_RELATION_UPDATE_CASCADE,
    DB_SYSTEM_OBJECT,
    DISPLAY_CONTROL_MAP,
    MSO_AUTOMATION_SECURITY_FORCE_DISABLE,
    map_dao_field_type,
)


@dataclass
class ExtractionOptions:
    """Configuration for COM/DAO extraction."""

    password: str | None = None
    include_system_tables: bool = False
    profile_data: bool = False
    profile_sample_limit: int = 5
    profile_max_rows: int = 100_000
    allow_autoexec: bool = False


def _sanitize_filename(name: str, index: int) -> str:
    """Convert an arbitrary Access object name into a safe Windows filename."""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip().strip(".")
    if not cleaned:
        cleaned = "unnamed"
    if len(cleaned) > 80:
        cleaned = cleaned[:80]
    return f"{index:04d}_{cleaned}.txt"


def _serialize_value(value: Any) -> Any:
    """Convert COM/DAO values into JSON-serializable primitives."""
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return value.isoformat()
    if isinstance(value, bytes):
        return f"<bytes:{len(value)}>"
    # pywintypes.datetime exposes isoformat()
    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except Exception:
            pass
    return str(value)


def _safe_prop(obj: Any, prop_name: str, default: Any = None) -> Any:
    """Read a DAO Property by name without raising if absent."""
    try:
        return obj.Properties(prop_name).Value
    except Exception:
        return default


def _safe_attr(obj: Any, attr_name: str, default: Any = None) -> Any:
    """Read a COM attribute safely."""
    try:
        return getattr(obj, attr_name)
    except Exception:
        return default


def _collect_all_properties(obj: Any) -> dict[str, Any]:
    """Read all readable scalar DAO properties from a COM object."""
    props: dict[str, Any] = {}
    try:
        for prop in obj.Properties:
            try:
                name = str(prop.Name)
                val = prop.Value
                if isinstance(val, (bytes, bytearray, memoryview)):
                    continue
                props[name] = _serialize_value(val)
            except Exception:
                continue
    except Exception:
        pass
    return props


class AccessComExtractor:
    """Extract DAO metadata and SaveAsText artifacts from a Microsoft Access database."""

    def __init__(self, options: ExtractionOptions | None = None) -> None:
        self.options = options or ExtractionOptions()

    def extract(self, db_path: str | Path, raw_output_dir: str | Path) -> Path:
        """Open Access via out-of-process COM and dump raw metadata + SaveAsText files.

        Returns the path to ``manifest.json`` inside ``raw_output_dir``.
        """
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise RuntimeError(
                "pywin32 がインストールされていません。Windows 上で `pip install pywin32` を実行してください。"
            ) from exc

        abs_db_path = Path(db_path).resolve()
        if not abs_db_path.exists():
            raise FileNotFoundError(f"Access ファイルが見つかりません: {abs_db_path}")

        out_dir = Path(raw_output_dir).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)
        for sub in ("forms", "reports", "macros", "modules"):
            (out_dir / sub).mkdir(exist_ok=True)

        warnings: list[str] = []
        app = None
        db_opened = False

        pythoncom.CoInitialize()
        try:
            # DispatchEx starts an isolated Access.Application instance so user windows are untouched.
            app = win32com.client.DispatchEx("Access.Application")
            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.UserControl = False
            except Exception:
                pass
            if not self.options.allow_autoexec:
                try:
                    app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
                except Exception as exc:
                    warnings.append(f"AutomationSecurity の設定に失敗しました: {exc}")

            if self.options.password is not None:
                app.OpenCurrentDatabase(str(abs_db_path), False, self.options.password)
            else:
                app.OpenCurrentDatabase(str(abs_db_path), False)
            db_opened = True

            db = app.CurrentDb()
            extracted_at = datetime.datetime.now(datetime.timezone.utc).isoformat()

            metadata = self._extract_metadata(app, db, abs_db_path, extracted_at, warnings)
            tables = self._extract_tables(db, warnings)
            relations = self._extract_relations(db, warnings)
            queries = self._extract_queries(db, warnings)

            # Collect linked database paths from linked tables
            linked_dbs: set[str] = set()
            for tbl in tables:
                connect = tbl.get("connect", "")
                match = re.search(r"DATABASE=([^;]+)", connect, re.IGNORECASE)
                if match:
                    linked_dbs.add(match.group(1).strip())
                elif connect and tbl.get("kind") == "linked_odbc":
                    linked_dbs.add(connect)
            metadata["linked_databases"] = sorted(linked_dbs)

            dao_schema = {
                "metadata": metadata,
                "tables": tables,
                "relations": relations,
                "queries": queries,
            }
            (out_dir / "dao_schema.json").write_text(
                json.dumps(dao_schema, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

            object_files: dict[str, dict[str, str]] = {
                "forms": self._export_saveastext_group(
                    app, db, AC_FORM, "AllForms", "Forms", out_dir / "forms", "forms", warnings
                ),
                "reports": self._export_saveastext_group(
                    app, db, AC_REPORT, "AllReports", "Reports", out_dir / "reports", "reports", warnings
                ),
                "macros": self._export_saveastext_group(
                    app, db, AC_MACRO, "AllMacros", "Scripts", out_dir / "macros", "macros", warnings
                ),
                "modules": self._export_saveastext_group(
                    app, db, AC_MODULE, "AllModules", "Modules", out_dir / "modules", "modules", warnings
                ),
            }

            manifest = {
                "format_version": "1.0",
                "extracted_at": extracted_at,
                "source_file": str(abs_db_path),
                "dao_schema_file": "dao_schema.json",
                "object_files": object_files,
                "warnings": warnings,
            }
            manifest_path = out_dir / "manifest.json"
            manifest_path.write_text(
                json.dumps(manifest, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            return manifest_path
        finally:
            if app is not None:
                if db_opened:
                    try:
                        app.CloseCurrentDatabase()
                    except Exception:
                        pass
                try:
                    app.Quit(AC_QUIT_SAVE_NONE)
                except Exception:
                    pass
                del app
            pythoncom.CoUninitialize()

    def _extract_metadata(
        self,
        app: Any,
        db: Any,
        abs_db_path: Path,
        extracted_at: str,
        warnings: list[str],
    ) -> dict[str, Any]:
        props = _collect_all_properties(db)
        startup_form = _safe_prop(db, "StartUpForm")
        if isinstance(startup_form, str) and startup_form.startswith("Form."):
            startup_form = startup_form[5:]

        vba_refs: list[dict[str, Any]] = []
        try:
            for ref in app.References:
                vba_refs.append(
                    {
                        "name": str(_safe_attr(ref, "Name", "")),
                        "description": str(_safe_attr(ref, "Description", "")),
                        "full_path": str(_safe_attr(ref, "FullPath", "")),
                        "guid": str(_safe_attr(ref, "Guid", "")),
                        "major": int(_safe_attr(ref, "Major", 0) or 0),
                        "minor": int(_safe_attr(ref, "Minor", 0) or 0),
                        "is_broken": bool(_safe_attr(ref, "IsBroken", False)),
                        "built_in": bool(_safe_attr(ref, "BuiltIn", False)),
                    }
                )
        except Exception as exc:
            warnings.append(f"VBA参照設定の取得に失敗しました: {exc}")

        return {
            "file_path": str(abs_db_path),
            "file_name": abs_db_path.name,
            "file_size_bytes": abs_db_path.stat().st_size,
            "access_version": str(_safe_attr(app, "Version", "")),
            "dao_version": str(_safe_attr(db, "Version", "")),
            "startup_form": str(startup_form) if startup_form else None,
            "app_title": str(_safe_prop(db, "AppTitle")) if _safe_prop(db, "AppTitle") else None,
            "allow_bypass_key": _safe_prop(db, "AllowBypassKey"),
            "properties": props,
            "vba_references": vba_refs,
            "linked_databases": [],
            "extracted_at": extracted_at,
        }

    def _extract_tables(self, db: Any, warnings: list[str]) -> list[dict[str, Any]]:
        tables: list[dict[str, Any]] = []
        for tdf in db.TableDefs:
            name = str(tdf.Name)
            attrs = int(_safe_attr(tdf, "Attributes", 0) or 0)

            # Always ignore transient temporary tables
            if name.startswith("~TMPCLP"):
                continue

            is_system = (
                name.startswith("MSys")
                or name.startswith("USys")
                or name.startswith("~")
                or bool(attrs & DB_SYSTEM_OBJECT)
            )
            if is_system and not self.options.include_system_tables:
                continue

            connect = str(_safe_attr(tdf, "Connect", "") or "")
            if is_system:
                kind = "system"
            elif bool(attrs & DB_ATTACHED_ODBC) or connect.upper().startswith("ODBC;"):
                kind = "linked_odbc"
            elif bool(attrs & DB_ATTACHED_TABLE) or connect != "":
                kind = "linked_access"
            else:
                kind = "local"

            indexes: list[dict[str, Any]] = []
            pk_fields: list[str] = []
            try:
                for idx in tdf.Indexes:
                    idx_fields: list[dict[str, Any]] = []
                    for idx_fld in idx.Fields:
                        idx_fields.append(
                            {
                                "name": str(idx_fld.Name),
                                "descending": bool(int(_safe_attr(idx_fld, "Attributes", 0) or 0) & 1),
                            }
                        )
                    is_primary = bool(_safe_attr(idx, "Primary", False))
                    if is_primary:
                        pk_fields = [f["name"] for f in idx_fields]
                    indexes.append(
                        {
                            "name": str(idx.Name),
                            "primary": is_primary,
                            "unique": bool(_safe_attr(idx, "Unique", False)),
                            "foreign": bool(_safe_attr(idx, "Foreign", False)),
                            "ignore_nulls": bool(_safe_attr(idx, "IgnoreNulls", False)),
                            "required": bool(_safe_attr(idx, "Required", False)),
                            "fields": idx_fields,
                        }
                    )
            except Exception as exc:
                warnings.append(f"テーブル '{name}' のインデックス取得に失敗しました: {exc}")

            record_count = self._get_record_count(db, tdf, name, kind)

            fields: list[dict[str, Any]] = []
            try:
                for fld in tdf.Fields:
                    fields.append(
                        self._extract_field(
                            db, name, fld, pk_fields, record_count, kind, warnings
                        )
                    )
            except Exception as exc:
                warnings.append(f"テーブル '{name}' のフィールド取得に失敗しました: {exc}")

            tables.append(
                {
                    "name": name,
                    "kind": kind,
                    "is_hidden": bool(attrs & DB_HIDDEN_OBJECT),
                    "connect": connect,
                    "source_table_name": str(_safe_attr(tdf, "SourceTableName", "") or ""),
                    "description": str(_safe_prop(tdf, "Description", "") or ""),
                    "validation_rule": str(_safe_attr(tdf, "ValidationRule", "") or ""),
                    "validation_text": str(_safe_attr(tdf, "ValidationText", "") or ""),
                    "record_count": record_count,
                    "fields": fields,
                    "indexes": indexes,
                    "primary_key_fields": pk_fields,
                    "date_created": _serialize_value(_safe_attr(tdf, "DateCreated", "")),
                    "last_updated": _serialize_value(_safe_attr(tdf, "LastUpdated", "")),
                }
            )
        return tables

    def _get_record_count(self, db: Any, tdf: Any, name: str, kind: str) -> int | None:
        if kind == "local":
            try:
                rc = int(tdf.RecordCount)
                if rc >= 0:
                    return rc
            except Exception:
                pass
        try:
            escaped = name.replace("]", "]]")
            rs = db.OpenRecordset(f"SELECT COUNT(*) FROM [{escaped}]")
            try:
                return int(rs.Fields(0).Value)
            finally:
                rs.Close()
        except Exception:
            return None

    def _extract_field(
        self,
        db: Any,
        table_name: str,
        fld: Any,
        pk_fields: list[str],
        record_count: int | None,
        table_kind: str,
        warnings: list[str],
    ) -> dict[str, Any]:
        fld_name = str(fld.Name)
        dao_type = int(_safe_attr(fld, "Type", 0) or 0)
        size = int(_safe_attr(fld, "Size", 0) or 0)
        attrs = int(_safe_attr(fld, "Attributes", 0) or 0)
        mapping = map_dao_field_type(dao_type, size)

        # Lookup properties (important in Access because combos/FKs are often hidden at table level)
        display_ctrl_raw = _safe_prop(fld, "DisplayControl")
        row_source = str(_safe_prop(fld, "RowSource", "") or "")
        row_source_type = str(_safe_prop(fld, "RowSourceType", "") or "")
        lookup: dict[str, Any] | None = None
        if row_source or (display_ctrl_raw in (110, 111)):
            lookup = {
                "display_control": DISPLAY_CONTROL_MAP.get(int(display_ctrl_raw or 0), str(display_ctrl_raw or "")),
                "row_source_type": row_source_type,
                "row_source": row_source,
                "bound_column": _safe_prop(fld, "BoundColumn"),
                "column_count": _safe_prop(fld, "ColumnCount"),
                "column_widths": str(_safe_prop(fld, "ColumnWidths", "") or ""),
                "limit_to_list": _safe_prop(fld, "LimitToList"),
                "allow_multiple_values": bool(_safe_prop(fld, "AllowMultipleValues", False)),
            }

        profile: dict[str, Any] | None = None
        if (
            self.options.profile_data
            and table_kind == "local"
            and not mapping.is_complex
            and dao_type not in (9, 11, 17)
            and (record_count is None or record_count <= self.options.profile_max_rows)
        ):
            profile = self._profile_field(db, table_name, fld_name, record_count, warnings)

        return {
            "name": fld_name,
            "ordinal_position": int(_safe_attr(fld, "OrdinalPosition", 0) or 0),
            "dao_type": dao_type,
            "type_name": mapping.dao_name,
            "python_type": mapping.python_type,
            "sqlalchemy_type": mapping.sqlalchemy_type,
            "size": size,
            "required": bool(_safe_attr(fld, "Required", False)),
            "allow_zero_length": bool(_safe_attr(fld, "AllowZeroLength", False)),
            "default_value": str(_safe_attr(fld, "DefaultValue", "") or ""),
            "validation_rule": str(_safe_attr(fld, "ValidationRule", "") or ""),
            "validation_text": str(_safe_attr(fld, "ValidationText", "") or ""),
            "input_mask": str(_safe_prop(fld, "InputMask", "") or ""),
            "format": str(_safe_prop(fld, "Format", "") or ""),
            "caption": str(_safe_prop(fld, "Caption", "") or ""),
            "description": str(_safe_prop(fld, "Description", "") or ""),
            "is_autonumber": bool(attrs & DB_AUTO_INCR_FIELD),
            "is_primary_key": fld_name in pk_fields,
            "is_hyperlink": bool(attrs & DB_HYPERLINK_FIELD),
            "is_multivalue": mapping.is_complex and dao_type != 101,
            "is_attachment": dao_type == 101,
            "attributes": attrs,
            "lookup": lookup,
            "profile": profile,
        }

    def _profile_field(
        self,
        db: Any,
        table_name: str,
        field_name: str,
        record_count: int | None,
        warnings: list[str],
    ) -> dict[str, Any] | None:
        t_esc = table_name.replace("]", "]]")
        f_esc = field_name.replace("]", "]]")
        try:
            rs = db.OpenRecordset(
                f"SELECT COUNT([{f_esc}]) AS nn, MIN([{f_esc}]) AS mn, MAX([{f_esc}]) AS mx FROM [{t_esc}]"
            )
            try:
                non_null = int(rs.Fields("nn").Value or 0)
                min_val = _serialize_value(rs.Fields("mn").Value)
                max_val = _serialize_value(rs.Fields("mx").Value)
            finally:
                rs.Close()

            total = record_count if record_count is not None else non_null
            null_count = max(total - non_null, 0)
            null_ratio = round(null_count / total, 4) if total > 0 else 0.0

            distinct_count = 0
            samples: list[str] = []
            rs_dist = db.OpenRecordset(
                f"SELECT DISTINCT [{f_esc}] FROM [{t_esc}] WHERE [{f_esc}] IS NOT NULL"
            )
            try:
                while not rs_dist.EOF:
                    distinct_count += 1
                    if len(samples) < self.options.profile_sample_limit:
                        samples.append(str(_serialize_value(rs_dist.Fields(0).Value)))
                    rs_dist.MoveNext()
            finally:
                rs_dist.Close()

            return {
                "non_null_count": non_null,
                "null_count": null_count,
                "null_ratio": null_ratio,
                "distinct_count": distinct_count,
                "min_value": str(min_val) if min_val is not None else None,
                "max_value": str(max_val) if max_val is not None else None,
                "sample_values": samples,
            }
        except Exception as exc:
            warnings.append(f"プロファイリング失敗 ({table_name}.{field_name}): {exc}")
            return None

    def _extract_relations(self, db: Any, warnings: list[str]) -> list[dict[str, Any]]:
        relations: list[dict[str, Any]] = []
        try:
            for rel in db.Relations:
                name = str(rel.Name)
                if name.startswith("MSys") and not self.options.include_system_tables:
                    continue
                attrs = int(_safe_attr(rel, "Attributes", 0) or 0)
                if attrs & DB_RELATION_LEFT:
                    join_type = "left"
                elif attrs & DB_RELATION_RIGHT:
                    join_type = "right"
                else:
                    join_type = "inner"

                field_mappings: list[dict[str, str]] = []
                for fld in rel.Fields:
                    field_mappings.append(
                        {
                            "source_field": str(fld.Name),
                            "target_field": str(fld.ForeignName),
                        }
                    )

                relations.append(
                    {
                        "name": name,
                        "table": str(rel.Table),
                        "foreign_table": str(rel.ForeignTable),
                        "enforce_referential_integrity": not bool(attrs & DB_RELATION_DONT_ENFORCE),
                        "cascade_update": bool(attrs & DB_RELATION_UPDATE_CASCADE),
                        "cascade_delete": bool(attrs & DB_RELATION_DELETE_CASCADE),
                        "is_one_to_one": bool(attrs & DB_RELATION_UNIQUE),
                        "join_type": join_type,
                        "attributes": attrs,
                        "fields": field_mappings,
                    }
                )
        except Exception as exc:
            warnings.append(f"リレーションシップ取得に失敗しました: {exc}")
        return relations

    def _extract_queries(self, db: Any, warnings: list[str]) -> list[dict[str, Any]]:
        queries: list[dict[str, Any]] = []
        try:
            for qdf in db.QueryDefs:
                name = str(qdf.Name)
                if name.startswith("~sq_"):
                    continue
                q_type_code = int(_safe_attr(qdf, "Type", 0) or 0)
                params: list[dict[str, Any]] = []
                try:
                    for p in qdf.Parameters:
                        p_type = int(_safe_attr(p, "Type", 0) or 0)
                        mapping = map_dao_field_type(p_type)
                        params.append(
                            {
                                "name": str(p.Name),
                                "dao_type": p_type,
                                "type_name": mapping.dao_name,
                                "python_type": mapping.python_type,
                            }
                        )
                except Exception:
                    pass

                queries.append(
                    {
                        "name": name,
                        "query_type_code": q_type_code,
                        "query_type": DAO_QUERY_TYPES.get(q_type_code, f"Unknown({q_type_code})"),
                        "sql": str(_safe_attr(qdf, "SQL", "") or ""),
                        "description": str(_safe_prop(qdf, "Description", "") or ""),
                        "connect": str(_safe_attr(qdf, "Connect", "") or ""),
                        "returns_records": bool(_safe_attr(qdf, "ReturnsRecords", True)),
                        "max_records": int(_safe_attr(qdf, "MaxRecords", 0) or 0),
                        "parameters": params,
                        "date_created": _serialize_value(_safe_attr(qdf, "DateCreated", "")),
                        "last_updated": _serialize_value(_safe_attr(qdf, "LastUpdated", "")),
                    }
                )
        except Exception as exc:
            warnings.append(f"クエリ定義の取得に失敗しました: {exc}")
        return queries

    def _export_saveastext_group(
        self,
        app: Any,
        db: Any,
        ac_type: int,
        project_collection_name: str,
        container_name: str,
        target_dir: Path,
        rel_dir: str,
        warnings: list[str],
    ) -> dict[str, str]:
        names: list[str] = []
        try:
            collection = getattr(app.CurrentProject, project_collection_name)
            for item in collection:
                names.append(str(item.Name))
        except Exception:
            try:
                container = db.Containers(container_name)
                for doc in container.Documents:
                    names.append(str(doc.Name))
            except Exception as exc:
                warnings.append(f"{project_collection_name} の列挙に失敗しました: {exc}")
                return {}

        exported: dict[str, str] = {}
        for idx, name in enumerate(names, start=1):
            filename = _sanitize_filename(name, idx)
            out_file = target_dir / filename
            try:
                app.SaveAsText(ac_type, name, str(out_file))
                exported[name] = f"{rel_dir}/{filename}"
            except Exception as exc:
                warnings.append(f"SaveAsText 失敗 ({project_collection_name}: {name}): {exc}")
        return exported
