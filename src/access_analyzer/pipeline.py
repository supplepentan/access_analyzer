"""End-to-end two-stage analysis pipeline (Stage 1: Extract -> Stage 2: Parse & Analyze)."""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

from access_analyzer import __version__
from access_analyzer.analyzers.dependency_analyzer import build_dependency_graph
from access_analyzer.analyzers.migration_analyzer import build_migration_inventory
from access_analyzer.extractors.com_extractor import AccessComExtractor, ExtractionOptions
from access_analyzer.models.analysis import DatabaseAnalysis
from access_analyzer.models.database import DatabaseMetadata, RelationInfo, TableInfo
from access_analyzer.models.logic import MacroInfo, VbaModuleInfo
from access_analyzer.models.query import QueryInfo
from access_analyzer.models.ui import FormInfo, ReportInfo
from access_analyzer.parsers.encoding import read_access_text_file
from access_analyzer.parsers.macro_parser import parse_macro
from access_analyzer.parsers.sql_parser import parse_access_sql
from access_analyzer.parsers.ui_parser import parse_form, parse_report
from access_analyzer.parsers.vba_parser import parse_vba_module


def build_analysis_from_raw(raw_dir: str | Path) -> DatabaseAnalysis:
    """Stage 2: Parse a raw extraction directory without requiring Microsoft Access."""
    raw_path = Path(raw_dir).resolve()
    manifest_path = raw_path / "manifest.json"
    if not manifest_path.exists():
        raise FileNotFoundError(f"manifest.json が見つかりません: {manifest_path}")

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    warnings: list[str] = list(manifest.get("warnings", []))

    dao_schema_path = raw_path / manifest.get("dao_schema_file", "dao_schema.json")
    dao_data = json.loads(dao_schema_path.read_text(encoding="utf-8"))

    metadata = DatabaseMetadata.model_validate(dao_data["metadata"])
    tables = [TableInfo.model_validate(t) for t in dao_data.get("tables", [])]
    relations = [RelationInfo.model_validate(r) for r in dao_data.get("relations", [])]

    queries: list[QueryInfo] = []
    for raw_q in dao_data.get("queries", []):
        q = QueryInfo.model_validate(raw_q)
        q.parsed = parse_access_sql(
            q.sql,
            query_type=q.query_type,
            is_passthrough=(q.query_type == "PassThrough" or bool(q.connect)),
        )
        queries.append(q)

    object_files: dict[str, dict[str, str]] = manifest.get("object_files", {})

    forms: list[FormInfo] = []
    for name, rel_file in object_files.get("forms", {}).items():
        fpath = raw_path / rel_file
        if not fpath.exists():
            warnings.append(f"フォーム定義ファイルが見つかりません: {rel_file}")
            continue
        forms.append(parse_form(name, read_access_text_file(fpath)))

    reports: list[ReportInfo] = []
    for name, rel_file in object_files.get("reports", {}).items():
        rpath = raw_path / rel_file
        if not rpath.exists():
            warnings.append(f"レポート定義ファイルが見つかりません: {rel_file}")
            continue
        reports.append(parse_report(name, read_access_text_file(rpath)))

    macros: list[MacroInfo] = []
    for name, rel_file in object_files.get("macros", {}).items():
        mpath = raw_path / rel_file
        if not mpath.exists():
            warnings.append(f"マクロ定義ファイルが見つかりません: {rel_file}")
            continue
        macros.append(parse_macro(name, read_access_text_file(mpath)))

    vba_modules: list[VbaModuleInfo] = []
    for name, rel_file in object_files.get("modules", {}).items():
        vpath = raw_path / rel_file
        if not vpath.exists():
            warnings.append(f"VBAモジュールファイルが見つかりません: {rel_file}")
            continue
        vba_modules.append(parse_vba_module(name, read_access_text_file(vpath), module_type="standard"))

    dep_graph = build_dependency_graph(
        metadata=metadata,
        tables=tables,
        relations=relations,
        queries=queries,
        forms=forms,
        reports=reports,
        macros=macros,
        vba_modules=vba_modules,
    )

    migration_inventory = build_migration_inventory(
        metadata=metadata,
        tables=tables,
        queries=queries,
        forms=forms,
        reports=reports,
        macros=macros,
        vba_modules=vba_modules,
        dep_graph=dep_graph,
    )

    return DatabaseAnalysis(
        analyzer_version=__version__,
        metadata=metadata,
        tables=tables,
        relations=relations,
        queries=queries,
        forms=forms,
        reports=reports,
        macros=macros,
        vba_modules=vba_modules,
        dependency_graph=dep_graph,
        migration_inventory=migration_inventory,
        warnings=warnings,
    )


def analyze_database(
    db_path: str | Path,
    output_dir: str | Path,
    options: ExtractionOptions | None = None,
    keep_raw: bool = True,
) -> DatabaseAnalysis:
    """Run full Stage 1 (COM extraction) + Stage 2 (analysis) and return DatabaseAnalysis."""
    out_path = Path(output_dir).resolve()
    out_path.mkdir(parents=True, exist_ok=True)
    extractor = AccessComExtractor(options=options)

    if keep_raw:
        raw_dir = out_path / "raw"
        extractor.extract(db_path, raw_dir)
        return build_analysis_from_raw(raw_dir)

    with tempfile.TemporaryDirectory(prefix="access_analyzer_raw_") as tmp:
        raw_dir = Path(tmp)
        extractor.extract(db_path, raw_dir)
        return build_analysis_from_raw(raw_dir)
