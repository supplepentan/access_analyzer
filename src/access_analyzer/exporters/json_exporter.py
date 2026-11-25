"""Export DatabaseAnalysis to structured JSON files."""

from __future__ import annotations

import json
from pathlib import Path

from access_analyzer.models.analysis import DatabaseAnalysis


def _write_json(path: Path, payload: object) -> None:
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def export_json(analysis: DatabaseAnalysis, json_dir: str | Path) -> dict[str, Path]:
    """Write unified analysis.json and per-category JSON files into json_dir."""
    out = Path(json_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)

    full_path = out / "analysis.json"
    _write_json(full_path, analysis.model_dump(mode="json"))

    schema_path = out / "schema.json"
    _write_json(
        schema_path,
        {
            "metadata": analysis.metadata.model_dump(mode="json"),
            "tables": [t.model_dump(mode="json") for t in analysis.tables],
            "relations": [r.model_dump(mode="json") for r in analysis.relations],
        },
    )

    queries_path = out / "queries.json"
    _write_json(queries_path, [q.model_dump(mode="json") for q in analysis.queries])

    ui_path = out / "ui.json"
    _write_json(
        ui_path,
        {
            "forms": [f.model_dump(mode="json") for f in analysis.forms],
            "reports": [r.model_dump(mode="json") for r in analysis.reports],
        },
    )

    logic_path = out / "logic.json"
    _write_json(
        logic_path,
        {
            "macros": [m.model_dump(mode="json") for m in analysis.macros],
            "vba_modules": [m.model_dump(mode="json") for m in analysis.vba_modules],
            "form_vba_modules": [
                f.vba_module.model_dump(mode="json") for f in analysis.forms if f.vba_module is not None
            ],
            "report_vba_modules": [
                r.vba_module.model_dump(mode="json") for r in analysis.reports if r.vba_module is not None
            ],
        },
    )

    deps_path = out / "dependencies.json"
    _write_json(
        deps_path,
        {
            "dependency_graph": analysis.dependency_graph.model_dump(mode="json"),
            "migration_inventory": analysis.migration_inventory.model_dump(mode="json"),
        },
    )

    return {
        "analysis": full_path,
        "schema": schema_path,
        "queries": queries_path,
        "ui": ui_path,
        "logic": logic_path,
        "dependencies": deps_path,
    }
