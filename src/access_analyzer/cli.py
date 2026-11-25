"""Command-line interface for access_analyzer."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from access_analyzer import __version__
from access_analyzer.exporters.json_exporter import export_json
from access_analyzer.exporters.markdown_exporter import export_markdown
from access_analyzer.extractors.com_extractor import AccessComExtractor, ExtractionOptions
from access_analyzer.pipeline import analyze_database, build_analysis_from_raw


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="access-analyzer",
        description="Microsoft Access (.accdb/.mdb) の構造を徹底解析し、Python再構築用JSON/Markdown設計書を生成します。",
    )
    parser.add_argument("--version", action="version", version=f"access-analyzer {__version__}")

    sub = parser.add_subparsers(dest="command", required=True)

    # 1. analyze (Stage 1 + Stage 2 + Export)
    p_analyze = sub.add_parser("analyze", help="AccessファイルをCOM抽出し、解析・設計書出力まで一括実行する")
    p_analyze.add_argument("db_path", type=Path, help="対象の .accdb / .mdb ファイルパス")
    p_analyze.add_argument("-o", "--output", type=Path, default=None, help="出力ディレクトリ（省略時: ./output/<db名>）")
    p_analyze.add_argument("--password", type=str, default=None, help="DBパスワード（設定されている場合）")
    p_analyze.add_argument("--include-system", action="store_true", help="MSys等のシステムテーブルも含める")
    p_analyze.add_argument("--profile-data", action="store_true", help="各列のNULL率・ユニーク数・サンプル値を集計する")
    p_analyze.add_argument("--no-keep-raw", action="store_true", help="中間ダンプ(raw/)を保存せず破棄する")
    p_analyze.add_argument("--allow-autoexec", action="store_true", help="起動時マクロ無効化をスキップする")

    # 2. extract (Stage 1 only)
    p_extract = sub.add_parser("extract", help="Stage 1のみ: Access COM/DAO/SaveAsTextで生ダンプ(raw/)を抽出する")
    p_extract.add_argument("db_path", type=Path, help="対象の .accdb / .mdb ファイルパス")
    p_extract.add_argument("-o", "--output", type=Path, required=True, help="生ダンプ出力先ディレクトリ")
    p_extract.add_argument("--password", type=str, default=None, help="DBパスワード")
    p_extract.add_argument("--include-system", action="store_true", help="システムテーブルも含める")
    p_extract.add_argument("--profile-data", action="store_true", help="データプロファイリングを実行する")
    p_extract.add_argument("--allow-autoexec", action="store_true", help="起動時マクロ無効化をスキップする")

    # 3. build (Stage 2 + Export only, no Access required)
    p_build = sub.add_parser("build", help="Stage 2のみ: 抽出済みrawディレクトリから解析・設計書を再生成する（Access不要）")
    p_build.add_argument("raw_dir", type=Path, help="extractコマンドで作成したrawディレクトリ")
    p_build.add_argument("-o", "--output", type=Path, required=True, help="JSON/Markdown出力先ディレクトリ")

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    if args.command == "extract":
        opts = ExtractionOptions(
            password=args.password,
            include_system_tables=args.include_system,
            profile_data=args.profile_data,
            allow_autoexec=args.allow_autoexec,
        )
        extractor = AccessComExtractor(options=opts)
        manifest_path = extractor.extract(args.db_path, args.output)
        print(f"[OK] 生ダンプ抽出完了: {manifest_path}")
        return 0

    if args.command == "build":
        analysis = build_analysis_from_raw(args.raw_dir)
        json_files = export_json(analysis, args.output / "json")
        md_files = export_markdown(analysis, args.output / "docs")
        print(f"[OK] 解析・出力完了:")
        print(f"  - JSON: {json_files['analysis']}")
        print(f"  - Docs: {md_files['README.md']}")
        return 0

    if args.command == "analyze":
        out_dir = args.output or (Path("output") / args.db_path.stem)
        opts = ExtractionOptions(
            password=args.password,
            include_system_tables=args.include_system,
            profile_data=args.profile_data,
            allow_autoexec=args.allow_autoexec,
        )
        analysis = analyze_database(
            db_path=args.db_path,
            output_dir=out_dir,
            options=opts,
            keep_raw=not args.no_keep_raw,
        )
        json_files = export_json(analysis, out_dir / "json")
        md_files = export_markdown(analysis, out_dir / "docs")
        inv = analysis.migration_inventory
        print(f"[OK] 解析完了: {analysis.metadata.file_name}")
        print(
            f"  Tables={inv.total_tables}, Queries={inv.total_queries}, "
            f"Forms={inv.total_forms}, Reports={inv.total_reports}, "
            f"Macros={inv.total_macros}, VBA Modules={inv.total_vba_modules} ({inv.total_vba_lines} lines)"
        )
        print(f"  Hotspots: {len(inv.hotspots)} 件")
        print(f"  JSON: {json_files['analysis']}")
        print(f"  Docs: {md_files['README.md']}")
        return 0

    parser.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())
