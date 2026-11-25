# access_analyzer

Microsoft Access（`.accdb` / `.mdb`）の内部構造を **COM / DAO / `SaveAsText`** で徹底的に抽出し、**Python で再構築するための構造化 JSON + Markdown 設計書（Mermaid ER図・画面遷移図・コールグラフ・横断依存関係図付き）** を自動生成する解析ツールです。

## 特長

- **2段階パイプライン（抽出と解析の分離）**:
  - **Stage 1 (`extract`)**: Windows 上の Microsoft Access COM (`Access.Application` Out-of-process) を用いて DAO スキーマ・リレーション・クエリ・参照設定および `SaveAsText` 定義（フォーム/レポート/マクロ/VBA）を `raw/` へダンプ。
  - **Stage 2 (`build`)**: Access を起動せず Pure Python のみで高速にパース・静的解析・依存グラフ構築・Python移行インベントリ生成・JSON/Markdown出力を実行。
- **徹底的な解析対象**:
  - テーブル・フィールド（DAO型→Python/SQLAlchemy型マッピング、ルックアップ定義、添付/複数値フィールド、データプロファイリング）・インデックス・リレーションシップ
  - クエリ（SQL 静的解析、参照ソース、JOIN、Access固有関数、`Forms!`/`TempVars!` 参照、クロス集計、アクションクエリ）
  - フォーム・レポート（セクション、コントロール階層、サブフォーム/サブレポートの Master/Child リンク、グループ化、イベント、埋め込みマクロ、コードビハインド VBA）
  - マクロ（XML / 旧形式両対応、`AutoExec` / `AutoKeys` 検知）および VBA（プロシージャ、`DoCmd.*`、埋め込み SQL、`DLookup`/`DSum`、Win32 API `Declare`、外部 COM 連携）
  - 全オブジェクト横断の依存関係グラフ、トポロジカル実装順序、未使用（孤立）候補、Python 再構築ホットスポット一覧、画面設計ブループリント

## セットアップ

```powershell
.\.venv\Scripts\Activate.ps1
pip install -e ".[dev]" --no-build-isolation
```

## 使い方

### 1. 一括解析（COM抽出 → 解析 → JSON & Markdown設計書出力）

```powershell
access-analyzer analyze C:\path\to\app.accdb -o .\output\app
```

主なオプション:
- `--password <pw>`: パスワード保護された DB のパスワード
- `--profile-data`: 各列の NULL 率・ユニーク数・最小/最大・サンプル値を集計
- `--include-system`: `MSys*` 等のシステムテーブルも含める
- `--no-keep-raw`: 中間ダンプ (`raw/`) を残さない

### 2. Stage 1 のみ（生ダンプ抽出）

```powershell
access-analyzer extract C:\path\to\app.accdb -o .\output\app\raw
```

### 3. Stage 2 のみ（生ダンプから設計書を再生成・Access不要）

```powershell
access-analyzer build .\output\app\raw -o .\output\app
```

## 出力構成

```text
output/<db_name>/
├── raw/                             # COM / SaveAsText の生ダンプ
├── json/
│   ├── analysis.json                # 統合構造データ (LLM・コード生成ツール連携向け)
│   ├── schema.json                  # テーブル・リレーション
│   ├── queries.json                 # クエリ・SQL解析
│   ├── ui.json                      # フォーム・レポート
│   ├── logic.json                   # マクロ・VBA
│   └── dependencies.json            # 依存グラフ・移行インベントリ
└── docs/
    ├── README.md                    # 全体サマリー
    ├── 01_er_diagram.md             # Mermaid ER図 + テーブル定義 + SQLAlchemyモデル案
    ├── 02_queries.md                # クエリ詳細・Access固有構文チェック
    ├── 03_forms_ui.md               # Mermaid 画面遷移図 + フォーム・コントロール仕様
    ├── 04_reports.md                # レポート・グループ化・集計仕様
    ├── 05_macros_and_vba.md         # マクロ・VBA解析 + Mermaid コールグラフ
    ├── 06_dependency_graph.md       # 全オブジェクト横断 Mermaid 依存図 + 影響分析マトリクス
    └── 07_python_migration_guide.md # Python再構築ガイド・ホットスポット・画面ブループリント
```

## テスト実行

```powershell
pytest
```
