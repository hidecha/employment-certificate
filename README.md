# 就労証明書 自動作成エージェント

日本の自治体の就労証明書Excelフォーマットを解析し、全テキスト入力フィールドとチェックボックスのセルアドレスをJSON形式で返すUiPathエージェントと、解析結果に基づいてExcelに自動入力するRPAワークフロー

[Qiita記事](https://qiita.com/hidecha/items/7e45924aeaeb37e197bd)

## 機能

- 自治体の就労証明書Excel形式ファイルをURLからダウンロード
- テキスト入力フィールドとチェックボックスのセルアドレスを自動検出
- 標準フォーマット全19項目に対応
- JSON形式で構造化された解析結果を出力
- CSVデータに基づくExcelへの自動入力（RPAワークフロー）

## セットアップ

### 1. 必要条件

- Python 3.11以上
- uv（Pythonパッケージマネージャー）
- UiPath Automation Cloud アカウント
- UiPath Studio 2024.10以上

### 2. インストール

#### uvのインストール（初回のみ）

```powershell
# Windows (PowerShell)
irm https://astral.sh/uv/install.ps1 | iex

# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh
```

#### プロジェクトのセットアップ

```powershell
# リポジトリのクローン
git clone https://github.com/hidecha/employment-certificate.git
cd employment-certificate/Agent

# Python仮想環境
uv venv

# 仮想環境の有効化
# Windows (PowerShell)
.venv\Scripts\Activate.ps1

# macOS/Linux
source .venv/bin/activate

# 依存関係のインストール
uv sync
```

### 3. 環境設定

UiPath CLIでログイン
```powershell
uipath auth
```

## 使用方法

### Agent（Excel解析エージェント）

```powershell
cd Agent

# input.jsonを使用して実行
uipath run agent -f input.json
```

#### 入力形式

`input.json`:
```json
{
  "url": "https://example.com/shuroushomei.xlsx"
}
```

#### 出力形式

```json
{
  "municipality": "○○市",
  "text": {"フィールド名": "セルアドレス", ...},
  "checkbox": {"チェックボックス名": "セルアドレス", ...},
  "error": null,
  "warning": null
}
```

- `municipality`: 自治体名（Excel内容から推測、不明な場合はURLドメインからフォールバック推測）
- `text`: テキスト入力フィールド名→セルアドレス
- `checkbox`: チェックボックス名→セルアドレス
- `error`: エラー時のみ文字列、正常時は`null`
- `warning`: 不足フィールドがある場合の警告文字列、正常時は`null`

### RPA（Excel自動入力ワークフロー）

UiPath Studioで `RPA/Main.xaml` を開いて実行

#### 入力パラメータ

| パラメータ | 説明 | 例 |
|-----------|------|-----|
| `url` | 就労証明書ExcelのURL | `https://example.com/shuroushomei.xlsx` |
| `csv` | 入力データCSVのパス | `input\sample_data.csv` |

#### CSVデータ形式

`RPA/input/sample_data.csv`:

```csv
type,item,value
text,証明日_年,"2026"
text,証明日_月,"3"
checkbox,業種_情報通信業,1
checkbox,雇用形態_正社員,1
...
```

| カラム | 説明 |
|-------|------|
| `type` | `text`（テキスト入力）または `checkbox`（チェックボックス） |
| `item` | フィールド名（Agentの出力キーと対応） |
| `value` | テキストの場合は入力値、チェックボックスの場合は `1`（チェック）/ `0`（未チェック） |

### 開発時のテスト実行

```powershell
# 解析スクリプトの直接実行（テスト用）
uv run python Agent/skill/scripts/analyze_shuroushomei.py [Excel file path] --json
```

## プロジェクト構成

```
analyze-shuroushomei/
├── README.md               # 本ファイル
├── Agent/                  # Excel解析エージェント
│   ├── main.py             # メインエージェント（LangGraph StateGraph + LLM修復）
│   ├── skill/              # スキルディレクトリ
│   │   ├── scripts/
│   │   │   └── analyze_shuroushomei.py  # Excel解析スクリプト（セクション別パーサー構成）
│   │   └── SKILL.md        # スキル定義
│   ├── local.settings.json # ローカル設定
│   ├── langgraph.json      # グラフ定義
│   ├── entry-points.json   # エントリポイント定義
│   ├── pyproject.toml      # プロジェクト設定と依存関係
│   ├── uv.lock             # 依存関係のロックファイル
│   └── .env                # 環境変数（要作成）
└── RPA/                    # Excel自動入力ワークフロー
    ├── Main.xaml            # メインワークフロー
    ├── project.json         # UiPathプロジェクト設定
    ├── entry-points.json    # エントリポイント定義
    └── input/
        └── sample_data.csv  # サンプル入力データ
```

## Agentアーキテクチャ

```
main.py (LangGraph StateGraph + uipath-langchain + LangChain tools)
  ├─ node_skill — React Agent (create_react_agent)
  │    ├─ download_excel: URLからExcelをダウンロード
  │    ├─ run_analyze_shuroushomei: 解析スクリプト実行
  │    ├─ read_text_file: ファイル読み取り
  │    └─ write_text_file: ファイル書き込み（エラー修復用）
  │
  ├─ node_municipality — 直接LLM呼び出し (llm.ainvoke)
  │    └─ 都道府県補完 / URLフォールバック推測
  │
  ├─ node_evaluator — React Agent (create_react_agent)
  │    └─ evaluate_fields: フィールド件数検証
  │
  └─ node_repair — 条件付き実行（ツールなし）
       └─ 不足フィールドをwarningとして報告
```

グラフフロー: `START → node_skill → node_municipality → node_evaluator → (node_repair) → END`

### 解析スクリプトの構成

`analyze_shuroushomei.py` は以下のセクション別パーサー関数で構成されています:

| 関数 | セクション |
|------|-----------|
| `_parse_addressee()` | 宛先 |
| `_parse_certification_date()` | 証明日 |
| `_parse_office_info()` | 事業所情報 |
| `_parse_industry()` | No.1 業種 |
| `_parse_personal_info()` | No.2 フリガナ・本人氏名・生年月日 |
| `_parse_employment_period()` | No.3 雇用期間 |
| `_parse_work_office()` | No.4 本人就労先事業所 |
| `_parse_employment_type()` | No.5 雇用形態 |
| `_parse_work_time()` | No.6 就労時間（固定+変則） |
| `_parse_work_record()` | No.7 就労実績 |
| `_parse_maternity_leave()` | No.8 産前産後休業 |
| `_parse_childcare_leave()` | No.9 育児休業 |
| `_parse_other_leave()` | No.10 産休育休以外 |
| `_parse_return_date()` | No.11 復職年月日 |
| `_parse_short_time_work()` | No.12 短時間勤務 |
| `_parse_tanshin_funin()` | No.17 単身赴任 |
| `_parse_remarks()` | No.18 備考欄 |
| `_parse_guardian_section()` | No.19 保護者記載欄 |

## 対応フィールド一覧

### テキスト入力フィールド

| セクション | フィールド数 | 主なフィールド名 |
|-----------|------------|----------------|
| 宛先 | 1 | `宛先` |
| 証明日 | 3 | `証明日_年`, `証明日_月`, `証明日_日` |
| 事業所情報 | 10 | `事業所名`, `代表者名`, `所在地`, `電話番号1〜3`, `担当者名`, `記載者連絡先1〜3` |
| No.1 業種 | 1 | `業種_その他` |
| No.2 個人情報 | 5 | `フリガナ`, `本人氏名`, `生年月日_年/月/日` |
| No.3 雇用期間 | 6 | `雇用期間_開始_年/月/日`, `雇用期間_終了_年/月/日` |
| No.4 就労先 | 2 | `事業所名称`, `事業所住所` |
| No.5 雇用形態 | 1 | `雇用形態_その他` |
| No.6 就労時間（固定） | 19 | `就労時間_月間_*`, `就労時間_平日_*`, `就労時間_土曜_*`, `就労時間_日祝_*` |
| No.6 就労時間（変則） | 9 | `変則就労_合計時間_*`, `変則就労_就労日数`, `変則就労_就労時間帯_*` |
| No.7 就労実績 | 12 | `就労実績_1月目_年/月/日数/時間数` 〜 `3月目` |
| No.8 産前産後休業 | 6 | `産前産後休業_開始_年/月/日`, `産前産後休業_終了_年/月/日` |
| No.9 育児休業 | 6 | `育児休業_開始_年/月/日`, `育児休業_終了_年/月/日` |
| No.10 産休育休以外 | 7 | `産休育休以外_その他理由`, `産休育休以外_開始_年/月/日`, `産休育休以外_終了_年/月/日` |
| No.11 復職年月日 | 3 | `復職年月日_年`, `復職年月日_月`, `復職年月日_日` |
| No.12 短時間勤務 | 11 | `短時間勤務_開始_*`, `短時間勤務_終了_*`, `短時間勤務_就労時間帯_*` |
| No.17 単身赴任 | 6 | `単身赴任_開始_年/月/日`, `単身赴任_終了_年/月/日` |
| No.18 備考欄 | 1 | `備考欄` |
| No.19 保護者記載欄 | 15 | `保護者記載欄_児童名1〜3`, `保護者記載欄_生年月日1〜3_年/月/日`, `保護者記載欄_施設名1〜3` |

### チェックボックス

| セクション | フィールド数 | 主なフィールド名 |
|-----------|------------|----------------|
| 業種 | 12+ | `業種_農業・林業`, `業種_情報通信業`, `業種_その他` 等 |
| 雇用期間 | 2 | `雇用期間_無期`, `雇用期間_有期` |
| 雇用形態 | 6+ | `雇用形態_正社員`, `雇用形態_パート・アルバイト` 等 |
| 就労時間 | 8 | `就労時間_月` 〜 `就労時間_祝日` |
| 変則就労 | 1+ | `変則就労_月間`, `変則就労_週間` |
| 産前産後休業 | 2+ | `産前産後休業_取得予定`, `産前産後休業_取得中`, `産前産後休業_取得済み` |
| 育児休業 | 3 | `育児休業_取得予定`, `育児休業_取得中`, `育児休業_取得済み` |
| 産休育休以外 | 5+ | `産休育休以外_取得予定/取得中/取得済み`, `産休育休以外_理由_*` |
| 復職年月日 | 2 | `復職年月日_復職予定`, `復職年月日_復職済み` |
| 短時間勤務 | 2 | `短時間勤務_取得予定`, `短時間勤務_取得中` |
| 保育士等勤務実態有無 | 2+ | `保育士等勤務実態有無_有`, `保育士等勤務実態有無_無` |
| 雇用契約満了後更新有無 | 3+ | `雇用契約満了後更新有無_有/無/未定` |
| 入所内定時育休短縮可否 | 2 | `入所内定時育休短縮可否_可`, `入所内定時育休短縮可否_否` |
| 育休延長可否 | 2 | `育休延長可否_可`, `育休延長可否_否` |
| 保護者記載欄 | 6 | `保護者記載欄_利用中1〜3`, `保護者記載欄_利用申込中1〜3` |
