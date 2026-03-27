# analyze-shuroushomei

日本の自治体の就労証明書Excelフォーマットを解析し、全テキスト入力フィールドとチェックボックスのセルアドレスをJSON形式で返すエージェント。

## アーキテクチャ

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

```
skill/scripts/analyze_shuroushomei.py (Excel解析)
  ├─ analyze_certificate(): セクション別パーサーを順次呼び出し
  │    ├─ _parse_addressee()         — 宛先
  │    ├─ _parse_certification_date() — 証明日
  │    ├─ _parse_office_info()       — 事業所情報
  │    ├─ _parse_industry()          — No.1 業種
  │    ├─ _parse_personal_info()     — No.2 フリガナ・本人氏名・生年月日
  │    ├─ _parse_employment_period() — No.3 雇用期間
  │    ├─ _parse_work_office()       — No.4 本人就労先事業所
  │    ├─ _parse_employment_type()   — No.5 雇用形態
  │    ├─ _parse_work_time()         — No.6 就労時間
  │    │    ├─ _parse_fixed_work_time()     — 固定就労
  │    │    └─ _parse_irregular_work_time() — 変則就労
  │    ├─ _parse_work_record()       — No.7 就労実績
  │    ├─ _parse_maternity_leave()   — No.8 産前産後休業
  │    ├─ _parse_childcare_leave()   — No.9 育児休業
  │    ├─ _parse_other_leave()       — No.10 産休育休以外
  │    ├─ _parse_return_date()       — No.11 復職年月日
  │    ├─ _parse_short_time_work()   — No.12 短時間勤務
  │    ├─ _parse_tanshin_funin()     — No.17 単身赴任
  │    ├─ _parse_remarks()           — No.18 備考欄
  │    └─ _parse_guardian_section()   — No.19 保護者記載欄
  ├─ find_all_checkboxes(): チェックボックス検出
  │    ├─ _detect_checkbox_label(): ラベル検出
  │    └─ _detect_row_context(): 行コンテキスト検出
  └─ verify_and_repair(): 検証・自己修復
       ├─ _repair_text_fields() → _repair_text_field_by_scan(): テキスト修復
       └─ _repair_checkboxes(): チェックボックス修復
```

- **LLM**: `UiPathAzureChatOpenAI` (`uipath_langchain.chat.models`) — URL構築・認証を自動処理
- **エージェントフレームワーク**: LangChain tools (`langchain_core.tools.tool`) + LangGraph React Agent (`langgraph.prebuilt.create_react_agent`)
- **グラフ**: LangGraph `START → node_skill → node_municipality → node_evaluator → (node_repair) → END`
- **解析スクリプト**: `skill/scripts/analyze_shuroushomei.py` — openpyxlでExcelセル構造を解析
- **期待セクション定義**: `_EXPECTED_TEXT_SECTIONS` / `_EXPECTED_CB_SECTIONS` は `analyze_shuroushomei.py` で定義（main.pyから `importlib` でインポート、単一ソース）

## エージェント構成

| エージェント | 役割 | ツール |
|-------------|------|--------|
| `skill_agent` | Excelダウンロード＋解析スクリプト実行 | `download_excel`, `run_analyze_shuroushomei`, `read_text_file`, `write_text_file` |
| `municipality_agent` | 自治体名の都道府県補完 | (ツールなし、LLM推論のみ) |
| `evaluator_agent` | 解析結果の件数検証 | `evaluate_fields` |
| `node_repair` | 不足フィールドをwarningとして報告 | (ツールなし) |

## ファイル構成

| ファイル | 役割 |
|---------|------|
| `main.py` | エージェント本体。LangGraph StateGraph + uipath-langchain + create_react_agent 構成 |
| `skill/scripts/analyze_shuroushomei.py` | Excel解析スクリプト（セクション別パーサー構成、`--json`フラグでJSON出力） |
| `skill/SKILL.md` | スキル定義（inner agentが参照） |
| `local.settings.json` | ローカル環境変数（`Values`配下に配置） |
| `.env` | UiPath認証情報（`UIPATH_URL`, `UIPATH_ACCESS_TOKEN`等） |
| `langgraph.json` | グラフ定義（`./main.py:graph`） |
| `entry-points.json` | UiPathエントリポイント定義 |
| `agent.mermaid` | グラフのMermaidフローチャート |
| `input.json` | テスト用入力（`{"url": "..."}`) |

## 実行方法

```bash
uipath run agent -f input.json
```

## 環境変数

| 変数 | 設定場所 | 説明 |
|------|---------|------|
| `UIPATH_CHAT_MODEL` | `local.settings.json` | LLMモデル名（デフォルト: `gpt-5.2-2025-12-11`） |
| `MAX_TURNS` | `local.settings.json` | エージェントの最大ターン数（デフォルト: 30） |
| `UIPATH_URL` | `.env` | UiPathプラットフォームURL |
| `UIPATH_ACCESS_TOKEN` | `.env` | UiPath認証トークン |

## UiPath LLM Gateway

`uipath-langchain` SDK (`UiPathAzureChatOpenAI`) 経由でLLMにアクセス:
- エンドポイント: `{UIPATH_URL}/llmgateway_/openai/deployments/{model}/chat/completions` (SDK が自動構築)
- 認証: `UIPATH_ACCESS_TOKEN` を SDK が自動設定
- 利用可能モデル: `gpt-5.2-2025-12-11` (動作確認済み), `gpt-4o-2024-08-06`, `gpt-4o-2024-11-20` 等

**注意**: リージョン(US等)によって利用可能モデルが異なる。

## 入出力

**Input**: `{"url": "https://...xlsx"}`

**Output**:
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

## 解析スクリプト (analyze_shuroushomei.py)

### 構成

- **セクション別パーサー**: `analyze_certificate()` は18個のセクション別関数 (`_parse_*`) を順次呼び出す
- **共通ヘルパー**: `_parse_leave_period()` (休業系共通), `_find_other_text_input()` (その他テキスト共通), `_get_label_end_col()` (ラベル結合範囲取得)
- **チェックボックス**: `find_all_checkboxes()` → `_detect_checkbox_label()` + `_detect_row_context()` でラベル・コンテキスト検出
- **自己修復**: `verify_and_repair()` → `_repair_text_field_by_scan()` (汎用スキャン修復) + `_repair_checkboxes()` (代替マーカー検出)
- **エクスポート**: `_EXPECTED_TEXT_SECTIONS`, `_EXPECTED_CB_SECTIONS` を main.py からインポート可能

### 対応セクション（標準フォーマット全19項目）

ヘッダー（宛先・証明日・事業所情報）、No.1業種〜No.19保護者記載欄まで。

### 自治体名の推測

ファイル名ではなくExcel内容から推測:
1. 宛先セルの「○○市長」「○○区長」等から抽出
2. 「宛」ラベルの左側セルを確認
3. 市区町村名パターンで検索

### テスト

Agent venv の Python を使うこと（Python 3.11+ の型ヒント構文を使用しているため）。

```bash
# プロジェクトルートから実行
Agent/.venv/bin/python Agent/skill/scripts/analyze_shuroushomei.py <file>.xlsx --json
```

## コーディング規約

- Python 3.11+前提: モダンな型ヒント使用（`dict`, `list`, `str | None`）
- LLM: `UiPathAzureChatOpenAI` (`uipath_langchain.chat.models`) を使用
- エージェント: `create_react_agent` (`langgraph.prebuilt`) + `langchain_core.tools.tool` を使用
- グラフ: LangGraph StateGraph で4ノード構成（`node_skill → node_municipality → node_evaluator → node_repair`）
- 環境変数は `local.settings.json` の `Values` 配下に配置
- UiPath認証情報は `.env` に配置（`local.settings.json`には含めない）
- 出力フィールド名の規約:
  - 構造的区切りは `_`（`証明日_年`, `業種_情報通信業`）
  - ラベル内の中点 `・` は保持（`業種_農業・林業`, `雇用形態_パート・アルバイト`）
  - 半角括弧は使用しない（`電話番号1` not `電話番号(1)`）
  - チェックボックスのセクション名はテキストフィールドと統一（`_SECTION_NAME_MAP` で変換）
  - 重複ラベルには連番を付与（`保護者記載欄_利用中1`, `利用中2`, `利用中3`）
  - 休憩時間は `休憩時間`（`うち休憩時間` ではなく）
