import os
import sys
import re
import json
import logging
import subprocess
import importlib.util
import urllib.request
from pathlib import Path
from urllib.parse import urlparse

from langgraph.graph import START, StateGraph, END
from langgraph.prebuilt import create_react_agent
from pydantic import BaseModel
from langchain_core.tools import tool
from langchain_core.messages import HumanMessage, SystemMessage
from uipath_langchain.chat.models import UiPathAzureChatOpenAI

# ---------------------------------------------------------------------------
# Import expected section definitions from the analysis script (single source)
# ---------------------------------------------------------------------------
_analyze_script_path = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "skill", "scripts", "analyze_shuroushomei.py"
)
_spec = importlib.util.spec_from_file_location("analyze_shuroushomei", _analyze_script_path)
_analyze_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_analyze_mod)
_EXPECTED_TEXT_SECTIONS = _analyze_mod._EXPECTED_TEXT_SECTIONS
_EXPECTED_CB_SECTIONS = _analyze_mod._EXPECTED_CB_SECTIONS

# ---------------------------------------------------------------------------
# Load environment variables from local.settings.json
# ---------------------------------------------------------------------------
_env_path = Path(__file__).parent / "local.settings.json"
if _env_path.exists():
    with open(_env_path) as _f:
        for _key, _value in json.load(_f).get("Values", {}).items():
            os.environ.setdefault(_key, str(_value))

# ---------------------------------------------------------------------------
# Configure logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

MAX_TURNS = int(os.getenv("MAX_TURNS", "30"))
logger.info(f"Initialized with MAX_TURNS={MAX_TURNS}")


# ---------------------------------------------------------------------------
# Schemas (LangGraph I/O)
# ---------------------------------------------------------------------------

class Input(BaseModel):
    url: str


class State(BaseModel):
    url: str
    municipality: str = "不明"
    sheet_name: str = ""
    text: dict = {}
    checkbox: dict = {}
    error: str | None = None
    warning: str | None = None
    # Inter-node communication fields
    excel_path: str = ""
    raw_municipality: str = "不明"
    is_complete: bool = False
    eval_data: dict = {}


class Output(BaseModel):
    municipality: str = "不明"
    sheet_name: str = ""
    text: dict = {}
    checkbox: dict = {}
    error: str | None = None
    warning: str | None = None


# ---------------------------------------------------------------------------
# UiPath LLM Gateway — via uipath-langchain SDK
# ---------------------------------------------------------------------------

_model_name = os.getenv("UIPATH_CHAT_MODEL", "gpt-5.2-2025-12-11")

llm = UiPathAzureChatOpenAI(
    model=_model_name,
    temperature=0,
    max_tokens=4096,
)

logger.info(f"LLM initialized: model={_model_name}")


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def _extract_json_from_text(content: str) -> dict | None:
    """テキストからJSON辞書を抽出する。"""
    json_match = re.search(r'\{[\s\S]*\}', content)
    if not json_match:
        return None
    try:
        return json.loads(json_match.group())
    except json.JSONDecodeError:
        return None


def _find_deficit_sections(
    fields: dict, expected_sections: list[tuple[str, int]]
) -> list[tuple[str, int, int]]:
    """不足セクションを検出する。(prefix, expected, actual) のリストを返す。"""
    deficits = []
    for prefix, expected in expected_sections:
        actual = len([k for k in fields if k.startswith(prefix)])
        if actual < expected:
            deficits.append((prefix, expected, actual))
    return deficits


def _extract_domain_hint(hostname: str) -> tuple[str, str] | None:
    """ドメインからローマ字部分と自治体種別を抽出する。"""
    parts = hostname.split('.')
    type_map = {'city': '市または特別区', 'town': '町', 'vill': '村'}
    for i, part in enumerate(parts):
        if part in type_map and i + 1 < len(parts):
            return parts[i + 1], type_map[part]
    return None


def _find_downloaded_excel() -> str | None:
    """カレントディレクトリから直近のダウンロード済みExcelファイルを探す。"""
    cwd = Path(os.getcwd())
    for ext in ("*.xlsx", "*.xls"):
        files = sorted(cwd.glob(ext), key=lambda p: p.stat().st_mtime, reverse=True)
        if files:
            return str(files[0])
    return None



def _build_deficit_lines(eval_data: dict) -> list[str]:
    """eval_dataから不足セクションの説明行リストを構築する。"""
    lines = []
    for d in eval_data.get("text_deficits", []):
        lines.append(f"テキスト不足 - {d['prefix']}: {d['actual']}/{d['expected']}件")
    for d in eval_data.get("cb_deficits", []):
        lines.append(f"チェックボックス不足 - {d['prefix']}: {d['actual']}/{d['expected']}件")
    return lines


def _get_final_ai_content(result: dict) -> str:
    """React agentの結果から最後のAIメッセージの内容を取得する。"""
    messages = result.get("messages", [])
    for msg in reversed(messages):
        if hasattr(msg, "content") and msg.content and not getattr(msg, "tool_calls", None):
            return str(msg.content)
    return ""


# ---------------------------------------------------------------------------
# Tools (LangChain @tool)
# ---------------------------------------------------------------------------

@tool
def download_excel(url: str) -> str:
    """URLからExcelファイルをダウンロードしてファイルパスを返す。"""
    try:
        filename = os.path.basename(url.split('?')[0])
        filepath = os.path.join(os.getcwd(), filename)
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req) as resp, open(filepath, 'wb') as f:
            f.write(resp.read())
        return json.dumps(
            {"status": "ok", "filepath": filepath, "size": os.path.getsize(filepath)},
            ensure_ascii=False,
        )
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False)


@tool
def run_analyze_shuroushomei(filepath: str) -> str:
    """就労証明書Excelファイルを解析してJSON結果を返す。"""
    script = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "skill", "scripts", "analyze_shuroushomei.py"
    )
    result = subprocess.run(
        [sys.executable, script, filepath, "--json"],
        capture_output=True, text=True,
        env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        cwd=os.getcwd(),
    )
    if result.stderr:
        logger.debug(f"analyze stderr:\n{result.stderr[-3000:]}")
    if result.returncode != 0:
        return json.dumps({"error": result.stderr[-2000:]}, ensure_ascii=False)
    return result.stdout


@tool
def read_text_file(filepath: str) -> str:
    """テキストファイルの内容を読む。"""
    with open(filepath, 'r', encoding='utf-8') as f:
        return f.read()[:10000]


@tool
def write_text_file(filepath: str, content: str) -> str:
    """テキストファイルに内容を書き込む。"""
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    return f"書き込み完了: {filepath}"


@tool
def evaluate_fields(analysis_result_json: str) -> str:
    """解析結果JSONを受け取り、テキストフィールドとチェックボックスの件数を検証する。"""
    try:
        data = json.loads(analysis_result_json)
    except json.JSONDecodeError:
        data = _extract_json_from_text(analysis_result_json)
        if not data:
            return json.dumps({"error": "JSONの解析に失敗しました"}, ensure_ascii=False)

    text = data.get("text", {})
    checkbox = data.get("checkbox", {})

    text_deficits = _find_deficit_sections(text, _EXPECTED_TEXT_SECTIONS)
    cb_deficits = _find_deficit_sections(checkbox, _EXPECTED_CB_SECTIONS)

    total = sum(e - a for _, e, a in text_deficits) + sum(e - a for _, e, a in cb_deficits)

    return json.dumps({
        "is_complete": total == 0,
        "text_count": len(text),
        "checkbox_count": len(checkbox),
        "text_deficits": [
            {"prefix": p, "expected": e, "actual": a} for p, e, a in text_deficits
        ],
        "cb_deficits": [
            {"prefix": p, "expected": e, "actual": a} for p, e, a in cb_deficits
        ],
        "total_deficit": total,
    }, ensure_ascii=False)



# ---------------------------------------------------------------------------
# Agent instructions
# ---------------------------------------------------------------------------

_SKILL_INSTRUCTIONS = (
    "あなたは日本の自治体の就労証明書Excelファイルを解析するエキスパートです。\n"
    "手順:\n"
    "1. download_excel でURLからExcelファイルをダウンロード\n"
    "2. run_analyze_shuroushomei でダウンロードしたファイルを解析\n"
    "3. 解析結果のJSONをそのまま最終回答として返却\n\n"
    "エラーが発生した場合:\n"
    "1. read_text_file でスクリプトやファイルの内容を確認\n"
    "2. write_text_file でスクリプトを修正\n"
    "3. run_analyze_shuroushomei を再実行\n\n"
    "最終回答はJSON形式のみを返してください（説明文は不要）。\n"
    'JSON形式: {"municipality": "自治体名", "sheet_name": "シート名", "text": {...}, "checkbox": {...}}'
)

_EVALUATOR_INSTRUCTIONS = (
    "あなたは就労証明書の解析結果を評価するエージェントです。\n"
    "evaluate_fields ツールを使って、テキストフィールドとチェックボックスの件数を検証してください。\n"
    "与えられた解析結果JSONをそのまま evaluate_fields ツールの analysis_result_json 引数に渡してください。\n"
    "ツールの結果をそのまま最終回答として返却してください（説明文は不要）。"
)


_MUNICIPALITY_INSTRUCTIONS = (
    "あなたは日本の自治体名を特定するエキスパートです。\n"
    "与えられた自治体名とURLから、都道府県＋市区町村名を返してください。\n"
    "例: 東京都文京区、大阪府大阪市、北海道札幌市\n"
    "都道府県が確定できない場合は市区町村名のみ返してください。\n"
    "推測できない場合は「不明」と返してください。\n"
    "回答は自治体名のみを返してください（説明文は不要）。"
)


# ---------------------------------------------------------------------------
# React agents (LangGraph prebuilt)
# ---------------------------------------------------------------------------

_skill_react = create_react_agent(
    llm,
    [download_excel, run_analyze_shuroushomei, read_text_file, write_text_file],
)

_evaluator_react = create_react_agent(
    llm,
    [evaluate_fields],
)



# ---------------------------------------------------------------------------
# Graph nodes
# ---------------------------------------------------------------------------

async def node_skill(state: State) -> dict:
    """Step 1: Skill Agent — Excelダウンロード＆解析"""
    logger.info(f"[node_skill] URL: {state.url}")

    result = await _skill_react.ainvoke(
        {"messages": [
            SystemMessage(content=_SKILL_INSTRUCTIONS),
            HumanMessage(content=(
                f"以下のURLから就労証明書Excelファイルをダウンロードし、解析してください。\n"
                f"URL: {state.url}"
            )),
        ]},
        config={"recursion_limit": MAX_TURNS * 2},
    )

    final_content = _get_final_ai_content(result)
    parsed = _extract_json_from_text(final_content)
    if not parsed:
        return {"error": "スキルエージェント: JSON解析に失敗しました"}

    text = parsed.get("text", {})
    checkbox = parsed.get("checkbox", {})
    raw_municipality = parsed.get("municipality", "不明")
    sheet_name = parsed.get("sheet_name", "")
    excel_path = _find_downloaded_excel() or ""

    logger.info(
        f"[node_skill] 結果: municipality={raw_municipality}, "
        f"sheet_name={sheet_name!r}, "
        f"text={len(text)}件, checkbox={len(checkbox)}件"
    )

    return {
        "text": text,
        "checkbox": checkbox,
        "raw_municipality": raw_municipality,
        "sheet_name": sheet_name,
        "excel_path": excel_path,
    }


async def node_municipality(state: State) -> dict:
    """Step 2: Municipality Agent — 自治体名解決"""
    logger.info(f"[node_municipality] raw={state.raw_municipality}")

    try:
        hostname = urlparse(state.url).hostname or ''
    except Exception:
        hostname = ''

    muni_prompt = None
    if state.raw_municipality != '不明':
        muni_prompt = (
            f"日本の自治体「{state.raw_municipality}」の都道府県名を補完してください。\n"
            f"参考URL: {hostname}\n"
            "都道府県＋市区町村名のみを返してください。"
        )
    else:
        hint = _extract_domain_hint(hostname)
        if hint:
            romaji, domain_type = hint
            muni_prompt = (
                f"日本の自治体ウェブサイトのドメイン「{hostname}」から自治体名を推測してください。\n"
                f"ローマ字部分: {romaji}\n"
                f"種別ヒント: {domain_type}\n"
                "都道府県＋市区町村名を返してください。推測できない場合は「不明」と返してください。"
            )

    municipality = state.raw_municipality
    if muni_prompt:
        try:
            response = await llm.ainvoke([
                SystemMessage(content=_MUNICIPALITY_INSTRUCTIONS),
                HumanMessage(content=muni_prompt),
            ])
            name = response.content.strip().strip('「」。')
            if name and name != '不明' and len(name) <= 15:
                municipality = name
        except Exception as e:
            logger.warning(f"[node_municipality] エラー: {e}")

    logger.info(f"[node_municipality] {state.raw_municipality} -> {municipality}")
    return {"municipality": municipality}


async def node_evaluator(state: State) -> dict:
    """Step 3: Evaluator Agent — 結果検証"""
    logger.info(f"[node_evaluator] text={len(state.text)}件, checkbox={len(state.checkbox)}件")

    eval_input_json = json.dumps(
        {"text": state.text, "checkbox": state.checkbox}, ensure_ascii=False
    )

    result = await _evaluator_react.ainvoke(
        {"messages": [
            SystemMessage(content=_EVALUATOR_INSTRUCTIONS),
            HumanMessage(content=(
                f"以下の解析結果を evaluate_fields ツールに渡して検証してください。\n"
                f"```json\n{eval_input_json}\n```"
            )),
        ]},
        config={"recursion_limit": 10},
    )

    final_content = _get_final_ai_content(result)
    eval_data = _extract_json_from_text(final_content)
    is_complete = eval_data.get("is_complete", True) if eval_data else True
    total_deficit = eval_data.get("total_deficit", 0) if eval_data else 0

    logger.info(f"[node_evaluator] complete={is_complete}, deficit={total_deficit}")

    return {
        "is_complete": is_complete,
        "eval_data": eval_data or {},
        "error": None,
    }


async def node_repair(state: State) -> dict:
    """Step 4: Repair Node — 不足フィールドをwarningとして報告"""
    deficit_lines = _build_deficit_lines(state.eval_data)
    total_deficit = state.eval_data.get("total_deficit", 0)
    logger.info(f"[node_repair] 不足{total_deficit}件をwarningとして報告")

    warning = f"以下のフィールドが不足しています（このExcelには該当セクションが存在しない可能性があります）: {'; '.join(deficit_lines)}"
    logger.warning(f"[node_repair] {warning}")
    return {"warning": warning, "error": None}


# ---------------------------------------------------------------------------
# Conditional edge: repair needed?
# ---------------------------------------------------------------------------

def should_repair(state: State) -> str:
    """Evaluator結果に基づいて修復ノードをスキップするか判定する。"""
    if state.error:
        return "end"
    if not state.is_complete and state.eval_data.get("total_deficit", 0) > 0:
        logger.info("[router] -> node_repair (不足フィールドあり)")
        return "repair"
    logger.info("[router] -> end (修復不要)")
    return "end"


# ---------------------------------------------------------------------------
# Build the graph (UiPath entry point)
#
#   START -> node_skill -> node_municipality -> node_evaluator
#                                                   |
#                                          should_repair?
#                                          /            \
#                                     repair            end
#                                       |                |
#                                   node_repair -------> END
# ---------------------------------------------------------------------------

logger.info("Building state graph...")
builder = StateGraph(State, input=Input, output=Output)

builder.add_node("node_skill", node_skill)
builder.add_node("node_municipality", node_municipality)
builder.add_node("node_evaluator", node_evaluator)
builder.add_node("node_repair", node_repair)

builder.add_edge(START, "node_skill")
builder.add_edge("node_skill", "node_municipality")
builder.add_edge("node_municipality", "node_evaluator")
builder.add_conditional_edges("node_evaluator", should_repair, {
    "repair": "node_repair",
    "end": END,
})
builder.add_edge("node_repair", END)

logger.info("Compiling graph...")
graph = builder.compile()
logger.info("Graph compilation complete")
