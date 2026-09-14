# -*- coding: utf-8 -*-
"""
表格智能体（基于 GLM / 智谱大模型）

功能
----
给定一个 .xlsx 表格和一句中文查询（例如“按总分从高到低排序”“筛选出八年级1班的学生”），
智能体会调用 GLM 大模型理解查询意图，通过一系列工具（过滤、排序、选列、分组聚合、取前 N 行等）
在原始表格上逐步操作，最终生成一张新的表格并保存为新的 .xlsx 文件。

依赖
----
- pandas、openpyxl（用于读写表格）
- 调用智谱 API 使用标准库 urllib，无需额外安装 requests / openai 等第三方库。

密钥配置
--------
智谱 API Key 的完整形式为 `{id}.{secret}`。本文件同时支持两种配置方式：

1. 环境变量（推荐）：
       GLM_API_KEY_ID    = 你的 api_key_ID（id 部分）
       GLM_API_KEY       = 你的 api_key（secret 部分）
       GLM_MODEL         = 可选，模型名，默认 glm-4-flash
   最终请求头 Authorization: Bearer {GLM_API_KEY_ID}.{GLM_API_KEY}

2. 直接改下方 CONFIG 中的占位符（把密钥写进代码，注意不要提交到公开仓库）。

如果 GLM_API_KEY 本身已经是完整 key（包含 '.' 的 `id.secret` 形式），
也可以直接把完整 key 放进 GLM_API_KEY，并把 GLM_API_KEY_ID 留空。

使用方式
--------
命令行：
    python agent.py --file "表格.xlsx" --query "按总分从高到低排序"

代码调用：
    from agent import TableAgent
    df, out_path = TableAgent().run("表格.xlsx", "筛选出八年级1班的学生")
"""

import json
import os
import re
import urllib.request
import urllib.error
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------
CONFIG = {
    # 把密钥填在这里即可（留空则读取环境变量 GLM_API_KEY_ID / GLM_API_KEY）
    "api_key_id": "95d61358b1c74a09982f9c7ffc7dba79",   # 智谱 api_key_ID（id 部分）
    "api_key": "PPsP9nBW0biPMbRN",      # 智谱 api_key（secret 部分；若已是完整 `id.secret` 也放这里）
}

ZHIPU_BASE_URL = "https://open.bigmodel.cn/api/paas/v4"
DEFAULT_MODEL = "glm-4.5-flash"   # 可换 glm-4-plus / glm-4-air / glm-4-flash 等
MAX_STEPS = 12                  # 智能体最多执行的工具调用轮数，防止死循环
PREVIEW_ROWS = 6                # 每一步反馈给模型的结果预览行数
TEMPERATURE = 0.1


# ---------------------------------------------------------------------------
# 密钥
# ---------------------------------------------------------------------------
def resolve_api_key():
    """返回用于 Authorization: Bearer 的完整 key（{id}.{secret}）。"""
    key_id = CONFIG["api_key_id"] or os.environ.get("GLM_API_KEY_ID", "")
    key = CONFIG["api_key"] or os.environ.get("GLM_API_KEY", "")

    if not key and not key_id:
        raise RuntimeError(
            "未找到 GLM API Key。请设置环境变量 GLM_API_KEY_ID / GLM_API_KEY，"
            "或在 agent.py 顶部 CONFIG 中填写密钥。"
        )

    # key 本身已经形如 `id.secret`，直接使用
    if key and "." in key and not key_id:
        return key.strip()

    if key_id and key:
        return f"{key_id.strip()}.{key.strip()}"

    # 只有一部分密钥，无法拼出完整 key
    raise RuntimeError(
        "GLM 密钥不完整：需要同时提供 api_key_ID 与 api_key（或直接提供完整 `id.secret` 形式的 key）。"
    )


def _http_post_json(url, payload, api_key, timeout=180):
    """使用标准库 urllib 发送 JSON POST 请求（兼容智谱 OpenAI 风格接口）。"""
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/json; charset=utf-8")
    req.add_header("Authorization", "Bearer " + api_key)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"GLM API HTTP {e.code}: {body}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"无法连接 GLM API：{e.reason}") from e


# ---------------------------------------------------------------------------
# 表格概览（供模型理解表格结构）
# ---------------------------------------------------------------------------
def build_schema(df, max_examples=8):
    """把表格结构压缩成一段中文描述，包含列名、类型和示例/可选值。"""
    lines = [f"表格共有 {len(df)} 行、{len(df.columns)} 列。"]
    lines.append("列名（按顺序）：" + "、".join(str(c) for c in df.columns))
    for c in df.columns:
        col = str(c)
        vals = df[col].dropna()
        nunique = vals.nunique()
        if pd.api.types.is_numeric_dtype(df[col]):
            examples = [str(v) for v in vals.head(max_examples).tolist()]
            lines.append(f"- {col}：数值列，示例值 [{', '.join(examples)}]")
        else:
            if nunique <= max_examples * 2:
                uniq = sorted({str(v) for v in vals.unique()})
                lines.append(f"- {col}：文本列，可选值 [{', '.join(uniq[:30])}]")
            else:
                examples = [str(v) for v in vals.head(max_examples).tolist()]
                lines.append(f"- {col}：文本列，示例值 [{', '.join(examples)}]")
    return "\n".join(lines)


def build_schema_multi(tables, table_names):
    """多表情况下的结构描述：列出每张表别名、来源和列结构。"""
    lines = [f"共加载 {len(tables)} 张表："]
    for alias, df in tables.items():
        lines.append(f"\n【{alias}】来源《{table_names.get(alias, alias)}》")
        lines.append(build_schema(df))
    lines.append("\n当前工作表格为第一张表；如需合并多张表，请用 merge_tables 工具。")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 工具定义（OpenAI / 智谱 function-calling 格式）
# ---------------------------------------------------------------------------
_FILTER_OPS = [
    "==", "!=", ">", "<", ">=", "<=",
    "contains", "not_contains", "startswith", "endswith",
    "is_empty", "not_empty",
]

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "filter_rows",
            "description": (
                "按条件过滤表格行。filters 为条件列表，多个条件之间是“且(AND)”关系。"
                "op 支持：==、!=、>、<、>=、<=（比较），contains/not_contains（包含/不包含子串），"
                "startswith/endswith（前缀/后缀），is_empty/not_empty（为空/非空，无需 value）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "filters": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "column": {"type": "string", "description": "要过滤的列名"},
                                "op": {"type": "string", "enum": _FILTER_OPS},
                                "value": {"type": "string", "description": "比较值（is_empty/not_empty 时省略）"},
                            },
                            "required": ["column", "op"],
                        },
                    }
                },
                "required": ["filters"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "sort_rows",
            "description": (
                "按指定列排序。by 为列名列表（可多列），ascending 为与 by 等长的布尔列表（True 升序 / False 降序）。"
                "若某列是等级（如 A+、A、B+），可用 value_order 给出该列从高到低的完整顺序，"
                "例如 {\"总分\": [\"A+\",\"A\",\"A-\",\"B+\",\"B\",\"B-\",\"C+\",\"C\",\"C-\",\"D\",\"E\"]}。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "by": {"type": "array", "items": {"type": "string"}, "description": "排序依据的列名列表"},
                    "ascending": {"type": "array", "items": {"type": "boolean"}, "description": "是否升序，与 by 对应"},
                    "value_order": {
                        "type": "object",
                        "description": "可选，列名 -> 该列的等级顺序（从高到低）",
                        "additionalProperties": {"type": "array", "items": {"type": "string"}},
                    },
                },
                "required": ["by"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "select_columns",
            "description": "只保留指定的若干列，其余列删除。",
            "parameters": {
                "type": "object",
                "properties": {
                    "columns": {"type": "array", "items": {"type": "string"}, "description": "要保留的列名列表"},
                },
                "required": ["columns"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "group_aggregate",
            "description": (
                "按 group_by 列分组，并对 aggregations 中指定的列做聚合。"
                "func 支持：sum/mean/count/max/min/median/nunique/first/last（或中文：求和/平均/计数/最大/最小/中位数/去重计数）。"
                "例如按“班级”分组，对“总分”求平均。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "group_by": {"type": "array", "items": {"type": "string"}, "description": "分组列名列表"},
                    "aggregations": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "column": {"type": "string", "description": "要聚合的列名"},
                                "func": {"type": "string", "description": "聚合函数"},
                            },
                            "required": ["column", "func"],
                        },
                    },
                },
                "required": ["group_by", "aggregations"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "take_top_n",
            "description": "只保留当前表格的前 n 行（常用于排序后取前几名）。",
            "parameters": {
                "type": "object",
                "properties": {
                    "n": {"type": "integer", "description": "保留的行数"},
                },
                "required": ["n"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "reset",
            "description": "撤销所有操作，恢复到原始表格（当之前的操作有误时可重新开始）。",
            "parameters": {"type": "object", "properties": {}},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "merge_tables",
            "description": (
                "把多张已加载的表格合并成一张。tables 为表格别名列表（如 [\"表1\",\"表2\"]），"
                "on 为用于对齐的公共键列名（如姓名、考试号），how 为合并方式（inner 只保留都有的行，outer 保留所有行，默认 outer）。"
                "合并结果会成为当前工作表格。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "tables": {"type": "array", "items": {"type": "string"}, "description": "要合并的表格别名列表"},
                    "on": {"type": "array", "items": {"type": "string"}, "description": "对齐用的公共键列名"},
                    "how": {"type": "string", "enum": ["left", "right", "inner", "outer"], "description": "合并方式，默认 outer"},
                },
                "required": ["tables", "on"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "compute_rank",
            "description": (
                "计算排名并新增一列。column 为要排名的数值列；new_column 为新列名（默认“列名+排名”）。"
                "ascending=False 表示分数越高排名越靠前（第1名最高分）。"
                "group_by 可选：分组排名，如按班级排名填 [\"班级\"]，年级排名则不填（留空）。"
                "method 可选 min/dense/average/max/first，默认 min（并列名次，如 1、1、3）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "column": {"type": "string", "description": "要排名的数值列"},
                    "new_column": {"type": "string", "description": "新排名列名，默认“列名+排名”"},
                    "group_by": {"type": "array", "items": {"type": "string"}, "description": "分组排名的分组列，年级排名留空"},
                    "ascending": {"type": "boolean", "description": "是否升序（True=低分第1，False=高分第1，默认 False）"},
                    "method": {"type": "string", "enum": ["min", "dense", "average", "max", "first"], "description": "并列名次处理方式，默认 min"},
                },
                "required": ["column"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "classify_grades",
            "description": (
                "按分数线把数值成绩划分成等级，新增一列。thresholds 为从高到低的分界值（如 [90,80,70,60]），"
                "labels 为从高到低对应的等级（长度须比 thresholds 多 1，如 [\"A\",\"B\",\"C\",\"D\",\"E\"]）。"
                "即 ≥90 得 A、80~89 得 B、70~79 得 C、60~69 得 D、<60 得 E。"
                "也可用优秀/良好/及格划分，如 thresholds=[90,75,60]、labels=[\"优秀\",\"良好\",\"及格\",\"不及格\"]。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "column": {"type": "string", "description": "要划分等级的数值列"},
                    "new_column": {"type": "string", "description": "新等级列名，默认“列名+等级”"},
                    "thresholds": {"type": "array", "items": {"type": "number"}, "description": "从高到低的分界值列表"},
                    "labels": {"type": "array", "items": {"type": "string"}, "description": "从高到低的等级标签，长度=thresholds+1"},
                },
                "required": ["column", "thresholds", "labels"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "score_segment_stats",
            "description": (
                "统计每个分数段的人数，生成一张“分数段/人数”的汇总表。"
                "bins 为从低到高的分界值列表（如 [0,60,70,80,90,100]），"
                "会自动生成“0-60、60-70、…、90-100”等分数段标签并统计每段人数。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "column": {"type": "string", "description": "要统计的数值列"},
                    "bins": {"type": "array", "items": {"type": "number"}, "description": "从低到高的分界值列表"},
                },
                "required": ["column", "bins"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "compute_stats",
            "description": (
                "统计平均分、及格率、优秀率等指标，生成汇总表。columns 为要统计的数值列列表（可多列，如各学科）。"
                "group_by 可选：按班级分组填 [\"班级\"]；统计各学科指标则 columns 填 [\"语文\",\"数学\",…]。"
                "pass_line 为及格线（默认 60），excellent_line 为优秀线（默认 90）。"
                "结果包含 人数、平均分、及格率(%)、优秀率(%)。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "columns": {"type": "array", "items": {"type": "string"}, "description": "要统计的数值列列表"},
                    "group_by": {"type": "array", "items": {"type": "string"}, "description": "分组列列表，可留空"},
                    "pass_line": {"type": "number", "description": "及格线，默认 60"},
                    "excellent_line": {"type": "number", "description": "优秀线，默认 90"},
                },
                "required": ["columns"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "compare_scores",
            "description": (
                "计算两次成绩的进退步：新增一列 = column（本次）减去 base_column（上次/基准）。"
                "正值表示进步，负值表示退步。常用于“本次考试与上次考试对比”。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "column": {"type": "string", "description": "本次成绩列"},
                    "base_column": {"type": "string", "description": "上次/基准成绩列"},
                    "new_column": {"type": "string", "description": "新列名，默认“进退步”"},
                },
                "required": ["column", "base_column"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "set_output_name",
            "description": (
                "设置结果文件的名称。当用户在查询中指定了输出文件名（如“保存为排名表.xlsx”“命名为xxx”）时调用。"
                "name 可带或不带 .xlsx 后缀。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "结果文件名"},
                },
                "required": ["name"],
            },
        },
    },
]

_FUNC_MAP = {
    "sum": "sum", "求和": "sum",
    "mean": "mean", "avg": "mean", "average": "mean", "平均": "mean", "平均值": "mean",
    "count": "count", "计数": "count",
    "max": "max", "最大": "max", "最大值": "max",
    "min": "min", "最小": "min", "最小值": "min",
    "median": "median", "中位数": "median",
    "nunique": "nunique", "去重计数": "nunique",
    "first": "first", "last": "last",
}

# 标准等级顺序（从高到低），用于排序时自动识别等级列
_GRADE_ORDER_DESC = ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D+", "D", "D-", "E"]
_GRADE_RE = re.compile(r"^[A-Ea-e][+-]?$")


def _is_grade_column(series):
    """判断一列是否基本由等级值（如 A+/A/B+/B/…）构成。"""
    vals = series.dropna().astype(str).str.strip()
    if vals.empty:
        return False
    return vals.str.match(_GRADE_RE).mean() >= 0.8


def _sanitize_output_name(name):
    """把用户/模型给出的文件名清理为合法 .xlsx 文件名，非法时返回 None。"""
    name = str(name).strip().strip('"').strip("“”'’")
    name = Path(name).name                      # 去掉任何路径部分
    for ch in r'\/:*?"<>|':
        name = name.replace(ch, "")
    name = name.strip().rstrip(".")
    if not name:
        return None
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


_OUTPUT_NAME_RE = re.compile(
    r"(?:保存为|命名为|输出为|输出到|导出为|导出到|另存为|文件名为|存为|写成)\s*[:：]?\s*([^\s，。；;、,]+)"
)


def _extract_output_name(query):
    """从查询里提取用户指定的输出文件名（如“保存为排名表.xlsx”），没有则返回 None。"""
    m = _OUTPUT_NAME_RE.search(str(query))
    if not m:
        return None
    return _sanitize_output_name(m.group(1))

SYSTEM_PROMPT = """你是一个表格数据处理助手。用户会给出表格的结构信息和一句中文查询。
你需要使用提供的工具完成用户想要的操作，最终得到一张新的表格（或一张汇总表）。

可用工具：
- filter_rows：按条件过滤行（找出某班级、某分数段、某学科不达标的学生）
- sort_rows：按一列或多列排序（支持多级优先级、等级列自动排序）
- select_columns：只保留指定列
- group_aggregate：分组聚合（求和/平均/计数等）
- take_top_n：取前 n 行
- reset：撤销所有操作，恢复初始表格
- merge_tables：把多张表合并成一张（按姓名/考试号等键对齐）
- compute_rank：计算排名（总分/单科/班级/年级排名）
- classify_grades：按分数线把成绩划分成等级（A/B/C/D 或 优秀/良好/及格）
- score_segment_stats：统计每个分数段的人数
- compute_stats：统计平均分、及格率、优秀率等指标
- compare_scores：计算两次成绩的进退步（本次减上次）
- set_output_name：设置结果文件名称（当用户在查询中指定了输出文件名时调用）

规则：
1. 每次工具调用后，你会收到结果表格的预览（行数、列数和前几行）。根据预览判断是否已满足查询，必要时继续调用工具。
2. 列名必须与“表格信息”中给出的列名完全一致，不要臆造列名。
3. 多个过滤条件若要“且”关系，就放在一次 filter_rows 的 filters 数组里；若需要“或”关系，请分多次调用并在结果间合并（如需要）。
4. 对等级列（A+、A、B+、B、C…）排序时，系统会自动按标准等级顺序处理，无需手动填 value_order。
5. 排名、等级划分、分数段、平均分/及格率/优秀率、进退步等计算只对“数值列”有效；若列是等级（A/B/C）而非分数，则无法计算，请改用 sort/filter。
6. 多表合并时，用表格别名（表1、表2…）引用，并用公共键列（如姓名、考试号）对齐。
7. 完成所有操作后，不再调用工具，用一句简短的中文说明你对表格做了什么。
8. 如果用户在查询中指定了输出文件名（如“保存为排名表.xlsx”“命名为xxx”），请调用 set_output_name 工具设置文件名。

标准等级高低顺序（从高到低，自动套用）：A+ > A > A- > B+ > B > B- > C+ > C > C- > D+ > D > D- > E。"""


# ---------------------------------------------------------------------------
# 工具执行器
# ---------------------------------------------------------------------------
def _apply_op(df, col, op, value):
    s = df[col]
    if op in ("is_empty", "not_empty"):
        empty = s.isna() | (s.astype(str).str.strip() == "")
        return ~empty if op == "not_empty" else empty

    # 数值列：把字符串 value 转为数值再比较
    if pd.api.types.is_numeric_dtype(s):
        try:
            value = float(value)
            s = s.astype(float)
        except (TypeError, ValueError):
            pass

    if op == "==":
        return s == value
    if op == "!=":
        return s != value
    if op == ">":
        return s > value
    if op == "<":
        return s < value
    if op == ">=":
        return s >= value
    if op == "<=":
        return s <= value
    if op == "contains":
        return s.astype(str).str.contains(str(value), na=False)
    if op == "not_contains":
        return ~s.astype(str).str.contains(str(value), na=False)
    if op == "startswith":
        return s.astype(str).str.startswith(str(value), na=False)
    if op == "endswith":
        return s.astype(str).str.endswith(str(value), na=False)
    raise ValueError(f"未知操作符：{op}")


class TableAgent:
    def __init__(self, api_key=None, model=None, verbose=True):
        self.api_key = api_key or resolve_api_key()
        self.model = model or os.environ.get("GLM_MODEL", DEFAULT_MODEL)
        self.verbose = verbose
        self.chat_url = ZHIPU_BASE_URL + "/chat/completions"

        self.original = None    # 原始表（第一张加载的表）
        self.df = None          # 当前工作表
        self.columns = []
        self.tables = {}        # 表格别名 -> DataFrame（支持多表合并）
        self.table_names = {}   # 表格别名 -> 原始文件名
        self.output_name = None  # 用户/模型指定的输出文件名

    # ---- 表格加载 ----
    def load_tables(self, table_paths):
        """加载一张或多张表到注册表，第一张作为当前工作表。"""
        if isinstance(table_paths, (str, os.PathLike)):
            table_paths = [table_paths]
        self.tables = {}
        self.table_names = {}
        for i, p in enumerate(table_paths, start=1):
            alias = f"表{i}"
            self.tables[alias] = pd.read_excel(p, header=0)
            self.table_names[alias] = str(Path(p).stem)
        first = list(self.tables.keys())[0]
        self.original = self.tables[first].copy()
        self.df = self.original.copy()
        self.columns = [str(c) for c in self.df.columns]
        return self.original

    def load_table(self, table_path):
        """加载单张表（兼容旧用法）。"""
        return self.load_tables(table_path)

    # ---- 与模型交互 ----
    def _chat(self, messages):
        payload = {
            "model": self.model,
            "messages": messages,
            "tools": TOOLS,
            "temperature": TEMPERATURE,
        }
        return _http_post_json(self.chat_url, payload, self.api_key)

    # ---- 工具调度 ----
    def _execute(self, name, args):
        df = self.df

        if name == "filter_rows":
            mask = pd.Series(True, index=df.index)
            for f in args.get("filters", []):
                col = f.get("column")
                if col not in df.columns:
                    return f"错误：列「{col}」不存在，可用列：{self.columns}"
                mask = mask & _apply_op(df, col, f.get("op", "=="), f.get("value"))
            df = df[mask].reset_index(drop=True)

        elif name == "sort_rows":
            by = args.get("by", [])
            if isinstance(by, str):          # 兼容模型误传单个字符串列名
                by = [by]
            ascending = args.get("ascending", True)
            value_order = dict(args.get("value_order") or {})
            missing = [c for c in by if c not in df.columns]
            if missing:
                return f"错误：列 {missing} 不存在，可用列：{self.columns}"
            if isinstance(ascending, bool):
                ascending = [ascending] * len(by)
            if len(ascending) != len(by):
                return f"错误：ascending 长度({len(ascending)})与 by 长度({len(by)})不一致"
            # 未显式指定顺序的排序列，若其值是等级（A+/A/B+/B…），自动套用标准等级顺序
            for c in by:
                if c not in value_order and _is_grade_column(df[c]):
                    value_order[c] = _GRADE_ORDER_DESC
            df2 = df.copy()
            for c, order in value_order.items():
                if c in df2.columns and order:
                    # value_order 语义为“从高到低”，而 pandas Categorical 的
                    # categories 索引 0 是最低值，因此反转为“从低到高”。
                    df2[c] = pd.Categorical(
                        df2[c].astype(str),
                        categories=[str(x) for x in reversed(order)],
                        ordered=True,
                    )
            df = df2.sort_values(by=by, ascending=ascending).reset_index(drop=True)

        elif name == "select_columns":
            cols = args.get("columns", [])
            missing = [c for c in cols if c not in df.columns]
            if missing:
                return f"错误：列 {missing} 不存在，可用列：{self.columns}"
            df = df[cols].copy()

        elif name == "group_aggregate":
            group_by = args.get("group_by", [])
            aggs = args.get("aggregations", [])
            missing = [c for c in group_by if c not in df.columns]
            if missing:
                return f"错误：分组列 {missing} 不存在，可用列：{self.columns}"
            agg_dict = {}
            for a in aggs:
                col = a.get("column")
                if col not in df.columns:
                    return f"错误：聚合列「{col}」不存在，可用列：{self.columns}"
                func = _FUNC_MAP.get(str(a.get("func", "")).strip().lower(), a.get("func"))
                agg_dict[col] = func
            df = df.groupby(group_by, as_index=False).agg(agg_dict)

        elif name == "take_top_n":
            n = int(args.get("n", 1))
            df = df.head(n).reset_index(drop=True)

        elif name == "reset":
            df = self.original.copy()

        elif name == "merge_tables":
            tables = args.get("tables", [])
            on = args.get("on", [])
            if isinstance(on, str):
                on = [on]
            how = args.get("how", "outer")
            if not tables:
                return "错误：请用 tables 指定要合并的表格别名"
            unknown = [t for t in tables if t not in self.tables]
            if unknown:
                return f"错误：表格别名 {unknown} 不存在，可用：{list(self.tables.keys())}"
            merged = self.tables[tables[0]]
            for t in tables[1:]:
                merged = merged.merge(self.tables[t], on=on, how=how)
            df = merged

        elif name == "compute_rank":
            col = args.get("column")
            new_col = args.get("new_column") or (str(col) + "排名")
            group_by = args.get("group_by", [])
            if isinstance(group_by, str):
                group_by = [group_by]
            ascending = args.get("ascending", False)
            method = args.get("method", "min")
            if col not in df.columns:
                return f"错误：列「{col}」不存在，可用列：{self.columns}"
            missing = [g for g in group_by if g not in df.columns]
            if missing:
                return f"错误：分组列 {missing} 不存在，可用列：{self.columns}"
            s = pd.to_numeric(df[col], errors="coerce")
            if group_by:
                tmp = df.copy()
                tmp["_rank_score"] = s
                r = tmp.groupby(group_by)["_rank_score"].rank(method=method, ascending=ascending)
            else:
                r = s.rank(method=method, ascending=ascending)
            df = df.copy()
            df[new_col] = r.astype("Int64")

        elif name == "classify_grades":
            col = args.get("column")
            new_col = args.get("new_column") or (str(col) + "等级")
            thresholds = args.get("thresholds", [])
            labels = args.get("labels", [])
            if col not in df.columns:
                return f"错误：列「{col}」不存在，可用列：{self.columns}"
            if len(labels) != len(thresholds) + 1:
                return f"错误：labels 长度应为 {len(thresholds) + 1}（比 thresholds 多 1），实际 {len(labels)}"
            s = pd.to_numeric(df[col], errors="coerce")
            bins = [-float("inf")] + sorted(thresholds) + [float("inf")]
            labels_asc = list(reversed(labels))
            df = df.copy()
            df[new_col] = pd.cut(s, bins=bins, labels=labels_asc, right=False).astype(object)

        elif name == "score_segment_stats":
            col = args.get("column")
            bins = args.get("bins", [])
            if col not in df.columns:
                return f"错误：列「{col}」不存在，可用列：{self.columns}"
            if len(bins) < 2:
                return "错误：bins 至少需要 2 个分界值"
            s = pd.to_numeric(df[col], errors="coerce")
            labels = [f"{bins[i]}-{bins[i + 1]}" for i in range(len(bins) - 1)]
            cut = pd.cut(s, bins=bins, labels=labels, right=True, include_lowest=True)
            counts = cut.value_counts().sort_index()
            df = pd.DataFrame({"分数段": [str(x) for x in counts.index], "人数": [int(x) for x in counts.values]})

        elif name == "compute_stats":
            columns = args.get("columns", [])
            if isinstance(columns, str):
                columns = [columns]
            group_by = args.get("group_by", [])
            if isinstance(group_by, str):
                group_by = [group_by]
            pass_line = float(args.get("pass_line", 60))
            excellent_line = float(args.get("excellent_line", 90))
            for c in columns + group_by:
                if c not in df.columns:
                    return f"错误：列「{c}」不存在，可用列：{self.columns}"
            long = df.melt(id_vars=group_by, value_vars=columns, var_name="科目", value_name="得分")
            long["得分"] = pd.to_numeric(long["得分"], errors="coerce")
            keys = group_by + ["科目"]

            def _agg(g):
                s = g["得分"].dropna()
                n = len(s)
                return pd.Series({
                    "人数": n,
                    "平均分": round(float(s.mean()), 2) if n else 0.0,
                    "及格率(%)": round(float((s >= pass_line).mean() * 100), 2) if n else 0.0,
                    "优秀率(%)": round(float((s >= excellent_line).mean() * 100), 2) if n else 0.0,
                })

            df = long.groupby(keys).apply(_agg).reset_index()

        elif name == "compare_scores":
            col_a = args.get("column")
            col_b = args.get("base_column")
            new_col = args.get("new_column") or "进退步"
            for c in (col_a, col_b):
                if c not in df.columns:
                    return f"错误：列「{c}」不存在，可用列：{self.columns}"
            a = pd.to_numeric(df[col_a], errors="coerce")
            b = pd.to_numeric(df[col_b], errors="coerce")
            df = df.copy()
            df[new_col] = a - b

        elif name == "set_output_name":
            nm = _sanitize_output_name(args.get("name", ""))
            if not nm:
                return "错误：输出文件名无效"
            self.output_name = nm
            return f"已设置结果文件名为：{nm}（当前表格不变）"

        else:
            return f"错误：未知工具 {name}"

        self.df = df
        self.columns = [str(c) for c in df.columns]
        return self._preview(df)

    @staticmethod
    def _preview(df):
        parts = [
            f"当前表格：{len(df)} 行 × {len(df.columns)} 列。",
            "列名：" + "、".join(str(c) for c in df.columns),
            "前几行预览：",
        ]
        parts.append(df.head(PREVIEW_ROWS).to_string(max_cols=30, max_colwidth=20))
        return "\n".join(parts)

    # ---- 主流程 ----
    def run(self, table_paths, query, on_event=None):
        """运行智能体。on_event 可选：回调函数，接收事件 dict（供 GUI 实时展示进度）。

        事件类型：
            {"type": "status",  "text": ...}   处理中的状态提示
            {"type": "step",    "step": n, "tool": ..., "args": {...}}  一次工具调用
            {"type": "preview", "text": ...}   工具调用后的表格预览
            {"type": "answer",  "text": ...}   最终说明
            {"type": "done",    "out_path": ..., "rows": n, "cols": n}  完成
            {"type": "error",   "text": ...}   错误
        """
        def emit(event):
            if on_event:
                on_event(event)

        paths = [table_paths] if isinstance(table_paths, (str, os.PathLike)) else list(table_paths)
        self.output_name = None
        emit({"type": "status", "text": f"正在加载 {len(paths)} 张表格…"})
        self.load_tables(paths)
        if len(self.tables) == 1:
            schema = build_schema(self.original)
        else:
            schema = build_schema_multi(self.tables, self.table_names)

        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": f"表格信息：\n{schema}\n\n用户查询：{query}\n请使用工具完成操作。",
            },
        ]

        final_answer = ""
        try:
            for step in range(MAX_STEPS):
                emit({"type": "status", "text": f"第 {step + 1} 轮：正在思考…"})
                resp = self._chat(messages)
                msg = resp["choices"][0]["message"]
                messages.append(msg)

                tool_calls = msg.get("tool_calls")
                if not tool_calls:
                    final_answer = msg.get("content") or ""
                    break

                for tc in tool_calls:
                    fn = tc.get("function", {})
                    name = fn.get("name", "")
                    try:
                        args = json.loads(fn.get("arguments") or "{}")
                    except json.JSONDecodeError:
                        args = {}
                    if self.verbose:
                        print(f"[步骤 {step + 1}] 调用工具 {name} 参数={json.dumps(args, ensure_ascii=False)}")
                    emit({"type": "step", "step": step + 1, "tool": name, "args": args})
                    result = self._execute(name, args)
                    emit({"type": "preview", "text": result})
                    messages.append({"role": "tool", "tool_call_id": tc.get("id", ""), "content": result})
            else:
                final_answer = "(达到最大工具调用轮数，已按当前结果输出)"
        except Exception as e:
            emit({"type": "error", "text": str(e)})
            raise

        # 保存结果：优先使用查询/工具指定的文件名，否则自动生成
        output_name = self.output_name or _extract_output_name(query)
        base_dir = Path(paths[0]).parent
        if output_name:
            out_path = base_dir / output_name
        else:
            out_path = base_dir / (Path(paths[0]).stem + "_结果.xlsx")
        self.df.to_excel(out_path, index=False)
        emit({"type": "answer", "text": final_answer or "（无说明）"})
        emit({"type": "done", "out_path": str(out_path),
              "rows": len(self.df), "cols": len(self.df.columns)})

        if self.verbose:
            print("\n" + "=" * 60)
            print("处理说明：", final_answer or "（无说明）")
            print(f"结果已保存到：{out_path}")
            print(self._preview(self.df))

        return self.df, str(out_path), final_answer


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------
def main(argv=None):
    import argparse

    parser = argparse.ArgumentParser(description="基于 GLM 的表格智能体：对 xlsx 表格做排序/过滤/合并/统计等操作")
    parser.add_argument("--file", "-f", action="append", help="输入 .xlsx 文件路径（可多次指定，用于多表合并）")
    parser.add_argument("--query", "-q", help="中文查询，例如“按总分从高到低排序”")
    args = parser.parse_args(argv)

    table_paths = args.file
    query = args.query

    if not table_paths:
        table_paths = [input("请输入表格路径（.xlsx）：").strip().strip('"')]
    if not query:
        query = input("请输入你的中文查询：").strip()

    agent = TableAgent()
    agent.run(table_paths, query)


if __name__ == "__main__":
    main()
