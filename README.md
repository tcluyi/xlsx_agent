# 表格智能体（xlsx_agent）

基于 **GLM（智谱大模型）** 的表格处理智能体：输入一句中文查询，即可对 Excel 表格进行**排序、过滤、合并、排名、等级划分、分数段统计、指标统计、成绩对比**等操作，并生成新的 `.xlsx` 文件。

支持两种使用方式：

- **命令行**：`python agent.py --file 表格.xlsx --query "…"`
- **图形界面**：`python gui.py`，浏览器打开一个类大模型聊天界面，实时展示处理过程

---

## 功能一览

智能体通过「工具调用（function calling）」完成操作，当前共 **13 个工具**：

| 功能 | 工具 | 说明 |
|------|------|------|
| 过滤 / 筛选 | `filter_rows` | 按条件过滤行（数值比较、包含、前缀、为空等，多条件为"且"） |
| 排序 | `sort_rows` | 单列/多列排序，支持**多级优先级**，等级列（A+/A/B+…）自动按标准等级顺序 |
| 选列 | `select_columns` | 只保留指定列 |
| 分组聚合 | `group_aggregate` | 按列分组求和/平均/计数/最大/最小等 |
| 取前 N 行 | `take_top_n` | 常用于排序后取前几名 |
| 撤销 | `reset` | 恢复初始表格 |
| 多表合并 | `merge_tables` | 把多张表按公共键（姓名/考试号）合并成一张 |
| 排名计算 | `compute_rank` | 总分/单科/班级/年级排名，支持并列名次 |
| 等级划分 | `classify_grades` | 按分数线划 A/B/C/D，或优秀/良好/及格 |
| 分数段统计 | `score_segment_stats` | 统计每个分数段的人数 |
| 指标统计 | `compute_stats` | 平均分、及格率、优秀率（按班级/按学科） |
| 成绩对比 | `compare_scores` | 本次减上次，计算进退步 |
| 设置输出文件名 | `set_output_name` | 当查询中指定了输出文件名时使用 |

---

## 环境与依赖

- Python 3.8+
- `pandas`、`openpyxl`（读写 Excel）
- 调用智谱 API 使用标准库 `urllib`，**无需**额外安装 `requests`/`openai`

```powershell
pip install pandas openpyxl
```

---

## 密钥配置

智谱 API Key 的完整形式为 `{id}.{secret}`（`32位十六进制` + `.` + `16位字母数字`），认证头为 `Authorization: Bearer {id}.{secret}`。

两种配置方式（二选一）：

**方式一：环境变量（推荐）**

```powershell
$env:GLM_API_KEY_ID = "你的 api_key_ID（id 部分）"
$env:GLM_API_KEY    = "你的 api_key（secret 部分）"
$env:GLM_MODEL      = "glm-4.5-flash"   # 可选，默认 glm-4.5-flash，可换 glm-4-plus / glm-4-air 等
```

**方式二：直接写入代码**

编辑 `agent.py` 顶部 `CONFIG`：

```python
CONFIG = {
    "api_key_id": "你的 api_key_ID",
    "api_key":    "你的 api_key（secret）",
}
```

> 如果 `api_key` 本身就是完整的 `id.secret` 形式，也可直接放进 `api_key` 并留空 `api_key_id`。

---

## 使用方式

### 1. 命令行

```powershell
# 单张表
python agent.py --file "成绩汇总.xlsx" --query "按总分从高到低排序"

# 多张表（合并）
python agent.py -f 语文.xlsx -f 数学.xlsx -f 英语.xlsx --query "按姓名合并成总表"
```

- `--file` / `-f`：输入 `.xlsx` 文件路径，可多次指定
- `--query` / `-q`：中文查询
- 不传参数会进入交互式输入

### 2. 图形界面（聊天式）

```powershell
python gui.py
```

浏览器自动打开 `http://127.0.0.1:8000`。功能：

- 左侧列出目录下的 `.xlsx` 文件，可**多选**（对应多表合并）
- 支持**拖拽** xlsx 文件进窗口，或点「上传 xlsx」选择文件
- 输入查询后，处理过程（加载、每一步工具调用、表格预览、最终结果）像大模型对话一样**流式展示**
- 结果可一键下载

### 3. Python 代码调用

```python
from agent import TableAgent

agent = TableAgent()
df, out_path, answer = agent.run("成绩汇总.xlsx", "按总分从高到低排序")
# df        结果 DataFrame
# out_path  输出文件路径
# answer    智能体的处理说明

# 需要实时获取处理进度时，传入 on_event 回调
agent.run("成绩汇总.xlsx", "按班级排序", on_event=lambda ev: print(ev))
```

`on_event` 收到的事件类型：`status`（状态）、`step`（工具调用）、`preview`（表格预览）、`answer`（说明）、`done`（完成）、`error`（错误）。

---

## 查询示例

| 想做什么 | 查询 |
|----------|------|
| 排序 | 按总分从高到低排序 |
| 多级排序 | 先按班级排序，再按总分、语文、数学…从高到低排序 |
| 筛选 | 找出八年级1班的学生 |
| 筛选（不达标） | 找出语文低于 60 分的学生 |
| 合并 | 把各科成绩按姓名合并成一张总表 |
| 排名 | 计算总分的年级排名 |
| 班级排名 | 计算每个学生的班级排名（按班级分组） |
| 等级划分 | 把成绩按 90/80/70/60 划分成 A/B/C/D/E |
| 等级划分（及格） | 按 90/75/60 划分成优秀/良好/及格/不及格 |
| 分数段 | 统计 90-100、80-89、70-79… 各有多少人 |
| 指标 | 统计各班的平均分、及格率、优秀率 |
| 指标（按学科） | 统计语文、数学各科的平均分、及格率 |
| 成绩对比 | 计算本次总分相对上次总分的进退步 |
| 指定输出名 | 按总分排序，保存为 年级排名表.xlsx |

> 若不指定输出文件名，默认保存为 `原文件名_结果.xlsx`；指定了（如"保存为 xxx"）则用该名称。

---

## 表格格式说明

- **等级制表**（成绩为 A+/A/B+/B…）：支持排序、筛选、选列、合并、等级排序等
- **百分制表**（成绩为数值分数）：支持上述全部功能，以及排名、等级划分、分数段、平均分/及格率/优秀率、进退步等**数值计算**

> 排名、等级划分、分数段、指标统计、成绩对比等**只对数值列有效**——若列是等级（A/B/C）而非分数，无法做这些计算（智能体会自动改用排序/筛选）。

---

## 运行测试

```powershell
python test_agent.py
```

`test_agent.py` 覆盖全部功能（用内存数据，不调用 API、不依赖文件），当前共 **21 项**断言。

---

## 文件结构

```
xlsx_agent/
├── agent.py        # 核心智能体（工具定义 + 执行逻辑 + 命令行入口）
├── gui.py          # 图形界面服务端（本地 HTTP + 流式接口，纯标准库）
├── index.html      # 图形界面前端（聊天 UI）
├── test_agent.py   # 功能测试
└── README.md       # 本文件
```

## 打包
```
python -m PyInstaller --onefile --add-data "index.html;." --name 表格智能体 gui.py
```