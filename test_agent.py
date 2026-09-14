# -*- coding: utf-8 -*-
"""
agent.py 的功能测试（不调用 GLM API，直接测试工具执行逻辑）。

运行方式：
    python test_agent.py

每个功能打印一条 [PASS] / [FAIL]，最后给出汇总。
"""
import agent
import pandas as pd


def make_agent(df, tables=None):
    """构造一个绕过 API 密钥的 TableAgent，用于直接测试 _execute。"""
    a = agent.TableAgent.__new__(agent.TableAgent)
    a.verbose = False
    a.df = df.copy()
    a.original = df.copy()
    a.columns = [str(c) for c in df.columns]
    a.tables = tables or {"表1": df.copy()}
    a.table_names = {k: k for k in a.tables}
    return a


PASS = 0
FAIL = 0


def check(name, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"[PASS] {name}")
    else:
        FAIL += 1
        print(f"[FAIL] {name}  {detail}")


# ---------------------------------------------------------------------------
# 1. 过滤 / 筛选（找出某分数段、某班级、某学科不达标的学生）
# ---------------------------------------------------------------------------
df = pd.DataFrame({
    "班级": ["1班", "2班", "1班", "2班"],
    "姓名": ["甲", "乙", "丙", "丁"],
    "语文": [90, 45, 72, 58],
    "总分": [500, 380, 410, 300],
})
a = make_agent(df)
a._execute("filter_rows", {"filters": [
    {"column": "班级", "op": "==", "value": "1班"},
    {"column": "语文", "op": "<", "value": "60"},
]})
# 期望只剩 1班 且 语文<60 —— 实际数据里没有，应为 0 行
check("筛选-班级且不达标(空结果)", len(a.df) == 0, f"实际 {len(a.df)} 行")

a = make_agent(df)
a._execute("filter_rows", {"filters": [
    {"column": "总分", "op": ">=", "value": "400"},
    {"column": "总分", "op": "<=", "value": "600"},
]})
check("筛选-分数段 400~600", set(a.df["姓名"]) == {"甲", "丙"}, f"实际 {list(a.df['姓名'])}")


# ---------------------------------------------------------------------------
# 2. 排序（多级优先级 + 等级自动排序）
# ---------------------------------------------------------------------------
df2 = pd.DataFrame({
    "班级": ["1班", "1班", "2班", "1班"],
    "姓名": ["甲", "乙", "丙", "丁"],
    "总分": ["A", "A", "A+", "B+"],
    "语文": ["A", "B+", "A", "A"],
})
a = make_agent(df2)
a._execute("sort_rows", {"by": ["班级", "总分", "语文"], "ascending": [True, False, False]})
order = list(a.df["姓名"])
check("排序-班级升序+总分降序(等级自动)", order == ["甲", "乙", "丁", "丙"], f"实际 {order}")


# ---------------------------------------------------------------------------
# 3. 多表合并（各科成绩表合并成总表）
# ---------------------------------------------------------------------------
t1 = pd.DataFrame({"姓名": ["甲", "乙"], "语文": [90, 80]})
t2 = pd.DataFrame({"姓名": ["甲", "乙"], "数学": [85, 95]})
a = make_agent(t1, tables={"表1": t1, "表2": t2})
a._execute("merge_tables", {"tables": ["表1", "表2"], "on": ["姓名"], "how": "outer"})
merged = a.df
check("合并-列齐备", set(merged.columns) == {"姓名", "语文", "数学"}, f"实际 {list(merged.columns)}")
check("合并-数据对齐", merged.set_index("姓名").loc["乙", "数学"] == 95)


# ---------------------------------------------------------------------------
# 4. 排名计算（总分/班级/年级排名）
# ---------------------------------------------------------------------------
df3 = pd.DataFrame({
    "班级": ["1班", "1班", "2班", "2班", "1班"],
    "姓名": ["甲", "乙", "丙", "丁", "戊"],
    "总分": [90, 80, 85, 75, 90],
})
a = make_agent(df3)
a._execute("compute_rank", {"column": "总分", "new_column": "年级排名", "ascending": False})
check("排名-年级排名", list(a.df["年级排名"]) == [1, 4, 3, 5, 1], f"实际 {list(a.df['年级排名'])}")

a = make_agent(df3)
a._execute("compute_rank", {"column": "总分", "new_column": "班级排名", "group_by": ["班级"], "ascending": False})
check("排名-班级排名", list(a.df["班级排名"]) == [1, 3, 1, 2, 1], f"实际 {list(a.df['班级排名'])}")


# ---------------------------------------------------------------------------
# 5. 等级划分（A/B/C/D 与 优秀/良好/及格）
# ---------------------------------------------------------------------------
df4 = pd.DataFrame({"姓名": ["甲", "乙", "丙", "丁", "戊"], "成绩": [95, 85, 75, 65, 55]})
a = make_agent(df4)
a._execute("classify_grades", {"column": "成绩", "new_column": "等级",
                               "thresholds": [90, 80, 70, 60], "labels": ["A", "B", "C", "D", "E"]})
check("等级-A/B/C/D", list(a.df["等级"]) == ["A", "B", "C", "D", "E"], f"实际 {list(a.df['等级'])}")

a = make_agent(df4)
a._execute("classify_grades", {"column": "成绩", "new_column": "评级",
                               "thresholds": [90, 75, 60], "labels": ["优秀", "良好", "及格", "不及格"]})
check("等级-优秀/良好/及格", list(a.df["评级"]) == ["优秀", "良好", "良好", "及格", "不及格"],
      f"实际 {list(a.df['评级'])}")


# ---------------------------------------------------------------------------
# 6. 分数段统计
# ---------------------------------------------------------------------------
df5 = pd.DataFrame({"成绩": [95, 85, 75, 65, 55, 100]})
a = make_agent(df5)
a._execute("score_segment_stats", {"column": "成绩", "bins": [0, 60, 70, 80, 90, 100]})
seg = dict(zip(a.df["分数段"], a.df["人数"]))
check("分数段-90~100 有 2 人", seg.get("90-100") == 2, f"实际 {seg}")
check("分数段-各段人数", seg.get("0-60") == 1 and seg.get("60-70") == 1 and seg.get("70-80") == 1
      and seg.get("80-90") == 1, f"实际 {seg}")


# ---------------------------------------------------------------------------
# 7. 平均分/及格率/优秀率（按学科、按班级）
# ---------------------------------------------------------------------------
df6 = pd.DataFrame({
    "班级": ["1班", "1班", "2班", "2班", "1班"],
    "语文": [90, 80, 70, 60, 50],
    "数学": [85, 75, 65, 55, 95],
})
a = make_agent(df6)
a._execute("compute_stats", {"columns": ["语文", "数学"], "group_by": []})
by_subj = a.df.set_index("科目")
check("统计-按学科平均分", abs(by_subj.loc["语文", "平均分"] - 70) < 1e-6
      and abs(by_subj.loc["数学", "平均分"] - 75) < 1e-6,
      f"实际 {dict(by_subj['平均分'])}")
check("统计-语文及格率80%", by_subj.loc["语文", "及格率(%)"] == 80.0)

a = make_agent(df6)
a._execute("compute_stats", {"columns": ["语文"], "group_by": ["班级"]})
by_cls = a.df.set_index("班级")
check("统计-按班级人数", by_cls.loc["1班", "人数"] == 3 and by_cls.loc["2班", "人数"] == 2)


# ---------------------------------------------------------------------------
# 8. 成绩对比（进退步）
# ---------------------------------------------------------------------------
df7 = pd.DataFrame({"姓名": ["甲", "乙"], "本次总分": [100, 90], "上次总分": [90, 95]})
a = make_agent(df7)
a._execute("compare_scores", {"column": "本次总分", "base_column": "上次总分", "new_column": "进退步"})
check("对比-进退步", list(a.df["进退步"]) == [10, -5], f"实际 {list(a.df['进退步'])}")


# ---------------------------------------------------------------------------
# 9. 输出文件名（正则提取 + 工具指定）
# ---------------------------------------------------------------------------
check("文件名-提取保存为", agent._extract_output_name("按总分排序，保存为 排名表.xlsx") == "排名表.xlsx")
check("文件名-无后缀自动补", agent._extract_output_name("命名为 一班学生") == "一班学生.xlsx")
check("文件名-未指定返回None", agent._extract_output_name("按总分从高到低排序") is None)
check("文件名-非法字符清理", agent._sanitize_output_name('结果表?.xlsx') == "结果表.xlsx")
a = make_agent(pd.DataFrame({"a": [1]}))
a._execute("set_output_name", {"name": "我的结果"})
check("文件名-工具设置", a.output_name == "我的结果.xlsx")


# ---------------------------------------------------------------------------
# 10. 工具数量自检
# ---------------------------------------------------------------------------
names = [t["function"]["name"] for t in agent.TOOLS]
expect = {"filter_rows", "sort_rows", "select_columns", "group_aggregate", "take_top_n",
          "reset", "merge_tables", "compute_rank", "classify_grades",
          "score_segment_stats", "compute_stats", "compare_scores", "set_output_name"}
check("工具数量=13", len(names) == 13 and set(names) == expect, f"实际 {names}")


print("\n" + "=" * 50)
print(f"共 {PASS + FAIL} 项，通过 {PASS}，失败 {FAIL}")
if FAIL:
    raise SystemExit(1)
