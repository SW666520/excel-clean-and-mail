#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""清洗前后 diff 校验。输出 JSON 结论，pass=false 时携带可回喂给模型的错误描述。

用法:
    python verify.py --before 原表.xlsx --after 清洗后.xlsx [--expect-rows-max-drop 0.5]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

# 中文 Windows 控制台默认 GBK，输出预览数据里的特殊字符会抛 UnicodeEncodeError
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

# 行数骤减是最危险的静默故障：代码"跑通了"但数据被误删
DEFAULT_MAX_DROP = 0.5


def load(path: str, raw: bool = True):
    """raw=True 保留原始文本形态（看原表脏在哪）；
    raw=False 让 pandas 自然推断类型（才能验证类型是否真被修好）。"""
    import pandas as pd
    return pd.read_excel(path, dtype=object) if raw else pd.read_excel(path)


def col_profile(df) -> dict:
    import pandas as pd

    out = {}
    for c in df.columns:
        s = df[c]
        non_null = s.dropna()
        numeric = pd.to_numeric(non_null, errors="coerce").notna().sum()
        datetime_like = 0
        if len(non_null):
            datetime_like = int(pd.to_datetime(
                non_null, errors="coerce", format="mixed").notna().sum())
        out[str(c)] = {
            "dtype": str(s.dtype),
            "non_null": int(non_null.shape[0]),
            "null": int(s.isna().sum()),
            "numeric_parsable": int(numeric),
            "datetime_parsable": datetime_like,
            "unique": int(non_null.nunique()),
        }
    return out


def blank_rows(df) -> int:
    return int(df.isna().all(axis=1).sum())


def check(before_path: str, after_path: str, max_drop: float) -> dict:
    result = {"pass": False, "errors": [], "warnings": [], "diff": {}, "summary": ""}

    if not os.path.exists(after_path):
        result["errors"].append(
            f"输出文件未生成: {after_path}。清洗代码可能执行失败或写入路径不对。")
        return result

    try:
        b = load(before_path, raw=True)
        a = load(after_path, raw=False)
    except Exception as e:
        result["errors"].append(f"读取待比对文件失败: {type(e).__name__}: {e}")
        return result

    b_rows, a_rows = int(b.shape[0]), int(a.shape[0])
    b_cols = [str(c) for c in b.columns]
    a_cols = [str(c) for c in a.columns]

    result["diff"] = {
        "rows": {"before": b_rows, "after": a_rows, "delta": a_rows - b_rows},
        "cols": {"before": len(b_cols), "after": len(a_cols),
                 "removed": [c for c in b_cols if c not in a_cols],
                 "added": [c for c in a_cols if c not in b_cols]},
        "blank_rows": {"before": blank_rows(b), "after": blank_rows(a)},
        "columns_after": col_profile(a),
        "preview_after": a.head(3).fillna("").astype(str).to_dict(orient="records"),
    }

    if a_rows == 0:
        result["errors"].append("清洗后数据为空（0 行）。过滤条件过强或表头行判断错误。")

    if b_rows > 0:
        drop = 1 - a_rows / b_rows
        if drop > max_drop:
            result["errors"].append(
                f"行数从 {b_rows} 降到 {a_rows}（减少 {drop:.0%}），"
                f"超过允许阈值 {max_drop:.0%}，疑似误删数据。")
        elif drop > 0:
            result["warnings"].append(
                f"行数减少 {b_rows - a_rows} 行（{drop:.0%}），"
                "若为去重/去空行则属预期。")

    if len(a_cols) == 0:
        result["errors"].append("清洗后没有任何列。")

    unnamed_after = [c for c in a_cols
                     if c.strip() == ""
                     or c.lower().startswith("unnamed")
                     or c.startswith("未命名列")]
    if unnamed_after:
        ratio = len(unnamed_after) / len(a_cols)
        if ratio > 0.5:
            # 过半列名缺失，几乎可以断定 header_row 选错了——这是本领域最高频的故障
            result["errors"].append(
                f"{len(unnamed_after)}/{len(a_cols)} 列没有有效列名（{unnamed_after[:5]}），"
                "表头行几乎肯定选错了。请调大 header_row 重新定位真表头。")
        else:
            result["warnings"].append(
                f"仍存在未命名列 {unnamed_after}，表头可能未正确识别。")

    # 真表头以文本为主；若列名本身长得像日期或纯数字，说明把数据行当成了表头
    header_looks_like_data = [
        c for c in a_cols
        if re.match(r"^\s*\d{4}[-/年]\d{1,2}", c) or re.match(r"^-?[\d.,]+$", c)]
    if header_looks_like_data:
        result["errors"].append(
            f"列名 {header_looks_like_data} 看起来是数据而不是字段名，"
            "说明 header_row 指到了数据行。请重新定位真表头。")

    if result["diff"]["blank_rows"]["after"] > 0:
        result["warnings"].append(
            f"清洗后仍有 {result['diff']['blank_rows']['after']} 个全空行。")

    # 表头行修正后前后列名可能完全不重叠，因此只按清洗后的结果判断类型是否到位
    fixed, still_text = [], []
    for c in a_cols:
        prof_a = result["diff"]["columns_after"][c]
        if prof_a["non_null"] == 0:
            continue
        is_object = prof_a["dtype"] == "object"
        num_ratio = prof_a["numeric_parsable"] / prof_a["non_null"]
        dt_ratio = prof_a["datetime_parsable"] / prof_a["non_null"]
        if not is_object:
            fixed.append(c)
        elif num_ratio > 0.9 or dt_ratio > 0.9:
            still_text.append(c)

    result["diff"]["typed_columns"] = fixed
    if still_text:
        result["warnings"].append(
            f"列 {still_text} 内容看起来是数字/日期但仍以文本存储，建议转换类型。")

    result["pass"] = not result["errors"]
    result["summary"] = (
        f"{'通过' if result['pass'] else '未通过'}｜"
        f"行 {b_rows}→{a_rows}，列 {len(b_cols)}→{len(a_cols)}，"
        f"空行 {result['diff']['blank_rows']['before']}→"
        f"{result['diff']['blank_rows']['after']}，"
        f"已定型列 {len(fixed)} 个，"
        f"错误 {len(result['errors'])} 项，提醒 {len(result['warnings'])} 项")
    return result


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--before", required=True)
    ap.add_argument("--after", required=True)
    ap.add_argument("--expect-rows-max-drop", type=float, default=DEFAULT_MAX_DROP)
    args = ap.parse_args()

    r = check(args.before, args.after, args.expect_rows_max_drop)
    print(json.dumps(r, ensure_ascii=False, indent=2, default=str))
    return 0 if r["pass"] else 1


if __name__ == "__main__":
    sys.exit(main())
