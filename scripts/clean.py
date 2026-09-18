#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""计划驱动的清洗执行器。

模型只输出结构化计划（JSON），本脚本用白名单算子执行——不 eval / 不 exec，
能做什么完全由 OPS 决定。这样清洗过程可审计、可回放、可单测。

用法:
    python clean.py --input 原表.xlsx --plan plan.json --output 清洗后.xlsx
    python clean.py --input 原表.xlsx --plan-json '{"header_row":4,...}' --output out.xlsx
    python clean.py --print-schema
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

FULLWIDTH_SPACE = "\u3000"

PLAN_SCHEMA = {
    "sheet": "字符串，可选。要处理的 sheet 名，缺省取第一个",
    "header_row": "整数，1-based。真表头所在行，跳过装饰行用",
    "drop_columns": "字符串数组，可选。要删除的列名（如幽灵空列）",
    "rename_columns": "对象，可选。{旧列名: 新列名}",
    "strip_whitespace": "布尔，默认 true。清理首尾空格与全角空格",
    "to_numeric": "字符串数组，可选。转数值的列（自动剥离 ¥ , % 元）",
    "to_datetime": "字符串数组，可选。转日期的列（兼容 2026年1月1日 / 2026-01-01 / 2026/1/1）",
    "drop_blank_rows": "布尔，默认 true。删除全空行",
    "drop_duplicates": "布尔，默认 false。整行去重",
    "drop_total_rows": "字符串数组，可选。在这些列里出现 合计/小计/总计 的行视为汇总行删除",
    "filter_rows": "对象数组，可选。[{column, op, value}]，op ∈ eq/ne/gt/gte/lt/lte/contains/notna",
    "pivot": "对象，可选。{index:[列], columns:[列], values:[列], agg:sum|mean|count|max|min}",
    "sort_by": "对象数组，可选。[{column, ascending}]",
}

TOTAL_WORDS = ("合计", "小计", "总计", "汇总", "total", "subtotal")
AGG_WHITELIST = {"sum", "mean", "count", "max", "min", "median", "nunique"}
FILTER_OPS = {"eq", "ne", "gt", "gte", "lt", "lte", "contains", "notna"}


class PlanError(ValueError):
    """计划本身不合法——属于可回喂给模型修正的错误。"""


def clean_text(v):
    if not isinstance(v, str):
        return v
    s = v.replace(FULLWIDTH_SPACE, " ").strip()
    return s


def to_num(v):
    """剥离货币符号与千分位后转数值；百分号按比例还原。"""
    if v is None or isinstance(v, (int, float)):
        return v
    s = str(v).replace(FULLWIDTH_SPACE, " ").strip()
    if s == "":
        return None
    pct = s.endswith("%")
    s = re.sub(r"[¥$￥,，\s元]", "", s)
    s = s.rstrip("%")
    if s in ("", "-", "--"):
        return None
    try:
        n = float(s)
    except ValueError:
        return None
    return n / 100 if pct else n


def to_date(v):
    import pandas as pd

    if v is None:
        return None
    if not isinstance(v, str):
        return pd.to_datetime(v, errors="coerce")
    s = v.replace(FULLWIDTH_SPACE, " ").strip()
    # 中文日期先归一成标准分隔符，再交给 pandas
    s = re.sub(r"[年月]", "-", s).rstrip("日").rstrip("-")
    return pd.to_datetime(s, errors="coerce")


def require_columns(df, cols, field: str):
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise PlanError(
            f"计划字段 {field} 引用了不存在的列 {missing}；"
            f"当前可用列为 {list(df.columns)}")


def dedup_column_names(cols: list[str]) -> list[str]:
    """同名列加后缀，避免后续按列名索引时拿到 DataFrame 而非 Series。"""
    seen, out = {}, []
    for c in cols:
        name = c if c else "未命名列"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        out.append(name)
    return out


def run(inp: str, plan: dict, out: str) -> dict:
    import pandas as pd

    log: list[str] = []

    header_row = plan.get("header_row", 1)
    if not isinstance(header_row, int) or header_row < 1:
        raise PlanError(f"header_row 必须是 >=1 的整数，收到 {header_row!r}")

    sheet = plan.get("sheet", 0)
    df = pd.read_excel(inp, sheet_name=sheet, header=header_row - 1, dtype=object)
    log.append(f"读取 {inp}｜sheet={sheet if sheet != 0 else '第一个'}｜"
               f"表头第 {header_row} 行｜{df.shape[0]} 行 {df.shape[1]} 列")

    df.columns = dedup_column_names(
        ["" if str(c).startswith("Unnamed:") else str(c).strip() for c in df.columns])

    # 全空的幽灵列先删，避免它们污染后续列名引用
    ghost = [c for c in df.columns if df[c].isna().all()]
    if ghost:
        df = df.drop(columns=ghost)
        log.append(f"删除全空列 {ghost}")

    if plan.get("drop_columns"):
        hit = [c for c in plan["drop_columns"] if c in df.columns]
        if hit:
            df = df.drop(columns=hit)
            log.append(f"按计划删除列 {hit}")

    if plan.get("rename_columns"):
        mapping = {k: v for k, v in plan["rename_columns"].items() if k in df.columns}
        if mapping:
            df = df.rename(columns=mapping)
            log.append(f"重命名列 {mapping}")

    if plan.get("strip_whitespace", True):
        obj_cols = [c for c in df.columns if df[c].dtype == object]
        for c in obj_cols:
            df[c] = df[c].map(clean_text)
        log.append(f"清理 {len(obj_cols)} 个文本列的首尾空格与全角空格")

        # 只含空格的单元格清理后变成空串，仍非 NA——归一后才能识别出幽灵列/空行
        df = df.replace({"": None})
        ghost2 = [c for c in df.columns if df[c].isna().all()]
        if ghost2:
            df = df.drop(columns=ghost2)
            log.append(f"删除仅含空格的幽灵列 {ghost2}")

    if plan.get("drop_total_rows"):
        require_columns(df, plan["drop_total_rows"], "drop_total_rows")
        mask = pd.Series(False, index=df.index)
        for c in plan["drop_total_rows"]:
            col = df[c].astype(str).str.lower()
            for w in TOTAL_WORDS:
                mask |= col.str.contains(w, na=False, regex=False)
        n = int(mask.sum())
        if n:
            df = df[~mask]
            log.append(f"删除汇总行 {n} 行（匹配 合计/小计/总计 等）")

    if plan.get("drop_blank_rows", True):
        before = df.shape[0]
        df = df.dropna(how="all")
        if before != df.shape[0]:
            log.append(f"删除全空行 {before - df.shape[0]} 行")

    for c in plan.get("to_numeric") or []:
        require_columns(df, [c], "to_numeric")
        df[c] = df[c].map(to_num)
        df[c] = pd.to_numeric(df[c], errors="coerce")
        log.append(f"列「{c}」转为数值，失败 {int(df[c].isna().sum())} 个置空")

    for c in plan.get("to_datetime") or []:
        require_columns(df, [c], "to_datetime")
        df[c] = df[c].map(to_date)
        df[c] = pd.to_datetime(df[c], errors="coerce")
        log.append(f"列「{c}」转为日期，失败 {int(df[c].isna().sum())} 个置空")

    for f in plan.get("filter_rows") or []:
        col, op, val = f.get("column"), f.get("op"), f.get("value")
        if op not in FILTER_OPS:
            raise PlanError(f"不支持的过滤操作 {op!r}，可用: {sorted(FILTER_OPS)}")
        require_columns(df, [col], "filter_rows")
        s = df[col]
        before = df.shape[0]
        if op == "eq":
            df = df[s == val]
        elif op == "ne":
            df = df[s != val]
        elif op == "gt":
            df = df[pd.to_numeric(s, errors="coerce") > val]
        elif op == "gte":
            df = df[pd.to_numeric(s, errors="coerce") >= val]
        elif op == "lt":
            df = df[pd.to_numeric(s, errors="coerce") < val]
        elif op == "lte":
            df = df[pd.to_numeric(s, errors="coerce") <= val]
        elif op == "contains":
            df = df[s.astype(str).str.contains(str(val), na=False, regex=False)]
        elif op == "notna":
            df = df[s.notna()]
        log.append(f"过滤 {col} {op} {val!r}：{before} → {df.shape[0]} 行")

    if plan.get("drop_duplicates"):
        before = df.shape[0]
        df = df.drop_duplicates()
        if before != df.shape[0]:
            log.append(f"整行去重，删除 {before - df.shape[0]} 行")

    for s in plan.get("sort_by") or []:
        require_columns(df, [s["column"]], "sort_by")
        df = df.sort_values(s["column"], ascending=s.get("ascending", True))
        log.append(f"按 {s['column']} {'升序' if s.get('ascending', True) else '降序'} 排序")

    cleaned = df.reset_index(drop=True)
    result = {"rows": int(cleaned.shape[0]), "cols": int(cleaned.shape[1]),
              "columns": [str(c) for c in cleaned.columns], "log": log,
              "output": out, "sheets_written": ["清洗后明细"]}

    pv = plan.get("pivot")
    pivot_df = None
    if pv:
        agg = pv.get("agg", "sum")
        if agg not in AGG_WHITELIST:
            raise PlanError(f"不支持的聚合方式 {agg!r}，可用: {sorted(AGG_WHITELIST)}")
        index, values = pv.get("index") or [], pv.get("values") or []
        columns = pv.get("columns") or []
        if not index or not values:
            raise PlanError("pivot 需要同时提供 index 与 values")
        require_columns(cleaned, index + columns + values, "pivot")
        pivot_df = cleaned.pivot_table(
            index=index, columns=columns or None, values=values,
            aggfunc=agg, margins=True, margins_name="合计")
        log.append(f"透视汇总：按 {index}"
                   f"{' × ' + str(columns) if columns else ''} "
                   f"对 {values} 做 {agg}")
        result["sheets_written"].append("汇总")

    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        cleaned.to_excel(w, sheet_name="清洗后明细", index=False)
        if pivot_df is not None:
            pivot_df.to_excel(w, sheet_name="汇总")
    log.append(f"写出 {out}")

    return result


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input")
    ap.add_argument("--plan", help="计划 JSON 文件路径")
    ap.add_argument("--plan-json", help="直接传入计划 JSON 字符串")
    ap.add_argument("--output")
    ap.add_argument("--print-schema", action="store_true")
    args = ap.parse_args()

    if args.print_schema:
        print(json.dumps(PLAN_SCHEMA, ensure_ascii=False, indent=2))
        return 0

    if not (args.input and args.output and (args.plan or args.plan_json)):
        ap.error("需要 --input、--output，以及 --plan 或 --plan-json")

    raw = args.plan_json
    if args.plan:
        with open(args.plan, encoding="utf-8") as f:
            raw = f.read()
    try:
        plan = json.loads(raw)
    except json.JSONDecodeError as e:
        print(json.dumps({"ok": False, "error": f"计划不是合法 JSON: {e}"},
                         ensure_ascii=False, indent=2))
        return 1

    try:
        res = run(args.input, plan, args.output)
    except PlanError as e:
        # 计划级错误：信息足够具体，模型据此改一版计划即可
        print(json.dumps({"ok": False, "error_type": "PlanError", "error": str(e)},
                         ensure_ascii=False, indent=2))
        return 1
    except Exception as e:
        print(json.dumps({"ok": False, "error_type": type(e).__name__,
                          "error": str(e)}, ensure_ascii=False, indent=2))
        return 1

    print(json.dumps({"ok": True, **res}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
