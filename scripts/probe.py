#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""环境与表格结构体检。输出 JSON，供上层决定执行档位与清洗策略。

用法:
    python probe.py <xlsx路径> [--sheet 名称] [--max-mb 50]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile

# 中文 Windows 控制台默认 GBK，输出 ¥/全角空格等字符会直接抛 UnicodeEncodeError
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

FULLWIDTH_SPACE = "\u3000"

# 体检结论档位，与 SKILL.md 的降级策略一一对应
TIER_FULL = "full"        # pandas 可用，走完整清洗 + 汇总
TIER_DEGRADE = "degrade"  # 缺 pandas，退到 openpyxl 基础清洗
TIER_PARTIAL = "partial"  # 文件过大，仅抽样处理并明确告知
TIER_BLOCK = "block"      # 加密/损坏/不存在，直接拒绝


def check_deps() -> dict:
    found = {}
    for mod in ("pandas", "openpyxl", "numpy"):
        try:
            m = __import__(mod)
            found[mod] = getattr(m, "__version__", "unknown")
        except ImportError:
            found[mod] = None
    return found


def is_encrypted(path: str) -> bool:
    """加密的 xlsx 是 OLE2 容器而非 zip，读取会抛异常而不是返回错误结果。"""
    try:
        with open(path, "rb") as f:
            if f.read(8) == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                return True
    except OSError:
        return False
    return False


def check_file(path: str, max_mb: float) -> dict:
    info = {"path": path, "exists": False, "readable": False, "encrypted": False,
            "size_mb": None, "zip_ok": False, "problems": []}

    if not os.path.exists(path):
        info["problems"].append("文件不存在")
        return info
    info["exists"] = True

    if not os.path.isfile(path):
        info["problems"].append("路径是目录，不是文件")
        return info

    info["size_mb"] = round(os.path.getsize(path) / 1024 / 1024, 2)

    if is_encrypted(path):
        info["encrypted"] = True
        info["problems"].append("文件已加密，需要密码才能打开")
        return info

    try:
        with zipfile.ZipFile(path) as z:
            info["zip_ok"] = z.testzip() is None
    except (zipfile.BadZipFile, OSError) as e:
        info["problems"].append(f"文件结构损坏或非 xlsx 格式: {e}")
        return info

    if info["size_mb"] > max_mb:
        info["problems"].append(
            f"文件 {info['size_mb']}MB 超过阈值 {max_mb}MB，全量加载可能耗尽内存")

    info["readable"] = True
    return info


def sniff_header_row(rows: list[list], scan: int = 8) -> int:
    """猜表头行：取前 scan 行里非空单元格最多、且字符串占比最高的一行。

    业务表常见前几行是标题、制表人、空行，真表头往往在第 2~5 行。
    """
    best_idx, best_score = 0, -1.0
    for i, row in enumerate(rows[:scan]):
        cells = [c for c in row if c is not None and str(c).strip() != ""]
        if not cells:
            continue
        str_ratio = sum(1 for c in cells if isinstance(c, str)) / len(cells)
        score = len(cells) * (0.5 + str_ratio)
        if score > best_score:
            best_idx, best_score = i, score
    return best_idx


def profile_sheet(path: str, sheet: str | None) -> dict:
    """只读前若干行做结构快照，不整表加载——大文件也能秒回。"""
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheet_names = wb.sheetnames
        ws = wb[sheet] if sheet and sheet in sheet_names else wb[sheet_names[0]]

        sample = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            sample.append(list(row))
            if i >= 39:
                break

        merged = []
        # read_only 模式拿不到 merged_cells，需要单独以常规模式打开表头区域
        try:
            wb2 = openpyxl.load_workbook(path, read_only=False, data_only=True)
            merged = [str(r) for r in wb2[ws.title].merged_cells.ranges][:20]
            wb2.close()
        except (MemoryError, OSError):
            merged = ["(文件过大，跳过合并单元格检测)"]

        header_idx = sniff_header_row(sample)
        header = [str(c).strip() if c is not None else "" for c in sample[header_idx]] \
            if sample else []
        body = sample[header_idx + 1:]

        return {
            "sheet_names": sheet_names,
            "active_sheet": ws.title,
            "guessed_header_row": header_idx + 1,  # 1-based，给人看
            "header": header,
            "merged_ranges": merged,
            "sample_rows": body[:5],
            "signals": detect_dirt(header, body, merged),
        }
    finally:
        wb.close()


def detect_dirt(header: list[str], body: list[list], merged: list[str]) -> dict:
    """脏数据信号扫描。每一项都对应 references/excel-pitfalls.md 里的一个坑。"""
    flat = [c for row in body for c in row if c is not None]
    text = [str(c) for c in flat if isinstance(c, str)]

    unnamed = sum(1 for h in header if h == "" or h.lower().startswith("unnamed"))
    named = [h for h in header if h and not h.lower().startswith("unnamed")]
    dup_header = len(named) - len(set(named))

    date_like = sum(1 for t in text
                    if re.match(r"^\s*\d{4}[-/年]\d{1,2}[-/月]\d{1,2}", t))
    num_with_sep = sum(1 for t in text
                       if re.match(r"^\s*[¥$]?-?[\d,]+\.?\d*\s*[%元]?\s*$", t)
                       and ("," in t or "%" in t or "¥" in t or "元" in t))
    padded = sum(1 for t in text if t != t.strip() or FULLWIDTH_SPACE in t)

    empty_rows = sum(1 for row in body
                     if all(c is None or str(c).strip() == "" for c in row))

    # 整行序列化后比对，等价于 pandas 的 duplicated()，但不依赖 pandas
    seen, dup_rows = set(), 0
    for row in body:
        key = tuple("" if c is None else str(c) for c in row)
        if any(k for k in key):
            if key in seen:
                dup_rows += 1
            seen.add(key)

    return {
        "unnamed_columns": unnamed,
        "duplicated_column_names": dup_header,
        "merged_cells_in_sheet": len([m for m in merged if ":" in m]),
        "text_stored_dates": date_like,
        "text_stored_numbers": num_with_sep,
        "whitespace_polluted_cells": padded,
        "blank_rows_in_sample": empty_rows,
        "duplicate_rows_in_sample": dup_rows,
    }


def decide_tier(deps: dict, file_info: dict, max_mb: float) -> tuple[str, str]:
    if not file_info["exists"]:
        return TIER_BLOCK, "文件不存在，无法继续。请确认路径。"
    if file_info["encrypted"]:
        return TIER_BLOCK, "文件已加密。本技能不尝试破解密码，请先手动解密后重试。"
    if not file_info["readable"]:
        return TIER_BLOCK, "; ".join(file_info["problems"]) or "文件不可读。"
    if deps["openpyxl"] is None:
        return TIER_BLOCK, "缺少 openpyxl，无法读取 xlsx。请先安装：pip install openpyxl"
    if file_info["size_mb"] and file_info["size_mb"] > max_mb:
        return TIER_PARTIAL, (
            f"文件 {file_info['size_mb']}MB 较大，将改为分块/抽样处理，"
            "并在报告中说明未覆盖的范围。")
    if deps["pandas"] is None:
        return TIER_DEGRADE, (
            "未检测到 pandas，降级为 openpyxl 基础清洗："
            "可做去空行/去首尾空格/类型转换，不做透视汇总。")
    return TIER_FULL, "环境完备，执行完整清洗 + 汇总流程。"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--max-mb", type=float, default=50.0)
    args = ap.parse_args()

    deps = check_deps()
    file_info = check_file(args.path, args.max_mb)
    tier, reason = decide_tier(deps, file_info, args.max_mb)

    report = {"tier": tier, "reason": reason, "deps": deps, "file": file_info,
              "structure": None}

    if tier != TIER_BLOCK:
        try:
            report["structure"] = profile_sheet(args.path, args.sheet)
        except Exception as e:
            report["tier"] = TIER_BLOCK
            report["reason"] = f"读取表格内容失败: {type(e).__name__}: {e}"

    print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    return 0 if report["tier"] != TIER_BLOCK else 2


if __name__ == "__main__":
    sys.exit(main())
