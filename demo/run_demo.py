#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""面试演示脚本：一条命令跑完四个场景，避免现场逐条敲命令。

用法:
    python demo/run_demo.py          # 全部场景
    python demo/run_demo.py 1        # 只跑场景 1
"""
from __future__ import annotations

import json
import os
import subprocess
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEMO = os.path.join(ROOT, "demo")
DIRTY = os.path.join(DEMO, "dirty_sales.xlsx")
PLAN = os.path.join(DEMO, "plan_sales.json")
OUT = os.path.join(DEMO, "out_cleaned.xlsx")

BAR = "=" * 66


def run(args: list[str], env: dict | None = None) -> tuple[int, dict | None, str]:
    e = dict(os.environ)
    if env:
        e.update(env)
    p = subprocess.run([sys.executable, *args], cwd=ROOT, env=e,
                       capture_output=True)
    text = p.stdout.decode("utf-8", errors="replace")
    try:
        return p.returncode, json.loads(text), text
    except json.JSONDecodeError:
        return p.returncode, None, text + p.stderr.decode("utf-8", errors="replace")


def section(n: int, title: str) -> None:
    print(f"\n{BAR}\n场景 {n}：{title}\n{BAR}")


def scene1() -> None:
    section(1, "正常清洗 —— 体检 → 计划 → 执行 → 校验")

    print("\n[1/3] probe.py 体检")
    _, d, raw = run(["scripts/probe.py", DIRTY])
    if not d:
        print(raw)
        return
    s = d["structure"]
    print(f"  档位          : {d['tier']}（{d['reason']}）")
    print(f"  sheet 列表    : {s['sheet_names']}")
    print(f"  真表头推断    : 第 {s['guessed_header_row']} 行")
    print("  脏数据信号    :")
    for k, v in s["signals"].items():
        if v:
            print(f"    - {k}: {v}")

    print("\n[2/3] clean.py 按计划执行")
    _, d, raw = run(["scripts/clean.py", "--input", DIRTY,
                     "--plan", PLAN, "--output", OUT])
    if not d or not d.get("ok"):
        print(raw)
        return
    for line in d["log"]:
        print(f"    {line}")

    print("\n[3/3] verify.py 前后校验")
    code, d, raw = run(["scripts/verify.py", "--before", DIRTY, "--after", OUT])
    if not d:
        print(raw)
        return
    print(f"  结论: {d['summary']}")
    print(f"  已定型列: {d['diff']['typed_columns']}")
    for w in d["warnings"]:
        print(f"  提醒: {w}")
    print(f"  退出码 {code}（0 = 通过）")


def scene2() -> None:
    section(2, "自愈 —— 故意把表头行设成 1，校验器应当拦住")

    bad = os.path.join(DEMO, "_demo_bad.xlsx")
    run(["scripts/clean.py", "--input", DIRTY,
         "--plan-json", json.dumps({"sheet": "销售明细", "header_row": 1}),
         "--output", bad])

    code, d, raw = run(["scripts/verify.py", "--before", DIRTY, "--after", bad])
    if d:
        print(f"  pass = {d['pass']}，退出码 {code}（1 = 已拦住）")
        for e in d["errors"]:
            print(f"  ERROR: {e}")
        print("\n  → 模型据此把 header_row 改成 4 重试，即回到场景 1 的成功路径。")
    else:
        print(raw)
    if os.path.exists(bad):
        os.remove(bad)


def scene3() -> None:
    section(3, "降级 —— 三档策略各自触发")

    fake = os.path.join(DEMO, "_fakelib")
    os.makedirs(fake, exist_ok=True)
    with open(os.path.join(fake, "pandas.py"), "w", encoding="utf-8") as f:
        f.write('raise ImportError("simulated: pandas not installed")\n')

    enc = os.path.join(DEMO, "encrypted_demo.xlsx")
    if not os.path.exists(enc):
        with open(enc, "wb") as f:
            f.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 512)

    checks = [
        ("缺 pandas", ["scripts/probe.py", DIRTY], {"PYTHONPATH": fake}),
        ("文件加密", ["scripts/probe.py", enc], None),
        ("文件超限", ["scripts/probe.py", DIRTY, "--max-mb", "0.001"], None),
        ("文件不存在", ["scripts/probe.py", os.path.join(DEMO, "nope.xlsx")], None),
    ]
    for label, args, env in checks:
        _, d, raw = run(args, env)
        if d:
            print(f"  {label:8} → tier={d['tier']:8} {d['reason']}")
        else:
            print(f"  {label:8} → 解析失败: {raw[:120]}")

    for f in os.listdir(fake):
        os.remove(os.path.join(fake, f))
    os.rmdir(fake)


def scene4() -> None:
    section(4, "计划校验 —— 引用不存在的列，返回可修正的 PlanError")

    _, d, raw = run(["scripts/clean.py", "--input", DIRTY, "--plan-json",
                     json.dumps({"sheet": "销售明细", "header_row": 4,
                                 "to_numeric": ["利润额"]}),
                     "--output", os.path.join(DEMO, "_demo_x.xlsx")])
    if d:
        print(f"  ok = {d.get('ok')}")
        print(f"  错误类型 = {d.get('error_type')}")
        print(f"  错误信息 = {d.get('error')}")
        print("\n  → 错误信息自带可用列名，模型改一版计划即可，无需人工介入。")
    else:
        print(raw)


def main() -> int:
    if not os.path.exists(DIRTY):
        print("未找到演示数据，先生成……")
        run(["demo/make_dirty_data.py", DIRTY])

    scenes = {1: scene1, 2: scene2, 3: scene3, 4: scene4}
    pick = [int(a) for a in sys.argv[1:] if a.isdigit()] or sorted(scenes)
    for n in pick:
        if n in scenes:
            scenes[n]()
    print(f"\n{BAR}\n演示结束\n{BAR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
