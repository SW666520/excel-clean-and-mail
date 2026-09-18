---
name: excel-clean-and-mail
description: |
  清洗并汇总本地 Excel/CSV 表格，并可将清洗结果通过网易邮箱发送给指定收件人。
  先体检表格结构再执行，全程本地运行、原文件只读、输出到新文件。
  用户说出「这个表好乱帮我整理一下」「表头不在第一行」「日期是文本没法排序」
  「金额带千分位算不了」「有重复行和空行」「按区域/渠道汇总」「做个透视表」
  「统计各部门合计」时使用。
  如果用户需要，把清洗结果发送到邮箱（如「把清洗结果发到某个邮箱」「清洗后发邮件给 xx」），
  发送仅支持网易邮箱（163 / 126 / yeah.net）。
  不触发：加密表格、云端在线表格（飞书/钉钉/Google Sheets）、
  Word/PPT/PDF 文档、需要写回原文件的场景、图表美化与格式排版。
  以及除网易邮箱邮件外的任何其他外部服务（如企业微信/钉钉/其他邮箱）。

metadata:
  dumate:
    displayName: "表格清洗官"
    summary: "体检并清洗本地 Excel/CSV 表格：修正表头、规范日期与金额格式、去重去空行，按需生成分组汇总或透视表，并将清洗结果通过网易邮箱发送给指定收件人。原文件只读，结果输出到新文件。"
    publisher: "DuMate"
    icon: "assets/icon.png"
    publishedAt: 1789660800000
    version: "1.0.1"
    level: L3
    category: 效率工具
    tags: ["Excel", "CSV", "数据清洗", "透视表", "表格汇总"]
    sensitive: false
    allow_implicit_invocation: true
---

# 表格清洗官

把结构混乱的业务表格整理成可直接分析的规范数据，并按需生成汇总表。

## 核心原则

1. **原文件只读**。所有结果写入新文件，绝不覆盖用户的原始数据。
2. **先体检后动手**。不看表结构就写清洗逻辑，等于蒙着眼睛做手术。
3. **模型只出计划，脚本负责执行**。不生成、不执行任意代码；
   能做什么由 `clean.py` 的算子白名单决定。
4. **结果必须校验**。不看 `verify.py` 的结论就不许说「已完成」。
5. **不虚构数字**。所有行数、列数、汇总值都来自脚本真实输出。

## 执行流程

### 第 1 步：体检

```bash
python scripts/probe.py <文件路径>
```

读它的输出，重点看三处：

- `tier` —— 决定走哪个档位（见下方「降级策略」）
- `structure.guessed_header_row` —— 真表头在第几行
- `structure.signals` —— 有哪些脏数据，每个信号的含义查
  `references/excel-pitfalls.md`

若 `tier` 为 `block`，把 `reason` 如实告诉用户并停止，不要尝试绕过。

### 第 2 步：出计划

根据体检结果写一份 JSON 清洗计划。字段速查、信号到字段的映射、
执行顺序陷阱，全部见 `references/plan-recipes.md`。

**只使用 `structure.header` 里真实存在的列名**，不要凭印象猜。

涉及以下操作前，必须先向用户说明并取得确认：

- `drop_duplicates`（可能删掉合法的同值记录）
- `filter_rows`（会缩小数据范围）
- 数据区合并单元格的向下填充（无法可靠区分「省略」与「真空」）

### 第 3 步：执行

```bash
python scripts/clean.py --input <原表> --plan <计划.json> --output <新文件>
```

返回 `PlanError` 说明计划本身有问题（例如引用了不存在的列），
错误信息里带着当前可用列名，改一版计划重试。

### 第 4 步：校验

```bash
python scripts/verify.py --before <原表> --after <新文件>
```

- `pass: true` → 把 `clean.py` 的操作日志和 `verify.py` 的 diff 摘要一起汇报给用户
- `pass: false` → 按 `errors` 修计划回到第 3 步，**最多重试 2 次**

两次仍不通过：停下来，告诉用户卡在哪一步、`verify.py` 报了什么错、
建议怎么处理。不要继续盲试，更不能把失败说成成功。

### 第 5 步：汇报

必须包含：

- 改了什么（直接引用 `clean.py` 的 `log`，逐条都是真实操作）
- 数据变化（行数、列数、空行数的前后对比）
- 遗留提醒（`verify.py` 的 `warnings`）
- 输出文件的完整路径

## 降级策略

`probe.py` 的 `tier` 字段直接对应处理档位：

| tier | 触发条件 | 行为 |
|---|---|---|
| `full` | pandas 可用、文件正常 | 完整清洗 + 透视汇总 |
| `degrade` | 缺 pandas | 退到 openpyxl 基础清洗（去空行、去空格、类型转换），**明确告知用户无法做透视汇总**，并给出 `pip install pandas` 的建议 |
| `partial` | 文件超过 50MB | 分块或抽样处理，**在报告里写清哪些范围没覆盖**，不谎称全量完成 |
| `block` | 文件加密 / 损坏 / 不存在 / 缺 openpyxl | 停止执行，说明原因和用户可采取的动作 |

降级时**必须主动告知用户当前处于降级档位**，而不是安静地少做一部分事。

## 边界

不做的事，遇到时直接说明并建议替代方案：

- 破解加密表格 —— 请用户先手动解密
- 读写云端在线表格 —— 请用户先导出为 xlsx
- 写回或覆盖原文件 —— 始终输出新文件
- 图表生成、单元格样式美化 —— 超出本技能范围
- 在数据不足时猜测业务含义 —— 应当询问用户

## 文件说明

```
SKILL.md                      本文件，只放路由与铁律
scripts/probe.py              环境与表结构体检，输出 tier + 脏数据信号
scripts/clean.py              计划驱动的清洗执行器，白名单算子，零 exec
scripts/verify.py             前后 diff 校验，自愈循环的裁判
references/excel-pitfalls.md  中文业务表格十个坑（按需加载）
references/plan-recipes.md    计划字段速查与失败修正表（按需加载）
demo/make_dirty_data.py       生成演示用脏表格
tests/cases.json              触发、输出、降级的测试用例
```

两个 `references` 文件**不要预先加载**，只在实际需要时才读——
体检发现具体脏数据信号时读 `excel-pitfalls.md`，
准备写计划时读 `plan-recipes.md`。
