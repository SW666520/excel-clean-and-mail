# excel-clean-and-mail

把结构混乱的业务表格整理成可直接分析的规范数据，并按需通过邮件发送清洗结果。

## 功能概述

本工具专为中文业务表格设计，解决日常工作中最常见的脏数据问题：

- **表头错位**：前几行是大标题/制表人，真表头埋在第 N 行
- **日期格式混乱**：同列混用 `2026/2/1`、`2026-02-22`、`2026年1月1日`
- **金额带货币符号**：`¥220,800`、`1,280元` 等文本型金额
- **空格污染**：首尾空格、全角空格导致同一值被分成多组
- **重复/空行**：幽灵行干扰统计
- **文本型数字**：排序变字典序、求和报错

处理流程遵循"先体检后动手"原则：不盲目清洗，每一步都有校验。

## 目录结构

```
excel-clean-and-mail/
├── README.md                  # 本文件
├── SKILL.md                   # 技能路由与执行铁律
├── _meta.json                 # 技能元数据
├── config/
│   ├── smtp_config.json       # 邮件配置（需手动填写）
│   └── smtp_config.example.json # 配置模板
├── scripts/
│   ├── probe.py               # 环境与表结构体检
│   ├── clean.py               # 计划驱动的清洗执行器
│   ├── verify.py              # 前后 diff 校验
│   └── send_mail.py           # SMTP 邮件发送
├── references/
│   ├── plan-recipes.md        # 清洗计划字段速查
│   └── excel-pitfalls.md      # 中文业务表格十个坑
├── demo/
│   ├── run_demo.py            # 演示入口
│   ├── make_dirty_data.py     # 生成演示用脏表格
│   └── plan_sales.json        # 演示清洗计划
└── tests/
    └── cases.json             # 触发、输出、降级的测试用例
```

## 快速开始

### 1. 准备邮件配置（可选）

若需邮件发送功能，复制配置模板并填写：

```bash
cp config/smtp_config.example.json config/smtp_config.json
```

编辑 `config/smtp_config.json`：

```json
{
  "smtp_host": "smtp.163.com",
  "smtp_port": 465,
  "use_ssl": true,
  "sender_email": "你的邮箱@163.com",
  "auth_code": "你的16位SMTP授权码",
  "default_to": "默认收件人@example.com"
}
```

> 只清洗不发送邮件时，可跳过此步骤。

### 2. 体检表格

```bash
python scripts/probe.py <文件路径>
```

输出示例：

```json
{
  "tier": "full",
  "structure": {
    "guessed_header_row": 4,
    "header": ["订单日期", "区域", "渠道", "产品型号", "销售数量", "销售额", "备注"],
    "signals": {
      "text_stored_dates": 35,
      "text_stored_numbers": 25,
      "whitespace_polluted_cells": 9,
      "blank_rows_in_sample": 1,
      "duplicate_rows_in_sample": 1
    }
  }
}
```

### 3. 编写清洗计划

根据体检结果写一份 JSON 清洗计划（见 `references/plan-recipes.md` 字段速查）：

```json
{
  "sheet": "销售明细",
  "header_row": 4,
  "drop_columns": ["", " "],
  "strip_whitespace": true,
  "to_datetime": ["订单日期"],
  "to_numeric": ["销售数量", "销售额"],
  "drop_blank_rows": true,
  "drop_duplicates": true,
  "drop_total_rows": ["订单日期"]
}
```

### 4. 执行清洗

```bash
python scripts/clean.py --input <原表.xlsx> --plan <计划.json> --output <新文件.xlsx>
```

### 5. 校验结果

```bash
python scripts/verify.py --before <原表.xlsx> --after <新文件.xlsx>
```

### 6. 发送邮件（可选）

```bash
python scripts/send_mail.py \
  --to <收件人> \
  --subject "清洗结果" \
  --body "清洗完成，见附件。" \
  --attach <新文件.xlsx>
```

## 核心脚本说明

| 脚本 | 职责 | 关键输出 |
|---|---|---|
| `probe.py` | 环境与表结构体检 | `tier`（处理档位）、`guessed_header_row`（真表头行号）、`signals`（脏数据信号） |
| `clean.py` | 计划驱动的清洗执行 | 新 xlsx 文件、操作日志（去重/去空行/类型转换的逐条记录） |
| `verify.py` | 前后 diff 校验 | `pass`（是否通过）、`diff`（行列数变化、类型分布）、`warnings` |
| `send_mail.py` | SMTP 邮件发送 | 邮件投递状态，仅支持网易邮箱（163/126/yeah.net） |

### probe.py 输出解读

- `tier: full` —— pandas + openpyxl 齐全，完整清洗 + 透视汇总
- `tier: degrade` —— 缺 pandas，退到 openpyxl 基础清洗（去空行/去空格/类型转换），**不做透视汇总**
- `tier: partial` —— 文件超过 50MB，分块处理
- `tier: block` —— 文件加密/损坏/不存在，停止执行

### clean.py 算子白名单

只支持以下操作，任何不在白名单内的配置都会返回 `PlanError`：

| 算子 | 说明 | 示例 |
|---|---|---|
| `header_row` | 真表头所在行（1-based） | `"header_row": 4` |
| `sheet` | 指定工作表 | `"sheet": "销售明细"` |
| `drop_columns` | 删除列（按列名） | `["", " "]` |
| `rename_columns` | 重命名列 | `{"旧名": "新名"}` |
| `strip_whitespace` | 清首尾空格 + 全角空格 | `true`（默认开启） |
| `to_datetime` | 转日期，兼容中文格式 | `["订单日期"]` |
| `to_numeric` | 转数值，自动剥离 `¥,元%` | `["销售数量", "销售额"]` |
| `drop_blank_rows` | 删全空行 | `true` |
| `drop_duplicates` | 整行去重 | `false`（需用户确认） |
| `drop_total_rows` | 删合计/小计/总计行 | `["订单日期"]` |
| `filter_rows` | 按条件过滤行 | `[{"column":"区域", "op":"eq", "value":"华南"}]` |
| `pivot` | 透视汇总（需 pandas） | `{"index":["区域"], "columns":["渠道"], "values":["销售额"], "agg":"sum"}` |
| `sort_by` | 排序 | `[{"column":"订单日期", "ascending":false}]` |

`filter_rows` 的 `op` 白名单：`eq` `ne` `gt` `gte` `lt` `lte` `contains` `notna`

`pivot` 的 `agg` 白名单：`sum` `mean` `count` `max` `min` `median` `nunique`

## 清洗计划编写指南

计划的核心原则：**只使用 `probe.py` 输出的 `structure.header` 里真实存在的列名**，不要凭印象猜。

标准模板（典型中文销售明细表）：

```json
{
  "sheet": "销售明细",
  "header_row": 4,
  "strip_whitespace": true,
  "to_datetime": ["订单日期"],
  "to_numeric": ["销售数量", "销售额"],
  "drop_total_rows": ["订单日期"],
  "drop_blank_rows": true,
  "drop_duplicates": true,
  "pivot": {
    "index": ["区域"],
    "columns": ["渠道"],
    "values": ["销售额"],
    "agg": "sum"
  }
}
```

详细字段说明、执行顺序陷阱和失败修正表，见 `references/plan-recipes.md`。

## 降级策略

`probe.py` 的 `tier` 字段直接决定处理档位，**降级时必须主动告知用户**：

| tier | 触发条件 | 行为 |
|---|---|---|
| `full` | pandas 可用、文件正常 | 完整清洗 + 透视汇总 |
| `degrade` | 缺 pandas | 退到 openpyxl 基础清洗，明确告知无法做透视汇总，给出 `pip install pandas` 建议 |
| `partial` | 文件超过 50MB | 分块或抽样处理，报告里写清哪些范围没覆盖 |
| `block` | 文件加密/损坏/不存在/缺 openpyxl | 停止执行，说明原因和用户可采取的动作 |

## 边界

本工具**不做**以下事情：

- 破解加密表格 —— 请用户先手动解密
- 读写云端在线表格 —— 请用户先导出为 xlsx
- 写回或覆盖原文件 —— 始终输出新文件
- 图表生成、单元格样式美化 —— 超出本技能范围
- 在数据不足时猜测业务含义 —— 应当询问用户

## 测试

测试用例位于 `tests/cases.json`，覆盖触发、输出、降级场景。运行演示：

```bash
cd demo
python run_demo.py
```

## 依赖

- **必装**：`openpyxl`
- **可选**：`pandas`、`numpy`（透视汇总与高级清洗需要）

安装：

```bash
pip install openpyxl pandas numpy
```

## 核心原则

1. **原文件只读**。所有结果写入新文件，绝不覆盖用户的原始数据。
2. **先体检后动手**。不看表结构就写清洗逻辑，等于蒙着眼睛做手术。
3. **模型只出计划，脚本负责执行**。不生成、不执行任意代码；能做什么由 `clean.py` 的算子白名单决定。
4. **结果必须校验**。不看 `verify.py` 的结论就不许说「已完成」。
5. **不虚构数字**。所有行数、列数、汇总值都来自脚本真实输出。

## 贡献与反馈

本技能仍在持续优化。如遇到 `verify.py` 两次仍不通过、或发现新的脏数据模式不在当前信号列表中，请保留原始样本并反馈。