# 清洗计划编写指南

> 本文件按需加载。只有在需要写一份清洗计划时才读它。

模型的职责是**输出一份计划**，不是写代码。计划是 JSON，由 `clean.py` 用固定算子执行。
这样做的好处：不需要 `exec` 任意代码，每个算子行为可预测、可单测，出错时错误信息定位明确。

查看完整字段说明：

```bash
python scripts/clean.py --print-schema
```

---

## 计划字段速查

| 字段 | 类型 | 默认 | 说明 |
|---|---|---|---|
| `sheet` | 字符串 | 第一个 sheet | 要处理的 sheet 名 |
| `header_row` | 整数(1-based) | 1 | 真表头所在行，取自 `probe.py` 的 `guessed_header_row` |
| `drop_columns` | 数组 | — | 显式删除的列 |
| `rename_columns` | 对象 | — | `{旧名: 新名}` |
| `strip_whitespace` | 布尔 | `true` | 清首尾空格 + 全角空格 |
| `to_numeric` | 数组 | — | 转数值，自动剥离 `¥ , % 元` |
| `to_datetime` | 数组 | — | 转日期，兼容中文格式 |
| `drop_blank_rows` | 布尔 | `true` | 删全空行 |
| `drop_duplicates` | 布尔 | `false` | 整行去重（**需用户确认**） |
| `drop_total_rows` | 数组 | — | 在这些列里查 `合计/小计/总计` |
| `filter_rows` | 数组 | — | `[{column, op, value}]` |
| `pivot` | 对象 | — | `{index, columns, values, agg}` |
| `sort_by` | 数组 | — | `[{column, ascending}]` |

`filter_rows` 的 `op` 白名单：`eq` `ne` `gt` `gte` `lt` `lte` `contains` `notna`

`pivot` 的 `agg` 白名单：`sum` `mean` `count` `max` `min` `median` `nunique`

不在白名单内的操作会返回 `PlanError`，而不是静默忽略。

---

## 标准计划模板

针对典型的中文销售明细表：

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

---

## 从 probe 信号映射到计划字段

读 `probe.py` 的输出，按下表填计划。**不要凭空猜列名**，只用 `structure.header` 里实际存在的名字。

| probe 信号 | 计划动作 |
|---|---|
| `guessed_header_row: 4` | `"header_row": 4` |
| `text_stored_dates > 0` | 把日期列加入 `to_datetime` |
| `text_stored_numbers > 0` | 把金额/数量列加入 `to_numeric` |
| `whitespace_polluted_cells > 0` | 保持 `strip_whitespace: true` |
| `duplicate_rows_in_sample > 0` | 询问用户后再设 `drop_duplicates: true` |
| `blank_rows_in_sample > 0` | 保持 `drop_blank_rows: true` |
| `unnamed_columns > 0` | 先核对 `header_row` 是否正确，而不是急着删列 |
| `merged_cells_in_sheet > 0` | 若在数据区，向用户确认是否需要向下填充 |
| `sheet_names` 有多个 | 名字不明确时询问用户，不要猜 |

---

## 计划执行顺序

`clean.py` 的执行顺序是固定的，写计划时需要知道，否则会踩到顺序陷阱：

```
读取(header_row) → 删全空列 → drop_columns → rename_columns
→ strip_whitespace → 空串归一 → 复查幽灵列
→ drop_total_rows → drop_blank_rows
→ to_numeric → to_datetime
→ filter_rows → drop_duplicates → sort_by
→ pivot → 写出
```

两个关键点：

1. **`drop_total_rows` 在类型转换之前**。合计行里的 `=SUM()` 缓存值是合法数字，
   若先转类型就无法靠「非数字」特征识别它了。
2. **`rename_columns` 很早执行**。后续所有字段（`to_numeric`、`pivot` 等）
   都必须使用**新列名**。

---

## 失败后如何修计划

`clean.py` 返回 `PlanError` 时，错误信息已包含当前可用列名，据此改一版即可。
`verify.py` 返回 `pass: false` 时，按错误类型对症处理：

| verify 错误 | 修正动作 |
|---|---|
| 「过半列没有有效列名」 | 调大 `header_row`（最常见的故障） |
| 「列名看起来是数据」 | `header_row` 指到数据行了，往上调 |
| 「清洗后数据为空」 | `filter_rows` 条件过强，或 `header_row` 错 |
| 「行数减少超过阈值」 | 检查是否误开 `drop_duplicates`，或过滤条件过宽 |
| 「仍以文本存储」 | 把该列补进 `to_numeric` / `to_datetime` |

**重试上限 2 次**。两次仍失败就停下来向用户说明卡在哪一步、
`verify.py` 报了什么，让用户决定——不要继续盲目试错，也不要假装成功。
