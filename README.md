
# 课程表转换工具

把 WakeUp 课程表模板整理成**可批量导入的 CSV** 的单文件脚本。
已把原来的 `test.py`（课表解析）和 `transformer.py`（xlsx → csv）合并为 `converter.py`。

---

## 1. 文件说明

| 文件 | 说明 |
| --- | --- |
| `converter.py` | **主脚本（推荐）**，合并版，包含课表解析 + xlsx→csv 全部功能 |
| `test.py` / `transformer.py` | 旧版两个脚本，功能已被 `converter.py` 完全覆盖，确认无误后可删除 |

---

## 2. 环境准备

需要 Python 3.8 及以上，并安装 openpyxl：

```bash
python --version
pip install openpyxl
```

---

## 3. 快速开始

输入既可以是**相对文件名**，也可以是**绝对路径**；路径含空格时用引号包起来。
省略输入时脚本会交互式询问。

```bash
# 相对文件名
python converter.py test.xlsx

# 绝对路径
python converter.py "D:\课表\test.xlsx"

# 交互式输入
python converter.py
```

默认输出（放在输入文件旁边）：

- `输入名_转换结果.csv` —— 最终结果
- `输入名_转换结果.xlsx` —— 规范化后的表格（加 `--no-xlsx` 可关闭）

例如 `python converter.py test.xlsx` 会生成 `test_转换结果.csv` 和 `test_转换结果.xlsx`。

---

## 4. 两种模式

| 模式 | 说明 | 对应旧脚本 |
| --- | --- | --- |
| `parse`（默认） | 把课程单元格里的多行文本拆成规范的 7 列，并转换地点格式 | `test.py` |
| `raw` | 原样导出整张表，自动用左上角的值填充合并单元格 | `transformer.py` |

```bash
# 解析模式（默认）
python converter.py test.xlsx

# 原样导出模式
python converter.py 课表.xlsx --mode raw
```

---

## 5. 全部参数

| 参数 | 说明 |
| --- | --- |
| `input` | 输入文件名或绝对路径，省略则交互式询问 |
| `-m, --mode {parse,raw}` | 转换模式，默认 `parse` |
| `-s, --sheet N` | 工作表索引，0 表示第一个表，默认 `0` |
| `-o, --output 路径` | 指定输出的 csv 路径，默认 `输入名_转换结果.csv` |
| `--xlsx-output 路径` | parse 模式下 xlsx 的输出路径，默认 `输入名_转换结果.xlsx` |
| `--no-xlsx` | parse 模式下只生成 csv，不生成 xlsx |
| `--header-rows N` | 表头占用行数，第 N+1 行对应第 1 节，默认 `2` |
| `--keep-first-col` | 不跳过 A 列（默认跳过 A 列的文字标签） |
| `--list-sheets` | 只列出文件里所有工作表名称后退出 |
| `-h, --help` | 查看帮助 |

常用示例：

```bash
# 查看所有工作表，确认要处理哪一个
python converter.py test.xlsx --list-sheets

# 处理第 2 个工作表，并指定输出文件名
python converter.py test.xlsx -s 1 -o 我的课表.csv

# 表头占 3 行的表格
python converter.py test.xlsx --header-rows 3
```

---

## 6. 输出列含义

`parse` 模式输出的 CSV 固定为 7 列：

| 列 | 含义 | 来源 |
| --- | --- | --- |
| 课程名称 | 课程名 | 单元格第 1 行 |
| 星期 | 1=周一 …… 7=周日 | 所在列（B 列 = 1） |
| 开始节数 | 第 1 节 = 1 | 所在行 |
| 结束节数 | 连堂时大于开始节数 | 合并单元格跨越的行数 |
| 老师 | 授课教师 | 单元格第 2 行 |
| 地点 | 如 `4号楼智慧教室1` | 单元格里的“地点：xxx”行 |
| 周数 | 如 `4-6、9-14` | 单元格第 3 行 |

地点转换规则：`智慧教室一4-355` 会变成 `4号楼智慧教室1`，
即 教室名+序号+楼号-房间号 变成 楼号+“号楼”+教室名+序号。

CSV 使用 `utf-8-sig` 编码写出，用 Excel / WPS 直接打开不会中文乱码。

---

## 7. 使用前必读（模板要求）

- **必须删除所有括号**，并把“，”替换为“、”
- **必须删除所有空白行、列**
- 本工具以 **B3 单元格作为周一的第一节课**，请注意对齐
- 本工具按照 **WakeUp 课程表**提供的模板制作，其他软件生成的表格可能无法直接使用

---

## 8. 常见问题

| 问题 | 解决办法 |
| --- | --- |
| `[错误] 找不到文件` | 检查路径拼写；路径含空格要用引号包起来 |
| `[错误] 不支持旧版 .xls 格式` | 先用 Excel / WPS 把 .xls 另存为 .xlsx |
| `[错误] 无法写入 xxx，该文件可能正被 Excel / WPS 打开` | 关闭正在打开该 csv/xlsx 的程序后重试 |
| 解析到的记录数为 0 | 用 `--list-sheets` 确认工作表；用 `--header-rows` 调整表头行数 |
| 星期/节数整体错位 | 检查表格左上角是否与模板一致（B3 = 周一第 1 节） |
| 中文乱码 | 本工具已用 `utf-8-sig` 输出，请勿用记事本另存为 ANSI |

---

## 9. 从旧脚本迁移

| 旧用法 | 新用法 |
| --- | --- |
| `python test.py` | `python converter.py test.xlsx` |
| `python transformer.py` | `python converter.py 课表.xlsx --mode raw` |
