# scripts

依赖：Python 3.10+，`openpyxl`、`pandas`（`pip install openpyxl pandas`）。只读输入文件，不改原件。输出中文。

| 脚本 | 干什么 | 典型用法 |
|---|---|---|
| `dupcount.py` | 在 6 种口径（精确 / 去首尾空白 / 去所有空白 / 忽略大小写 / 忽略大小写＋去空白 / NFKC 全半角＋忽略大小写＋去空白＋去零宽）× 2 种策略（先到先得 `duplicated()` / 同键全拒 `keep=False`）下数重复；给 `--target` 就标出恰好等于平台数字的组合 | `python dupcount.py 失败文件.xlsx --sheet 字段信息 --header-rows 4 --keys A,B,C,D,F --target 2013` |
| `combos.py` | 列出若干文件在同一批列上的取值组合与行数，给出「只在 A 有 / 只在 B 有」的差集 | `python combos.py "成功.xlsx::字段信息::4" "失败.xlsx::字段信息::4" --cols B,C` |
| `probe.py` | 生成探针：候选写法做笛卡尔积，每种组合配一行**不同的**真实样本（主键不重复），写进模板副本，另出行号对照表；样本行不够就报错 | `python probe.py --template 空模板.xlsx --sheet 表信息 --header-rows 4 --sample 真实行.xlsx --vary vary.json --pk E --out 探针.xlsx --map-out 对照表.xlsx` |
| `xlsxdiff.py` | 对照两个 xlsx 的结构：sheet 名与顺序、dimension、表头逐格、合并单元格、数据验证、公式格数、每列非空数/最大长度/类型分布、首尾空白与零宽字符格数；只列不同处 | `python xlsxdiff.py 成功.xlsx 失败.xlsx --header-rows 4` |

`vary.json` 示例（列字母 → 候选值）：

```json
{"C": ["XX公司系统A", "系统A"], "D": ["INST1", "S1"]}
```

提醒：多数平台没有测试环境，探针里能过的行会真的写进生产——`--sample` 只放本来就要导的真实行。

测试：仓库根目录 `python tests/test_scripts.py`（不依赖 pytest，合成数据，7 项）。
