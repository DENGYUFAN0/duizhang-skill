# 对账 · duizhang-skill

> 平台说了什么不重要，平台做了什么才重要。 · *Debug bulk imports into opaque enterprise platforms by reconciling the platform's own numbers.*

一个 agent skill（适用于 Claude Code 及任何支持 SKILL.md 的 agent），专治「把 Excel 按模板填好、导进一个规则不透明的企业系统，结果成批失败」。典型场景是央企/国企的数据治理平台、资源目录、主数据、报送系统——往往由多家供应商先后建设，模板说明和真实校验对不上，报错也不说人话。

它的核心方法只有一条：**以平台自己的导入记录、错误报告和曾经成功过的文件为唯一真相，用数字对账反推隐藏规则**。每条假说都必须精确复现平台报过的数字（成功数、失败数、错误条数），对上了才算找到规则。

## 为什么需要它

这个 skill 来自一次真实战役：AI 照着模板字面规则交了 5 版，平台 0 行成功。最后从导入记录和同事一份「部分成功」的文件里做对账：

| 平台数字 | 对账 |
|---|---|
| 同事版失败 7,800 行 | 表信息 5,787 行（填了 Schema，全败）＋ 字段信息按**忽略大小写**判出的重复 2,013 行 |
| 候选版错误 129,455 条，失败 126,648 行 | 多出的 2,807 ＝ 忽略大小写后的重复 123 ＋ 2,684 |
| 历史成功 5,788 与 117,277 | 按实例名、忽略大小写去重后的表数与字段数 |

三组独立数字精确吻合，才确定：模板表头写的「数据库名称(schema)」其实要填**实例名**，平台判重**忽略大小写**。完整复盘见 [case-study.md](skill/duizhang/references/case-study.md)。

## 核心方法

| 步骤 | 做什么 |
|---|---|
| 0 收证 | 导入记录全表、错误报告全量导出、任何曾经成功过的文件（自己的、同事的、历史的） |
| 1 比对 | 成功与失败文件在「查找列」上的取值差集；结构差异 |
| 2 对账 | 把平台数字列成等式；多口径数重复，找恰好等于平台数字的口径 |
| 3 探针 | 未证实的规则用几十行真实数据做多变体探针，单独交付 |
| 4 生成 | 只按实测规则改；按平台口径预去重、自己决定留哪行；每处改动记账 |
| 5 自检 | 回放平台规则而不是模板规则；写明范围与未覆盖 |
| 6 回收 | 上传后拿成功数对账，对不上回第 2 步 |

另附三份参考：[外包拼凑型系统烂点目录](skill/duizhang/references/shishan-patterns.md)（25 条，每条带识别方法与应对）、[国企数据工作协作与汇报对策](skill/duizhang/references/soe-playbook.md)、[AI 在这类任务里的典型失误](skill/duizhang/references/ai-pitfalls.md)。

## 安装

### Claude Code（macOS / Linux）

```bash
git clone https://github.com/DENGYUFAN0/duizhang-skill.git
cp -r duizhang-skill/skill/duizhang ~/.claude/skills/
pip install openpyxl pandas
```

### Windows（PowerShell）

```powershell
git clone https://github.com/DENGYUFAN0/duizhang-skill.git
New-Item -ItemType Directory -Force "$env:USERPROFILE\.claude\skills" | Out-Null
Copy-Item -Recurse duizhang-skill\skill\duizhang "$env:USERPROFILE\.claude\skills\"
pip install openpyxl pandas
```

### 其他 agent

把 `skill/duizhang/SKILL.md` 放进 system prompt 或 AGENTS.md，`references/` 按需读取；`scripts/` 是独立的 Python 脚本，任何环境都能跑。

## 使用示例

```
/duizhang 表元数据导入模板传了三次，12 万行全失败，报「未查找到数据库配置信息」，截图和文件在这
导入对账：同事那版成功了 11 万多行，我这版 0 行，帮我看差在哪
/duizhang 要把三份清单整合后按平台模板导入，先告诉我要准备什么
```

## 适用 / 不适用

- **适用**：按某系统的导入模板填数据并上传；导入成批失败、报错看不懂；多版本、多人并行、没人说得清规则。
- **不适用**：写系统本身的代码；绕过权限或校验；伪造数据凑成功率；普通的财务对账。

## 已知局限

- 方法依赖平台给出**可对账的数字**（成功数、失败数、错误条数）。平台什么都不报时，只能靠探针。
- 案例只有一次，烂点目录里有些条目是「强推断」或「推断」，文中已逐条标注。
- 多数平台没有测试环境，探针会写进生产。本 skill 要求只用本来就要导的真实数据做探针，但风险要由使用者自己评估。

## 免责声明

案例已脱敏：单位、系统、人员、库名均已泛化，不针对任何具体机构或供应商，也不构成对任何具体产品的评价。文中关于「国企工作方式」的内容是一次项目的经验观察和少量有出处的公开资料，不是统计结论。

## English

**duizhang** ("reconciliation") is an agent skill for importing spreadsheet data into opaque enterprise platforms whose documented template rules don't match their real validation. Instead of trusting the template, it treats the platform's own import log, error report and any previously successful file as ground truth, and reverse-engineers hidden rules by reconciling counts: a hypothesis only stands if it reproduces the platform's numbers exactly. Includes a catalogue of common platform flaws, a playbook for working inside large state-owned organizations, and four Python helpers (multi-normalization duplicate counter, value-combination diff, probe-file generator, structural xlsx diff). Content is in Chinese.

## License

[MIT](LICENSE) © 2026 DENG YUFAN
