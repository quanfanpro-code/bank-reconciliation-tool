# 银行对账 Skill

在 Codex 中核对银行流水和公司银行日记账。每个账户生成一份 Excel，列出逐笔对应、收支差额和日末、月末余额，可追查到原工作表和行号。

## 安装和使用

1. 把 [bank-reconciliation-agent 文件夹](.agents/skills/bank-reconciliation-agent) 整个复制到 Codex 使用的 Skill 目录。核对所需文件都在这个文件夹里。
2. 在 Codex 调用 `$bank-reconciliation-agent`，提供同一账户的银行流水和日记账。核对多个账户时，可以分别提供银行流水目录和日记账目录。
3. 如有指定的日期容差、重要性金额或存放报告的位置，可以一并告诉 Agent。其余设置采用默认值。

运行需要 Python 3.11 或更新版本。Agent 会检查 [Python 依赖](.agents/skills/bank-reconciliation-agent/requirements.txt)。本机 Excel 可用时，Skill 会用它打开报告并重算。Agent 的运行指令见 [SKILL.md](.agents/skills/bank-reconciliation-agent/SKILL.md)。

## Skill 怎么核对

- Agent 查看两份原表的工作表、表头和内容，确认账户以及各列的含义。同一个文件里的多张交易明细页一起读取，每笔保留原始位置。
- 按设置的日期范围比较金额、摘要、对方、编号和余额，选定逐笔、跨日和合笔的对应。几笔都能凑出同一金额时，会继续比较哪一组更符合两份原表。
- 银行收入、支出和余额分别核对。明确期间内无收付的日期，余额沿用前一日。某户只找到一侧时，报告列出已读到的收支及覆盖范围，其他账户照常核对。

## 报告怎么看

每户的 Excel 包含 **核对总览、月度核对、日度核对、差异说明、差异明细和交易明细**。先看月度和日度的收支、余额及差额，再从「差异说明」点进相应原行。已查清的差额直接列在表里。

## 开源许可

本项目按 [GNU AGPL v3.0](LICENSE) 发布。
