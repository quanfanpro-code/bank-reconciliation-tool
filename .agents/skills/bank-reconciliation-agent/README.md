# 银行流水核对 Skill

复制本文件夹到使用的 Skill 目录，然后调用 `bank-reconciliation-agent`。整个文件夹包含入口和核对引擎。另一台电脑从 GitHub 拿到这个文件夹即可使用。执行由那台电脑完成。

需要 Python 3.11 或更新版本。运行 Skill 的 Agent 会检查 Python 依赖并按本目录的 `requirements.txt` 补齐。Python 由运行环境提供。可用时程序会用本机 Excel 重算和校验。其他环境会检查工作簿结构、填写区和明细跳转。

## 核对流程

- Agent 通过问答收集每户的两份文件，或银行流水和日记账两个目录，以及日期容差、重要性设置和报告路径。
- Skill 自带的引擎读取原表，处理逐笔、跨日和合笔。当前 Agent 在需要时判断交易用途。日度、月度分别比较收入、支出和余额。
- 每户生成一份 Excel，包含核对总览、月度核对、日度核对、待核对、待核对明细和交易明细。点待核对表里的链接可跳到整组原始记录。

## 文件

- `SKILL.md` 为 Agent 运行说明。
- `scripts/运行核对.py` 为程序入口。
- `scripts/engine/` 带齐读表、核对、出表及验证代码。
- `requirements.txt` 列出 Python 依赖。
- `LICENSE` 为随附代码的许可。
