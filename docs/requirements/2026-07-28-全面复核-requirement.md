# 银行流水核对工具 v3.0 全面复核 — 需求文档

日期：2026-07-28
工作流：requirement-workflow 六阶段（阶段 1 产出）

## 一、背景与问题

银行流水核对工具 v3.0 是一个桌面应用（customtkinter GUI + pandas + openpyxl），
自动核对银行对账单与企业日记账，输出 Excel 核对报告。代码约 6400 行、9 个业务模块，
readme 声称已经过多轮修复（46 个测试通过）。

用户要求：对该程序做代码层面和业务逻辑层面的全面复核，消灭所有致命 bug，
并在两个层面尽所能优化到最佳。

复核发现（实证）的现状问题：

### 致命（P0，已实证）

1. **默认参数无法启动**：GUI 默认值「组合窗口=31、最大深度=30」，
   而 `validate.py` 上限为「窗口≤30、深度≤20」，`start()` 首先调用校验，
   导致用户用出厂默认参数点击「开始核对」必然报错，程序事实上不可用。
2. **批量聚合匹配崩溃**：`matcher.match_batch_aggregation` 执行
   `group['summary'].fillna('')`，当 summary 列被 `_downcast_dtypes`
   转为 category（重复度高的数据，如代发工资，恰恰是该功能的目标场景）时
   抛 `TypeError: Cannot setitem on a Categorical`，整个核对流程中断。
3. **报告生成崩溃**：`BalanceRecalculator.extract_initial_balance` 对余额值直接
   `Decimal(str(v))`，遇到千分位格式（如 `"1,234.56"`）抛 `InvalidOperation`，
   `generate_report` 失败，用户拿不到报告。

### 严重（P1）

4. **空余额单元格被写成 Decimal('0.00')**：`standardize_data` 的 `std_balance`
   用 `clean_amount`（空值→0.00），把「没填余额」和「余额为 0」混为一谈。
   实证：第一天首笔余额为空时，期初余额被错误推算为 `-100.00`（应为「无法确定」），
   触发虚假的期初不一致警告，并污染每日统计与余额连续性判断。
5. **期初余额提取不使用用户列映射**：`extract_initial_balance` 内部用精确匹配猜列名
   （'日期'/'余额'/'摘要'），对「交易日期」「账户余额」「业务说明」等真实列名失效；
   Reporter 持有 mapping 却未传入。
6. **随机种子 -1 被校验拒绝**：readme 明确文档化「-1 = 每次随机」，
   `_validate_int(allow_zero=True)` 拒绝一切负数，文档功能不可达。
7. **GUI「允许异号」默认勾选**，与 readme 默认「不勾选」直接矛盾
   （readme 在 2026-06-10 明确收紧过该默认值）。
8. **「余额连续性异常」sheet 是死代码**：readme 承诺第六个 sheet，
   `Reporter.check_balance_continuity` 已定义但从未被调用，报告里没有该 sheet。
9. **汇总表数字全部 `astype(str)`**：readme 的 P1 修复记录声称已解决
   「数字转字符串导致 Excel 中无法计算」，代码第 420 行依然存在。
10. **报告顺序不可复现**：Reporter 用 `for mid in set(...)` 遍历匹配 ID，
    每次运行匹配明细行序不同，与 readme 强调的「可复现」冲突。

### 中等（P2）

11. `_match_total` 中 30% 分布容差是死代码（`_total_structure_matches` 已要求
    分桶完全相等）；实现比 readme 更严——方向是对的（更安全），readme 需同步。
12. `match_exact_1to1` 并行分支 `except Exception: pass` 静默吞异常，
    不像容差/组合阶段那样记入 `exception_logger`。
13. `_randomized_greedy` 先 `shuffle` 后 `sort`，洗牌被排序中和，
    「随机化」名存实亡，贪心兜底多样性不足。
14. `single_amount_with_direction` 模式下方向列缺失（API 层绕过 GUI 时）
    会静默按正数入账；应在数据层防御。
15. `utils.round_decimal` 的 `decimals` 参数在 Decimal 路径被忽略（硬编码 0.01）。
16. `_postprocess_summary` 警告涂色行范围错位 1 行（表头被涂、末行警告漏涂）。
17. readme 与实现多处不一致：匹配实为 7 步（批量聚合未在六级流程图中）、
    默认参数表、第六 sheet 名称等。
18. 死代码/死参数：`load_file(chunksize)` 从未分块、`is_non_empty_amount` 等
    死 import、`_greedy_fallback` 未被调用。

## 二、用户与使用场景

- 财务/会计人员（readme 作者 CPA-Q 的目标用户），非程序员，按月核对银行账与日记账。
- 数据为 Excel/CSV，几千至几万行，含合并标题行、汇总行、续行空日期等真实账务格式。
- 核心诉求：默认参数下能用、不崩溃、不误报、结果可复现、报告可直接用于审计底稿。

## 三、功能范围（本次复核）

- 消灭上述 P0/P1 全部问题；P2 在低风险前提下尽量处理。
- 每个修复配套回归测试（TDD：先红后绿）。
- 同步 readme 与真实实现（readme 是需求基线之一，两者必须一致）。
- 保持现有模块架构、匹配算法核心策略与公开 API 不变。

## 四、明确不做的内容

- 不重写架构、不引入新依赖、不改 GUI 框架。
- 不调整匹配算法的业务策略（六级/七级顺序、阈值默认值语义），只修缺陷。
- 不做性能大改（倒排索引/折半枚举已存在，仅修正确性）。
- 不执行任何 git 写操作。

## 五、约束与风险

- 宿主环境为 Windows + Git Bash，控制台 GBK（验证脚本需 `PYTHONIOENCODING=utf-8`）。
- 现有 46 个测试必须保持通过；新测试用 pytest，放入 `tests/`。
- 改动需最小化、风格与周边代码一致。
- 风险：修改 `std_balance` 语义（空→None）会影响 balance/reporter 的下游判断，
  需全链路回归。

## 六、验收标准

1. 默认参数下 `validate_config_params` 通过，程序可启动核对。
2. P0 三个崩溃场景各有回归测试且通过（category summary 聚合、千分位余额期初提取）。
3. 空余额单元格标准化后为 None；期初余额推断不被空单元格污染（有测试）。
4. 随机种子 -1 通过校验；GUI 默认值与 readme 一致。
5. 报告包含「余额连续性异常」类 sheet（或明确以「余额差异明细」替代并更新 readme）。
6. 汇总表数字列在 Excel 中为数值类型。
7. 同一数据+同参数运行两次，匹配明细行序一致。
8. `pytest tests/` 全绿；readme 更新记录如实登记本轮改动。

## 七、调研摘要与来源

- 需求基线：`readme.md`（v3.0，1065 行，含历次修复记录）。
- 代码证据：全部 12 个源文件 + 4 个测试文件已通读；
  关键 bug 均经 Python 实证（见上文「实证」标注）。
- 现有测试基线：46 passed（2026-07-28 复核开始时实测）。

## 八、已确认与待确认

已确认（用户指令与实证）：
- 复核范围=全模块；目标=消灭致命 bug + 代码/业务双层优化；必须走六阶段流程。

基于宿主规则记录的假设（auto 模式下按最合理决策推进，不再逐项追问）：
- A1：readme 与代码冲突时，「文档化的默认参数/功能」以 readme 为准
  （修代码迁就 readme）；「代码比 readme 更严格的安全行为」以代码为准（改 readme）。
- A2：「余额连续性异常」恢复为独立 sheet 生成（修代码迁就 readme），
  同时保留「余额差异明细」。
- A3：月/日核销保持当前的严格结构校验（不放宽到 readme 的 30%），readme 同步改述。
