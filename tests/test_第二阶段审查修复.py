"""独立审查回归：凭证续行、范围缺失、金额口径和批次证据边界。"""
import re
import hashlib
import json
import time
from decimal import Decimal
from pathlib import Path

import pandas as pd
import pytest

from application import run_reconciliation
from data_loader import DataLoader, ParseErrorCollector
from data_structures import MatcherConfig
from input_precheck import InputPrecheckBlockedError, TableStructure, build_input_precheck
from matcher import Matcher
from reporter import Reporter
import matcher as 匹配模块


def 映射(借方="借方金额", 贷方="贷方金额", **额外):
    return {"date": "日期", "voucher": "凭证号", "summary": "摘要",
            "debit": 借方, "credit": 贷方, "mode": "debit_credit",
            "auxiliary_text_columns": ["摘要", "对方户名"], **额外}


def 原始(金额, 来源="journal", 摘要="项目结算", 对方="甲公司", 借方="借方金额", 贷方="贷方金额"):
    return pd.DataFrame([
        {"日期": "2026-08-10", "凭证号": "", "摘要": 摘要, "对方户名": 对方,
         "本方账号": "A001", "币种": "CNY",
         借方: abs(额) if 来源 == "bank" else 0,
         贷方: abs(额) if 来源 == "journal" else 0}
        for 额 in 金额
    ])


def 输入检查(银行, 账, 银行映射=None, 账映射=None):
    银行映射 = 银行映射 or 映射()
    账映射 = 账映射 or 映射()
    加载器 = DataLoader()
    return build_input_precheck(
        raw_bank=银行, raw_journal=账,
        bank=加载器.standardize_data(银行, 银行映射, "bank"),
        journal=加载器.standardize_data(账, 账映射, "journal"),
        bank_mapping=银行映射, journal_mapping=账映射,
        bank_structure=TableStructure(0, 1, list(银行.columns), 10),
        journal_structure=TableStructure(0, 1, list(账.columns), 10),
    )


def 运行匹配(银行, 账):
    加载器 = DataLoader()
    实例 = Matcher(加载器.standardize_data(银行, 映射(), "bank"),
                  加载器.standardize_data(账, 映射(), "journal"), MatcherConfig())
    实例.run()
    return 实例


def test_日期凭证同时留空的34行续行继承首行并形成35行完整凭证():
    账 = 原始([100] * 35)
    账.loc[0, "凭证号"] = "记001"
    账.loc[1:, ["日期", "凭证号"]] = ""
    账["摘要"] = [f"结算明细{i}" for i in range(35)]
    标准账 = DataLoader().standardize_data(账, 映射(), "journal")

    assert len(标准账) == 35
    assert 标准账["voucher_no"].tolist() == ["记001"] * 35
    assert 标准账["original_file_row"].tolist() == list(range(2, 37))
    实例 = 运行匹配(原始([3500], "bank"), 账)
    完整 = [候选 for 候选 in 实例.selected_candidates
            if len(候选.bank_idxs) == 1 and len(候选.journal_idxs) == 35]
    assert len(完整) == 1
    assert 完整[0].processing_status.value == "整组勾稽一致"
    assert "凭证" in 完整[0].evidence.get("business_basis", "")


@pytest.mark.parametrize("来源", ["journal", "bank"])
def test_新日期无凭证不能沿用旧凭证且银行空凭证不能前填(来源):
    数据 = 原始([100, 200, 300], 来源)
    数据.loc[0, "凭证号"] = "记001"
    数据.loc[1:, "日期"] = "2026-08-11"
    标准 = DataLoader().standardize_data(数据, 映射(), 来源)
    assert len(标准) == 3
    assert all(pd.isna(值) or str(值).strip() == "" for 值 in 标准["voucher_no"].iloc[1:])


@pytest.mark.parametrize(("列名", "项目", "其他值"),
                         [("本方账号", "核对账户", "A002"), ("币种", "核对币种", "USD")])
@pytest.mark.parametrize("情况", ["部分空白", "部分空白且双方冲突", "一侧混合另一侧全空"])
def test_有效交易身份部分缺失必须披露且明确冲突仍优先阻断(列名, 项目, 其他值, 情况):
    银行 = 原始([100, 200, 300], "bank")
    账 = 原始([100, 200, 300])
    银行.loc[1, 列名] = ""
    if 情况 == "部分空白且双方冲突":
        账[列名] = 其他值
    elif 情况 == "一侧混合另一侧全空":
        银行.loc[2, 列名] = 其他值
        账[列名] = ""
    检查项 = next(项 for 项 in 输入检查(银行, 账).items if 项.name == 项目)

    if 情况 == "部分空白":
        assert 检查项.status == "疑点"
        assert re.search(r"(?:缺失|空白|未提供|缺少).{0,12}1\s*(?:行|笔)",
                         检查项.explanation + 检查项.bank_result)
    else:
        assert 检查项.status == "无法计算"


@pytest.mark.parametrize(
    ("银行前缀", "账前缀", "银行元数据", "账元数据", "预期"),
    [("原币", "本位币", None, None, "无法计算"),
     ("原币", "原币", None, None, "通过"),
     ("本位币", "本位币", None, None, "通过"),
     ("", "", None, None, "疑点"),
     ("", "", "原币", "本位币", "无法计算"),
     ("", "", "本位币", "本位币", "通过")],
)
def test_原币本位币口径独立检查且未注明不得视为已验证(银行前缀, 账前缀, 银行元数据, 账元数据, 预期):
    银行借, 银行贷 = 银行前缀 + "借方金额", 银行前缀 + "贷方金额"
    账借, 账贷 = 账前缀 + "借方金额", 账前缀 + "贷方金额"
    银行 = 原始([100, 200], "bank", 借方=银行借, 贷方=银行贷)
    账 = 原始([100, 200], 借方=账借, 贷方=账贷)
    银行映射, 账映射 = 映射(银行借, 银行贷), 映射(账借, 账贷)
    if 银行元数据:
        银行映射["amount_basis"] = 银行元数据
    if 账元数据:
        账映射["amount_basis"] = 账元数据
    检查 = 输入检查(银行, 账, 银行映射, 账映射)
    项目 = [项 for 项 in 检查.items if 项.name == "金额口径"]

    assert len(项目) == 1, "即使双方币种都是CNY，仍须独立披露所选金额列的原币/本位币口径"
    assert 项目[0].status == 预期
    assert 检查.has_blockers is (预期 == "无法计算")


@pytest.mark.parametrize(("摘要", "应有完整组"), [("货款", False), ("供应商批付", False)])
def test_同供应商同日普通货款不等于完整业务批次(摘要, 应有完整组):
    实例 = 运行匹配(原始([400, 600], "bank", 摘要), 原始([500, 500], 摘要=摘要))
    完整组 = [候选 for 候选 in 实例.candidates
              if len(候选.bank_idxs) == 2 and len(候选.journal_idxs) == 2
              and 候选.evidence.get("resolves_full_group")]
    assert bool(完整组) is 应有完整组


@pytest.mark.parametrize(("银行金额", "账金额", "方向"),
                         [([100, 200], [100, 100, 100], "少"),
                          ([100, 100, 100, 100], [100, 100, 100], "多")])
def test_工资笔数线索明确相对方向和差一笔但不确定原因(银行金额, 账金额, 方向):
    实例 = 运行匹配(原始(银行金额, "bank", "8月工资", ""), 原始(账金额, 摘要="8月工资汇总", 对方=""))
    完整 = [候选 for 候选 in 实例.selected_candidates
            if len(候选.bank_idxs) == len(银行金额) and len(候选.journal_idxs) == len(账金额)]
    assert len(完整) == 1
    线索 = 完整[0].evidence.get("batch_review_hint", "")
    assert re.search(rf"银行(?:流水)?.{{0,8}}(?:序时账|日记账).{{0,6}}{方向}\s*1\s*笔", 线索), 线索
    assert any(词 in 线索 for 词 in ("核查", "待查", "未确定", "不能确定"))


def test_大量同额精确单笔不重复触发深度30拆分搜索(monkeypatch):
    原求解器 = 匹配模块._solve_combination
    调用深度 = []

    def 记录深度(values, dates, indices, target, max_depth=30, allow_mixed_sign=False, date_window=31):
        调用深度.append(max_depth)
        return 原求解器(values, dates, indices, target, max_depth, allow_mixed_sign, date_window)

    monkeypatch.setattr(匹配模块, "_solve_combination", 记录深度)
    日期 = pd.Timestamp("2026-08-10")
    目标 = {"view_dict": [{"index": i, "date": 日期, "amount_decimal": 1000000, "matched": False}
                          for i in range(30)]}
    配置 = MatcherConfig(allow_greedy_fallback=False)
    for i in range(40):
        assert 匹配模块._process_single_source((i, 日期, 1000000, 目标, 配置)) is None
    assert not 调用深度 or max(调用深度) <= 4, (
        f"全是等额精确单笔时可直接不补搜；若补搜也不应回退到深度30：{调用深度}"
    )


@pytest.mark.parametrize("非完整别名", ["C-NY", "R_MB"])
def test_币种完整别名不能通过删除分隔符伪造成一致(非完整别名):
    银行, 账 = 原始([100], "bank"), 原始([100])
    银行["币种"] = 非完整别名
    项 = next(项 for 项 in 输入检查(银行, 账).items if 项.name == "核对币种")
    assert 项.status == "无法计算"


@pytest.mark.parametrize("显式选择", [None, "原币币种", "本位币币种"])
def test_多个币种列换序不能改变核对范围且显式映射有效(显式选择):
    银行, 账 = 原始([100], "bank"), 原始([100])
    银行 = 银行.drop(columns="币种").assign(原币币种="USD", 本位币币种="CNY")
    账["币种"] = "USD" if 显式选择 == "原币币种" else "CNY"
    状态 = []
    for 数据 in [银行, 银行[list(reversed(银行.columns))]]:
        银行映射 = 映射(**({"currency": 显式选择} if 显式选择 else {}))
        项 = next(项 for 项 in 输入检查(数据, 账, 银行映射).items if 项.name == "核对币种")
        状态.append(项.status)
    assert 状态 == (["无法计算", "无法计算"] if 显式选择 is None else ["通过", "通过"])


def 应用入口(tmp_path, 银行, 账, 银行映射=None, 账映射=None, 捕获=None):
    银行路径, 账路径 = tmp_path / "银行流水.xlsx", tmp_path / "银行存款序时账.xlsx"
    银行.to_excel(银行路径, index=False)
    账.to_excel(账路径, index=False)
    return run_reconciliation(
        str(银行路径), str(账路径), 银行映射 or 映射(), 账映射 or 映射(), MatcherConfig(),
        bank_skiprows=0, journal_skiprows=0, bank_header_rows=1, journal_header_rows=1,
        output_path=tmp_path / "正式核对结果.xlsx", matcher_ready=捕获,
    )


@pytest.mark.parametrize("冲突", ["账户", "金额口径"])
def test_硬闸门在匹配前保留独立输入问题报告而不生成正式结果(tmp_path, 冲突):
    银行, 账 = 原始([100], "bank"), 原始([100])
    银行映射, 账映射 = 映射(), 映射()
    if 冲突 == "账户":
        账["本方账号"] = "A002"
    else:
        银行映射["amount_basis"], 账映射["amount_basis"] = "原币", "本位币"
    匹配器 = []
    with pytest.raises(InputPrecheckBlockedError) as 错:
        应用入口(tmp_path, 银行, 账, 银行映射, 账映射, 匹配器.append)

    assert not 匹配器, "冲突应先于Matcher运行"
    assert not (tmp_path / "正式核对结果.xlsx").exists()
    问题路径 = getattr(错.value, "report_path", None)
    assert 问题路径, "输入冲突不能只留弹窗；异常应提供可复核的独立问题报告路径"
    表 = pd.read_excel(Path(问题路径), sheet_name=None)
    assert "输入检查" in 表
    assert any("运行" in 名 or "映射" in 名 for 名 in 表)
    assert any("原行" in 名 or "处置" in 名 or "冲突" in 名 for 名 in 表)


def test_同一源行多种解析错误只计一行且人口分类闭合():
    银行, 账 = 原始([100, 200], "bank"), 原始([100])
    银行["日期"] = 银行["日期"].astype(object)
    银行["借方金额"] = 银行["借方金额"].astype(object)
    银行.loc[1, ["日期", "借方金额"]] = ["错误日期", "金额不明"]
    收集器 = ParseErrorCollector()
    加载器 = DataLoader(error_collector=收集器)
    标准银行 = 加载器.standardize_data(银行, 映射(), "bank")
    标准账 = 加载器.standardize_data(账, 映射(), "journal")
    # 同一原文件第3行已被两项检查发现问题；事件条数不是数据行数。
    异常 = [{"type": "日期解析失败", "source_type": "bank", "row": 3,
             "original_file_row": 3, "original_value": "错误日期", "column": "日期"},
            {"type": "金额解析失败", "source_type": "bank", "row": 3,
             "original_file_row": 3, "original_value": "金额不明", "column": "借方金额"}]
    检查 = build_input_precheck(raw_bank=银行, raw_journal=账, bank=标准银行, journal=标准账,
        bank_mapping=映射(), journal_mapping=映射(),
        bank_structure=TableStructure(0, 1, list(银行.columns), 10),
        journal_structure=TableStructure(0, 1, list(账.columns), 10), parse_errors=异常)
    项 = next(项 for 项 in 检查.items if 项.name == "数据入口")
    assert "解析异常1行" in 项.bank_result, 项.bank_result
    for 文本 in ["原始2行", "有效交易1行", "非交易0行", "其他排除0行"]:
        assert 文本 in 项.bank_result


def test_入口报告保留人口金额去向及文件工作表表头映射证据(tmp_path):
    银行, 账 = 原始([100, 200], "bank"), 原始([100])
    银行["借方金额"] = 银行["借方金额"].astype(object)
    银行.loc[1, "借方金额"] = "金额不明"
    输出 = 应用入口(tmp_path, 银行, 账)
    表 = pd.read_excel(输出, sheet_name=None)
    全文 = " ".join(名 + " " + " ".join(数据.fillna("").astype(str).to_numpy().ravel()) for 名, 数据 in 表.items())
    指纹 = hashlib.sha256((tmp_path / "银行流水.xlsx").read_bytes()).hexdigest()
    assert 指纹 in 全文, "交付报告须绑定实际输入文件指纹"
    assert "Sheet1" in 全文 and "借方金额" in 全文
    assert "表头" in 全文 and "映射" in 全文
    assert "金额不明" in 全文
    assert any("处置" in 名 or "人口" in 名 or "去向" in 名 for 名 in 表), "原始行应有互斥分类及金额去向台账"


def 余额样本():
    银行, 账 = 原始([100, 200], "bank"), 原始([100, 200])
    for 表 in [银行, 账]:
        表["日期"] = ["2026-08-10", "2026-08-11"]
        表["摘要"] = ["甲公司设备款", "甲公司材料款"]
        表["期初余额"] = [1000, None]
    银行["余额"], 账["余额"] = [900, 650], [900, 700]
    return 银行, 账


def test_余额与期间收支总体控制在匹配前形成结构化输入检查():
    银行, 账 = 余额样本()
    检查 = 输入检查(银行, 账, 映射(balance="余额"), 映射(balance="余额"))
    余额项 = [项 for 项 in 检查.items if "余额" in 项.name]
    assert 余额项, "银行期初1000减支出300应为700，实际650，须在Matcher前形成结构化异常"
    assert any(项.status != "通过" and "50" in (项.bank_result + 项.comparison + 项.explanation) for 项 in 余额项)
    assert any("日期" in 项.name or "期间" in 项.name for 项 in 检查.items)
    assert any("金额" in 项.name or "收支" in 项.name for 项 in 检查.items)


def test_总体余额异常不允许依赖该总体的候选无保留自动确认(tmp_path):
    银行, 账 = 余额样本()
    捕获 = []
    输出 = 应用入口(tmp_path, 银行, 账, 映射(balance="余额"), 映射(balance="余额"), 捕获.append)
    assert 捕获[0].selected_candidates
    assert all(候选.processing_status.value != "自动确认" for 候选 in 捕获[0].selected_candidates)
    结论 = pd.read_excel(输出, sheet_name="核对结论").fillna("").astype(str)
    assert "核对已自动完成" not in " ".join(结论.to_numpy().ravel())


def test_换行后候选稳定键最终关系号及报告竞争引用不随位置漂移():
    银行, 账 = 原始([1000], "bank"), 原始([1000, 400, 600])
    两次 = [运行匹配(银行, 账), 运行匹配(银行, 账.iloc[[2, 0, 1]].reset_index(drop=True))]

    def 语义引用(实例):
        return {(tuple(sorted(实例.bank.at[i, "amount"] for i in c.bank_idxs)),
                 tuple(sorted(实例.journal.at[i, "amount"] for i in c.journal_idxs))):
                (c.candidate_id, c.final_match_id, tuple(sorted(c.evidence.get("alternative_candidate_ids", []))))
                for c in 实例.candidates if c.metrics.total_diff_li == 0 and not c.evidence.get("business_conflicts")}

    assert 语义引用(两次[0]) == 语义引用(两次[1]), "相同金额组成的候选键与关系引用不得编码输入下标"
    for 实例 in 两次:
        表 = Reporter(实例).build_report_tables(实例.config)
        引用 = " ".join(pd.concat([表["逐笔匹配"], 表["整组勾稽"]])["其他可能对应"].fillna("").astype(str))
        替代 = [编号 for c in 实例.selected_candidates for 编号 in c.evidence.get("alternative_candidate_ids", [])]
        assert 替代 and all(编号 in 引用 for 编号 in 替代), "竞争说明必须带稳定引用，原行号只作为定位辅助"


@pytest.mark.parametrize("用途", ["8月工资", "员工报销", "供应商批付", "客户批收"])
def test_同日同用途无批次边界的两套真实批次不能合成自动完整组(用途):
    # 构造真值为第一批400+600、第二批300+700；导出未提供边界证据。
    银行 = 原始([400, 600, 300, 700], "bank", 用途, "")
    账 = 原始([1000, 1000], 摘要=用途 + "汇总", 对方="")
    if 用途 == "客户批收":
        for 数据 in [银行, 账]:
            数据[["借方金额", "贷方金额"]] = 数据[["贷方金额", "借方金额"]].to_numpy()
    实例 = 运行匹配(银行, 账)
    assert not any(len(c.bank_idxs) == 4 and len(c.journal_idxs) == 2
                   and c.evidence.get("resolves_full_group")
                   and c.processing_status.value in {"自动确认", "整组勾稽一致"}
                   for c in 实例.selected_candidates), "总额2000相等不能证明两批业务的对应边界"
    assert any(c.processing_status.value == "疑点事项" or c.is_ambiguous for c in 实例.selected_candidates)


def test_报告组成和双方待查保留原金额方向口径及凭证继承证据(tmp_path):
    银行, 账 = 原始([100, 777], "bank"), 原始([100, 888])
    银行.loc[1, ["日期", "摘要", "对方户名"]] = ["2026-01-01", "银行独有历史款", "乙公司"]
    账.loc[1, ["日期", "摘要", "对方户名"]] = ["2026-12-31", "账簿独有年末款", "丙公司"]
    输出 = 应用入口(tmp_path, 银行, 账, 映射(amount_basis="本位币"), 映射(amount_basis="本位币"))
    表 = pd.read_excel(输出, sheet_name=None)
    必需 = {"原借方列名", "原借方值", "原贷方列名", "原贷方值", "原金额列名", "原金额值",
            "原方向列名", "原方向值", "采用金额口径", "凭证是否继承"}
    for 名 in ["匹配组成", "银行侧待查", "日记账侧待查"]:
        assert not 表[名].empty
        assert 必需 <= set(表[名].columns), f"{名}缺少原始金额及凭证证据：{必需-set(表[名].columns)}"
        assert set(表[名]["原借方列名"]) == {"借方金额"}
        assert set(表[名]["原贷方列名"]) == {"贷方金额"}
        assert set(表[名]["采用金额口径"]) == {"本位币"}


def test_大额精确单笔之后的小额等价补搜有独立小预算(monkeypatch):
    原求解器, 深度 = 匹配模块._solve_combination, []

    def 监测(values, dates, indices, target, max_depth=30, allow_mixed_sign=False, date_window=31):
        深度.append(max_depth)
        assert max_depth <= 8, f"精确单笔后的等价补搜深度{max_depth}超过独立小预算8"
        return 原求解器(values, dates, indices, target, max_depth, allow_mixed_sign, date_window)

    monkeypatch.setattr(匹配模块, "_solve_combination", 监测)
    日期 = pd.Timestamp("2026-08-10")
    目标 = {"view_dict": [{"index": i, "date": 日期, "amount_decimal": 29000000 if i == 0 else 1000000, "matched": False}
                          for i in range(30)]}
    匹配模块._process_single_source((0, 日期, 29000000, 目标, MatcherConfig(allow_greedy_fallback=False)))
    assert 深度


def test_千行重复金额批次在合理时间完成并披露搜索预算及耗尽数():
    开始 = time.perf_counter()
    实例 = 运行匹配(原始([100] * 1000, "bank", "8月工资", ""), 原始([100000], 摘要="8月工资汇总", 对方=""))
    耗时 = time.perf_counter() - 开始
    assert 耗时 < 30, f"1000行等额工资完整组不应退化为子集穷举，实际{耗时:.2f}秒"
    assert any(len(c.bank_idxs) == 1000 and len(c.journal_idxs) == 1 and c.metrics.total_diff_li == 0 for c in 实例.selected_candidates)
    参数 = json.dumps(实例.run_parameters, ensure_ascii=False)
    assert re.search(r"budget|预算", 参数, re.I), 参数
    assert re.search(r"exhaust|耗尽", 参数, re.I), 参数


def test_同供应商同日普通具体摘要仍不能证明多对多业务边界():
    实例 = 运行匹配(
        原始([400, 600], "bank", "8月材料货款", "甲公司"),
        原始([500, 500], 摘要="8月材料货款", 对方="甲公司"),
    )
    assert not any(
        len(候选.bank_idxs) == 2
        and len(候选.journal_idxs) == 2
        and 候选.processing_status.value in {"自动确认", "整组勾稽一致"}
        for 候选 in 实例.selected_candidates
    )


def test_无编号同日批次双方笔数相等也不能吞并两个真实批次():
    银行 = 原始([400, 600, 300, 700], "bank", "员工报销", "")
    账 = 原始([500, 500, 500, 500], 摘要="员工报销", 对方="")
    实例 = 运行匹配(银行, 账)
    assert not any(
        len(候选.bank_idxs) == 4
        and len(候选.journal_idxs) == 4
        and 候选.processing_status.value in {"自动确认", "整组勾稽一致"}
        for 候选 in 实例.selected_candidates
    )
    assert any(
        候选.processing_status.value == "疑点事项"
        or 候选.evidence.get("batch_boundary_uncertain")
        for 候选 in 实例.selected_candidates
    )


@pytest.mark.parametrize(("银行回单", "账回单", "预期冲突"), [("R001", "R999", True), ("R001", "R001", False)])
def test_回单号或交易流水号明确不同必须成为关键冲突(银行回单, 账回单, 预期冲突):
    银行, 账 = 原始([100], "bank", "设备款", "甲公司"), 原始([100], 摘要="设备款", 对方="甲公司")
    银行["回单号"], 账["回单号"] = 银行回单, 账回单
    列映射 = 映射()
    列映射["auxiliary_text_columns"] = ["摘要", "对方户名", "回单号"]
    加载器 = DataLoader()
    实例 = Matcher(
        加载器.standardize_data(银行, 列映射, "bank"),
        加载器.standardize_data(账, 列映射, "journal"),
        MatcherConfig(),
    )
    实例.run()
    assert 实例.selected_candidates
    候选 = 实例.selected_candidates[0]
    if 预期冲突:
        assert 候选.processing_status.value != "自动确认"
        assert "交易流水号" in 候选.processing_reason or 候选.evidence.get("business_conflicts")
    else:
        assert 候选.processing_status.value == "自动确认"


def test_超过一千条解析异常仍完整保留供数据入口唯一分类():
    收集器 = ParseErrorCollector()
    for 行号 in range(2, 1007):
        收集器.record_amount_error(行号, "金额不明", "bank", "借方金额")
    错误 = 收集器.get_all_errors()
    assert len(错误) == 1005
    assert len({(项["source_type"], 项["row"]) for 项 in 错误}) == 1005
