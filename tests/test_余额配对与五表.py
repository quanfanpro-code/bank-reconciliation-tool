from decimal import Decimal
from unittest.mock import patch

import pandas as pd

from data_structures import MatcherConfig
from matcher import Matcher
from precision_engine import PrecisionEngine
from reporter import Reporter
from 直读XML import 读簿
from Excel验证 import 检查人工填写区


def _rows(entries):
    return pd.DataFrame([
        dict(date=pd.Timestamp(day), amount=Decimal(amount), amount_decimal=PrecisionEngine.to_integer_li(amount),
             balance=Decimal(balance), summary=summary, original_file_row=i + 2, original_idx=i,
             voucher_no=str(i + 1), aux_text_fields={"摘要": summary, **fields})
        for i, (day, amount, balance, summary, fields) in enumerate(entries)
    ])


def test_先确认逐笔再核对同段的合并记录(tmp_path):
    bank = _rows([
        ("2026-02-01", "-1980", "8020", "公积金", {}),
        ("2026-02-02", "-150", "7870", "转账", {"摘要": "报销", "对方户名": "张三"}),
        ("2026-02-02", "-147", "7723", "转账", {"摘要": "报销", "对方户名": "张三"}),
        ("2026-02-02", "-4000", "3723", "天然气款", {}),
    ])
    journal = _rows([
        ("2026-02-01", "-1980", "8020", "公积金", {}),
        ("2026-02-02", "-4000", "4020", "天然气款", {}),
        ("2026-02-03", "-297", "3723", "接待餐费", {}),
    ])
    matcher = Matcher(bank, journal, MatcherConfig(), balance_scope_verified=True,
                      verified_balance_pairs={(0, 0): "双方同日支付公积金", (3, 1): "双方同日支付天然气款"})
    matches = matcher.run()
    assert {tuple(item['bank_idxs']): tuple(item['journal_idxs']) for item in matches} == {(0,): (0,), (3,): (1,), (1, 2): (2,)}
    assert all(item['processing_status'] == '自动确认' for item in matches)
    output = tmp_path / '核对.xlsx'
    Reporter(matcher)._generate_report(str(output))
    workbook = 读簿(output)
    assert list(workbook.表) == ['核对总览', '月度核对', '日度核对', '待核对', '待核对明细', '交易明细']
    assert 检查人工填写区(output) == 1


def test_余额能对上但付款用途有明确冲突时不合并():
    bank = _rows([
        ("2026-02-01", "-120", "880", "报销", {"对方户名": "张三"}),
        ("2026-02-01", "-180", "700", "报销", {"对方户名": "李四"}),
    ])
    journal = _rows([("2026-02-01", "-300", "700", "住房公积金", {})])
    matcher = Matcher(bank, journal, MatcherConfig(), balance_scope_verified=True)
    assert not any(item['processing_status'] == '自动确认' for item in matcher.run())


def test_实际差额用自定阈值并能定位交易(tmp_path):
    bank = _rows([
        ("2026-02-01", "-100", "900", "小额付款", {}),
        ("2026-03-01", "-600", "300", "另一笔付款", {}),
    ])
    journal = _rows([("2026-02-01", "10", "1010", "收款", {})])
    matcher = Matcher(bank, journal, MatcherConfig(clearly_trivial_threshold=100, performance_materiality=500))
    matcher.run()
    output = tmp_path / '差额.xlsx'
    Reporter(matcher)._generate_report(str(output))
    assert 检查人工填写区(output) == 1
    book = 读簿(output)
    assert book['待核对'].值(2, 7) == 600
    assert 'MATCH(A2' in book['待核对'].值(2, 10)
    from Excel验证 import _验证五表结构
    assert _验证五表结构(output)['人工填写行数'] == 1


def test_启动Excel被拒绝时仍可验证五表报告(tmp_path):
    from Excel验证 import 验证Excel报告
    import pywintypes

    source = _rows([("2026-02-01", "-10", "990", "收付款", {})])
    matcher = Matcher(source, source, MatcherConfig())
    matcher.run()
    output = tmp_path / '报告.xlsx'
    Reporter(matcher)._generate_report(str(output))
    with patch('win32com.client.DispatchEx', side_effect=pywintypes.com_error(-2147024891, '拒绝访问', None, None)):
        result = 验证Excel报告(output)
    assert result['工作表数'] == 6
    assert result['验证方式'] == '工作簿结构检查'
