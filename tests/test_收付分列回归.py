from decimal import Decimal

import pandas as pd
import pytest

from data_loader import DataLoader
from gui import _auto_mapping_for_columns
from input_precheck import _mapping_problems
from application import _verified_balance_pairs
from data_structures import OverallControlResult


def test_收付分列保留付款日期金额与流水余额():
    source = pd.DataFrame({
        '交易日期': ['2026-01-12', '2026-01-14'],
        '收入金额': [None, 3200], '支出金额': [500, None],
        '交易后余额': [1500, 4700], '借贷标记': ['支出', '收入'],
    })
    mapping = _auto_mapping_for_columns(source.columns, is_bank=True)
    assert (mapping['mode'], mapping['debit'], mapping['credit']) == ('debit_credit','支出金额','收入金额')
    assert DataLoader().standardize_data(source, mapping, 'bank')['amount'].tolist() == [Decimal('-500.00'), Decimal('3200.00')]


def test_单金额加借贷标记仍使用方向列():
    mapping = _auto_mapping_for_columns(['日期','交易金额','借贷标记','余额'], is_bank=True)
    assert (mapping['mode'], mapping['amount'], mapping['direction']) == ('single_amount_with_direction','交易金额','借贷标记')


@pytest.mark.parametrize('debit,credit',[('借贷标记','收入金额'),('收入金额','支出金额')])
def test_预检查拦截方向文本或收付映反(debit,credit):
    mapping = {'mode':'debit_credit','date':'交易日期','debit':debit,'credit':credit}
    assert _mapping_problems(mapping, ['交易日期','收入金额','支出金额','借贷标记'], 'bank')


def test_预检查拦截只读取收款金额():
    mapping = {'mode':'signed_amount','date':'交易日期','amount':'收入金额'}
    assert _mapping_problems(mapping, ['交易日期','收入金额','支出金额','借贷标记'], 'bank')


def test_同名划款凭证金额和逐笔余额一致才可补强():
    raw_bank = pd.DataFrame({'对方户名': ['示例有限公司'], '对方账号': ['00000000020002'],
                             '对方开户行': ['乙银行']})
    raw_journal = pd.DataFrame({'辅助核算名称': ['甲银行00000000010001']})
    row = {'date': pd.Timestamp('2026-09-03'), 'amount': Decimal('3200'), 'balance': Decimal('4500')}
    bank = pd.DataFrame([{**row,'original_file_row':2,'summary':'大额来账'}])
    journal = pd.DataFrame([{**row,'original_file_row':6,'summary':'转账','voucher_no':'9'}])
    control = OverallControlResult(bank_balance_status='通过',journal_balance_status='通过',
        period_status='通过',amount_status='通过',initial_balance_diff=Decimal(0),ending_balance_diff=Decimal(0))
    paths = ('示例有限公司-甲银行-00000000010001.xlsx','示例有限公司-甲银行-00000000010001-日记账.xlsx')
    args = (*paths,raw_bank,raw_journal,bank,journal,control)
    evidence = _verified_balance_pairs(*args)
    assert '乙银行尾号0002' in evidence[(0,0)] and '第9号凭证' in evidence[(0,0)]
    bank.loc[0,'balance'] = Decimal('4499')
    assert _verified_balance_pairs(*args) == {}
    bank.loc[0,'balance'] = Decimal('4500')
    raw_bank.loc[0,'对方户名'] = '其他有限公司'
    assert '乙银行尾号0002' not in _verified_balance_pairs(*args)[(0,0)]
    raw_bank.loc[0,'对方户名'] = '示例有限公司'
    assert _verified_balance_pairs(*args[:5],pd.concat([journal,journal],ignore_index=True),args[-1]) == {}
