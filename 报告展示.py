"""只组织已有核对结果的阅读展示；不参与匹配、风险分流或结果计算。"""
from collections import defaultdict
from decimal import Decimal

import pandas as pd


阅读表 = ('核对结论', '非工作日交易', '月度核对', '逐笔核对', '整组核对', '人工全查', '人工抽样', '未对应记录')
组成金额列 = ('银行收入', '银行支出', '序时账收入', '序时账支出')
组展示列 = ['核对编号', '行别', '银行记录', '银行收入', '银行支出', '序时账记录', '序时账收入', '序时账支出', '核对差额', '人工核对结果', '当前核对结论', '当前记录结果', '核对原因', '备注', '原始出处', '事项编号', '原始记录号', '填报编号', '银行金额', '序时账金额', '差额']


def _文字(record):
    return '\n'.join(str(x) for x in (pd.Timestamp(record['日期']).strftime('%Y-%m-%d'), record['摘要'], record['对方及凭证']) if x)


def _收支(records):
    amounts = [Decimal(str(record['金额'])) for record in records]
    return sum((max(x, Decimal(0)) for x in amounts), Decimal(0)), sum((max(-x, Decimal(0)) for x in amounts), Decimal(0))


def _汇总(records):
    if not records:
        return '未找到对应记录'
    income, expense = _收支(records)
    dates = sorted({pd.Timestamp(record['日期']).strftime('%Y-%m-%d') for record in records})
    period = dates[0] if len(dates) == 1 else dates[0] + '至' + dates[-1]
    summaries = list(dict.fromkeys(str(record['摘要']) for record in records if record['摘要']))
    description = '、'.join(summaries[:2]) + (f'等{len(summaries)}种摘要' if len(summaries) > 2 else '')
    return f'{period}\n{description}\n{len(records)}笔\n收入 {income:,.2f}\n支出 {expense:,.2f}'


def _金额(row, side, record):
    amount = float(record['金额'])
    row[side + '收入'] = max(amount, 0)
    row[side + '支出'] = max(-amount, 0)


def _组成(item, records):
    rows = []
    for record in records:
        side = '银行' if record['来源'] == '银行流水' else '序时账'
        row = {'核对编号': item['核对编号'], '事项编号': item['事项编号'], '原始记录号': record['原始记录号'], '行别': side + '组成', side + '记录': _文字(record), side + '金额': record['金额'], '原始出处': record['原始出处']}
        _金额(row, side, record)
        rows.append(row)
    return rows


def _主行(item, records, original=None):
    banks = [r for r in records if r['来源'] == '银行流水']
    journals = [r for r in records if r['来源'] == '序时账']
    grouped = len(banks) > 1 or len(journals) > 1
    bi, be = _收支(banks)
    ji, je = _收支(journals)
    row = dict(original or {})
    row.update({'核对编号': item['核对编号'], '事项编号': item['事项编号'], '行别': '核对事项', '核对原因': row.get('核对原因') or item['判断依据'], '核对差额': f'收入差 {bi-ji:,.2f}\n支出差 {be-je:,.2f}', '原始出处': '' if grouped else '；'.join(r['原始出处'] for r in records)})
    for side, source in (('银行', banks), ('序时账', journals)):
        row[side + '记录'] = _汇总(source) if grouped or not source else _文字(source[0])
        if not grouped and source:
            _金额(row, side, source[0])
    return row, grouped


def build_display_tables(tables):
    """从既有逐条原记录产生阅读视图；技术表和原计算数据继续保留。"""
    by_item = defaultdict(list)
    for record in tables['核对明细'].to_dict('records'):
        by_item[record['事项编号']].append(record)
    index = tables['复核事项索引'].to_dict('records')
    originals = {row['事项编号']: row for name in ('人工全查', '人工抽样') for row in tables[name].to_dict('records') if row['行别'] == '核对事项'}
    alternatives = defaultdict(list)
    for record in tables['其他对应供选择'].to_dict('records'):
        alternatives[record['事项编号']].append(record)
    pair_rows, group_rows, single_rows = [], [], []
    manual_rows = {'人工全查': [], '人工抽样': []}
    for item in index:
        records = by_item[item['事项编号']]
        banks = [r for r in records if r['来源'] == '银行流水']
        journals = [r for r in records if r['来源'] == '序时账']
        if len(banks) == len(journals) == 1:
            bank, journal = banks[0], journals[0]
            row = {'核对编号': item['核对编号'], '事项编号': item['事项编号'], '银行日期': bank['日期'], '银行记录': '\n'.join(str(x) for x in (bank['摘要'], bank['对方及凭证']) if x), '银行金额': bank['金额'], '序时账日期': journal['日期'], '序时账记录': '\n'.join(str(x) for x in (journal['摘要'], journal['对方及凭证']) if x), '序时账金额': journal['金额'], '差额': float(Decimal(str(bank['金额'])) - Decimal(str(journal['金额']))), '核对依据': item['判断依据'], '核对方式': item['核对方式'], '银行原始出处': bank['原始出处'], '序时账原始出处': journal['原始出处'], '银行记录号': bank['原始记录号'], '序时账记录号': journal['原始记录号']}
            pair_rows.append(row)
        elif banks and journals:
            main, _ = _主行(item, records)
            group_rows.extend([main, *_组成(item, records)])
        else:
            single_rows.extend(records)
        mode = item['核对方式']
        if mode in manual_rows:
            main, grouped = _主行(item, records, originals[item['事项编号']])
            manual_rows[mode].append(main)
            if grouped:
                manual_rows[mode].extend(_组成(item, records))
            # 一套候选只列一行摘要，完整组成集中保留在其他对应供选择。
            候选组 = defaultdict(list)
            for other in alternatives[item['事项编号']]:
                候选组[other['可选结果']].append(other)
            for label, group in 候选组.items():
                行 = {'核对编号': item['核对编号'], '事项编号': item['事项编号'],
                     '行别': label.replace('采用', ''), '核对原因': group[0]['关系说明'],
                     '原始出处': '完整组成集中见其他对应供选择；候选不是新增交易'}
                for side, source in (('银行', '银行流水'), ('序时账', '序时账')):
                    组成 = [r for r in group if r['来源'] == source]
                    收, 支 = _收支(组成)
                    行[side+'记录'] = f'{len(组成)}笔；收入 {收:,.2f}；支出 {支:,.2f}'
                manual_rows[mode].append(行)
    tables['逐笔核对'] = pd.DataFrame(pair_rows, columns=['核对编号', '银行日期', '银行记录', '银行金额', '序时账日期', '序时账记录', '序时账金额', '差额', '当前核对结论', '核对依据', '核对方式', '银行原始出处', '序时账原始出处', '事项编号', '银行记录号', '序时账记录号'])
    tables['整组核对'] = pd.DataFrame(group_rows, columns=[x for x in 组展示列 if x not in ('人工核对结果', '备注', '填报编号', '银行金额', '序时账金额', '差额')])
    tables['未对应记录'] = pd.DataFrame(single_rows, columns=['核对编号', '来源', '日期', '摘要', '对方及凭证', '金额', '当前核对结果', '核对方式', '原始出处', '事项编号', '原始记录号'])
    for name, rows in manual_rows.items():
        tables[name] = pd.DataFrame(rows, columns=组展示列)
    month = tables['月度核对'].copy()
    for name in ('银行-月末余额', '序时账-月末余额', '余额差额'):
        if name not in month:
            month[name] = None
    first = ['月份', '银行-收入金额', '序时账-收入金额', '收入金额差额', '银行-支出金额', '序时账-支出金额', '支出金额差额', '银行-月末余额', '序时账-月末余额', '余额差额', '余额衔接差', '余额说明']
    tables['月度核对'] = month[[x for x in first if x in month] + [x for x in month if x not in first]]
    return tables
