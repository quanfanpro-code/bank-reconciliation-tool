"""在原报告上组织月度、原始明细和可直接选择的人工核对结果。"""
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

import pandas as pd
from 报告列识别 import identify_columns

from 复核事项 import build_review_items
from 报告展示 import build_display_tables


主表 = ('核对结论', '月度核对', '核对明细', '人工全查', '人工抽样', '每日统计')
人工列 = ['事项编号', '行别', '核对原因', '银行记录', '银行金额', '序时账记录', '序时账金额', '差额', '人工核对结果', '当前核对结论', '备注', '原始出处', '填报编号']
原始列 = ['原始记录号', '事项编号', '来源', '日期', '对方及凭证', '摘要', '金额', '程序结论', '核对方式', '当前对应', '当前核对结果', '原始出处', '月份', '已对平金额', '已确认收入差', '已确认支出差', '收入差贡献', '支出差贡献']
索引列 = ['核对编号', '事项编号', '匹配ID', '程序结论', '风险等级', '核对方式', '影响金额', '差异金额', '首次日期', '入选原因', '判断依据', '跨期', '类型', '组金额', '最早日期', '人工核对结果', '当前核对结论']
选项列 = ['选项编号', '事项编号', '选择文字', '自动确认', '差异金额', '跨期', '关系说明', '人工选中', '有效对应', '选择冲突', '显示结果']
成员列 = ['选项编号', '事项编号', '原始记录号', '人工占用', '实际生效', '占用次数', '生效记录号', '显示结果', '已对平']


def _frame(rows, columns):
    return pd.DataFrame(rows, columns=columns)


def _record_key(side, row, index):
    return ('银行' if side == 'bank' else '账面') + str(int(row.get('original_file_row', row.get('original_idx', index))))


def _brief(row):
    fields = row.get('aux_text_fields', {})
    fields = fields if isinstance(fields, dict) else {}
    parts = [pd.Timestamp(row['date']).strftime('%Y-%m-%d'), str(row.get('summary', ''))]
    parts.extend(str(v) for k, v in fields.items() if v is not None and str(v).strip() and str(v) not in parts and any(s in k for s in ('对方', '单位', '凭证', '业务', '批次', '客户', '供应商')))
    voucher = str(row.get('voucher_no', '') or '')
    if voucher and voucher not in parts:
        parts.append('凭证' + voucher)
    return '；'.join(parts)


def build_review_tables(reporter, tables):
    matcher = reporter.matcher
    items = build_review_items(matcher)
    display_numbers = {item['事项编号']: f'{number:03d}' for number,item in enumerate(items,1)}
    source_info = {}
    for source in getattr(reporter.precheck_report, 'source_info', ()):
        side = 'bank' if source.get('来源') == '银行流水' else 'journal'
        source_info[side] = f"{Path(source.get('文件路径','')).name} · {source.get('工作表','')}"
    def origin(side,row,index):
        label = source_info.get(side, '银行流水' if side == 'bank' else '序时账')
        return f"{label} · 第{int(row.get('original_file_row',row.get('original_idx',index)))}行"
    candidates = {x.candidate_id: x for x in getattr(matcher, 'candidates', [])}
    selected = {x.final_match_id: x for x in getattr(matcher, 'selected_candidates', [])}
    index_rows, records, options, members, comparisons = [], [], [], [], []
    manual = {'人工全查': [], '人工抽样': []}
    owner = {}
    for item in items:
        item_id = item['事项编号']
        candidate = selected.get(item.get('匹配ID'))
        index = {k: item.get(k, '') for k in 索引列}
        index.update({'核对编号': display_numbers[item_id], '类型': candidate.match_type if candidate else '单边记录', '组金额': float(max(item['银行收入'] + item['银行支出'], item['日记账收入'] + item['日记账支出'])), '最早日期': item['首次日期'], '人工核对结果': '', '当前核对结论': ''})
        index_rows.append(index)
        for side, field in (('bank', '银行索引'), ('journal', '日记账索引')):
            for idx in item[field]:
                if (side, idx) in owner:
                    raise ValueError('同一原始记录进入多个复核事项')
                owner[(side, idx)] = item
        choices = []
        if candidate is not None:
            choices.append(('确认对应并保留差额' if item['差异金额'] else '确认对应', candidate))
            for alternative in candidate.evidence.get('alternative_candidate_ids', ()):
                other = candidates.get(alternative)
                if other and (other.bank_idxs, other.journal_idxs) != (candidate.bank_idxs, candidate.journal_idxs):
                    if all((other.bank_idxs, other.journal_idxs) != (x.bank_idxs, x.journal_idxs) for _, x in choices):
                        choices.append((f'采用其他对应{len(choices)}', other))
        for number, (label, choice) in enumerate(choices):
            option_id = item_id + f'|{number}'
            difference = float(choice.metrics.total_diff_li) / 10000
            cross = len({pd.Timestamp(x).strftime('%Y-%m') for x in choice.bank_dates + choice.journal_dates}) > 1
            result = ('已确认对应，差额保留' if difference else '已确认对应') + ('；跨期保留' if cross else '')
            if number == 0 and item['核对方式'] == '自动确认':
                result = '程序已确认对应' + ('；跨期保留' if cross else '')
            options.append({'选项编号': option_id, '事项编号': item_id, '选择文字': label, '自动确认': int(number == 0 and item['核对方式'] == '自动确认'), '差异金额': difference, '跨期': int(cross), '关系说明': choice.processing_reason or item['判断依据'], '显示结果': result})
            for side, indexes in (('bank', choice.bank_idxs), ('journal', choice.journal_idxs)):
                frame = getattr(matcher, side)
                for idx in indexes:
                    source_row = frame.loc[idx]
                    key = _record_key(side, source_row, idx)
                    members.append({'选项编号': option_id, '事项编号': item_id, '原始记录号': key})
                    if number:
                        comparisons.append({'事项编号': item_id, '可选结果': label, '来源': '银行流水' if side == 'bank' else '序时账', '日期': source_row['date'], '摘要及对方': _brief(source_row), '金额': float(source_row['amount']), '关系说明': choice.processing_reason or choice.evidence.get('business_basis', ''), '原始出处': origin(side,source_row,idx)})
        mode = item['核对方式']
        if mode in manual:
            bank_rows = [matcher.bank.loc[x] for x in item['银行索引']]
            journal_rows = [matcher.journal.loc[x] for x in item['日记账索引']]
            manual[mode].append({'事项编号': item_id, '填报编号': item_id, '行别': '核对事项', '核对原因': item['入选原因'] + ('；跨期' if item['跨期'] else '') + ('；' + item['判断依据'] if candidate else ''), '银行记录': _brief(bank_rows[0]) if len(bank_rows) == 1 else f'{len(bank_rows)}笔，完整组成如下' if bank_rows else '未找到银行对应', '银行金额': float(sum((x['amount'] for x in bank_rows), Decimal(0))) if bank_rows else None, '序时账记录': _brief(journal_rows[0]) if len(journal_rows) == 1 else f'{len(journal_rows)}笔，完整组成如下' if journal_rows else '未找到序时账对应', '序时账金额': float(sum((x['amount'] for x in journal_rows), Decimal(0))) if journal_rows else None, '差额': float(item['差异金额']), '人工核对结果': '', '当前核对结论': '', '备注': '', '原始出处': '；'.join(origin(side,getattr(matcher,side).loc[idx],idx) for side,field in (('bank','银行索引'),('journal','日记账索引')) for idx in item[field])})
            if len(bank_rows) > 1 or len(journal_rows) > 1:
                for label, side_rows in (('银行', bank_rows), ('序时账', journal_rows)):
                    for row in side_rows:
                        manual[mode].append({'事项编号': item_id, '行别': label+'组成', label+'记录': _brief(row), label+'金额': float(row['amount']), '原始出处': str(int(row.get('original_file_row', row.get('original_idx', 0))))})
    for side, source in (('bank', '银行流水'), ('journal', '序时账')):
        frame = getattr(matcher, side)
        for idx, row in frame.sort_values(['date', 'original_idx'], kind='stable').iterrows():
            item = owner[(side, idx)]
            amount = Decimal(str(row['amount']))
            sign = 1 if side == 'bank' else -1
            fields = row.get('aux_text_fields', {})
            fields = fields if isinstance(fields, dict) else {}
            party = '；'.join(str(v) for k, v in fields.items() if v is not None and any(s in k for s in ('对方', '单位', '凭证', '业务', '批次', '客户', '供应商')))
            voucher = str(row.get('voucher_no', '') or '')
            if voucher and voucher not in party:
                party += ('；' if party else '') + '凭证' + voucher
            records.append({'原始记录号': _record_key(side,row,idx), '事项编号': item['事项编号'], '来源': source, '日期': row['date'], '对方及凭证': party, '摘要': row.get('summary',''), '金额': float(amount), '程序结论': item['程序结论'], '核对方式': item['核对方式'], '当前对应': '', '当前核对结果': '', '原始出处': origin(side,row,idx), '月份': row['date'].strftime('%Y-%m'), '收入差贡献': float(max(amount,Decimal(0))*sign), '支出差贡献': float(max(-amount,Decimal(0))*sign)})
    records.sort(key=lambda row: (display_numbers[row['事项编号']], row['来源'] != '银行流水', row['日期'], row['原始记录号']))
    totals = defaultdict(lambda: {'收入差额':Decimal(0), '支出差额':Decimal(0)})
    for record in records:
        key = (record['月份'],record['事项编号'])
        totals[key]['收入差额'] += Decimal(str(record['收入差贡献']))
        totals[key]['支出差额'] += Decimal(str(record['支出差贡献']))
    contributions = [{'月份': month, '事项编号': item_id, **values, '净发生额差': values['收入差额']-values['支出差额'], '当前核对结论': ''} for (month,item_id),values in sorted(totals.items())]
    monthly = tables['月度统计'].copy()
    monthly = monthly.rename(columns={c: c.replace('日记账','序时账') for c in monthly.columns})
    # 月度原值不变；确认与未确认贡献使用实际原记录计算。
    for column in ['已确认收入差', '未确认收入差', '已确认支出差', '未确认支出差', '余额衔接差', '余额说明']:
        monthly[column] = ''
    result = {
        '月度核对': monthly,
        '核对明细': _frame(records, 原始列),
        '人工全查': _frame(manual['人工全查'],人工列),
        '人工抽样': _frame(manual['人工抽样'],人工列),
        '月度差异组成': _frame(contributions,['月份','事项编号','收入差额','支出差额','净发生额差','已确认收入差','未确认收入差','已确认支出差','未确认支出差','当前核对结论']),
        '其他对应供选择': _frame(comparisons,['事项编号','可选结果','来源','日期','摘要及对方','金额','关系说明','原始出处']),
        '复核事项索引': _frame(index_rows,索引列),
        '复核候选选项': _frame(options,选项列),
        '复核候选组成': _frame(members,成员列),
    }
    for name in ('核对明细','人工全查','人工抽样','月度差异组成','其他对应供选择'):
        result[name].insert(0,'核对编号',result[name]['事项编号'].map(display_numbers))
    result['核对明细'].insert(result['核对明细'].columns.get_loc('当前对应'),'当前核对编号','')
    return build_display_tables(result)
