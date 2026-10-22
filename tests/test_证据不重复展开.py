"""候选方案不是新增交易；阅读表不重复展开同一套候选的所有原始行。"""
from contextlib import closing

from tests._excel假面 import load_workbook
from tests.test_人工结果计算 import _报告


def _记录(sheet):
    行 = list(sheet.values)
    return [dict(zip(行[0], row)) for row in 行[1:]]


def test_证据仅列原始交易且候选集中跳转(tmp_path):
    文件 = _报告(tmp_path, [150, 450], [100, 200, 300, 600],
                 [((0, 1), (0, 1, 2))], [(0, (0, 1), (3,))])
    with closing(load_workbook(文件)) as 簿:
        证据 = _记录(簿['事项证据'])
        assert len([r for r in 证据 if r['行别'] == '银行组成']) == 2
        assert len([r for r in 证据 if r['行别'] == '序时账组成']) == 4
        assert not any('采用其他' in r['行别'] for r in 证据)
        主行 = next(r for r in 证据 if r['事项编号'] == 'M1' and r['行别'] == '事项结论与合计')
        assert '1套' in 主行['其他方案']
        人工 = _记录(簿['人工全查'])
        候选 = [r for r in 人工 if r['行别'] == '其他对应1']
        assert len(候选) == 1
        assert '2笔' in 候选[0]['银行记录']
        assert '1笔' in 候选[0]['序时账记录']
        assert len(_记录(簿['其他对应供选择'])) == 3
        assert len(_记录(簿['核对明细'])) == 6
        表 = 簿['事项证据']
        hs = {c.value: c.column for c in 表[1] if c.value}
        assert 表.cell(2, hs['其他方案']).hyperlink.target.startswith("#'其他对应供选择'!")
