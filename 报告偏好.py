"""保存新报告列排列；全部原始列保留，不移动已有报告的人工输入。"""
import json
import os
from pathlib import Path

可配置表 = ('逐笔核对', '整组核对', '未对应记录', '人工全查', '人工抽样')
默认列 = {
    '逐笔核对': ['银行日期','银行记录','银行金额','序时账日期','序时账记录','序时账金额','差额','当前核对结论','核对依据','核对方式','银行原始出处','序时账原始出处'],
    '整组核对': ['行别','银行记录','银行收入','银行支出','序时账记录','序时账收入','序时账支出','核对差额','当前核对结论','核对原因','原始出处'],
    '未对应记录': ['来源','日期','摘要','对方及凭证','金额','当前核对结果','核对方式','原始出处'],
}
for _name in ('人工全查','人工抽样'):
    默认列[_name] = ['行别','银行记录','银行收入','银行支出','序时账记录','序时账收入','序时账支出','核对差额','人工核对结果','当前核对结论','核对原因','备注','原始出处']


def _path(path=None):
    return Path(path) if path is not None else Path(os.environ.get('LOCALAPPDATA', Path.home())) / '银行流水核对工具' / '报告列顺序.json'


def load_column_preferences(path=None):
    target = _path(path)
    if not target.exists():
        return {}
    try:
        data = json.loads(target.read_text(encoding='utf-8-sig'))
    except (ValueError, OSError):
        return {}
    if not isinstance(data, dict):
        return {}
    return {name: list(dict.fromkeys(column for column in columns if isinstance(column, str))) for name, columns in data.items() if name in 可配置表 and isinstance(columns, list)}


def save_column_preferences(preferences, path=None):
    target = _path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_suffix('.new')
    temporary.write_text(json.dumps(preferences, ensure_ascii=False, indent=2), encoding='utf-8-sig')
    temporary.replace(target)


def apply_column_preferences(tables, preferences=None):
    settings = load_column_preferences() if preferences is None else preferences
    for name in 可配置表:
        if name not in tables or not settings.get(name):
            continue
        frame = tables[name]
        first = ['核对编号'] if '核对编号' in frame else []
        order = list(dict.fromkeys(first + [x for x in settings[name] if x in frame] + list(frame.columns)))
        tables[name] = frame[order]
    return tables
