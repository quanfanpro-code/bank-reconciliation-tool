# -*- coding: utf-8 -*-
"""读取本工具报告的列名，列位置只用于定位已识别字段的单元格。"""

from 直读XML import 列转字母


def identify_columns(表, required=()):
    """报告第一行为表头；允许换序、额外列和两端空格，不猜测歧义列。

    表：直读XML.表读 对象。
    """
    columns = {}
    行 = next(表.行值(1, 1), [])
    for 列号, 值 in enumerate(行, 1):
        label = str(值).strip() if 值 is not None else ''
        if not label:
            continue
        if label in columns:
            positions = f'{列转字母(columns[label])}、{列转字母(列号)}'
            raise ValueError(f'工作表“{表.名称}”存在重复列名“{label}”（{positions}列），请区分列名后再操作')
        columns[label] = 列号
    missing = [name for name in required if name not in columns]
    if missing:
        raise ValueError(f'工作表“{表.名称}”缺少必要列：{"、".join(missing)}；请恢复列名或使用完整报告')
    return columns
