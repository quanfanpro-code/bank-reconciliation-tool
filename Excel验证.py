# -*- coding: utf-8 -*-
"""报告交付前由独立的 Microsoft Excel 打开、重算并保存；失败即拒绝交付。"""
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
import gc
import psutil
import xml.etree.ElementTree as ET
import zipfile


def 检查人工填写区(路径):
    """程序公式不得混入选择格；重算后也不能把无效选择当作正常报告交付。"""
    from 直读XML import _表目标, _读共享串
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    with zipfile.ZipFile(路径) as 包:
        目标 = next((部件 for 名, 部件, _ in _表目标(包) if 名 == '待处理事项'), None)
        if 目标 is None:
            return 0
        串 = _读共享串(包)
        def 值(c):
            v = c.findtext(ns+'v')
            if c.get('t') == 's' and v is not None:
                return 串[int(v)]
            if c.get('t') == 'inlineStr':
                return ''.join(t.text or '' for t in c.iter(ns+'t'))
            return v or ''
        数据 = ET.fromstring(包.read(目标)).find(ns+'sheetData')
        if 数据 is None or not len(数据):
            return 0
        列名 = {''.join(x for x in c.get('r') if x.isalpha()): 值(c) for c in 数据[0]}
        if '人工核对结果' not in 列名.values():
            raise RuntimeError('填写区缺少人工核对结果列，未交付报告')
        行数 = 0
        for row in list(数据)[1:]:
            行数 += 1
            for c in row:
                字段 = 列名.get(''.join(x for x in c.get('r') if x.isalpha()))
                文 = str(值(c)).lstrip("' \t\r\n")
                if 字段 == '人工核对结果' and (c.find(ns+'f') is not None or 文.startswith('=')):
                    raise RuntimeError(f'填写区 待处理事项!{c.get("r")} 包含公式或公式文本，未交付报告')
                if 字段 == '备注' and 文.startswith('=IF(') and 'INDEX(' in 文 and '待处理事项' in 文:
                    raise RuntimeError(f'填写区 待处理事项!{c.get("r")} 包含程序回显公式，未交付报告')
                if 字段 == '当前结果' and '选择无效' in 文:
                    raise RuntimeError(f'填写区 待处理事项!{c.get("r")} 存在无效选择，未交付报告')
        return 行数


def 验证Excel报告(路径, *, logger=None, cancel_check=None):
    检查人工填写区(路径)
    try:
        import pythoncom
        import win32com.client
        import win32process
    except ImportError as exc:
        raise RuntimeError('缺少 Excel 验证组件 pywin32，请按 README 安装依赖；未交付报告') from exc
    if cancel_check:
        cancel_check()
    if logger:
        logger('开始用本机 Excel 实际打开、重算并验证报告')
    def 打开并重算():
        # COM 的初始化和释放独占一个线程，不能注销调用者已有的 Excel 连接。
        pythoncom.CoInitialize()
        app = book = None
        自有进程 = None
        try:
            原有进程 = {p.pid for p in psutil.process_iter(['name']) if (p.info['name'] or '').lower() == 'excel.exe'}
            app = win32com.client.DispatchEx('Excel.Application')
            pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
            if pid in 原有进程:
                app = None
                raise RuntimeError('未能创建独立 Excel 进程，已保留原有 Excel 窗口')
            自有进程 = psutil.Process(pid)
            app.Visible = False
            app.IgnoreRemoteRequests = True
            app.DisplayAlerts = False
            app.EnableEvents = False
            app.AskToUpdateLinks = False
            app.AutomationSecurity = 3
            book = app.Workbooks.Open(Filename=str(Path(路径).resolve()), UpdateLinks=0,
                                      ReadOnly=False, IgnoreReadOnlyRecommended=True,
                                      AddToMru=False, CorruptLoad=0)
            if book.ReadOnly:
                raise RuntimeError('报告被只读打开，无法保存公式计算结果')
            app.CalculateFullRebuild()
            if cancel_check:
                cancel_check()
            book.Save()
            return {'Excel版本': str(app.Version), '工作表数': int(book.Worksheets.Count)}
        except InterruptedError:
            raise
        except Exception as exc:
            raise RuntimeError(f'本机 Excel 打开或重算验证失败，未交付报告：{exc}') from exc
        finally:
            try:
                if book is not None:
                    book.Close(SaveChanges=False)
            finally:
                try:
                    if app is not None:
                        app.Quit()
                finally:
                    book = app = None
                    gc.collect()   # 在所属线程释放 COM 包装对象，避免后台 Excel 残留
                    pythoncom.CoUninitialize()
                    if 自有进程 is not None:
                        try:
                            自有进程.wait(timeout=1)
                        except psutil.TimeoutExpired:
                            # Excel 已关簿并 Quit；只回收本次新建、仍被 COM 引用留住的进程。
                            自有进程.terminate()
                            自有进程.wait(timeout=5)
    with ThreadPoolExecutor(max_workers=1) as 线程:
        结果 = 线程.submit(打开并重算).result()

    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    relns = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    公式数 = 0
    with zipfile.ZipFile(路径) as 包:
        rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(包.read('xl/_rels/workbook.xml.rels'))}
        for 表 in ET.fromstring(包.read('xl/workbook.xml')).find(ns+'sheets'):
            部件 = rels[表.get(relns+'id')]
            部件 = 部件.lstrip('/') if 部件.startswith('/') else 'xl/' + 部件
            with 包.open(部件) as 流:
                for _, 元 in ET.iterparse(流, events=('end',)):
                    if 元.tag == ns+'c':
                        if 元.get('t') == 'e':
                            raise RuntimeError(f'Excel重算发现错误：{表.get("name")}!{元.get("r")} {元.findtext(ns+"v")}；未交付报告')
                        if 元.find(ns+'f') is not None:
                            公式数 += 1
                        元.clear()
                    elif 元.tag == ns+'row':
                        元.clear()
    结果['公式数'] = 公式数
    结果['人工填写行数'] = 检查人工填写区(路径)
    if logger:
        logger(f'Excel实际打开和重算验证通过：{结果["工作表数"]} 张表，{公式数:,} 个公式，无公式错误')
    return 结果
