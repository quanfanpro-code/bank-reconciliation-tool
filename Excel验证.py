# -*- coding: utf-8 -*-
"""交付前验证 Excel 报告。Windows 拒绝启动 Excel 时，检查报表结构和明细跳转。"""
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
import gc
import psutil
import xml.etree.ElementTree as ET
import zipfile


class _Excel启动被拒绝(RuntimeError):
    pass


def 明细跳转公式(行, 表名, 末行, 标题):
    return f'HYPERLINK("#\'{表名}\'!A"&(MATCH(A{行},\'{表名}\'!$A$2:$A${末行},0)+1),"{标题}")'


def _验证五表结构(路径):
    from collections import defaultdict
    from 直读XML import 读簿

    with zipfile.ZipFile(路径) as 包:
        if 包.testzip() is not None:
            raise RuntimeError('报告文件损坏，未交付报告')
    簿 = 读簿(路径, 读样式=False)
    旧版 = ['核对总览', '月度核对', '日度核对', '待核对', '交易明细']
    新版 = [*旧版[:-1], '待核对明细', '交易明细']
    if 簿.表名列表 not in (旧版, 新版):
        raise RuntimeError('无法验证报告的工作表结构，未交付报告')
    if any(表.状态 != 'visible' or not 表.值(1, 1) for 表 in 簿):
        raise RuntimeError('报告缺少可见工作表或表头，未交付报告')
    待核对, 交易明细 = 簿['待核对'], 簿['交易明细']
    有专表 = 簿.表名列表 == 新版
    明细表 = 簿['待核对明细'] if 有专表 else 交易明细
    全量分组 = defaultdict(list)
    for 记录 in 交易明细.行值(2):
        全量分组[记录[0]].append(记录)
    预期公式 = {}
    预期明细 = []
    for 行 in range(2, 待核对.最大行 + 1):
        当前编号 = 待核对.值(行, 1)
        if not 当前编号:
            continue
        记录组 = 全量分组.get(当前编号)
        if not 记录组:
            raise RuntimeError(f'待核对第 {行} 行找不到交易明细，未交付报告')
        标题 = f"查看全部{len(记录组)}笔" if 有专表 else "查看明细"
        预期公式[行, 10] = 明细跳转公式(行, '待核对明细' if 有专表 else '交易明细', 明细表.最大行, 标题)
        if 有专表:
            预期明细.extend(记录组)
    if 有专表 and list(明细表.行值(2)) != 预期明细:
        raise RuntimeError('待核对明细与完整交易记录不一致，未交付报告')
    if 待核对.公式 != 预期公式 or any(表.公式 for 表 in 簿 if 表 is not 待核对):
        raise RuntimeError('报告含有异常公式或明细跳转，未交付报告')
    return {'工作表数': len(簿.表名列表), '公式数': len(预期公式), '验证方式': '工作簿结构检查',
            '人工填写行数': len(预期公式)}


def 检查人工填写区(路径):
    """程序公式不得混入选择格；重算后也不能把无效选择当作正常报告交付。"""
    from 直读XML import _表目标, _读共享串
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    with zipfile.ZipFile(路径) as 包:
        目标 = next((部件 for 名, 部件, _ in _表目标(包) if 名 in {'待核对', '待处理事项'}), None)
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
        根 = ET.fromstring(包.read(目标))
        数据 = 根.find(ns+'sheetData')
        if 数据 is None or not len(数据):
            return 0
        列名 = {''.join(x for x in c.get('r') if x.isalpha()): 值(c) for c in 数据[0]}
        if not {'人工核对结果', '核查结论'} & set(列名.values()):
            raise RuntimeError('填写区缺少核查结论列，未交付报告')
        if '核查结论' in 列名.values():
            from 直读XML import 字母转列
            填写列 = {字母转列(c) for c, name in 列名.items() if name in ('核查结论', '备注')}
            for v in 根.findall(ns+'dataValidations/'+ns+'dataValidation'):
                for ref in v.get('sqref', '').split():
                    两端 = ref.split(':')
                    起 = 字母转列(''.join(x for x in 两端[0] if x.isalpha()))
                    末 = 字母转列(''.join(x for x in 两端[-1] if x.isalpha()))
                    if any(起 <= c <= 末 for c in 填写列):
                        raise RuntimeError(f'填写区 {ref} 仍有下拉或输入限制，未交付报告')
        行数 = 0
        for row in list(数据)[1:]:
            行数 += 1
            for c in row:
                字段 = 列名.get(''.join(x for x in c.get('r') if x.isalpha()))
                文 = str(值(c)).lstrip("' \t\r\n")
                if 字段 in ('人工核对结果', '核查结论') and (c.find(ns+'f') is not None or 文.startswith('=')):
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
            try:
                app = win32com.client.DispatchEx('Excel.Application')
            except Exception as exc:
                if getattr(exc, 'hresult', None) == -2147024891:  # E_ACCESSDENIED
                    raise _Excel启动被拒绝() from exc
                raise
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
        except (InterruptedError, _Excel启动被拒绝):
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
    try:
        with ThreadPoolExecutor(max_workers=1) as 线程:
            结果 = 线程.submit(打开并重算).result()
    except _Excel启动被拒绝:
        if cancel_check:
            cancel_check()
        结果 = _验证五表结构(路径)
        if logger:
            logger('Windows 拒绝启动 Excel，工作表、填写区和明细跳转已通过文件检查')
        return 结果

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
