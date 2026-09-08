"""集中筛选条件及新报告列顺序设置。"""
import calendar
from datetime import date
from decimal import Decimal, InvalidOperation
import queue
import re
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from 底稿筛选 import FilterCriteria, validate_filter_criteria, _describe_criteria, preview_filter
from 报告偏好 import 默认列, 可配置表, load_column_preferences, save_column_preferences


class FilterDialog(tk.Toplevel):
    def __init__(self, parent, source_path=None, preferences_path=None):
        super().__init__(parent)
        self.title('筛选底稿与报告列顺序')
        self.geometry('850x760')
        self.minsize(760, 680)
        self.transient(parent)
        self.result = None
        self.source_path = source_path
        self.preferences_path = preferences_path
        self.events = queue.Queue()
        self.preview_started = None
        self.preview_criteria = None
        self.variables = {name: tk.StringVar(self, value='不限' if name == 'amount_basis' else '') for name in ('business_types','statuses','reasons','start_date','end_date','include_text','exclude_text','coverage_ratio','amount_basis','min_amount','max_amount')}
        self.absolute = tk.BooleanVar(self, value=True)
        self.preferences = load_column_preferences(preferences_path)
        self.protocol('WM_DELETE_WINDOW', self.cancel)
        tabs = ttk.Notebook(self)
        tabs.pack(fill='both', expand=True, padx=16, pady=12)
        body, columns = ttk.Frame(tabs, padding=14), ttk.Frame(tabs, padding=14)
        tabs.add(body, text='筛选条件')
        tabs.add(columns, text='以后新报告的列顺序')
        body.columnconfigure(1, weight=1)
        ttk.Label(body, text='各项条件同时满足；留空不限。选中任何一笔，保留所属整项及双方组成。', wraplength=750).grid(row=0,column=0,columnspan=3,sticky='w',pady=(0,14))
        fields = [('business_types','业务类型',()),('statuses','核对结果',()),('reasons','判断依据包含',()),('start_date','开始日期',()),('end_date','结束日期',()),('include_text','包含文字',()),('exclude_text','排除文字',()),('amount_basis','金额口径',('不限','银行单笔','序时账单笔','整组金额')),('min_amount','金额下限（含，元）',()),('max_amount','金额上限（含，元）',()),('coverage_ratio','条件内累计覆盖（%）',())]
        self.combos = {}
        for row,(key,label,options) in enumerate(fields,1):
            ttk.Label(body,text=label).grid(row=row,column=0,sticky='w',pady=5,padx=(0,12))
            if options or key in ('business_types','statuses'):
                widget=ttk.Combobox(body,textvariable=self.variables[key],values=options,state='readonly' if key=='amount_basis' else 'normal')
                self.combos[key]=widget
            else:
                widget=ttk.Entry(body,textvariable=self.variables[key])
            widget.grid(row=row,column=1,sticky='ew',pady=5)
            if key in ('start_date','end_date'):
                ttk.Button(body,text='选日期',command=lambda k=key:self.pick_date(k)).grid(row=row,column=2,padx=(8,0))
        ttk.Checkbutton(body,text='单笔取绝对金额（包括负数）；取消后按收入正、支出负比较',variable=self.absolute).grid(row=12,column=0,columnspan=3,sticky='w',pady=8)
        ttk.Label(body,text='整组金额始终取双方收支绝对金额合计的较大值；多项文字用分号分隔。\n列顺序在另一标签页独立保存，对以后新生成报告生效。',wraplength=750).grid(row=13,column=0,columnspan=3,sticky='w',pady=5)
        self.description=tk.StringVar(self)
        ttk.Label(body,textvariable=self.description,wraplength=720,foreground='#24445B').grid(row=14,column=0,columnspan=3,sticky='ew',pady=8)
        self.preview_text=tk.StringVar(self,value='先预览可读取报告中的业务类型和核对结果供选择；也可直接填写条件。预览和导出均在后台执行。')
        ttk.Label(body,textvariable=self.preview_text,wraplength=720).grid(row=15,column=0,columnspan=3,sticky='ew',pady=8)
        actions=ttk.Frame(self,padding=(16,0,16,14));actions.pack(fill='x')
        self.preview_button=ttk.Button(actions,text='预览数量',command=self.preview);self.preview_button.pack(side='left')
        ttk.Button(actions,text='清空筛选',command=self.clear).pack(side='left',padx=8)
        ttk.Button(actions,text='取消',command=self.cancel).pack(side='right')
        ttk.Button(actions,text='另存筛选底稿',command=self.confirm).pack(side='right',padx=8)
        self._columns_ui(columns)
        for variable in [*self.variables.values(),self.absolute]:
            variable.trace_add('write',self._refresh_description)
        self._refresh_description()
        self.poll_id=self.after(150,self._poll)
        self.grab_set()

    def build_criteria(self):
        values={name:variable.get().strip() for name,variable in self.variables.items()}
        multi=lambda key:tuple(x.strip() for x in re.split('[;；]',values[key]) if x.strip())
        try:
            result=FilterCriteria(coverage_ratio=float(Decimal(values['coverage_ratio'])/100) if values['coverage_ratio'] else None,
                start_date=values['start_date'] or None,end_date=values['end_date'] or None,
                include_text=multi('include_text'),exclude_text=multi('exclude_text'),business_types=multi('business_types'),statuses=multi('statuses'),reasons=multi('reasons'),
                amount_basis=values['amount_basis'],min_amount=Decimal(values['min_amount']) if values['min_amount'] else None,max_amount=Decimal(values['max_amount']) if values['max_amount'] else None,amount_absolute=self.absolute.get())
        except (InvalidOperation, ValueError) as exc:
            raise ValueError('金额和覆盖比例请填写有效数字') from exc
        validate_filter_criteria(result)
        return result

    def _refresh_description(self,*_):
        if self.preview_criteria is not None:
            self.preview_text.set('条件已修改，请重新预览数量。')
        try:
            self.description.set(_describe_criteria(self.build_criteria()))
        except ValueError as exc:
            self.description.set(str(exc))

    def clear(self):
        for name,var in self.variables.items():
            var.set('不限' if name=='amount_basis' else '')
        self.absolute.set(True)

    def _source(self):
        if not self.source_path:
            self.source_path=filedialog.askopenfilename(parent=self,title='选择全量核对报告',filetypes=[('Excel','*.xlsx')]) or None
        return self.source_path

    def preview(self):
        try:
            criteria=self.build_criteria()
        except ValueError as exc:
            messagebox.showerror('筛选条件',str(exc),parent=self);return
        if not self._source():
            return
        self.preview_button.configure(state='disabled')
        self.preview_started = time.monotonic()
        self.preview_text.set('正在读取报告并预览，窗口可继续操作…')
        source=self.source_path
        def work():
            try:
                self.events.put((criteria,preview_filter(source,criteria),None))
            except Exception as exc:
                self.events.put((criteria,None,str(exc)))
        threading.Thread(target=work,daemon=True).start()

    def _poll(self):
        try:
            while True:
                criteria,value,error=self.events.get_nowait()
                self.preview_button.configure(state='normal')
                self.preview_started = None
                self.preview_criteria = criteria
                if error:
                    self.preview_text.set('预览失败：'+error)
                else:
                    current=None
                    try:current=self.build_criteria()
                    except ValueError:pass
                    prefix='' if criteria==current else '条件已修改，以下是上次预览：'
                    self.preview_text.set(prefix+f"选中 {value['selected']} / {value['total']} 项；整项金额 {value['amount']:,.2f} 元。\n"+value['description'])
                    for key,options in value.get('options',{}).items():
                        if key in self.combos:self.combos[key].configure(values=options)
        except queue.Empty:
            pass
        if self.preview_started is not None:
            elapsed = int(time.monotonic() - self.preview_started)
            if elapsed >= 15:
                self.preview_text.set(f'正在读取报告并计算预览，已运行 {elapsed} 秒，窗口可继续操作…')
        self.poll_id=self.after(150,self._poll)

    def confirm(self):
        try:
            result=self.build_criteria()
        except ValueError as exc:
            messagebox.showerror('筛选条件',str(exc),parent=self);return
        if not self._source():return
        self.result=result
        self._close()

    def cancel(self):
        self.result=None
        self._close()

    def _close(self):
        if getattr(self,'poll_id',None):self.after_cancel(self.poll_id)
        self.grab_release()
        self.destroy()

    def _columns_ui(self,body):
        ttk.Label(body,text='选择表和列，点击上移或下移。核对编号固定在最前，所有证据列继续保留。\n保存后用于以后新生成的报告；已经填写的报告及本次筛选版不搬动列。',wraplength=720).pack(anchor='w',pady=(0,12))
        self.table_var=tk.StringVar(self,value=可配置表[0])
        combo=ttk.Combobox(body,textvariable=self.table_var,values=可配置表,state='readonly');combo.pack(fill='x')
        self.column_list=tk.Listbox(body,exportselection=False,font=('微软雅黑',11));self.column_list.pack(fill='both',expand=True,pady=12)
        buttons=ttk.Frame(body);buttons.pack(fill='x')
        ttk.Button(buttons,text='上移',command=lambda:self.move_column(-1)).pack(side='left')
        ttk.Button(buttons,text='下移',command=lambda:self.move_column(1)).pack(side='left',padx=8)
        ttk.Button(buttons,text='恢复本表默认顺序',command=self.reset_columns).pack(side='left')
        ttk.Button(buttons,text='保存列顺序',command=self.save_columns).pack(side='right')
        self.columns_message=tk.StringVar(self)
        ttk.Label(body,textvariable=self.columns_message).pack(anchor='w',pady=12)
        combo.bind('<<ComboboxSelected>>',lambda _:self._show_columns())
        self._show_columns()

    def _show_columns(self):
        name=self.table_var.get()
        order=list(dict.fromkeys([x for x in self.preferences.get(name,[]) if x in 默认列[name]]+默认列[name]))
        self.column_list.delete(0,'end')
        for column in order:self.column_list.insert('end',column)

    def move_column(self,step):
        selected=self.column_list.curselection()
        if not selected:return
        row=selected[0];target=row+step
        if not 0<=target<self.column_list.size():return
        value=self.column_list.get(row);self.column_list.delete(row);self.column_list.insert(target,value);self.column_list.selection_set(target)
        self.preferences[self.table_var.get()]=list(self.column_list.get(0,'end'))
        self.columns_message.set('列顺序已调整，点击保存后生效。')

    def reset_columns(self):
        self.preferences.pop(self.table_var.get(),None);self._show_columns()
        self.columns_message.set('已恢复本表默认顺序，点击保存后生效。')

    def save_columns(self):
        try:
            save_column_preferences(self.preferences,self.preferences_path)
            self.columns_message.set('已保存；以后新生成的报告使用此列顺序。')
        except OSError as exc:
            messagebox.showerror('列顺序保存失败',str(exc),parent=self)

    def pick_date(self,key):
        popup=tk.Toplevel(self);popup.title('选择日期');popup.transient(self);popup.grab_set()
        try:chosen=date.fromisoformat(self.variables[key].get())
        except ValueError:chosen=date.today()
        month=[chosen.year,chosen.month]
        label=tk.StringVar(popup)
        bar=ttk.Frame(popup,padding=8);bar.pack(fill='x')
        grid=ttk.Frame(popup,padding=8);grid.pack()
        def pick(day):
            self.variables[key].set(date(month[0],month[1],day).isoformat());close()
        def close():
            popup.destroy();self.grab_set()
        def show(step=0):
            number=month[0]*12+month[1]-1+step;month[:]=[number//12,number%12+1]
            label.set(f'{month[0]}年{month[1]}月')
            for child in grid.winfo_children():child.destroy()
            for i,title in enumerate('一二三四五六日'):ttk.Label(grid,text=title).grid(row=0,column=i)
            for row,week in enumerate(calendar.monthcalendar(*month),1):
                for col,day in enumerate(week):
                    if day:ttk.Button(grid,text=str(day),width=4,command=lambda d=day:pick(d)).grid(row=row,column=col,padx=1,pady=2)
        ttk.Button(bar,text='上月',command=lambda:show(-1)).pack(side='left')
        ttk.Label(bar,textvariable=label).pack(side='left',padx=16)
        ttk.Button(bar,text='下月',command=lambda:show(1)).pack(side='right')
        popup.protocol('WM_DELETE_WINDOW',close);show()
