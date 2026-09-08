"""集中筛选条件及新报告列顺序设置。"""
import calendar
from datetime import date
from decimal import Decimal, InvalidOperation
import queue
from pathlib import Path
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from 底稿筛选 import FilterCriteria, validate_filter_criteria, _describe_criteria, read_filter_report, summarize_filter
from 报告偏好 import 默认列, 可配置表, load_column_preferences, save_column_preferences


class FilterDialog(tk.Toplevel):
    condition_fields = {'业务类型':'business_types', '核对结果':'statuses', '判断依据包含':'reasons', '包含文字':'include_text', '排除文字':'exclude_text'}

    def __init__(self, parent, source_path=None, preferences_path=None):
        super().__init__(parent)
        self.title('筛选底稿与报告列顺序')
        self.geometry('900x800')
        self.minsize(820, 760)
        self.transient(parent)
        self.result = None
        self.source_path = None
        self.preferences_path = preferences_path
        self.events = queue.Queue()
        self.preview_started = None
        self.preview_criteria = None
        self.report_data = None
        self.report_options = {}
        self.source_error = ''
        self._loaded_stamp = None
        self._generation = 0
        self._busy = False
        self._pending = None
        self.variables = {name:tk.StringVar(self,value='不限' if name=='amount_basis' else '') for name in ('start_date','end_date','coverage_ratio','amount_basis','min_amount','max_amount')}
        self.absolute = tk.BooleanVar(self, value=True)
        self.preferences = load_column_preferences(preferences_path)
        self.protocol('WM_DELETE_WINDOW', self.cancel)
        source_bar = ttk.Frame(self, padding=(16,12,16,0)); source_bar.pack(fill='x')
        ttk.Button(source_bar, text='选择 / 更换报告', command=self.choose_source).pack(side='left')
        self.source_text = tk.StringVar(self, value='先选择全量核对报告，业务类型和核对结果会自动加载。')
        ttk.Label(source_bar,textvariable=self.source_text,wraplength=640).pack(side='left',padx=12)
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill='both',expand=True,padx=16,pady=12)
        body, self.column_tab = ttk.Frame(self.tabs,padding=12), ttk.Frame(self.tabs,padding=14)
        self.tabs.add(body,text='筛选条件'); self.tabs.add(self.column_tab,text='以后新报告的列顺序')
        body.columnconfigure(1,weight=1)
        self._rules_ui(body)
        fields=[('start_date','开始日期',()),('end_date','结束日期',()),('amount_basis','金额口径',('不限','银行单笔','序时账单笔','整组金额')),('min_amount','金额下限（含，元）',()),('max_amount','金额上限（含，元）',()),('coverage_ratio','条件内累计覆盖（%）',())]
        for row,(key,label,options) in enumerate(fields,1):
            ttk.Label(body,text=label).grid(row=row,column=0,sticky='w',padx=(0,12),pady=4)
            widget=ttk.Combobox(body,textvariable=self.variables[key],values=options,state='readonly') if options else ttk.Entry(body,textvariable=self.variables[key])
            widget.grid(row=row,column=1,sticky='ew',pady=4)
            if key in ('start_date','end_date'):
                ttk.Button(body,text='选日期',command=lambda k=key:self.pick_date(k)).grid(row=row,column=2,padx=(8,0))
        ttk.Checkbutton(body,text='单笔取绝对金额；取消后按收入正、支出负比较',variable=self.absolute).grid(row=7,column=0,columnspan=3,sticky='w',pady=7)
        ttk.Label(body,text='金额上下限留空不限；整组金额取双方收支绝对金额合计的较大值。',wraplength=780).grid(row=8,column=0,columnspan=3,sticky='w')
        self.description = tk.StringVar(self)
        description_box=ttk.Frame(body)
        description_box.grid(row=9,column=0,columnspan=3,sticky='ew',pady=8)
        self.description_view=tk.Text(description_box,height=3,wrap='word',font=('微软雅黑',10),foreground='#24445B',state='disabled')
        self.description_view.pack(side='left',fill='x',expand=True)
        description_scroll=ttk.Scrollbar(description_box,command=self.description_view.yview)
        description_scroll.pack(side='right',fill='y'); self.description_view.configure(yscrollcommand=description_scroll.set)
        self.description.trace_add('write',self._show_description)
        self.preview_text = tk.StringVar(self,value='选好报告后自动读取；添加条件，再预览数量或另存底稿。')
        ttk.Label(body,textvariable=self.preview_text,wraplength=770).grid(row=10,column=0,columnspan=3,sticky='ew',pady=5)
        actions=ttk.Frame(self,padding=(16,0,16,14)); actions.pack(side='bottom',fill='x',before=self.tabs)
        self.preview_button=ttk.Button(actions,text='预览数量',command=self.preview,state='disabled'); self.preview_button.pack(side='left')
        ttk.Button(actions,text='清空筛选',command=self.clear).pack(side='left',padx=8)
        ttk.Button(actions,text='取消',command=self.cancel).pack(side='right')
        self.export_button=ttk.Button(actions,text='另存筛选底稿',command=self.confirm,state='disabled'); self.export_button.pack(side='right',padx=8)
        self._columns_ui(self.column_tab)
        for variable in [*self.variables.values(),self.absolute]:
            variable.trace_add('write',self._refresh_description)
        self._refresh_description()
        self.poll_id=self.after(120,self._poll)
        self.grab_set()
        if source_path:
            self.set_source(source_path)

    def _rules_ui(self,body):
        box=ttk.LabelFrame(body,text='文字与结果条件：逐条添加，无需分隔符',padding=8)
        box.grid(row=0,column=0,columnspan=3,sticky='ew',pady=(0,10)); box.columnconfigure(1,weight=1)
        self.rule_kind=tk.StringVar(self,value='包含文字')
        self.rule_value=tk.StringVar(self)
        ttk.Combobox(box,textvariable=self.rule_kind,values=tuple(self.condition_fields),state='readonly',width=14).grid(row=0,column=0,padx=(0,8))
        self.rule_entry=ttk.Combobox(box,textvariable=self.rule_value)
        self.rule_entry.grid(row=0,column=1,sticky='ew')
        self.rule_entry.bind('<Return>',lambda _:self.add_rule())
        buttons=ttk.Frame(box); buttons.grid(row=0,column=2,padx=(8,0))
        ttk.Button(buttons,text='添加',command=self.add_rule,width=6).pack(side='left')
        ttk.Button(buttons,text='修改所选',command=self.edit_rule,width=9).pack(side='left',padx=4)
        ttk.Button(buttons,text='移除所选',command=self.remove_rule,width=9).pack(side='left')
        self.rule_list=ttk.Treeview(box,columns=('kind','value'),show='headings',selectmode='extended',height=4)
        self.rule_list.heading('kind',text='条件类型'); self.rule_list.heading('value',text='具体内容（点击一行可修改）')
        self.rule_list.column('kind',width=135,stretch=False); self.rule_list.column('value',width=460)
        self.rule_list.grid(row=1,column=0,columnspan=3,sticky='ew',pady=6)
        scroll=ttk.Scrollbar(box,orient='vertical',command=self.rule_list.yview); scroll.grid(row=1,column=3,sticky='ns')
        self.rule_list.configure(yscrollcommand=scroll.set)
        self.rule_list.bind('<<TreeviewSelect>>',self._select_rule)
        self.rule_list.bind('<Delete>',lambda _:self.remove_rule())
        self.rule_kind.trace_add('write',self._rule_choices)
        self.rule_message=tk.StringVar(self,value='输入后点击“添加”，条件才会进入下面的清单。')
        ttk.Label(box,textvariable=self.rule_message,wraplength=770).grid(row=2,column=0,columnspan=4,sticky='w')
        ttk.Label(box,text='包含文字须全部满足；排除文字任一出现即排除。\n业务类型、核对结果、判断依据各自满足任一项即可；不同类别同时满足。',wraplength=770).grid(row=3,column=0,columnspan=4,sticky='w')

    def _rule_choices(self,*_):
        field=self.condition_fields[self.rule_kind.get()]
        selectable=field in ('business_types','statuses')
        self.rule_entry.configure(values=self.report_options.get(field,()),state='readonly' if selectable else 'normal')

    def _select_rule(self,_=None):
        selected=self.rule_list.selection()
        if len(selected)==1:
            kind,value=self.rule_list.item(selected[0],'values')
            self.rule_kind.set(kind); self.rule_value.set(value)

    def _rule_input(self):
        kind,value=self.rule_kind.get(),self.rule_value.get().strip()
        if not value:
            raise ValueError('请填写或选择一个条件内容。')
        field=self.condition_fields[kind]
        if field in ('business_types','statuses') and value not in self.report_options.get(field,()):
            raise ValueError('请先选择报告，再从实际选项中选择此条件。')
        return kind,value

    def add_rule(self):
        self._save_rule()

    def edit_rule(self):
        selected=self.rule_list.selection()
        if len(selected)!=1:
            self.rule_message.set('请先选中一条要修改的条件。'); return
        self._save_rule(selected[0])

    def _save_rule(self,identity=None):
        try: values=self._rule_input()
        except ValueError as exc:
            self.rule_message.set(str(exc)); return
        if any(tuple(self.rule_list.item(row,'values'))==values for row in self.rule_list.get_children() if row!=identity):
            self.rule_message.set('此条件已在清单中，无需重复添加。'); return
        if identity:self.rule_list.item(identity,values=values)
        else:self.rule_list.insert('','end',values=values)
        self.rule_list.selection_remove(*self.rule_list.selection())
        self.rule_value.set(''); self.rule_message.set('条件已更新。清单中的全部条件会参与筛选。')
        self._refresh_description()

    def remove_rule(self):
        selected=self.rule_list.selection()
        if selected:self.rule_list.delete(*selected)
        self.rule_value.set(''); self._refresh_description()

    def build_criteria(self):
        values={name:variable.get().strip() for name,variable in self.variables.items()}
        rules={field:[] for field in self.condition_fields.values()}
        for row in self.rule_list.get_children():
            kind,value=self.rule_list.item(row,'values'); rules[self.condition_fields[kind]].append(value)
        try:
            result=FilterCriteria(coverage_ratio=float(Decimal(values['coverage_ratio'])/100) if values['coverage_ratio'] else None,
                start_date=values['start_date'] or None,end_date=values['end_date'] or None,
                **{key:tuple(items) for key,items in rules.items()},
                amount_basis=values['amount_basis'],min_amount=Decimal(values['min_amount']) if values['min_amount'] else None,
                max_amount=Decimal(values['max_amount']) if values['max_amount'] else None,amount_absolute=self.absolute.get())
        except (InvalidOperation,ValueError) as exc:
            raise ValueError('金额和覆盖比例请填写有效数字') from exc
        validate_filter_criteria(result)
        return result

    def _show_description(self,*_):
        self.description_view.configure(state='normal')
        self.description_view.delete('1.0','end'); self.description_view.insert('1.0',self.description.get())
        self.description_view.configure(state='disabled')

    def _refresh_description(self,*_):
        if self.preview_criteria is not None:
            self.preview_text.set('条件已修改，请重新预览数量。')
        try:self.description.set(_describe_criteria(self.build_criteria()))
        except ValueError as exc:self.description.set(str(exc))

    def clear(self):
        for name,var in self.variables.items():var.set('不限' if name=='amount_basis' else '')
        self.absolute.set(True)
        rows=self.rule_list.get_children()
        if rows:self.rule_list.delete(*rows)
        self.rule_value.set(''); self._refresh_description()

    def choose_source(self):
        path=filedialog.askopenfilename(parent=self,title='选择全量核对报告',filetypes=[('Excel','*.xlsx')])
        if path:self.set_source(path)

    @staticmethod
    def _file_stamp(path):
        value=Path(path).stat()
        return value.st_mtime_ns,value.st_size

    def set_source(self,path):
        path=Path(path).resolve()
        if self.source_path and path!=self.source_path:self.clear()
        self.source_path=path; self._generation+=1
        self.report_data=None; self.report_options={}; self.source_error=''; self._loaded_stamp=None
        self._rule_choices()
        self.source_text.set(str(path))
        self.preview_text.set('正在读取报告并自动加载选项，窗口可继续操作…')
        self.preview_button.configure(state='disabled'); self.export_button.configure(state='disabled')
        self._queue_job('load',FilterCriteria())

    def _queue_job(self,kind,criteria):
        # 同时只读取一份报告；快速换文件时仅保留最新待办，避免反复占用大内存。
        self._pending=(self._generation,kind,self.source_path,criteria,self.report_data,self._loaded_stamp)
        self._launch_pending()

    def _launch_pending(self):
        if self._busy or self._pending is None:return
        generation,kind,path,criteria,data,stamp=self._pending
        self._pending=None; self._busy=True; self.preview_started=time.monotonic()
        events, file_stamp = self.events, self._file_stamp
        def work():
            try:
                if kind=='load':
                    stamp_before=file_stamp(path)
                    data_read=read_filter_report(path)
                    if stamp_before!=file_stamp(path):
                        raise ValueError('读取过程中报告被修改，请保存完成后重新选择报告。')
                    value=summarize_filter(*data_read,criteria)
                else:
                    data_read=data; stamp_before=stamp
                    value=summarize_filter(*data_read,criteria)
                events.put((generation,kind,criteria,data_read,stamp_before,value,None))
            except Exception as exc:
                events.put((generation,kind,criteria,None,None,None,str(exc)))
        threading.Thread(target=work,daemon=True).start()

    def _ready(self):
        if self.report_data is None:
            self.preview_text.set(self.source_error or '请先选择报告并等待读取完成。'); return False
        try:fresh=self._file_stamp(self.source_path)==self._loaded_stamp
        except OSError:fresh=False
        if not fresh:
            self.set_source(self.source_path)
            self.preview_text.set('报告已更新，正在重新读取；完成后请再次操作。')
        return fresh

    def _criteria_for_action(self):
        value=self.rule_value.get().strip()
        if value and not any(tuple(self.rule_list.item(row,'values'))==(self.rule_kind.get(),value) for row in self.rule_list.get_children()):
            raise ValueError('输入的条件尚未加入清单，请点击“添加”或“修改所选”。')
        result=self.build_criteria()
        for label,field in (('业务类型','business_types'),('核对结果','statuses')):
            if any(item not in self.report_options.get(field,()) for item in getattr(result,field)):
                raise ValueError(f'当前报告不再包含所选{label}，请在条件清单中修改或移除。')
        return result

    def preview(self):
        if not self._ready():return
        try:criteria=self._criteria_for_action()
        except ValueError as exc:
            self.preview_text.set(str(exc)); return
        self.preview_button.configure(state='disabled')
        self.preview_text.set('正在按当前条件计算选中数量…')
        self._queue_job('preview',criteria)

    def _poll(self):
        try:
            while True:
                generation,kind,criteria,data,stamp,value,error=self.events.get_nowait()
                self._busy=False; self.preview_started=None
                if generation!=self._generation:continue
                if error:
                    self.preview_text.set('读取或预览失败：'+error)
                    if kind=='load':self.source_error=error
                else:
                    if kind=='load':
                        self.report_data=data; self._loaded_stamp=stamp
                        self.report_options=value['options']; self._rule_choices()
                        self.export_button.configure(state='normal')
                    self.preview_criteria=criteria
                    current=None
                    try:current=self.build_criteria()
                    except ValueError:pass
                    prefix='' if current==criteria else '条件已修改，以下为读取时的结果：'
                    self.preview_text.set(prefix+f"选中 {value['selected']} / {value['total']} 项；整项金额 {value['amount']:,.2f} 元。" )
                if self.report_data is not None:self.preview_button.configure(state='normal')
        except queue.Empty:pass
        self._launch_pending()
        if self.preview_started is not None:
            elapsed=int(time.monotonic()-self.preview_started)
            if elapsed>=15:self.preview_text.set(f'正在读取或计算，已运行 {elapsed} 秒，窗口可继续操作…')
        self.poll_id=self.after(120,self._poll)

    def confirm(self):
        if not self._ready():return
        try:self.result=self._criteria_for_action()
        except ValueError as exc:
            self.preview_text.set(str(exc)); return
        self._close()

    def cancel(self):
        self.result=None; self._close()

    def _close(self):
        if getattr(self,'poll_id',None):self.after_cancel(self.poll_id)
        self._pending=None
        self.grab_release(); self.destroy()

    def _columns_ui(self,body):
        ttk.Label(body,text='按住列名拖动即可排序，也可用上移、下移或Alt+上下键。核对编号固定在最前，所有证据列继续保留。\n保存后用于以后新生成的报告；已经填写的报告及本次筛选版不搬动列。',wraplength=720).pack(anchor='w',pady=(0,12))
        self.table_var=tk.StringVar(self,value=可配置表[0])
        combo=ttk.Combobox(body,textvariable=self.table_var,values=可配置表,state='readonly');combo.pack(fill='x')
        self.column_list=tk.Listbox(body,exportselection=False,font=('微软雅黑',11));self.column_list.pack(fill='both',expand=True,pady=12)
        self._dragging=False
        self.column_list.bind('<ButtonPress-1>',self._drag_start)
        self.column_list.bind('<B1-Motion>',self._drag_move)
        self.column_list.bind('<ButtonRelease-1>',self._drag_end)
        self.column_list.bind('<Alt-Up>',lambda e:(self.move_column(-1),'break')[1])
        self.column_list.bind('<Alt-Down>',lambda e:(self.move_column(1),'break')[1])
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

    def _drag_start(self,event):
        self._dragging=True
        self.column_list.focus_set()
        self.column_list.selection_clear(0,'end')
        self.column_list.selection_set(self.column_list.nearest(event.y))
        return 'break'

    def _drag_move(self,event):
        if not self._dragging:return
        if event.y<0:self.column_list.yview_scroll(-1,'units')
        elif event.y>=self.column_list.winfo_height():self.column_list.yview_scroll(1,'units')
        self._move_column_to(self.column_list.nearest(event.y))
        return 'break'

    def _drag_end(self,event):
        self._drag_move(event); self._dragging=False
        return 'break'

    def move_column(self,step):
        selected=self.column_list.curselection()
        if selected:self._move_column_to(selected[0]+step)

    def _move_column_to(self,target):
        selected=self.column_list.curselection()
        if not selected or not 0<=target<self.column_list.size():return
        row=selected[0]
        if row==target:return
        value=self.column_list.get(row); self.column_list.delete(row); self.column_list.insert(target,value)
        self.column_list.selection_clear(0,'end'); self.column_list.selection_set(target); self.column_list.see(target)
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
