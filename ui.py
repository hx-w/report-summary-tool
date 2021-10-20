# ======================
# imports
# ======================
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox as mBox

from script import *

win = tk.Tk()
win.title('周报汇总工具')

VERSION = 'v0.3-alpha'


def create_ui():
    main_frame = tk.Frame(win)

    # ----------------- tab  -----------------
    notebook = ttk.Notebook(main_frame)
    # single week single group
    tab_swsg = tk.Frame(notebook)
    # single week multiple group
    tab_swmg = tk.Frame(notebook)
    tab_mw = tk.Frame(notebook)  # multiple week
    notebook.add(tab_swsg, text='单周单组汇总')
    notebook.add(tab_swmg, text='单周全组汇总')
    notebook.add(tab_mw, text='多周汇总')
    notebook.pack()

    _copyright = ttk.Label(main_frame, text=f'{VERSION} @hx')
    _copyright.pack(side=tk.RIGHT)

    def tab_changed_callback(event):
        tab = event.widget.tab('current')['text']
        if tab == '单周单组汇总':
            pass
        elif tab == '单周全组汇总':
            cb_swmg_week['value'] = get_swmg_week_list(True, True)
            text_swmg_preview.configure(state="normal")
            text_swmg_preview.delete(1.0, tk.END)
            if len(cb_swmg_week['value']) == 0:
                btn_swmg_confirm.configure(state="disabled")
                cb_swmg_week.set('')
                cb_swmg_week.configure(state="disabled")
                text_swmg_preview.insert(tk.END, '<空>')
                text_swmg_preview.configure(state="disabled")
                return
            cb_swmg_week.current(0)
            if judge_swmg_empty(cb_swmg_week.get()):
                btn_swmg_confirm.configure(state="disabled")
                text_swmg_preview.insert(tk.END, '<空>')
                text_swmg_preview.configure(state="disabled")
                # cb_swmg_week.configure(state="disabled")
                return
            cb_swmg_week.configure(state="readonly")
            btn_swmg_confirm.configure(state="ready")
            text_swmg_preview.insert(
                tk.END, get_swmg_preview(cb_swmg_week.get()))
            text_swmg_preview.configure(state="disabled")
        elif tab == '多周汇总':
            cb_mw_start_week['value'] = get_swmg_week_list(False)
            if len(cb_mw_start_week['value']) > 0:
                cb_mw_start_week.current(0)
                cb_mw_start_week.configure(state="readonly")
            else:
                cb_mw_start_week.set('')
                cb_mw_start_week.configure(state="disabled")

            cb_mw_end_week['value'] = get_swmg_week_list(True)
            if len(cb_mw_end_week['value']) > 0:
                cb_mw_end_week.current(0)
                cb_mw_end_week.configure(state="readonly")
                label_week_count.configure(text=str(get_week_count(
                    cb_mw_start_week.get(), cb_mw_end_week.get())))
            else:
                label_week_count.configure(text="0")
                cb_mw_end_week.set('')
                cb_mw_end_week.configure(state="disabled")
    notebook.bind("<<NotebookTabChanged>>", tab_changed_callback)

    # ----------------- tab 1 tab_swsg ----------------
    ttk.Label(tab_swsg, text='选择小组').grid(row=0, column=0, padx=10, pady=10)

    def cb_swsg_group_callback(event):
        cb_swsg_week['value'] = get_swmg_week_list(True, True)
        text_swsg_preview.configure(state="normal")
        text_swsg_preview.delete(1.0, tk.END)
        label_swsg_count.configure(text=f'{get_swsg_count(cb_swsg_group.get(), cb_swsg_week.get())}人')
        if judge_swsg_empty(cb_swsg_group.get(), cb_swsg_week.get()):
            btn_swsg_confirm.configure(state="disabled")
            text_swsg_preview.insert(tk.END, '<空>')
        else:
            btn_swsg_confirm.configure(state="ready")
            text_swsg_preview.insert(tk.END, get_swsg_preview(
                cb_swsg_group.get(), cb_swsg_week.get()))
        text_swsg_preview.configure(state="disabled")

    def cb_swsg_week_callback(event):
        choose_week = cb_swsg_week.get()
        text_swsg_preview.configure(state="normal")
        text_swsg_preview.delete(1.0, tk.END)
        label_swsg_count.configure(text=f'{get_swsg_count(cb_swsg_group.get(), cb_swsg_week.get())}人')
        if judge_swsg_empty(cb_swsg_group.get(), choose_week):
            btn_swsg_confirm.configure(state="disabled")
            text_swsg_preview.insert(tk.END, '<空>')
        else:
            btn_swsg_confirm.configure(state="ready")
            text_swsg_preview.insert(tk.END, get_swsg_preview(
                cb_swsg_group.get(), cb_swsg_week.get()))
        text_swsg_preview.configure(state="disabled")

    cb_swsg_group = ttk.Combobox(tab_swsg, state="readonly")
    cb_swsg_group['value'] = global_config.group
    cb_swsg_group.grid(row=0, column=1)
    cb_swsg_group.current(0)
    cb_swsg_group.bind("<<ComboboxSelected>>", cb_swsg_group_callback)

    ttk.Label(tab_swsg, text='选择周次').grid(row=1, column=0)
    cb_swsg_week = ttk.Combobox(tab_swsg, state="readonly")
    # cb_swsg_week['value'] = get_swsg_week_list(cb_swsg_group['value'][0])
    cb_swsg_week['value'] = get_swmg_week_list(True, True)
    if len(cb_swsg_week['value']) > 0:
        cb_swsg_week.current(0)
    cb_swsg_week.grid(row=1, column=1, pady=10)
    cb_swsg_week.bind("<<ComboboxSelected>>", cb_swsg_week_callback)

    ttk.Label(tab_swsg, text='成员预览').grid(row=2, column=0, pady=10)
    text_swsg_preview = tk.Text(tab_swsg, width=28, height=5, state='normal')
    text_swsg_preview.grid(row=2, column=1, rowspan=3, sticky='NW', pady=10)
    text_swsg_preview.insert(tk.END, get_swsg_preview(
        cb_swsg_group.get(), cb_swsg_week.get()))
    text_swsg_preview.configure(state="disabled")

    ttk.Label(tab_swsg, text='成员计数').grid(row=6, column=0, pady=10)
    label_swsg_count = ttk.Label(tab_swsg, text='0人')
    label_swsg_count.configure(text=f'{get_swsg_count(cb_swsg_group.get(), cb_swsg_week.get())}人')
    label_swsg_count.grid(row=6, column=1, pady=10, sticky=tk.W)

    def btn_swsg_callback():
        if judge_swsg_empty(cb_swsg_group.get(), cb_swsg_week.get()):
            mBox.showerror('错误', '小组周报文件夹为空，拒绝执行')
            return
        if judge_swsg_result_exist(cb_swsg_group.get(), cb_swsg_week.get()):
            if global_config.safemode:
                mBox.showerror('错误', '目标文件已存在，因开启安全模式，拒绝执行')
                return
            else:
                if not mBox.askyesno('警告', '目标文件已存在，是否覆盖现有文件'):
                    return
                # mBox.showwarning('警告', '目标文件已存在，因关闭安全模式，继续执行')
        distpath = exec_swsg_merge(cb_swsg_group.get(), cb_swsg_week.get())
        mBox.showinfo('提示', '已汇总至：\n' + distpath)

    btn_swsg_confirm = ttk.Button(
        tab_swsg, text="执行", command=btn_swsg_callback)
    btn_swsg_confirm.grid(row=8, column=0, columnspan=2, pady=20)

    if len(cb_swsg_week['value']) == 0:
        text_swsg_preview.configure(state="normal")
        text_swsg_preview.insert(tk.END, '<空>')
        cb_swsg_week.configure(state="disabled")
        btn_swsg_confirm.configure(state="disabled")
        text_swsg_preview.configure(state="disabled")
        
    # ----------------- tab 2 tab_swmg ----------------
    def cb_swmg_week_callback(event):
        choose_week = cb_swmg_week.get()
        text_swmg_preview.configure(state="normal")
        text_swmg_preview.delete(1.0, tk.END)
        if judge_swmg_empty(choose_week):
            btn_swmg_confirm.configure(state="disabled")
            text_swmg_preview.insert(tk.END, '<空>')
        else:
            btn_swmg_confirm.configure(state="ready")
            text_swmg_preview.insert(tk.END, get_swmg_preview(choose_week))
        text_swmg_preview.configure(state="disabled")

    ttk.Label(tab_swmg, text='选择周次').grid(row=0, column=0, padx=10, pady=10)
    cb_swmg_week = ttk.Combobox(tab_swmg, state="readonly")
    cb_swmg_week['value'] = get_swmg_week_list()

    if len(cb_swmg_week['value']) > 0:
        cb_swmg_week.current(0)
    cb_swmg_week.grid(row=0, column=1)
    cb_swmg_week.bind("<<ComboboxSelected>>", cb_swmg_week_callback)

    ttk.Label(tab_swmg, text='成员预览').grid(row=1, column=0, pady=10)
    text_swmg_preview = tk.Text(tab_swmg, width=28, height=5, state='normal')
    text_swmg_preview.grid(row=1, column=1, rowspan=3, sticky='NW', pady=10)
    text_swmg_preview.insert(tk.END, get_swmg_preview(cb_swmg_week.get()))
    text_swmg_preview.configure(state="disabled")

    def btn_swmg_callback():
        if judge_swmg_empty(cb_swmg_week.get()):
            mBox.showerror('错误', '周报文件夹为空，拒绝执行')
            return
        if not judge_swmg_group_complete(cb_swmg_week.get()):
            if global_config.safemode:
                mBox.showerror('错误', '已汇总的小组名单与当前小组名单不匹配，因开启安全模式，拒绝执行')
                return
            else:
                if not mBox.askyesno('警告', '已汇总的小组名单与当前小组名单不匹配，是否继续执行'):
                    return
        if judge_swmg_result_exist(cb_swmg_week.get()):
            if global_config.safemode:
                mBox.showerror('错误', '目标文件已存在，因开启安全模式，拒绝执行')
                return
            else:
                if not mBox.askyesno('警告', '目标文件已存在，是否覆盖现有文件'):
                    return
        distpath = exec_swmg_merge(cb_swmg_week.get())
        mBox.showinfo('提示', '已汇总至：\n' + distpath)
    btn_swmg_confirm = ttk.Button(
        tab_swmg, text="执行", command=btn_swmg_callback)
    btn_swmg_confirm.grid(row=4, column=0, columnspan=2, pady=20)

    # ----------------- tab 3 tab_mw ----------------
    def cb_mw_start_week_callback(event):
        start_week = cb_mw_start_week.get()
        end_week = cb_mw_end_week.get()
        reverse_week = get_swmg_week_list(True)
        cb_mw_end_week['value'] = reverse_week[:reverse_week.index(
            start_week) + 1]
        if end_week > start_week:
            cb_mw_end_week.current(cb_mw_end_week['value'].index(end_week))
        else:
            cb_mw_end_week.current(cb_mw_end_week['value'].index(start_week))
        label_week_count.configure(
            text=str(get_week_count(start_week, cb_mw_end_week.get())))

    def cb_mw_end_week_callback(event):
        label_week_count.configure(text=str(get_week_count(
            cb_mw_start_week.get(), cb_mw_end_week.get())))

    ttk.Label(tab_mw, text='开始周次（升序）').grid(row=0, column=0, padx=10, pady=10)
    cb_mw_start_week = ttk.Combobox(
        tab_mw, state="readonly", values=get_swmg_week_list(False), width=14)
    if len(cb_mw_start_week['value']) > 0:
        cb_mw_start_week.current(0)
    else:
        cb_mw_start_week.set('')
        cb_mw_start_week.configure(state="disabled")
    cb_mw_start_week.grid(row=0, column=1, sticky=tk.W)
    cb_mw_start_week.bind("<<ComboboxSelected>>", cb_mw_start_week_callback)

    ttk.Label(tab_mw, text='结束周次（降序）').grid(row=1, column=0, padx=10, pady=10)
    cb_mw_end_week = ttk.Combobox(
        tab_mw, state="readonly", values=get_swmg_week_list(True), width=14)
    if len(cb_mw_end_week['value']) > 0:
        cb_mw_end_week.current(0)
    else:
        cb_mw_end_week.set('')
        cb_mw_end_week.configure(state="disabled")
    cb_mw_end_week.grid(row=1, column=1, sticky=tk.W)
    cb_mw_end_week.bind("<<ComboboxSelected>>", cb_mw_end_week_callback)

    ttk.Label(tab_mw, text='总计周数').grid(
        row=2, column=0, padx=10, pady=10, sticky=tk.W)
    label_week_count = ttk.Label(tab_mw)
    if len(cb_mw_start_week['value']) > 0:
        label_week_count.configure(text=str(get_week_count(
            cb_mw_start_week.get(), cb_mw_end_week.get())))
    else:
        label_week_count.configure(text="0")
    label_week_count.grid(row=2, column=1, sticky=tk.W)

    def btn_mw_callback():
        if judge_mw_result_exist(cb_mw_start_week.get(), cb_mw_end_week.get()):
            if global_config.safemode:
                mBox.showerror('错误', '目标文件已存在，因开启安全模式，拒绝执行')
                return
            else:
                if not mBox.askyesno('警告', '目标文件已存在，是否覆盖现有文件'):
                    return
        distpath = exec_mw_merge(cb_mw_start_week.get(), cb_mw_end_week.get())
        mBox.showinfo('提示', '已汇总至：\n' + distpath)

    btn_mw_confirm = ttk.Button(
        tab_mw, text="执行", command=btn_mw_callback)
    btn_mw_confirm.grid(row=3, column=0, columnspan=2, pady=20)
    # # ----------------- log ------------------
    # log_frame = tk.Frame(win)
    # m_text = tk.Text(log_frame)
    # m_text.pack()
    # log_frame.pack(padx=10, pady=10)

    main_frame.pack(padx=10, pady=10)
