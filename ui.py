# ======================
# imports
# ======================
import tkinter as tk
from tkinter import Label, ttk
from tkinter import scrolledtext
from tkinter import Menu
from tkinter import Spinbox
from tkinter import messagebox as mBox

from script import *

win = tk.Tk()
win.title('周报汇总工具')


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

    def tab_changed_callback(event):
        tab = event.widget.tab('current')['text']
        if tab == '单周单组汇总':
            pass
        elif tab == '单周全组汇总':
            cb_swmg_week['value'] = get_swmg_week_list()
            text_swmg_preview.configure(state="normal")
            text_swmg_preview.delete(1.0, tk.END)
            if len(cb_swmg_week['value']) == 0:
                btn_swmg_confirm.configure(state="disabled")
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
            text_swmg_preview.insert(tk.END, get_swmg_preview(cb_swmg_week.get()))
            text_swmg_preview.configure(state="disabled")
        elif tab == '多周汇总':
            pass
    notebook.bind("<<NotebookTabChanged>>", tab_changed_callback)

    # ----------------- tab 1 tab_swsg ----------------
    ttk.Label(tab_swsg, text='选择小组').grid(row=0, column=0, padx=10, pady=10)

    def cb_swsg_group_callback(event):
        choose_group = cb_swsg_group.get()
        cb_swsg_week['value'] = get_swsg_week_list(choose_group)
        text_swsg_preview.configure(state="normal")
        text_swsg_preview.delete(1.0, tk.END)
        if len(cb_swsg_week['value']) > 0:
            cb_swsg_week.current(0)
            cb_swsg_week.configure(state="readonly")
            btn_swsg_confirm.configure(state="ready")
            text_swsg_preview.insert(tk.END, get_swsg_preview(cb_swsg_group.get(), cb_swsg_week.get()))
        else:
            cb_swsg_week.configure(state="disabled")
            btn_swsg_confirm.configure(state="disabled")
            text_swsg_preview.insert(tk.END, '<空>')
        text_swsg_preview.configure(state="disabled")

    def cb_swsg_week_callback(event):
        choose_week = cb_swsg_week.get()
        text_swsg_preview.configure(state="normal")
        text_swsg_preview.delete(1.0, tk.END)
        if judge_swsg_empty(cb_swsg_group.get(), choose_week):
            btn_swsg_confirm.configure(state="disabled")
            text_swsg_preview.insert(tk.END, '<空>')
        else:
            btn_swsg_confirm.configure(state="ready")
            text_swsg_preview.insert(tk.END, get_swsg_preview(cb_swsg_group.get(), cb_swsg_week.get()))
        text_swsg_preview.configure(state="disabled")

    cb_swsg_group = ttk.Combobox(tab_swsg, state="readonly")
    cb_swsg_group['value'] = global_config.group
    cb_swsg_group.grid(row=0, column=1)
    cb_swsg_group.current(0)
    cb_swsg_group.bind("<<ComboboxSelected>>", cb_swsg_group_callback)

    ttk.Label(tab_swsg, text='选择周次').grid(row=1, column=0)
    cb_swsg_week = ttk.Combobox(tab_swsg, state="readonly")
    cb_swsg_week['value'] = get_swsg_week_list(cb_swsg_group['value'][0])
    if len(cb_swsg_week['value']) > 0:
        cb_swsg_week.current(0)
    cb_swsg_week.grid(row=1, column=1, pady=10)
    cb_swsg_week.bind("<<ComboboxSelected>>", cb_swsg_week_callback)

    ttk.Label(tab_swsg, text='成员预览').grid(row=2, column=0, pady=10)
    text_swsg_preview = tk.Text(tab_swsg, width=28, height=5, state='normal')
    text_swsg_preview.grid(row=2, column=1, rowspan=3, sticky='NW', pady=10)
    text_swsg_preview.insert(tk.END, get_swsg_preview(cb_swsg_group.get(), cb_swsg_week.get()))
    text_swsg_preview.configure(state="disabled")

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

    btn_swsg_confirm = ttk.Button(tab_swsg, text="执行", command=btn_swsg_callback)
    btn_swsg_confirm.grid(row=5, column=0, columnspan=2, pady=20)

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
        # if judge_swsg_result_exist(cb_swsg_group.get(), cb_swsg_week.get()):
        #     if global_config.safemode:
        #         mBox.showerror('错误', '目标文件已存在，因开启安全模式，拒绝执行')
        #         return
        #     else:
        #         if not mBox.askyesno('警告', '目标文件已存在，是否覆盖现有文件'):
        #             return
        #         # mBox.showwarning('警告', '目标文件已存在，因关闭安全模式，继续执行')
        distpath = exec_swmg_merge(cb_swmg_week.get())
        mBox.showinfo('提示', '已汇总至：\n' + distpath)
    btn_swmg_confirm = ttk.Button(tab_swmg, text="执行", command=btn_swmg_callback)
    btn_swmg_confirm.grid(row=4, column=0, columnspan=2, pady=20)

    # ----------------- tab 3 tab_mw ----------------

    # # ----------------- log ------------------
    # log_frame = tk.Frame(win)
    # m_text = tk.Text(log_frame)
    # m_text.pack()
    # log_frame.pack(padx=10, pady=10)

    main_frame.pack(padx=10, pady=10)
