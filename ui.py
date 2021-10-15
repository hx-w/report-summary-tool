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

    # ----------------- tab 1 tab_swsg ----------------
    ttk.Label(tab_swsg, text='选择小组').grid(row=0, column=0, padx=10, pady=10)

    def cb_sisg_group_callback(event):
        choose_group = cb_sisg_group.get()
        logger.debug(choose_group)
        cb_sisg_week['value'] = get_week_list(choose_group)
        if len(cb_sisg_week['value']) > 0:
            cb_sisg_week.current(0)
            cb_sisg_week.configure(state="readonly")
            btn_confirm.configure(state="ready")
            label_sisg_preview.configure(text=get_sisg_preview(
                cb_sisg_group.get(), cb_sisg_week.get()))
        else:
            cb_sisg_week.configure(state="disabled")
            btn_confirm.configure(state="disabled")
            label_sisg_preview.configure(text='<空>')

    def cb_sisg_week_callback(event):
        choose_week = cb_sisg_week.get()
        if judge_sisg_empty(cb_sisg_group.get(), choose_week):
            btn_confirm.configure(state="disabled")
            label_sisg_preview.configure(text='<空>')
        else:
            btn_confirm.configure(state="ready")
            label_sisg_preview.configure(text=get_sisg_preview(
                cb_sisg_group.get(), cb_sisg_week.get()))

    cb_sisg_group = ttk.Combobox(tab_swsg, state="readonly")
    cb_sisg_group['value'] = global_config.group
    cb_sisg_group.grid(row=0, column=1)
    cb_sisg_group.current(0)
    cb_sisg_group.bind("<<ComboboxSelected>>", cb_sisg_group_callback)

    ttk.Label(tab_swsg, text='选择周次').grid(row=1, column=0)
    cb_sisg_week = ttk.Combobox(tab_swsg, state="readonly")
    cb_sisg_week['value'] = get_week_list(cb_sisg_group['value'][0])
    if len(cb_sisg_week['value']) > 0:
        cb_sisg_week.current(0)
    cb_sisg_week.grid(row=1, column=1, pady=10)
    cb_sisg_week.bind("<<ComboboxSelected>>", cb_sisg_week_callback)

    ttk.Label(tab_swsg, text='成员预览').grid(row=2, column=0, pady=10)
    label_sisg_preview = ttk.Label(tab_swsg, text=get_sisg_preview(
        cb_sisg_group.get(), cb_sisg_week.get()))
    label_sisg_preview.grid(row=2, column=1, rowspan=2, sticky='W', pady=10)

    def btn_swsg_callback():
        if judge_sisg_empty(cb_sisg_group.get(), cb_sisg_week.get()):
            mBox.showerror('错误', '小组周报文件夹为空，拒绝执行')
            return
        if judge_sisg_result_exist(cb_sisg_group.get(), cb_sisg_week.get()):
            if global_config.safemode:
                mBox.showerror('错误', '目标文件已存在，因开启安全模式，拒绝执行')
                return
            else:
                if not mBox.askyesno('警告', '目标文件已存在，是否覆盖现有文件'):
                    return
                # mBox.showwarning('警告', '目标文件已存在，因关闭安全模式，继续执行')
        distpath = exec_sisg_merge(cb_sisg_group.get(), cb_sisg_week.get())
        mBox.showinfo('提示', '已汇总至：\n' + distpath)

    btn_confirm = ttk.Button(tab_swsg, text="执行", command=btn_swsg_callback)
    btn_confirm.grid(row=4, column=0, columnspan=2, pady=20)

    # ----------------- tab 2 tab_swmg ----------------
    ttk.Label(tab_swmg, text='选择周次').grid(row=0, column=0, padx=10, pady=10)
    cb_simg_week = ttk.Combobox(tab_swmg, state="readonly")
    cb_simg_week['value'] = get_simg_week_list()

    if len(cb_sisg_week['value']) > 0:
        cb_simg_week.current(0)
    cb_simg_week.grid(row=0, column=1)

    # ----------------- tab 3 tab_mw ----------------

    # # ----------------- log ------------------
    # log_frame = tk.Frame(win)
    # m_text = tk.Text(log_frame)
    # m_text.pack()
    # log_frame.pack(padx=10, pady=10)

    main_frame.pack(padx=10, pady=10)
