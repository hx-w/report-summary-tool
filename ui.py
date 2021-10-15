# ======================
# imports
# ======================
import tkinter as tk
from tkinter import ttk
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

    def cb_group_callback(event):
        choose_group = cb_group.get()
        logger.debug(choose_group)
        cb_week['value'] = get_week_list(choose_group)
        if len(cb_week['value']) > 0:
            cb_week.current(0)
            cb_week.configure(state="readonly")
            btn_confirm.configure(state="ready")
        else:
            cb_week.configure(state="disabled")
            btn_confirm.configure(state="disabled")

    cb_group = ttk.Combobox(tab_swsg, state="readonly")
    cb_group['value'] = global_config.group
    cb_group.grid(row=0, column=1)
    cb_group.current(0)
    cb_group.bind("<<ComboboxSelected>>", cb_group_callback)

    ttk.Label(tab_swsg, text='选择周次').grid(row=1, column=0)
    cb_week = ttk.Combobox(tab_swsg, state="readonly")
    cb_week['value'] = get_week_list(cb_group['value'][0])
    if len(cb_week['value']) > 0:
        cb_week.current(0)
    cb_week.grid(row=1, column=1)

    def btn_swsg_callback():
        if judge_sisg_empty(cb_group.get(), cb_week.get()):
            mBox.showerror('错误', '小组周报文件夹为空，拒绝执行')
            return
        if judge_sisg_result_exist(cb_group.get(), cb_week.get()):
            if global_config.safemode:
                mBox.showerror('错误', '目标文件已存在，因开启安全模式，拒绝执行')
                return
            else:
                if not mBox.askyesno('警告', '目标文件已存在，是否覆盖现有文件'):
                    return
                # mBox.showwarning('警告', '目标文件已存在，因关闭安全模式，继续执行')
        distpath = exec_sisg_merge(cb_group.get(), cb_week.get())
        mBox.showinfo('提示', '已汇总至：\n' + distpath)

    btn_confirm = ttk.Button(tab_swsg, text="执行", command=btn_swsg_callback)
    btn_confirm.grid(row=4, column=0, columnspan=2)
    # ----------------- tab 2 tab_swmg ----------------


    # ----------------- tab 3 tab_mw ----------------

    # # ----------------- log ------------------
    # log_frame = tk.Frame(win)
    # m_text = tk.Text(log_frame)
    # m_text.pack()
    # log_frame.pack(padx=10, pady=10)

    main_frame.pack(padx=10, pady=10)
