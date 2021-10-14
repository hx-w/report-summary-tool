#!/usr/bin/env python3
 
import sys
import tkinter
import tkinter.ttk
import pprint
 
## 0: grid geometry
## 1: pack geometry
bgrid = True
 
def guiInit():
    col = []
    datav = []
    data = {}
 
    for i in range(1, 100):
        col.append("col" + str(i))
        datav.append("data" + str(i) )
 
    for i in range(1, 100):
        data["line"+str(i)] = datav
 
    top = tkinter.Tk()
    # 窗口大小和位置 WxH+X+Y
    # 窗口大小 宽800, 高480
    # 窗口位置 距左边200, 距上边100
    top.geometry('800x480+200+100')
    tree = tkinter.ttk.Treeview(top, columns = col)
 
    tree.heading('#0', text="city")
    for colv in col:
        tree.column(colv, width=150, anchor='w')
        tree.heading(colv, text=colv)
 
    for k,v in data.items():
        itm = tree.insert("", "end", text=k, values=v)
        #pprint.pprint(tree.item(itm))
 
    yscrollbar = tkinter.ttk.Scrollbar(top, orient=tkinter.VERTICAL, command=tree.yview)
    xscrollbar = tkinter.ttk.Scrollbar(top, orient=tkinter.HORIZONTAL, command=tree.xview)
 
    tree.configure(yscrollcommand = yscrollbar.set)
    tree.configure(xscrollcommand = xscrollbar.set)
 
    if bgrid:
        ## 1. 设置位置
        tree.grid(row=0, column=0, sticky='nsew')
        yscrollbar.grid(row=0, column=1, sticky='nsew')
        xscrollbar.grid(row=1, column=0, sticky='nsew')
 
        ## 2. 设置大小
        ## tree@(0, 0):r.weight=1, c.weight=1
        ## 表格, row.weight = 1 高度随着窗口动态变化
        ## 表格, column.weight = 1 宽度随着窗口动态变化
        top.rowconfigure(0, weight=1)
        top.columnconfigure(0, weight=1)
 
        ## xball@(1, 0): r.weight=0, c.weight=1
        ## 横向滚动条, row.weight = 0 高度不变
        ## 横向滚动条, column.weight = 1 宽度随着窗口动态变化
        top.rowconfigure(1, weight=0)
        ## yball@(0, 1):r.weight=1, c.weight=0
        ## 竖向滚动条, row.weight = 1 高度随着窗口动态变化
        ## 竖向滚动条, column.weight = 0 宽度不变
        top.columnconfigure(1, weight=0)
    else:
        yscrollbar.pack(side='right', fill='y')
        xscrollbar.pack(side='bottom', fill='x')
        tree.pack(side='top', expand = 1, fill = 'both')
    top.mainloop()
 
def main():
    guiInit()
 
if __name__ == '__main__' :
    main()