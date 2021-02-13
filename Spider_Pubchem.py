#!/usr/bin/python
# coding=utf-8
# ----------------------
# @Time    : 2020/7/20 11:54
# @Author  : hwf
# @File    : Spectrum-Chemical.py
# ----------------------
import time
import threading
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import xlrd
import pubchem
import Save_Excel_Pubchem

def thread_it(func, *args):
    '''将函数打包进线程,解决运行耗时程序是窗口未响应情况'''
    # 创建
    t = threading.Thread(target=func, args=args)
    # 守护 !!!
    t.setDaemon(True)
    # 启动
    t.start()

    # 阻塞，界面未响应
    t.join()

    messagebox.showinfo("info", "Data acquisition completed！")

def run():
    search_list = read_excel(path.get())
    print(search_list)
    pubchem.main(search_list)
    Save_Excel_Pubchem.main()
    messagebox.showinfo("info", "Data acquisition completed！")

def runMerged(search_list):
    pubchem.main(search_list)
    Save_Excel_Pubchem.main()

def open_file():
    # 选择打开一个文件
    open_file_path = askopenfilename(title="Please select an Excel file to open",
                                     filetypes=[("Microsoft Excel", "*.xlsx"),
                                                ("Microsoft Excel 97-20003", "*.xls")])
    path.set(open_file_path)

def read_excel(excel_path):
    # 打开excel表格，获取搜索词
    data = xlrd.open_workbook(excel_path)
    # 获取表格
    table = data.sheet_by_name(data.sheet_names()[0])
    # 获取第一列所有数据
    a = table.col(0, start_rowx=0, end_rowx=None)
    search_list = [i.value for i in a]

    return list(set(search_list))

if __name__ == '__main__':
    # multiprocessing.freeze_support()  # 在Windows下编译需要加这行
    # 创建tkinter对象
    window = tk.Tk()
    # 设置窗口标题
    window.title("Pubchem-Spider")
    # 设置窗口大小
    window.geometry("420x160")
    # 放置控件

    path = tk.StringVar()
    tk.Label(window, text='Search file: ').place(x=30, y=40)
    tk.Entry(window, textvariable=path, width=30).place(x=110, y=40)
    tk.Button(window, text="Path", command=open_file).place(x=330, y=40)

    start_button = tk.Button(window, text="Start", width=15, command=run)
    start_button.place(x=155, y=100)
    # 主窗口循环显示
    window.mainloop()
