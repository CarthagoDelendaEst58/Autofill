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
import abcam, collected_excel, collected_with_sku
import gevent.monkey
# 先执行这步再往下导入，不然出错
gevent.monkey.patch_all()
import queue
import Save_Excel
import os

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
    start_button['state'] = tk.NORMAL
    # save_button['state'] = tk.NORMAL

def remove_file(dir_name):
    path = "./{}/".format(dir_name)
    for i in os.listdir(path):
        try:
            os.remove(path + i)
        except:
            print('cannot remove file')

def save():
    html_name_list = []
    html_name_list.append("Abcam")

    for html_name in html_name_list:
        Save_Excel.main(html_name)
    # messagebox.showinfo("info", "Generate excel completed！")

def del_excel():
    html_name_list = []
    html_name_list.append("Abcam")
    for html_name in html_name_list:
        path = "./{}.xls".format(html_name)
        if os.path.exists(path):
            os.remove(path)

def task1(name, search_list):
    remove_file(name)
    for search_word in search_list:
        with open('Abcam/'+str(search_word)+'.json', 'w') as f:
            abcam.main(search_word)
            time.sleep(1)


def run():
    search_list = read_excel(path.get())
    t2 = threading.Thread(target=task1,args=("Abcam",search_list,))
    t2.start()
    t2.join()

    collect()

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


def collect1():
    del_excel()
    save()
    collected_with_sku.main(path.get())
    messagebox.showinfo("info", "Collected with sku completed！")

def collect():
    del_excel()
    save()
    collected_excel.main()
    messagebox.showinfo("info", "Collected excel completed！")

if __name__ == '__main__':
    # multiprocessing.freeze_support()  # 在Windows下编译需要加这行
    # 创建tkinter对象
    window = tk.Tk()
    # 设置窗口标题
    window.title("Abcam-Spider")
    # 设置窗口大小
    window.geometry("420x280")
    # 放置控件

    path = tk.StringVar()
    tk.Label(window, text='Search file: ').place(x=30, y=40)
    tk.Entry(window, textvariable=path, width=30).place(x=110, y=40)
    tk.Button(window, text="Path", command=open_file).place(x=330, y=40)

    # CheckVar1 = tk.IntVar()
    # CheckVar2 = tk.IntVar()
    # CheckVar3 = tk.IntVar()
    # C1 = tk.Checkbutton(window, text="AlfaAesar", variable=CheckVar1, onvalue=1, offvalue=0)
    # C2 = tk.Checkbutton(window, text="Abcam", variable=CheckVar2,onvalue=1, offvalue=0)
    # C3 = tk.Checkbutton(window, text="ThermoFisher", variable=CheckVar3, onvalue=1, offvalue=0)
    # C1.place(x=30, y=60)
    # C2.place(x=185, y=60)
    # C3.place(x=265, y=60)


    start_button = tk.Button(window, text="Start", width=15, command=lambda :thread_it(run), state=tk.NORMAL)
    start_button.place(x=155, y=100)
    save_button = tk.Button(window, text="Collected with sku", width=15, command=collect)
    save_button.place(x=155, y=140)

    save1_button = tk.Button(window, text="Collected excel", width=15, command=collect1)
    save1_button.place(x=155, y=180)
    tk.Label(window, text="PS:  If you modify the json file,\nyou can use this button to regenerate excel\n(need to copy the selected record)", state=tk.DISABLED).place(x=90, y=210)
    # 主窗口循环显示
    window.mainloop()
