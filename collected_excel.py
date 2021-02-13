#!/usr/bin/python
# coding=utf-8
# ----------------------
# @Time    : 2020/8/15 13:59
# @Author  : hwf
# @File    : data_collect_with_sku.py
# ----------------------
import pandas as pd
import numpy as np
import os

#2.1.3
def main():
    e_list = ["Abcam.xls"]
    df_list = []
    for l in range(len(e_list)):
        e_name = e_list[l]
        try:
            cas_cmb = pd.concat(pd.read_excel(e_name, sheet_name=None), ignore_index=True)
        except FileNotFoundError:
            df_list.insert(l, e_name[:-4] + "company")
            df_list[l] = pd.DataFrame()
            print(e_name + " not found")
        else:
            df_list.insert(l, e_name[:-4] + "company")
            if e_name[:-4] == "Abcam":
                cas_cmb = cas_cmb.dropna(how="all")
                cas_cmb = cas_cmb.reset_index(drop=True)
                cas_cmb["sku number"] = np.nan
                cas_cmb[cas_cmb.columns[0]] = cas_cmb[cas_cmb.columns[0]].fillna(method="ffill")
                for i in range(len(cas_cmb.index)):
                    cas_cmb["sku number"][i] = cas_cmb["product_name"][i].split("(")[-1][0:-1]
                    cas_cmb["product_name"][i] = cas_cmb["product_name"][i][0:-(len(cas_cmb["sku number"][i]) + 2)]
            else:
                cas_cmb = cas_cmb
            # 删除所有值为空的行
            cas_cmb = cas_cmb.dropna(how='all')
            # 删除空行后重置索引
            cas_cmb = cas_cmb.reset_index(drop=True)
            # 分组填充
            cas_cmb = cas_cmb.groupby('product_name')[cas_cmb.columns].fillna(method="ffill")
            # cas_cmb = cas_cmb.groupby('product_name').ffill().groupby('product_name').bfill()
            cas_cmb['company'] = pd.Series([e_name[:-4] for x in range(len(cas_cmb.index))])
            col_with_num = [col for col in cas_cmb.columns if "number" in col]  # find the column with "number"
            cas_cmb.rename(columns={col_with_num[0]: "Clone number/product_number/catalog_number/SKU"},
                           inplace=True)  # change the name with "number" to same column name
            df_list[l] = cas_cmb
    final_cmb = pd.concat(df_list, axis=0)
    final_cmb = final_cmb.sort_values(by=["search_name","product_name","product_size","company"])
    final_cmb = final_cmb.reset_index(drop=True)
    final_cmb.to_excel("collected_data.xlsx")


def add(x, df_dict):
    for i in df_dict["data"]:
        if x == i["Example Key words"]:
            df_dict["data"].remove(i)
            print(i["mp_sku"])
            return i["mp_sku"]

def read_excel():

    # 创建最终返回的空字典
    df_dict = {}
    # 读取Excel文件
    # 设置读取某列数据类型
    dtype = {
        'mp_sku': str,
    }
    df = pd.read_excel('Example SKU.xlsx', dtype=dtype)

    # 替换Excel表格内的空单元格，否则在下一步处理中将会报错
    df.fillna("", inplace=True)

    df_list = []
    for i in df.index.values:
        # loc为按列名索引 iloc 为按位置索引，使用的是 [[行号], [列名]]
        df_line = df.loc[i, ['mp_sku', 'Example Key words',]].to_dict()
        # 将每一行转换成字典后添加到列表
        df_list.append(df_line)
    df_dict['data'] = df_list
    return df_dict

# if __name__ == '__main__':
#     main()