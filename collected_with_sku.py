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
# 2.1.2
# def main():
#     e_list  = ["AlfaAesar.xls","ThermoFisher.xls","Abcam.xls"]
#     e_list1  = []
#     for i in e_list:
#         if os.path.exists("./" + i):
#             e_list1.append(i)
#     df_list = []
#     for l in range(0,len(e_list1)):
#         e_name = e_list1[l]
#         df_list.insert(l, e_list1[l][:-4]+"company")
#         cas_cmb = pd.concat(pd.read_excel("./"+e_name ,sheet_name = None), ignore_index = True)
#         cas_cmb = cas_cmb.dropna(how = 'all')
#         cas_cmb = cas_cmb.reset_index(drop = True)
#         cas_cmb = cas_cmb.fillna(method = "ffill")
#         for i in range(0,cas_cmb['product_size'].count()):
#             j = cas_cmb['product_size'][i]
#             if pd.isna(j):
#                 unit = ''
#             else:
#                 unit = j[-2:]
#             if unit == "LB":
#                 number = pd.to_numeric(j[:-2],errors = "ignore")
#                 new_num = round(number * 0.453592,1)
#                 cas_cmb['product_size'][i] = str(new_num) + ' KG'
#             elif unit == "EA" or "ea":
#                 cas_cmb.drop([i])
#                 cas_cmb = cas_cmb.reset_index(drop = True)
#         cas_cmb['company'] = pd.Series([e_name[:-4] for x in range(len(cas_cmb.index))])
#         col_with_num = [col for col in cas_cmb.columns if "number" in col]                 # find the column with \number\,
#         if col_with_num != []:
#             cas_cmb.rename(columns = {col_with_num[0]: "catalog_number/product_stock_number"},inplace = True)   #change the name with \number\ to same column name,
#         try:
#             df_list[l] = cas_cmb
#         except:
#             print(l)
#             # print(cas_cmb)
#             # print(df_list)
#         df_list[l] = cas_cmb
#
#     if len(df_list) > 0:
#         final_cmb = pd.concat(df_list)
#         # final_cmb = final_cmb.sort_values(by=['product_name'])
#         final_cmb = final_cmb.sort_values(by=['search_name'])
#
#         # 使用apply函数, 如果search_name字段包含'***'关键词，则'mp_sku'这一列赋值为1,否则为0
#         # final_cmb['mp_sku'] = final_cmb.search_name.apply(lambda x: 1 if 'API' in x else 0)
#         df_dict = read_excel()
#         final_cmb['mp_sku'] = final_cmb["search_name"].apply(add, **{"df_dict":df_dict})
#
#         final_cmb.to_excel("./collected_data.xlsx")

#2.1.3
def main(path):
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
            cas_cmb = cas_cmb.dropna(how='all')
            cas_cmb = cas_cmb.reset_index(drop=True)
            # cas_cmb = cas_cmb.fillna(method="ffill")
            cas_cmb = cas_cmb.groupby('product_name')[cas_cmb.columns].fillna(method="ffill")
            cas_cmb['company'] = pd.Series([e_name[:-4] for x in range(len(cas_cmb.index))])
            col_with_num = [col for col in cas_cmb.columns if "number" in col]  # find the column with "number"
            cas_cmb.rename(columns={col_with_num[0]: "Clone number/product_number/catalog_number/SKU"},
                           inplace=True)  # change the name with "number" to same column name
            df_list[l] = cas_cmb
    final_cmb = pd.concat(df_list, axis=0)
    final_cmb = final_cmb.sort_values(by=["search_name","product_name","product_size","company"])
    final_cmb = final_cmb.reset_index(drop=True)

    # 添加mp_sku列
    # 使用apply函数, 如果search_name字段包含'***'关键词，则'mp_sku'这一列赋值为1,否则为0
    # final_cmb['mp_sku'] = final_cmb.search_name.apply(lambda x: 1 if 'API' in x else 0)
    # df_dict = read_excel(path)
    df_dict = read_excel('./Example SKU.xlsx')
    final_cmb['mp_sku'] = final_cmb["search_name"].apply(add, **{"df_dict":df_dict})


    cas_sku_des = pd.read_excel('CAS_SKU_DES_NAME.xlsx')
    final_cmb = final_cmb.reset_index(drop = True)
    for k in range(3, len(cas_sku_des.columns)):
        final_cmb['mp_' + cas_sku_des.columns[k]] = pd.Series([0 for x in range(len(final_cmb.index))])
    final_cmb['mp_'+ cas_sku_des.columns[3]] = final_cmb["mp_sku"].apply(lambda x: 1 if x is not None else 0)
    for i in range(len(final_cmb.index)):
        if final_cmb['mp_'+ cas_sku_des.columns[3]][i]== 1 :
            print(i)
            for j in range(len(cas_sku_des.index)):
                if final_cmb["mp_sku"][i] == cas_sku_des["sku"][j]:
                    for k in range(3, len(cas_sku_des.columns)):
                        final_cmb['mp_' + cas_sku_des.columns[k]][i] = cas_sku_des[cas_sku_des.columns[k]][j]
                        print(i, final_cmb["mp_"+ cas_sku_des.columns[k]][i])
        else:
            continue
    for h in range(len(final_cmb.columns)):
        final_cmb[final_cmb.columns[h]].loc[final_cmb[final_cmb.columns[h]] == 0] = np.nan
    final_cmb.to_excel("collected_data_with_sku.xlsx")


def add(x, df_dict):
    for i in df_dict["data"]:
        if x == i["Example Key words"]:
            df_dict["data"].remove(i)
            print(i["mp_sku"])
            return i["mp_sku"]

def read_excel(path):

    # 创建最终返回的空字典
    df_dict = {}
    # 读取Excel文件
    # 设置读取某列数据类型
    dtype = {
        'mp_sku': str,
    }
    # sheetName = "Example keyword with sku"
    # df = pd.read_excel(path, dtype=dtype, sheet_name=sheetName)

    df = pd.read_excel(path, dtype=dtype)

    # 替换Excel表格内的空单元格，否则在下一步处理中将会报错
    df.fillna("", inplace=True)

    df_list = []
    for i in df.index.values:
        # loc为按列名索引 iloc 为按位置索引，使用的是 [[行号], [列名]]
        df_line = df.loc[i, ['mp_sku', 'Example Key words ',]].to_dict()
        # 将每一行转换成字典后添加到列表
        df_list.append(df_line)
    df_dict['data'] = df_list
    return df_dict

# if __name__ == '__main__':
#     main('./Example SKU.xlsx')