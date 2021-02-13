#!/usr/bin/python
# coding=utf-8
# ----------------------
# @Time    : 2020/7/21 22:01
# @Author  : hwf
# @File    : Save_Excel.py
# ----------------------
import xlwt
import os
import yaml

class SaveExcel:

    def create_book(self):
        # 创建workbook对象
        book = xlwt.Workbook(encoding='utf-8', style_compression=0)
        data_dict = self.read_json()
        sheet = book.add_sheet("result", cell_overwrite_ok=True)
        self.new_write_data(sheet, data_dict)
        try:
            book.save('result.xls')
        except:
            print('Pubchem save error')


    def read_json(self):
        with open('./Pubchem/result.json', 'r', encoding='utf-8') as r:
            # data_dict = json.load(r)
            a = r.read().encode()
            b = a.replace(b'\xc2\x99', b'\x2d').decode()
            data_dict = yaml.load(b, Loader=yaml.FullLoader)
        return data_dict

    def new_write_data(self, sheet, data_dict):
        first_row_list = ["search_name", "cid", "Molecular Weight", "Monoisotopic Mass", "Physical Description", "Color/Form", "Boiling Point", "Melting Point", "Density", "LogP"]
        for i in first_row_list:
            sheet.write(0, first_row_list.index(i), i, self.cell_style())

        for index, product in enumerate(data_dict):
            for j in first_row_list:
                sheet.write(index + 1, first_row_list.index(j), product[j])

    def cell_style(self):
        font = xlwt.Font()
        font.bold = True
        style = xlwt.XFStyle()
        style.font = font
        return style

def main():
    save_excel = SaveExcel()
    save_excel.create_book()

if __name__ == '__main__':
    main()