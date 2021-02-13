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
    def __init__(self, html_name):
        self.html_name = html_name

    def create_book(self):
        # 创建workbook对象
        book = xlwt.Workbook(encoding='utf-8', style_compression=0)
        for search_json_file in os.listdir('./{}'.format(self.html_name)):
            search_name = search_json_file.split('.')[0]
            sheet = book.add_sheet(search_name, cell_overwrite_ok=True)
            data_dict = self.read_json(search_json_file)
            # self.write_data(sheet, data_dict)
            self.new_write_data(sheet, data_dict)
        book.save('{}.xls'.format(self.html_name))


    def read_json(self, search_json_file):
        print('./{}/{}'.format(self.html_name, search_json_file))
        with open('./{}/{}'.format(self.html_name, search_json_file), 'r', encoding='utf-8') as r:
            # data_dict = json.load(r)
            a = r.read().encode()
            b = a.replace(b'\xc2\x99', b'\x2d').decode()
            data_dict = yaml.load(b, Loader=yaml.FullLoader)
        return data_dict


    def write_data(self, sheet, data_dict):
        if self.html_name == "Abcam":
            first_row_list = ["product_name", "purity", "product_size", "product_price"]
        else:
            first_row_list = ["product_name", "product_CAS", "product_desc", "product_size", "product_price"]
        for i in first_row_list:
            sheet.write(0, first_row_list.index(i), i, self.cell_style())

        if self.html_name == "ThermoFisher":
            price_list_len = 0
            for product in data_dict:
                for j in first_row_list:
                    if j == "product_name" or j == "product_CAS":
                        sheet.write(data_dict.index(product) + price_list_len + 1,
                                    first_row_list.index(j), product[j])
                for index1, product_desc in enumerate(product["product_desc_list"]):
                    for j in first_row_list:
                        if j != "product_name" and j != "product_CAS":
                            if j == "product_desc":
                                sheet.write(data_dict.index(product) + price_list_len + index1 + 1,
                                            first_row_list.index(j), product_desc[j])
                            else:
                                for index2, product_price in enumerate(product_desc["price_list"]):
                                    sheet.write(data_dict.index(product) + price_list_len + index1 + index2 + 1,
                                                first_row_list.index(j), product_price[j])

                    price_list_len += len(product_desc["price_list"])
                price_list_len += len(product["product_desc_list"]) - 1


        else:
            price_list_len = 0
            for product in data_dict:
                for j in first_row_list:
                    if j != "product_size" and j != "product_price":
                        sheet.write(data_dict.index(product)+ price_list_len +1, first_row_list.index(j), product[j])
                    else:
                        # for price in product["price_list"]:
                        for index, price in enumerate(product["price_list"]):
                            sheet.write(data_dict.index(product)+ price_list_len + index +1, first_row_list.index(j), price[j])

                price_list_len += len(product["price_list"])

    def new_write_data(self, sheet, data_dict):
        first_row_list = ["product_name", "purity", "product_size", "product_price", "search_name","Buffer Requirements for Conjugation","Clonality","Clone number","Concentration","Function","Host species","Immunogen","Isotype","Light chain type","Purity","Species reactivity"]
        for i in first_row_list:
            sheet.write(0, first_row_list.index(i), i, self.cell_style())

        price_list_len = 0
        for product in data_dict:
            for j in first_row_list:
                if j != "product_size" and j != "product_price" and j != "product_stock_number" and j != "catalog_number":
                    sheet.write(data_dict.index(product)+ price_list_len +1, first_row_list.index(j), product[j])
                else:
                    # for price in product["price_list"]:
                    for index, price in enumerate(product["price_list"]):
                        sheet.write(data_dict.index(product)+ price_list_len + index +1, first_row_list.index(j), price[j])

            price_list_len += len(product["price_list"])

    def cell_style(self):
        font = xlwt.Font()
        font.bold = True
        style = xlwt.XFStyle()
        style.font = font
        return style

def main(html_name):
    save_excel = SaveExcel(html_name)
    save_excel.create_book()

# if __name__ == '__main__':
#     html_name = "Abcam"
#     main(html_name)