#!/usr/bin/python
# coding=utf-8
# ----------------------
# @Time    : 2020/7/24 16:07
# @Author  : hwf
# @File    : abcam.py
# ----------------------
from gevent import monkey

monkey.patch_all()
# monkey的patch_all()能把程序变成协作式运行，可以帮助程序实现异步
# 一定要先执行这一步，再执行接下来的import，不然会报错

import gevent
import requests
from lxml import etree
import json
import random
from itertools import groupby

class Abcam:
    def __init__(self, search_name):
        user_agent_list = [
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1",
            "Mozilla/5.0 (X11; CrOS i686 2268.111.0) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.57 Safari/536.11",
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1092.0 Safari/536.6",
            "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1090.0 Safari/536.6",
            "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/19.77.34.5 Safari/537.1",
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.9 Safari/536.5",
            "Mozilla/5.0 (Windows NT 6.0) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.36 Safari/536.5",
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
            "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_0) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
            "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1062.0 Safari/536.3",
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1062.0 Safari/536.3",
            "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
            "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
            "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.0 Safari/536.3",
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24",
            "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24"
        ]
        self.headers = {
            # "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36",
            "User-Agent": random.choice(user_agent_list),
        }
        # self.session = requests.session()
        self.data_json = []
        self.search_name = search_name
        self.num_objects = 0
        with open('./Abcam/{}.json'.format(self.search_name), 'w', encoding='utf-8') as w:
            w.write('[\n')

    def parse_url(self, url):
        i = 1
        while i<6:
            try:
                response = requests.get(url, headers=self.headers, timeout=20)
                i = 6
            except:
                print("请求超时-{}".format(str(i)))
                i += 1
                print(url)
        # with gevent.Timeout(20, False) as t:
        #     response = requests.get(url, headers=self.headers)
        return response.content.decode()

    def parse_url1(self, url):
        headers = {
            "x-requested-with": "XMLHttpRequest"
        }
        headers.update(self.headers)
        i = 1
        while i<6:
            try:
                response = requests.get(url, headers=headers, timeout=20)
                i = 6
            except:
                print("请求超时-{}".format(str(i)))
                i += 1
                print(url)
        # with gevent.Timeout(20, False) as t:
        #     response = requests.get(url, headers=self.headers)
        return response.content.decode()

    # 协程任务函数
    def task(self, product):
        a = {}
        a["search_name"] = self.search_name
        # 产品描述
        product_name = product.xpath('.//@data-productname')[0] if len(
            product.xpath('.//@data-productname')) > 0 else ''
        a["product_name"] = product_name.replace('\xa0',' ')

        # 产品浓度
        a["purity"] = product.xpath('.//div[@class="clearfix pws_item Purity"]/div[@class="pws_value"]/text()')[
            0].replace('\xa0',' ') if len(
            product.xpath('.//div[@class="clearfix pws_item Purity"]/div[@class="pws_value"]/text()')) > 0 else ''

        # 获取产品编码
        data_productcode = product.xpath('.//@data-productcode')[0]
        abid = [''.join(list(g)) for k, g in groupby(data_productcode, key=lambda x: x.isdigit())]
        # 获取产品规格
        product_size_list = self.get_price(abid[1])

        # 产品url
        product_url = "https://www.abcam.com/" + product.xpath('.//div[@class="pws-item-info"]//h3/a/@href')[0]
        # 进入商品详情页获取其他11个字段
        self.get_other(product_url, a, abid[1])
        
        a["price_list"] = []
        for size in product_size_list:
            b = {}
            # 商品规格
            b["product_size"] = size["Size"].replace('&micro;', 'µ')
            # 商品价格
            b["product_price"] = size["Price"]
            a["price_list"].append(b)

        print(a)
        print('')
        self.save_data(a)

    def get_other(self, product_url, a, productID):
        response_str = self.parse_url(product_url)
        html_obj = etree.HTML(response_str)
        a["Buffer Requirements for Conjugation"] = html_obj.xpath('.//*[contains(text(), "Buffer Requirements for Conjugation")]/../following-sibling::*[1]/text()')[0].replace('\r\n','').strip().replace('\xa0',' ') if len(html_obj.xpath('.//*[contains(text(), "Buffer Requirements for Conjugation")]/../following-sibling::*[1]/text()')) > 0 else ''

        a["Clonality"] = html_obj.xpath('.//h3[contains(text(), "Clonality")]/following-sibling::*[1]/text()')[0].replace('\r\n','').strip().replace('\xa0',' ') if len(html_obj.xpath('.//h3[contains(text(), "Clonality")]/following-sibling::*[1]/text()')) > 0 else ''

        a["Clone number"] = html_obj.xpath('.//h3[contains(text(), "Clone number")]/following-sibling::*[1]/text()')[0].replace('\r\n','').strip().replace('\xa0',' ') if len(html_obj.xpath('.//h3[contains(text(), "Clone number")]/following-sibling::*[1]/text()')) > 0 else ''

        try:
            ConcentrationList = self.get_Concentration(productID)
            a["Concentration"] = '\n'.join([i.replace('&micro;','μ') for i in ConcentrationList])
        except:
            a["Concentration"] = ''
        # a["Concentration"] = '\n'.join([i.replace('\r\n','').strip().replace('\xa0',' ') for i in html_obj.xpath('.//div[@id="concentration-information"]//div[contains(text(), "Concentration")]/following-sibling::*[1]//li/text()') if i.replace('\r\n','').strip()!='']) if len(html_obj.xpath('.//div[@id="concentration-information"]//div[contains(text(), "Concentration")]/following-sibling::*[1]//li/text()')) > 0 else ''


        a["Function"] = html_obj.xpath('.//h3[contains(text(), "Function")]/following-sibling::*[1]/text()')[0].replace('\r\n','').strip().replace('\xa0',' ') if len(html_obj.xpath('.//h3[contains(text(), "Function")]/following-sibling::*[1]/text()')) > 0 else ''

        a["Host species"] = html_obj.xpath('.//h3[contains(text(), "Host species")]/following-sibling::*[1]/text()')[0].replace('\r\n','').strip().replace('\xa0',' ') if len(html_obj.xpath('.//h3[contains(text(), "Host species")]/following-sibling::*[1]/text()')) > 0 else ''

        a["Immunogen"] = '\n'.join([i.replace('\r\n','').strip().replace('\xa0',' ') for i in html_obj.xpath('.//h3[contains(text(), "Immunogen")]/following-sibling::*[1]//text()') if i.replace('\r\n','').strip()!='']) if len(html_obj.xpath('.//h3[contains(text(), "Immunogen")]/following-sibling::*[1]//text()')) > 0 else ''

        a["Isotype"] = html_obj.xpath('.//h3[contains(text(), "Isotype")]/following-sibling::*[1]//text()')[0].replace('\r\n','').strip().replace('\xa0',' ') if len(html_obj.xpath('.//h3[contains(text(), "Isotype")]/following-sibling::*[1]//text()')) > 0 else ''

        a["Light chain type"] = html_obj.xpath('.//h3[contains(text(), "Light chain type")]/following-sibling::*[1]//text()')[0].replace('\r\n','').strip().replace('\xa0',' ') if len(html_obj.xpath('.//h3[contains(text(), "Light chain type")]/following-sibling::*[1]//text()')) > 0 else ''

        a["Purity"] = html_obj.xpath('.//h3[contains(text(), "Purity")]/following-sibling::*[1]//text()')[0].replace('\r\n','').strip().replace('\xa0',' ') if len(html_obj.xpath('.//h3[contains(text(), "Purity")]/following-sibling::*[1]//text()')) > 0 else ''

        a["Species reactivity"] = '\n'.join([i.replace('\r\n','').strip().replace('\xa0',' ') for i in html_obj.xpath('.//h3[contains(text(), "Species reactivity")]/following-sibling::*[1]//text()') if i.replace('\r\n','').strip()!='']) if len(html_obj.xpath('.//h3[contains(text(), "Species reactivity")]/following-sibling::*[1]//text()')) > 0 else ''

    def only_one(self, html_obj):
        a = {}
        a["search_name"] = self.search_name
        # 产品描述
        product_name = html_obj.xpath('.//h1[@class="title"]/text()')[0] if len(
            html_obj.xpath('.//h1[@class="title"]/text()')) > 0 else ''
        a["product_name"] = product_name.replace('\xa0', ' ')

        # 产品浓度
        a["purity"] = ''

        # 获取产品编码
        data_productcode = html_obj.xpath('.//h1[@class="title"]/text()')[0].split('(')[-1].split(')')[0]
        abid = [''.join(list(g)) for k, g in groupby(data_productcode, key=lambda x: x.isdigit())]
        # 获取产品规格
        product_size_list = self.get_price(abid[1])

        a["Buffer Requirements for Conjugation"] = html_obj.xpath(
            './/*[contains(text(), "Buffer Requirements for Conjugation")]/../following-sibling::*[1]/text()')[
            0].replace('\r\n', '').strip().replace('\xa0', ' ') if len(html_obj.xpath(
            './/*[contains(text(), "Buffer Requirements for Conjugation")]/../following-sibling::*[1]/text()')) > 0 else ''

        a["Clonality"] = html_obj.xpath('.//h3[contains(text(), "Clonality")]/following-sibling::*[1]/text()')[
            0].replace('\r\n', '').strip().replace('\xa0', ' ') if len(
            html_obj.xpath('.//h3[contains(text(), "Clonality")]/following-sibling::*[1]/text()')) > 0 else ''

        a["Clone number"] = html_obj.xpath('.//h3[contains(text(), "Clone number")]/following-sibling::*[1]/text()')[
            0].replace('\r\n', '').strip().replace('\xa0', ' ') if len(
            html_obj.xpath('.//h3[contains(text(), "Clone number")]/following-sibling::*[1]/text()')) > 0 else ''

        try:
            ConcentrationList = self.get_Concentration(abid[1])
            a["Concentration"] = '\n'.join([i.replace('&micro;', 'μ') for i in ConcentrationList])
        except:
            a["Concentration"] = ''
        # a["Concentration"] = '\n'.join([i.replace('\r\n','').strip().replace('\xa0',' ') for i in html_obj.xpath('.//div[@id="concentration-information"]//div[contains(text(), "Concentration")]/following-sibling::*[1]//li/text()') if i.replace('\r\n','').strip()!='']) if len(html_obj.xpath('.//div[@id="concentration-information"]//div[contains(text(), "Concentration")]/following-sibling::*[1]//li/text()')) > 0 else ''

        a["Function"] = html_obj.xpath('.//h3[contains(text(), "Function")]/following-sibling::*[1]/text()')[0].replace(
            '\r\n', '').strip().replace('\xa0', ' ') if len(
            html_obj.xpath('.//h3[contains(text(), "Function")]/following-sibling::*[1]/text()')) > 0 else ''

        a["Host species"] = html_obj.xpath('.//h3[contains(text(), "Host species")]/following-sibling::*[1]/text()')[
            0].replace('\r\n', '').strip().replace('\xa0', ' ') if len(
            html_obj.xpath('.//h3[contains(text(), "Host species")]/following-sibling::*[1]/text()')) > 0 else ''

        a["Immunogen"] = '\n'.join([i.replace('\r\n', '').strip().replace('\xa0', ' ') for i in html_obj.xpath(
            './/h3[contains(text(), "Immunogen")]/following-sibling::*[1]//text()') if
                                    i.replace('\r\n', '').strip() != '']) if len(
            html_obj.xpath('.//h3[contains(text(), "Immunogen")]/following-sibling::*[1]//text()')) > 0 else ''

        a["Isotype"] = html_obj.xpath('.//h3[contains(text(), "Isotype")]/following-sibling::*[1]//text()')[0].replace(
            '\r\n', '').strip().replace('\xa0', ' ') if len(
            html_obj.xpath('.//h3[contains(text(), "Isotype")]/following-sibling::*[1]//text()')) > 0 else ''

        a["Light chain type"] = \
        html_obj.xpath('.//h3[contains(text(), "Light chain type")]/following-sibling::*[1]//text()')[0].replace('\r\n',
                                                                                                                 '').strip().replace(
            '\xa0', ' ') if len(
            html_obj.xpath('.//h3[contains(text(), "Light chain type")]/following-sibling::*[1]//text()')) > 0 else ''

        a["Purity"] = html_obj.xpath('.//h3[contains(text(), "Purity")]/following-sibling::*[1]//text()')[0].replace(
            '\r\n', '').strip().replace('\xa0', ' ') if len(
            html_obj.xpath('.//h3[contains(text(), "Purity")]/following-sibling::*[1]//text()')) > 0 else ''

        a["Species reactivity"] = '\n'.join([i.replace('\r\n', '').strip().replace('\xa0', ' ') for i in html_obj.xpath(
            './/h3[contains(text(), "Species reactivity")]/following-sibling::*[1]//text()') if
                                             i.replace('\r\n', '').strip() != '']) if len(
            html_obj.xpath('.//h3[contains(text(), "Species reactivity")]/following-sibling::*[1]//text()')) > 0 else ''

        a["price_list"] = []
        for size in product_size_list:
            b = {}
            # 商品规格
            b["product_size"] = size["Size"].replace('&micro;', 'µ')
            # 商品价格
            b["product_price"] = size["Price"]
            a["price_list"].append(b)

        print(a)
        print('')
        self.save_data(a)

    def get_price(self, abid):
        url = "https://www.abcam.com/datasheetproperties/availability?abId={}".format(abid)
        price_size_res = self.parse_url1(url)
        price_size_dict = json.loads(price_size_res)
        return price_size_dict["size-information"]["Sizes"]

    def get_Concentration(self, abid):
        url = "https://www.abcam.com/datasheetproperties/concentrations?productId={}".format(abid)
        price_size_res = self.parse_url1(url)
        price_size_dict = json.loads(price_size_res)
        return price_size_dict["Concentrations"]

    def get_product(self, product_list):
        # 协程任务列表
        task_list = []
        for product in product_list:
            # 协程作业
            task = gevent.spawn(self.task, product)
            task_list.append(task)
        gevent.joinall(task_list)

    def start(self, response_str):
        html_obj = etree.HTML(response_str)
        product_list = html_obj.xpath('.//div[@class="search_results"]/div')
        isonlyone = html_obj.xpath('.//h1[@class="title"]')
        if product_list == [] and isonlyone == []:
            with open('./Abcam/{}.json'.format(self.search_name), 'a', encoding='utf-8') as w:
                w.write(']')
        elif isonlyone != []:
            self.only_one(html_obj)
            with open('./Abcam/{}.json'.format(self.search_name), 'a', encoding='utf-8') as w:
                w.write(']')
        else:
            product_items = html_obj.xpath('.//div[@class="search_results"]//@data-total-items')[0]
            # import time
            # print(product_list)
            # time.sleep(1222)
            self.get_product(product_list)
            # 判断是否有下一页
            page_num = int(product_items) // 20 if int(product_items) % 20 == 0 else (int(product_items) // 20) +1
            if page_num > 1:
                for num in range(2, page_num+1):
                    next_page_url = "https://www.abcam.com/products/loadmore?keywords={}&pagenumber={}".format(self.search_name,str(num))
                    self.next_page(next_page_url)

            # 没有下一页了保存数据
            # self.save_data()
            with open('./Abcam/{}.json'.format(self.search_name), 'a', encoding='utf-8') as w:
                w.write(']')

    def next_page(self, next_page_url):
        self.headers.update({"x-requested-with": "XMLHttpRequest"})
        response_str = self.parse_url(next_page_url)
        html_obj = etree.HTML(response_str)
        product_list = html_obj.xpath('.//div[contains(@class,"selection-item")]')
        self.get_product(product_list)

    def save_data(self, a):
        with open('./Abcam/{}.json'.format(self.search_name), 'a', encoding='utf-8') as w:
            if self.num_objects > 0:
                w.write(',\n')
            w.write(json.dumps(a, indent=4, ensure_ascii=False))
            self.num_objects = self.num_objects + 1

def main(search_name):
    url = "https://www.abcam.com/products?keywords={}".format(search_name)
    s = Abcam(search_name)
    res = s.parse_url(url)
    s.start(res)


# if __name__ == '__main__':
#     # search_name = "92-71-7"
#     search_name = "1344-28-1"
#     url = "https://www.abcam.com/products?keywords={}".format(search_name)
#     s = Abcam(search_name)
#     res = s.parse_url(url)
#     # print(res)
#     s.start(res)