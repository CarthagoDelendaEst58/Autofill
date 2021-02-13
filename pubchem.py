#!/usr/bin/python
# coding=utf-8
# ----------------------
# @Time    : 2020/7/24 16:07
# @Author  : hwf
# @File    : abcam.py
# ----------------------
import requests
import json
import random
import jsonpath

class Pubchem:
    def __init__(self):
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
            "User-Agent": random.choice(user_agent_list),
        }
        self.data_json = []
        # with open('./Pubchem/result.json', 'w', encoding='utf-8') as w:
        #     w.write('[\n')

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

        return json.loads(response.text)

    def get_node(self, parent_node, node_char):
        c = '$.Section[?(@.TOCHeading == "'+ node_char +'")]'
        node = jsonpath.jsonpath(parent_node[0], '$.Section[?(@.TOCHeading == "'+ node_char +'")]')
        def func(i):
            try:
                result = str(i["Value"]["StringWithMarkup"][0]["String"])
            except:
                c_node = jsonpath.jsonpath(i, '$.Value.Unit')
                result = str(i["Value"]["Number"][0]) + i["Value"]["Unit"] if c_node != False else str(i["Value"]["Number"][0])
            return result
        # Physical_Description = '\n'.join([i["Value"]["StringWithMarkup"][0]["String"] for i in node[0]["Information"]]) if node != False else ''

        node_char_result = '\n'.join([func(i) for i in node[0]["Information"]]) if node != False else ''

        return node_char_result

    def get_data(self, cid_list, search_name):
        for cid in cid_list:
            a = {}
            url3 = "https://pubchem.ncbi.nlm.nih.gov/rest/pug_view/data/compound/{}/JSON/".format(cid)
            data3 = self.parse_url(url3)

            node1 = jsonpath.jsonpath(data3,'$.Record.Section[?(@.TOCHeading == "Chemical and Physical Properties")]')
            if node1 != False:
                node1_1 = jsonpath.jsonpath(node1[0],'$.Section[?(@.TOCHeading == "Computed Properties")]')
                if node1_1 != False:
                    node1_1_1 = jsonpath.jsonpath(node1_1[0],'$.Section[?(@.TOCHeading == "Molecular Weight")]')
                    a["Molecular Weight"] = str(node1_1_1[0]["Information"][0]["Value"]["Number"][0]) + node1_1_1[0]["Information"][0]["Value"]["Unit"] if node1_1_1 != False else ''
                    
                    node1_1_2 = jsonpath.jsonpath(node1_1[0], '$.Section[?(@.TOCHeading == "Monoisotopic Mass")]')
                    a["Monoisotopic Mass"] = str(node1_1_2[0]["Information"][0]["Value"]["Number"][0]) + node1_1_2[0]["Information"][0]["Value"]["Unit"] if node1_1_2 != False else ''
                else:
                    a["Molecular Weight"], a["Monoisotopic Mass"] = '', ''

                node1_2 = jsonpath.jsonpath(node1[0], '$.Section[?(@.TOCHeading == "Experimental Properties")]')

                # if node1_2 != False:
                #     node1_2_1 = jsonpath.jsonpath(node1_2[0],'$.Section[?(@.TOCHeading == "Physical Description")]')
                #     Physical_Description = '\n'.join([i["Value"]["StringWithMarkup"][0]["String"] for i in node1_2_1[0]["Information"]]) if node1_2_1 != False else ''
                #     a["Physical Description"] = Physical_Description if Physical_Description != '\n' else ''
                #
                #     node1_2_2 = jsonpath.jsonpath(node1_2[0], '$.Section[?(@.TOCHeading == "Color/Form")]')
                #     Color_Form = '\n'.join([i["Value"]["StringWithMarkup"][0]["String"] for i in node1_2_2[0]["Information"]]) if node1_2_2 != False else ''
                #     a["Color/Form"] = Color_Form if Color_Form != '\n' else ''
                #
                #     node1_2_3 = jsonpath.jsonpath(node1_2[0], '$.Section[?(@.TOCHeading == "Boiling Point")]')
                #     Boiling_Point = '\n'.join([i["Value"]["StringWithMarkup"][0]["String"] for i in node1_2_3[0]["Information"]]) if node1_2_3 != False else ''
                #     a["Boiling Point"] = Boiling_Point if Boiling_Point != '\n' else ''
                #
                #     node1_2_4 = jsonpath.jsonpath(node1_2[0], '$.Section[?(@.TOCHeading == "Melting Point")]')
                #     Melting_Point = '\n'.join([i["Value"]["StringWithMarkup"][0]["String"] for i in node1_2_4[0]["Information"]]) if node1_2_4 != False else ''
                #     a["Melting Point"] = Melting_Point if Melting_Point != '\n' else ''
                #
                #     node1_2_5 = jsonpath.jsonpath(node1_2[0], '$.Section[?(@.TOCHeading == "Density")]')
                #     Density = '\n'.join([i["Value"]["StringWithMarkup"][0]["String"] for i in node1_2_5[0]["Information"]]) if node1_2_5 != False else ''
                #     a["Density"] = Density if Density != '\n' else ''
                #
                #     node1_2_6 = jsonpath.jsonpath(node1_2[0], '$.Section[?(@.TOCHeading == "LogP")]')
                #     LogP = '\n'.join([i["Value"]["StringWithMarkup"][0]["String"] for i in node1_2_6[0]["Information"]]) if node1_2_6 != False else ''
                #     a["LogP"] = LogP if LogP != '\n' else ''
                if node1_2 != False:
                    Physical_Description = self.get_node(node1_2, "Physical Description")
                    a["Physical Description"] = Physical_Description if Physical_Description != '\n' else ''

                    Color_Form = self.get_node(node1_2, "Color/Form")
                    a["Color/Form"] = Color_Form if Color_Form != '\n' else ''

                    Boiling_Point = self.get_node(node1_2, "Boiling Point")
                    a["Boiling Point"] = Boiling_Point if Boiling_Point != '\n' else ''


                    Melting_Point = self.get_node(node1_2, "Melting Point")
                    a["Melting Point"] = Melting_Point if Melting_Point != '\n' else ''


                    Density = self.get_node(node1_2, "Density")
                    a["Density"] = Density if Density != '\n' else ''


                    LogP = self.get_node(node1_2, "LogP")
                    a["LogP"] = LogP if LogP != '\n' else ''

                else:
                    a["Physical Description"], a["Color/Form"], a["Boiling Point"], a["Melting Point"], a["Density"], a["LogP"] = '','','','','',''

            a["search_name"] = search_name
            a["cid"] = cid
            print(a)
            print('')
            if a != {}:
                self.data_json.append(a)

    def save_data(self, data):
        with open('./Pubchem/result.json', 'w', encoding='utf-8') as w:
            w.write(json.dumps(data, indent=4, ensure_ascii=False))
            # w.write(',\n')

    def start(self, serach_list):
        for serach in serach_list:
            print(serach)
            url1 = "https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{}/cids/JSON".format(serach)
            data1 = self.parse_url(url1)
            try:
                base_cid = data1["IdentifierList"]["CID"][0]
                cid_list = [base_cid]
            except:
                # 没有最优搜索，查看有无下方列表
                url2 = 'https://pubchem.ncbi.nlm.nih.gov/sdq/sdqagent.cgi?infmt=json&outfmt=json&query={"select":"*","collection":"compound","where":{"ands":[{"*":"'+ serach +'"}]},"order":["relevancescore,desc"],"start":1,"limit":10,"width":1000000,"listids":0}'
                data2 = self.parse_url(url2)
                if len(data2["SDQOutputSet"][0]["rows"]) > 0:
                    cid_list = [product["cid"] for product in data2["SDQOutputSet"][0]["rows"][:5]]
                else:
                    # 没有搜索结果
                    continue

            self.get_data(cid_list, serach)
            self.save_data(self.data_json)


"https://pubchem.ncbi.nlm.nih.gov/"
'''
92-71-7
1344-28-1
7631-86-9
1344-28-1
53332-27-7
N-Acetyl-D-galactosamine
Adenosine 
'''
# if __name__ == '__main__':
#     # search_name = "92-71-7"
#     search_name = "1344-28-1"
#
#     # 商品详情页url
#     cid = '7105'
#     url = "https://pubchem.ncbi.nlm.nih.gov/compound/{}".format(cid)
#
#     # 最优搜索url（获取最优的cid）
#     url1 = "https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/N-Acetyl-D-galactosamine/cids/JSON"
#
#     # 最优搜索商品详情
#     url2 = '''https://pubchem.ncbi.nlm.nih.gov/sdq/sdqagent.cgi?infmt=json&outfmt=json&query={"select":"*","collection":"compound","where":{"ands":[{"cid":"7105"}]},"order":["cid,asc"],"start":1,"limit":10,"width":1000000,"listids":0}'''
#
#
#     # 下方商品列表各商品详情
#     url3 = '''https://pubchem.ncbi.nlm.nih.gov/sdq/sdqagent.cgi?infmt=json&outfmt=json&query={"select":"*","collection":"compound","where":{"ands":[{"*":"92-71-7"}]},"order":["relevancescore,desc"],"start":1,"limit":10,"width":1000000,"listids":0}'''
#
#     # 商品详情页数据
#     url4 = "https://pubchem.ncbi.nlm.nih.gov/rest/pug_view/data/compound/{}/JSON/".format(cid)
#     # 还有一种情况是“ https://pubchem.ncbi.nlm.nih.gov/rest/pug_view/data/patent/WO-02059742-A1/JSON/ ”
#
#     # 获取下方列表各个分类的数量（用于判断有无搜索结果）
#     url5 = '''https://pubchem.ncbi.nlm.nih.gov/sdq/sdqagent.cgi?infmt=json&outfmt=json&query={"hide":["*"],"collection":"compound,substance,bioassay,gene,protein,pathway,pubmed,patent","where":{"ands":[{"*":"123-123-123"}]}}'''
#
#     '''
#     判断有无搜索结果
#         有(依据：有最优搜索或者有下方商品列表)：
#             url1判断有无最优搜索
#                 有：
#                     url4获取最终数据
#                 无：
#                     url3获取前5个商品的最终数据
#         无：
#             pass
#     '''
#
#     "123-123-123-123"
#     url_test1 = '''https://pubchem.ncbi.nlm.nih.gov/sdq/sdqagent.cgi?infmt=json&outfmt=json&query={"select":"*","collection":"compound","where":{"ands":[{"*":"N-Acetyl-D-galactosamine"}]},"order":["relevancescore,desc"],"start":1,"limit":10,"width":1000000,"listids":0}'''
#     print(requests.get(url_test1).text)

def main(serach_list):
    pubchem = Pubchem()
    pubchem.start(serach_list)

if __name__ == '__main__':
    pubchem = Pubchem()
    serach_list = ["92-71-7","57-27-2","1344-28-1","7631-86-9","53332-27-7","N-Acetyl-D-galactosamine","Adenosine"]
    pubchem.start(serach_list)
