# -*- coding: utf-8 -*-
# surpassing Feb-2016
# 指定搜索方式1-23页的所有内容, 每一页25家公司, 三种类型(supplier，customer，competitor), 25 * 3 * 23 = 1725个excel下载
# code version 2

"""
1. 可改进内容1: 控制请求周期
2. 可改进内容2: 判断每次搜索返回的数据排列顺序是否相同
"""

import requests
import re

page_dict = {}

def search_parse():

    # open every single page of the search result (23 page), parse company_id and company_name, save as list.
    global page_dict
    page_id = range(1, 24)

    for each_page_id in page_id:

        headers = {
            # "GET /advancedsearchresults.php?action=basic&page=1&sector=&industrytype=sic&industrycode=36&Index=null&Exchange=null&Country=USA HTTP/1.1"
            "Host": "www.mergentonline.com.proxy.lib.fsu.edu",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Upgrade-Insecure-Requests": "1",
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.109 Safari/537.36",
            "Referer": "http://www.mergentonline.com.proxy.lib.fsu.edu/basicsearch.php",
            "Accept-Encoding": "gzip, deflate, sdch",
            "Accept-Language": "zh-CN,zh;q=0.8,en;q=0.6",
            # fill it "Cookie": ""
            }

        r = requests.get(
            # "http://www.mergentonline.com.proxy.lib.fsu.edu/companyhorizon.php?pagetype=supplier&compnumber=100176&isexcel=1",
            "http://www.mergentonline.com.proxy.lib.fsu.edu/advancedsearchresults.php?action=basic&page=" + str(each_page_id) + "&sector=&industrytype=sic&industrycode=36&Index=null&Exchange=null&Country=USA&sort=ASC&sortcol=5",
            headers=headers
        )
        # print(r.status_code)

        my_content = r.content.decode("utf-8")  # decoding here for regex
        # print(my_content)

        print("正在尝试第: ", each_page_id, " 页")

        # regex find all company name and company id within the page
        comp_id_list = re.findall("value='([^<]+)' ></input>", my_content)
        comp_name_list = re.findall("=synopsis'>([^<]+)</a></td><td>", my_content)
        print("已获取到的该页公司ID及名称: ", comp_id_list, comp_name_list)

        # save data from each page to a dictionary
        page_dict[each_page_id] = [[comp_id_list], [comp_name_list]]
        # print(page_dict)

    # print(page_dict)
    # write all information to local file
    with open("23页搜索内容信息汇总", 'wb') as sum_output:
        sum_output.write(page_dict)

    print("获取信息保存成功")


def download_xls():
    # pay attention to the logic when iter

    comp_type_list = ["supplier", "customer", "competitor"]
    company_count = 0

    for k, v in page_dict.items():
        # page_dict[k] = [[comp_id_1, comp_id_2], [comp_name_1, comp_name_2]]
        # page_dict[k][0] = [comp_id_1, comp_id_2]
        # page_dict[k][1] = [comp_name_1, comp_name_2]]

        for each_comp_id in page_dict[k][0]:
            for each_comp_name in page_dict[k][1]:
                for each_comp_type in comp_type_list:

                    company_count += 1

                    # Download start
                    # Refer_url = "http://www.mergentonline.com.proxy.lib.fsu.edu/companyhorizon.php?pagetype=" + each_comp_type + "&compnumber=" + each_comp_id
                    request_url = "http://www.mergentonline.com.proxy.lib.fsu.edu/companyhorizon.php?pagetype=" + each_comp_type + "&compnumber=" + each_comp_id + "&isexcel=1"

                    headers ={
                        # "GET /companyhorizon.php?pagetype=supplier&compnumber=100176&isexcel=1 HTTP/1.1"
                        "Host": "www.mergentonline.com.proxy.lib.fsu.edu",
                        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
                        "Upgrade-Insecure-Requests": "1",
                        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.109 Safari/537.36",
                        #"Referer": Refer_url,
                        "Accept-Encoding": "gzip, deflate, sdch",
                        "Accept-Language": "zh-CN,zh;q=0.8,en;q=0.6",
                        # fill it "Cookie": ""
                    }

                    # r_2 = requests.get(request_url, headers=headers)
                    # my_excel = r_2.content

                    # comp_name = comp_name_list.index(each_comp_id)
                    # urllib2.urlretrieve(request_url, str(each_comp_type) + str(each_comp_id) + ".xls")

                    resp = requests.get(request_url, headers=headers)

                    # solve page 10 issue with "/" in company name
                    clean_comp_name = page_dict[k][1][page_dict[k][1].index(each_comp_id)].replace("/", "-")

                    with open(clean_comp_name + "_" + each_comp_type + ".xls", 'wb') as output:
                        output.write(resp.content)

                    print("成功下载第" + company_count + "个表格: " + clean_comp_name + " " + each_comp_type + ".xls")


def start():
    if __name__ == '__main__':

        search_parse()
        download_xls()
