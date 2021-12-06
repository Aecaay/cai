import requests
import random
import time
import xlwt
import xlrd
import bs4
import re
import os
from xlutils.copy import copy
from requests import exceptions



#请求搜索页
def Find_url(url):
    head = {
        "User-Agent":"Mozilla/5.0 (Windows NT 6.1; WOW64; rv:6.0) Gecko/20100101 Firefox/6.0"
    }
    # proxies = {
    #     "https":"https://20.187.79.54:30001"
    # }
    all_link = requests.get(url,headers=head)
    print(all_link.status_code)
    web_url = bs4.BeautifulSoup(all_link.text,"html.parser")
    link_pag = str(web_url.findAll("ul",class_="gl-warp clearfix"))
    find_link = re.compile(r'href="//item.jd.com/(.*?).html" onclick="searchlog')
    link = re.findall(find_link,link_pag)
    return link

#请求详情页
def Request(url):
    proxies = [
        {"http":"http:222.78.6.2:8083"},
        {"http":"159.75.25.176:81"},
        {"http":"111.231.86.149:7890"},
        {"http":"182.84.145.138:3256"},
        {"http":"27.191.60.100:3256"},
        {"http":"121.232.148.71:9000"},
        {"http":"210.75.240.136:1080"},
        {"http":"163.125.112.90:8118"},
        {"http":"58.20.232.245:9091"},
        {"http":"111.231.86.149:7890}"}
    ]
    ip = random.choice(proxies)
    # proxies = {
    #     "https":"https://20.187.79.54:30001"
    # }
    user_list = [
        "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:6.0) Gecko/20100101 Firefox/6.0"
        "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1",
        "Mozilla/5.0 (X11; CrOS i686 2268.111.0) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.57 Safari/536.11",
        "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1092.0 Safari/536.6",
        "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1090.0 Safari/536.6",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36"
    ]
    head = {
        'User-Agent': random.choice(user_list)
    }
    print(url)
    try:
        time.sleep(0.5)
        web_pag = requests.get(url,headers=head,proxies=ip)
    except exceptions.ConnectionError as error:
        print(error)
    except exceptions.HTTPError as error:
        print(error)
    else:
        # print(web_pag.text)
        return web_pag.text

#筛选数据
def Find_data(data):
    data_list = []
    web_date = bs4.BeautifulSoup(data,"html.parser")
    web_under = str(web_date.findAll("div",class_="p-parameter"))
    web_top = str(web_date.findAll("div",id="choose-attrs"))
    # print(web_top)

    find_info = re.compile(r'<li title="(.*?)">')

    find_type = re.compile(r'data-value="(.*?)">')
    try:
        brand = re.findall(find_info,web_under)[0]#品牌
        title = re.findall(find_info,web_under)[1]#商品名称
        number = re.findall(find_info,web_under)[2]#编号
        shop = re.findall(find_info,web_under)[3]#店铺
        weight = re.findall(find_info,web_under)[4]#重量
        place = re.findall(find_info,web_under)[5]#产地
        effect = re.findall(find_info,web_under)[6]#功效
        formula = re.findall(find_info,web_under)[7]#配方
        age = re.findall(find_info,web_under)[8]#年龄
        taste = re.findall(find_info,web_under)[9]#口味
        univalent = re.findall(find_info,web_under)[10]#每斤单价
        coff = re.findall(find_info,web_under)[11]#国产或进口

        type = re.findall(find_type,web_top)#类型
        

        data_list.append(brand)
        data_list.append(title)
        data_list.append(number)
        data_list.append(shop)
        data_list.append(weight)
        data_list.append(place)
        data_list.append(effect)
        data_list.append(formula)
        data_list.append(age)
        data_list.append(taste)
        data_list.append(univalent)
        data_list.append(coff)
        data_list.append(type)
    except Exception as error:
        return
    else:
        return data_list

#写入excel表
def Write_Excel(data_list,save_path,page):
    col = ["编号","品牌","商品名称","编号","店铺","商品毛重","商品产地","功效","配方","使用阶段","口味","每斤价格","国产/进口","商品分类"]
    try:
        if page == 0:
            excel_book = xlwt.Workbook(encoding="utf-8")
            excel_sheet = excel_book.add_sheet("cat")
            for i in range(len(col)):
                excel_sheet.write(0,i,col[i])
        else:
            excel_book = xlrd.open_workbook("%s"%save_path)
            excel_book = copy(excel_book)
            excel_sheet = excel_book.get_sheet("cat")
        i = 0
        pages = 0
        
        while pages < len(data_list):
            if data_list[i] == None:
                i += 1
                pages += 1
                continue
            else:
                for col in range(len(data_list[i])):
                    if col == 0:
                        excel_sheet.write(page+1,0,page+1)
                    excel_sheet.write(page+1,col+1,data_list[i][col])
            i += 1
            pages += 1
            print(page)
            page += 1
            excel_book.save("%s"%save_path)
        return page
    except Exception as error:
        print("写表错了")
        

def Agent_Write(data_list,save_path):
    try:
        excel_book=xlrd.open_workbook("%s"%save_path)
        excel_sheet = excel_book.sheet_by_name("cat")
        page = excel_sheet.nrows
        excel_book = copy(excel_book)
        excel_sheet = excel_book.get_sheet("cat")
        i = 0
        pages = 0
        while pages < len(data_list):
            if data_list[i] == None:
                i += 1
                pages += 1
                continue
            else:
                for col in range(len(data_list[i])):
                    if col == 0:
                        excel_sheet.write(page,0,page)
                    excel_sheet.write(page,col+1,data_list[i][col])
            i += 1
            pages += 1
            page += 1
            excel_book.save("%s"%save_path)
    except Exception as error:
        print("error")





def main():
    save_path = "./cat.xls"
    page = 0
    num = 0
    s = random.randrange(0,101,1)
    while(num<500): #30
        link = []
        data_list = []
        url = "https://search.jd.com/Search?keyword=%E7%8C%AB%E7%B2%AE&qrst=1&wq=%E7%8C%AB%E7%B2%AE&page="+str(num+s)
        print(url)
        link.extend(Find_url(url))
        for i in range(0,len(link)//2):     #去重
            del link[i]
        o = 1
        for i in link:
            url = "https://item.jd.com/"+i+".html"
            time.sleep(2)
            data = Request(url)
            if data != None:
                data_list.append(Find_data(data))
            print(o)
            o+=1
        if os.path.exists(save_path) == False:
            if os.path.getsize(save_path) == 0:
                page = Write_Excel(data_list,save_path,page)
        elif os.path.getsize(save_path) != 0:
                Agent_Write(data_list,save_path)
        num+=1
    print("完成，%s"%save_path)


if __name__ == "__main__":
    main()
