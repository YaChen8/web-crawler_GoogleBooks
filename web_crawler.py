from lxml import etree
import requests
import sys, time
import random
import xlwt
from selenium import webdriver
from fake_useragent import UserAgent
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import os
from selenium.webdriver.common.keys import Keys


# 随机使用useragent
def getUseragent():
  myuseragent = UserAgent().chrome.split()
  myuseragent[-2] = 'Chrome/101.0.4951.41' #必须要改成这个版本的，不然google的版本太低无法爬取
  myuseragent = ' '.join(myuseragent)
  # print(myuseragent)
  return myuseragent

# 随机获取不同的google域名
def google_main():
    filepath = "./google main.txt"
    googleUrl = [line.strip() for line in open(filepath, 'r', encoding='utf-8').readlines()]
    url = random.choice(googleUrl)  # 随机抽取
    return url

# 获取每个图书详情页面的内容
def crawling_paragraph(urls, years, workbook, worksheet, row):
    for i in range(len(urls)):
        url = urls[i]
        year = years[i]
        driver.get(url)
        driver.implicitly_wait(15)
        # button：继续使用Google搜索前的须知
        try:
            driver.find_element_by_xpath("//button[@aria-label='同意 Google 出于所述的目的使用 Cookie 和其他数据']").click()
        except:
            print("success")
        html = driver.page_source
        j = 7
        while (1):
            try:
                paragraph_xpath = "/html/body/div[" + str(j) + "]/div[2]/text()"
                paragraph_data = etree.HTML(html).xpath(paragraph_xpath)
                if paragraph_data == []:
                    break;
                # print(paragraph_data)
                paragraph = ""  # 爬取得到的上下文语段
                for x in range(len(paragraph_data)):
                    if x != 0:
                        paragraph = paragraph + ("virus" + paragraph_data[x])
                    else:
                        paragraph = paragraph + (paragraph_data[x])
                print(paragraph)  # 获得上下文语段
                # 保存数据
                worksheet.write(row, 0, year, style)  # 带样式的写入
                worksheet.write(row, 1, paragraph, style)
                workbook.save('data.xls')  # 保存文件
            except:
                break
            j += 1
            row += 1
    # driver.quit()  # 爬虫完毕，关掉浏览器
    return row

# 获取每个图书的具体地址
def crawling(searchkey, start_year,end_year,page):
    all_year = []
    all_url = []
    googleUrl = google_main()
    # url = "https://"+googleUrl+"/search?q=%22virus%22&tbs=cdr:1,cd_min:1/1/"+str(start_year)+",cd_max:12/31/"+str(end_year)+",bkv:f&tbm=bks&sxsrf=ALiCzsaK6JsxRogXNyN1G6XsfjV1ULmvwQ:1653197488258&ei=sMqJYse6D8TIgQaixoqYCQ&start="+str((page-1)*10)+"&sa=N&ved=2ahUKEwjH0p_IsPL3AhVEZMAKHSKjApMQ8tMDegQIBxBE&biw=1492&bih=791&dpr=1.5"
    # url = "https://www.google.com/search?q=%22virus%22&tbs=bkv:f,cdr:1,cd_min:1/1/"+str(start_year)+",cd_max:12/31/"+str(end_year)+"&tbm=bks&sxsrf=ALiCzsZats_D7g4_PzK7_mw8KLGsRDw2w:1652840644356&ei=xFiEYoSPFdjukgWuz6HYCA&start="+str((page-1)*10)+"&sa=N&ved=2ahUKEwiEgeib_-f3AhVYt6QKHa5nCIs4UBDy0wN6BAgBEEM&biw=1492&bih=791&dpr=1.5"
    url = "https://"+googleUrl+"/search?q=%22"+searchkey+"%22&tbs=cdr:1,cd_min:1/1/"+str(start_year)+",cd_max:12/31/"+str(end_year)+",bkv:p&tbm=bks&sxsrf=ALiCzsbj-3aIf9qe6oP8fmC-VnnFzEVIaA:1653452752779&ei=0K-NYruVL_2T9u8PrdixkAI&start="+str((page-1)*10)+"&sa=N&ved=2ahUKEwi7ve2_5_n3AhX9if0HHS1sDCI4WhDy0wN6BAgBEFg&biw=1492&bih=791&dpr=1.5"
    driver.get(url)
    # 因为谷歌页面是动态加载的，需要给予页面加载时间
    driver.implicitly_wait(15)
    # button：继续使用Google搜索前的须知
    try:
        driver.find_element_by_xpath("//button[@aria-label='同意 Google 出于所述的目的使用 Cookie 和其他数据']").click()
    except:
        print("success")
    html = driver.page_source
    # print(html)
    for i in range(1,10,1): #循环一个页面上的十本书籍
        # 获取检索书籍的url,href_data[0]
        href_xpath = "/html/body[@id='gsr']/div[@id='main']/div[@id='cnt']/div[@id='rcnt']/div[@id='center_col']/div[@id='res']/div[@id='search']/div/div[@id='rso']/div[@class='Yr5TG']["+str(i)+"]/div[@class='bHexk Tz5Hvf']/a/@href"
        href_data = etree.HTML(html).xpath(href_xpath)
        # print(href_data[0])
        href_data = href_data[0].split('?')
        googleUrl = google_main()
        href_data[0] = "https://"+googleUrl+"/books?"+href_data[1]
        # print(href_data[0])
        all_url.append(href_data[0])
        # 获取检索书籍的出版时间,year_data[0]
        year_xpath = "/html/body[@id='gsr']/div[@id='main']/div[@id='cnt']/div[@id='rcnt']/div[@id='center_col']/div[@id='res']/div[@id='search']/div/div[@id='rso']/div[@class='Yr5TG']["+str(i)+"]/div[@class='bHexk Tz5Hvf']/div[@class='N96wpd']/span/text()"
        year_data = etree.HTML(html).xpath(year_xpath)
        if year_data[0].isdigit()==False:
            year_xpath = "/html/body[@id='gsr']/div[@id='main']/div[@id='cnt']/div[@id='rcnt']/div[@id='center_col']/div[@id='res']/div[@id='search']/div/div[@id='rso']/div[@class='Yr5TG']["+str(i)+"]/div[@class='bHexk Tz5Hvf']/div[@class='N96wpd']/span[2]/text()"
            year_data = etree.HTML(html).xpath(year_xpath)
        # print(year_data[0])
        all_year.append(year_data[0])
        # 获取详细书籍内的段落上下文
        # newrow = crawling_paragraph(href_data[0], year_data[0], workbook, worksheet, row)
        # row = newrow
    # driver.quit()  # 爬虫完毕，关掉浏览器
    return all_year,all_url


# 初始化一个对象，ChromeOptions代表浏览器的操作
options = Options()  #这是一个空对象
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument('lang=zh-CN,zh,zh-TW,en-US,en')
# 关闭webrtc
preferences = {
    "webrtc.ip_handling_policy": "disable_non_proxied_udp",
    "webrtc.multiple_routes_enabled": False,
    "webrtc.nonproxied_udp_enabled": False
}
options.add_experimental_option("prefs", preferences)
# 随机产生user_agent
options.add_argument('User-Agent=%s'%getUseragent())
# 使用socks代理
options.add_argument("proxy-server=socks5://127.0.0.1:4781")
# 去掉webdriver痕迹
options.add_argument("disable-blink-features=AutomationControlled")
#加载cookies中已经保存的账号和密码,运行的时候要关掉所有谷歌网站
options.add_argument(r"user-data-dir=C:\Users\cy111\AppData\Local\Google\Chrome\User Data")
#设置谷歌浏览器的页面无可视化
# options.add_argument('--headless')
# options.add_argument('--disable-gpu')
# 插件使用
driver_path = r'C:\Program Files\Google\Chrome\Application\chromedriver.exe'
driver = webdriver.Chrome(options = options,executable_path = driver_path)
# 设置日期与地理位置
def get_timezone_geolocation(ip):
    url = f"http://ip-api.com/json/{ip}"
    response = requests.get(url)
    return response.json()
res_json = get_timezone_geolocation("xxx.xxx.xxx.xxx")
# print(res_json)
geo = {
    "latitude": res_json["lat"],
    "longitude": res_json["lon"],
    "accuracy": 1
}
tz = {
    "timezoneId": res_json["timezone"]
}
driver.execute_cdp_cmd("Emulation.setGeolocationOverride", geo)
driver.execute_cdp_cmd("Emulation.setTimezoneOverride", tz)
# driver.get("http://www.google.com")

global row
row = 0
workbook = xlwt.Workbook(encoding = 'ascii')
worksheet = workbook.add_sheet('Sheet1')
style = xlwt.XFStyle() # 初始化样式
font = xlwt.Font() # 为样式创建字体
font.name = 'Times New Roman'
style.font = font # 设定样式


years = [2017, 2018, 2019, 2020, 2021]
key = "virus"

for year in years:
    start_year = year
    end_year = year
    page = 1
    while(1):
        try:
            all_year,all_url = crawling(key,start_year,end_year,page)
            print(all_year)
            print(all_url)
            row = crawling_paragraph(all_url, all_year, workbook, worksheet, row)
            print(row)
            page += 1
        except:
            break

















