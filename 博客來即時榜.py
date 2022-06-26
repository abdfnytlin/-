import requests
from bs4 import BeautifulSoup
import openpyxl
import time
from time import sleep
from datetime import date

lists = []
print("博客來即時榜")
print("="*50)
# 博客來即時榜連結
url = 'https://www.books.com.tw/web/sys_tdrntb/books/'
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'}

html = requests.get(url, headers=headers).text
soup = BeautifulSoup(html,'lxml')
title = soup.find('div',class_='mod type02_s002').text
res = soup.find_all('div',{'class':'mod type02_m035 clearfix'})[0]
items = res.select('.item')
n = 0  # count top number
print(title)
for item in items:
    n = n + 1
    name = item.select('a')[1].text
    msg = item.select('.msg')[0]
    try:
        author = msg.select('a')[0].text
    except:
        author = ""
    price = msg.select('.price_a')[0].text 
    print('No.', n)
    print("　書名：{}".format(name))
    print("　作者：{}".format(author))
    print(price)
    print()
    lists.append([n,name,author,price])
# 轉存為 excel
print("資料寫入Excel中，請稍候...")
workbook = openpyxl.Workbook()
sheet = workbook.worksheets[0]
listtitle=['排行','書名','作者','價格']
sheet['A1']="博客來即時榜"
sheet['B1']=time.ctime()
sheet.append(listtitle)
for item in lists:
    sheet.append(item)
    sleep(1)  # Delay 防止寫入錯誤
# 設定以儲存完成時間作為檔名
L = str(date.today()).split("-")  # 日期做成串列
S = "".join(L)[2:]  # 合併為字串，取西元年末兩數字
tx = int(time.localtime().tm_mday)  # 時間做成串列
if tx < 10:
    tm = time.ctime().split(" ")[4].split(":")
else:
    tm = time.ctime().split(" ")[4].split(":")
tm = "".join(tm)  # 合併為字串
books_top = "books_top_"+S+tm+".xlsx"
# 儲存
workbook.save(books_top)
print("儲存完成！")
