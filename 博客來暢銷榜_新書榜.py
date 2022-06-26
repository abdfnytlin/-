# 列印排行榜的副程式
def showpage(url,kind):
    html = requests.get(url, headers=headers).text
    soup = BeautifulSoup(html,'lxml')
    res = soup.find_all('div',{'class':'mod type02_m035 clearfix'})[0]
    items = res.select('.item')
    #print(items[0])
    n = 0  # count top number
    for item in items:
        n = n + 1
        name = item.select('a')[1].text
        msg = item.select('.msg')[0]
        try:
            author = msg.select('a')[0].text
        except:
            author = ""
        price = msg.select('.price_a')[0].text 
        #print('No.', n)
        #print("　書名：", name)
        #print("　作者：", author)
        #print(price)
        #print()
        lists.append([kind,n,name,author,price])

# 主程式
import requests
from bs4 import BeautifulSoup
import re
import openpyxl
import time
from time import sleep
from datetime import date

kindno = 0
lists = []
url_sale =  'https://www.books.com.tw/web/sys_saletopb/books'
url_new = 'https://www.books.com.tw/web/sys_newtopb/books'

while True:
    choice_top = int(input("請輸入要查詢的排行榜: (0:暢銷榜; 1:新書榜)"))
    if choice_top == 0:
        url = url_sale
        att = int(input("請選擇暢銷榜('7'日/'30'日):"))
        if att == 7:
            urls = url + "?attribute=7"
        elif att == 30:
            urls = url + "?attribute=30"
        else:
            print("輸入格式不符!")
            continue
        break
    elif choice_top == 1:
        url = url_new
        urls = url
        break
    else:
        print("輸入格式不符!")
        continue

# 擷取網頁
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'}
html = requests.get(urls, headers=headers).text
soup = BeautifulSoup(html,'lxml')

res = soup.find('div',class_='mod_b type02_l001-1 clearfix')
hrefs = res.select('a')
# print(res,hrefs)
# 建立不同分類的網址
pattern = url +"/"+"(\d{2})?"
kind_index = dict()
for i in range(len(hrefs)):
    m  = re.findall(pattern,str(hrefs[i]))[0]
    kind_index[i] = m
#print(kind_index)
kindno = int(input("請輸入要下載的分類:"))
if 0 <= kindno <= (len(hrefs)-1):
    kind = hrefs[kindno].text
    print("下載的分類編號:{} 分類名稱:{}".format(kindno,kind))
    if att == 7: attrib = "?attribute=7"
    if att == 30: attrib = "?attribute=30"
    if choice_top == 0:
        kindurl = url + "/" + kind_index[kindno] + attrib
    else:
        kindurl = url + "/" + kind_index[kindno]
    print(kindurl)
    showpage(kindurl,kind)
    # 轉存為 excel
    print("資料寫入Excel中，請稍候...")
    workbook = openpyxl.Workbook()
    sheet = workbook.worksheets[0]
    listtitle=["分類",'排行','書名','作者','價格']
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
        tm = time.ctime().split(" ")[3].split(":")
    tm = "".join(tm)  # 合併為字串
    books_top = "books_top_"+S+tm+".xlsx"
    # 儲存
    workbook.save(books_top)
    print("儲存完成！")
else:
    print("分類不存在！")
    print("請重新執行本程式！")
