import requests
from bs4 import BeautifulSoup
import re
import time
from time import sleep
from datetime import date
import openpyxl

def main():
    # 擷取網頁，設定初始路徑
    homepage = "https://www.books.com.tw/web/books_newbook"
    homeurl = "https://www.books.com.tw/web/books_nbtopm_01?v=1&o=1"
    url = 'https://www.books.com.tw/web/books_nbtopm_'
    mode = choose_mode(homeurl)
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'}
    html = requests.get(homeurl, headers=headers).text
    soup = BeautifulSoup(html,'lxml')
    res = soup.find('div',class_='mod_b type02_l001-1 clearfix')
    hrefs = res.select('a')
    # input_mode 選擇輸入分類編號或分類名稱
    if input_mode == 0:
        # 分類編號並非連續，使用字典確認各分類路徑
        pattern = r"https://www.books.com.tw/web/books_nbtopm_(\d{2})"
        kind_index = dict()
        for i in range(1,len(hrefs)+1):
            m  = re.findall(pattern,str(hrefs[i-1]))[0]
            kind_index[i] = m
        kindno = int(input("請輸入要下載的分類編號:"))
        if 0 < kindno <= len(hrefs):
            kind = hrefs[kindno-1].text
            print("下載的分類編號:{} 分類名稱:{}".format(kindno,kind))
            kindurl = url + kind_index[kindno] + mode
            print(kindurl)
            showkind(kindurl,kind)
            save_to_xls()
        else:
            print("分類不存在，請重新輸入!")
    if input_mode == 1:
        # 分類編號並非連續，使用字典確認各分類路徑
        pattern = url +"(\d{2})"
        kind_index = dict()
        for i in range(1,len(hrefs)+1):
            m  = re.findall(pattern,str(hrefs[i-1]))[0]
            kind_index[i] = m
        # 儲存分類名稱，做成字典
        kinds = []
        for href in hrefs:
            kinds.append(href.text)
        # 使用字典資料，以名稱指定查詢路徑
        ind = [i for i in range(1,len(hrefs)+1)]
        kindlink = [kind_index[i] for i in ind]
        kindkey = {kind:kind_lnk for kind,kind_lnk in zip(kinds,kindlink)}
        # 分類選擇
        kindname = input("請輸入要下載的分類名稱:")
        if kindname in kindkey:
            print("下載的分類名稱:{}".format(kindname))
            kindurl = url + kindkey[kindname] + mode
            print(kindurl)
            showkind(kindurl,kindname)
            save_to_xls()
        else:
            print("分類不存在，請重新輸入!")
    print("="*50)
        
def menu():
    tx = int(time.localtime().tm_mday)  # 時間做成串列
    if tx < 10:
        clock = time.ctime().split(" ")[4]
    else:
        clock = time.ctime().split(" ")[3]
    print("歡迎使用「博客來近期中文書新書查詢」")
    print("https://www.books.com.tw/web/books_newbook")
    print("現在時間：　{}　{}".format(str(date.today()),clock))
    print("-"*50)
    print("1. 下載新書資訊")
    print("2. 分 類 列 表")
    print("0. 結 束 程 式")
    print("="*50)

def display_kind():
    url = "https://www.books.com.tw/web/books_newbook"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'}
    html = requests.get(url,headers=headers).text
    soup = BeautifulSoup(html,'lxml')
    res = soup.find('div',class_='mod_b type02_l001-1 clearfix')
    hrefs = res.select('a')
    for index,href in enumerate(hrefs):
        print("{}\t{}".format(index+1,href.text))  # 顯示所有分類
    print("="*50)

def showkind(url,kind):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'}
    html = requests.get(url, headers = headers).text
    soup = BeautifulSoup(html, 'lxml')
    # 檢查分類頁數
    try:
        pages = int(soup.select('.cnt_page span')[0].text)
        print("共 {} 頁".format(pages))
        for page in range(1,pages+1):
            pageurl = url + '&page=' + str(page).strip()
            print("第", page, "頁", pageurl)   # 印出每一頁的網址
            showpage(pageurl,kind)
    except:
        print("共 1 頁")
        showpage(url,kind)

def choose_mode(url):
    print("-"*50)
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'}
    html = requests.get(url, headers=headers).text
    soup = BeautifulSoup(html,'lxml')
    ress = soup.find('div',class_='type02_m047 clearfix')
    opts = ress.select('option')
    pattern_option = 'https://[a-zA-Z0-9\./_-]+\?v=[0-9]&amp;o=([0-9])'
    imgs = re.findall(pattern_option,str(opts))
    optlist = []
    for opt in opts:
        optlist.append(opt.text)
    for i in range(len(imgs)-1):
        for j in range(i+1,len(imgs)):
            if imgs[j]<imgs[i]:
                temp_count = imgs[j]; temp_word = optlist[j]
                imgs[j] = imgs[i]; optlist[j] = optlist[i]
                imgs[i] = temp_count; optlist[i] = temp_word
    opt_dict = {opt:img for opt,img in zip(optlist,imgs)}
    # 選擇排序方式
    for opt in opt_dict.keys():
        print("{}\t{}".format(opt_dict[opt], opt))
    print("-"*50)
    o_selection = input("請選擇排序方式：")
    if o_selection in  opt_dict.values():
        print("排序方式：{}".format(optlist[int(o_selection)-1]))
        mode = '?v=1&o=' + o_selection
    elif o_selection in  opt_dict.keys():
        print("排序方式：{}".format(o_selection))
        mode = '?v=1&o=' + opt_dict[o_selection]        
    else:
        print("格式不符！將切換為排序方式「上市日期(新→舊)」。")
        mode = '?v=1&o=1'
    print("-"*50)
    return mode        
        
def showpage(url,kind):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'}
    html = requests.get(url, headers = headers).text
    soup = BeautifulSoup(html, 'lxml')
    res = soup.find_all('div', {'class': 'mod type02_m012 clearfix'})[0]
    items = res.select('.item')
    n = 0
    print('共有',len(items),'本')
    # 儲存頁面中的新書資料
    for item in items:
        n = n + 1
        msg = item.select('.msg')[0]
        title = msg.select('a')[0].text  # 書名
        author = msg.select('a')[1].text # 作者
        publisher = msg.select('a')[2].text # 出版社
        dates = msg.select('.info span')[0].text 
        release = re.findall(r"(\d{4}-\d{2}-\d{2}|\d{4}/\d{2}/\d{2})",dates)[0] # 出版日期
        onsale = item.select('.price .set2')[0].text
        price = re.findall(r"(\d{1,2}折\s{1}?\d{2,4}元|\d{2,4}元)",onsale)[0] # 價格
        content = item.select('.txt_cont')[0].text # 內容
        #print('No.',n)
        #print("書名:",title)
        #print("作者:",author)
        #print("出版社:",publisher)
        #print('出版日期:',release)
        #print('Onsale:',price)
        #print('內容:',content)
        #print()
        book_data.append([kind,title,author,publisher,release,price,content])

def save_to_xls():
    # 轉存為 excel
    print("資料寫入Excel中，請稍候...")
    workbook = openpyxl.Workbook()
    sheet = workbook.worksheets[0]
    listtitle=["分類",'書名','作者','出版社','出版日期','價格','內容']
    sheet.append(listtitle)
    for item in book_data:
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
    books_new = "books_new_"+S+tm+".xlsx"
    # 儲存
    workbook.save(books_new)
    print("儲存完成！")
        
# 主程式
while True:
    book_data = []
    menu()
    choice_menu = int(input("請輸入您要選擇的指令:"))
    print()
    if choice_menu == 1:
        input_mode = int(input("請選擇分類輸入模式: (0: 編號模式，1: 名稱模式)"))
        if input_mode == 0 or input_mode == 1:
            main()
        else: 
            print("格式不符，請重新輸入!")
            continue
    elif choice_menu == 2:
        display_kind()
    else:
        break
print("結束程式！")
