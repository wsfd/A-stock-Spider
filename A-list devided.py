#数据处理
import xlrd
from threading import Thread 
import xlwt
from time import sleep
import urllib.request
import gzip
import random
import zlib
from urllib import request 
from xlwt import Workbook
import time
import re
import socket
from bs4 import BeautifulSoup

timeout = 30
socket.setdefaulttimeout(timeout)
date = time.strftime('%Y-%m-%d',time.localtime(time.time()))
def get_proxies(url_ip):
    res = urllib.request.urlopen(url_ip)
    html = res.read()
    soup = BeautifulSoup(html)
    contents = soup.find_all('tr')
    regex = re.compile('\d+')
    proxies = []
    for each in contents:
        sock = each.find_all('td')
        if sock:
            ip = sock[0].text
            port = sock[1].text
            if re.findall(regex,ip):
                proxy = '%s:%s' %(ip,port)
                proxies.append(proxy)
    return proxies

#url_ip = 'http://cn-proxy.com/archives/218'
#proxy_list = get_proxies(url_ip)
proxy_list=[]

	
for line in open("下载.txt"):
    line = line.strip('\n')#去掉换行符   
    proxy_list.append(line)

x=[]
e=[]
b=[]
m={}

def proxyset():
    proxy = random.choice(proxy_list)
    proxy_support = request.ProxyHandler({'http':proxy})  
    opener = request.build_opener(proxy_support, request.HTTPHandler)  
    request.install_opener(opener)    

def main_function(order):
    names = locals()
    names['x%s' % order] = 0
    global m
    global e
    #读取股票代码
    data = xlrd.open_workbook('2016-1-3.xls')
    table = data.sheets()[0]
    nrows = table.nrows
    colnames = table.col_values(0) #某一列数据
    b = colnames[order::10]
    #获得当下时间戳
    nowtime = time.time()
    nowtime = "%.3f"%nowtime 
    nowtime = str(nowtime )
    nowtime = nowtime.replace('.','')
    
    proxyset() 
 
    for stockcode in b :    	
        url = 'http://xueqiu.com/recommend/pofriends.json?type=1&code={number}&start=0&count=14&_={time}'
        url = url.format(number=stockcode,time=nowtime)
        req = urllib.request.Request(url, headers = {
            'Accept':'application/json, text/javascript, */*; q=0.01',
            'Accept-Encoding':'gzip, deflate, sdch',
            'Accept-Language':'zh-CN,zh;q=0.8',
            'Cache-Control':'no-cache',
            'Connection':'keep-alive',
            'Cookie':'s=9e711qyz8y; webp=0; __utma=1.442364867.1448436216.1451681673.1451762655.28; __utmz=1.1450307253.24.3.utmcsr=google|utmccn=(organic)|utmcmd=organic|utmctr=(not%20provided); Hm_lvt_1db88642e346389874251b5a1eded6e3=1449331377,1449337330,1450307255,1451558257; xq_a_token=657d828b5d78bd9e62bc778ab7ac0bcdbc1b9337; xq_r_token=fff6555fe028c4588b961866cd2bb1e0147e04ab',
            'Host':'xueqiu.com',
            'RA-Sid':'655102BF-20150723-085431-c809af-3fa054',
            'RA-Ver':'3.0.7',
            'Referer':'http://xueqiu.com/S/SZ001979',
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.101 Safari/537.36',
            'X-Requested-With':'XMLHttpRequest',
        })
        try:
            oper = urllib.request.urlopen(req)
            html = oper.read()
            html = zlib.decompress(html, 16+zlib.MAX_WBITS)
            html = html.decode('utf-8','ignore') 
            c = re.search(r'"totalcount":(\d{1,6})',html)
            names['x%s' % order] = c.group(1)
            time.sleep(5)
        except:
            print(stockcode)
            e.append(stockcode)
            b.append(stockcode)
            proxyset() 
        else:

            if names['x%s' % order] == '0':
                continue
            else:
                m[stockcode] = names['x%s' % order]  
                output = '股票代码:{stockcode} 关注人数:{d}'
                output = output.format(stockcode=stockcode,d=names['x%s' % order])
                print (output)	

starttime=time.time(); #记录开始时间				
threads = [] 
t1 = Thread(target=main_function,args=(0,))
threads.append(t1)#将这个子线程添加到线程列表中
t2 = Thread(target=main_function,args=(1,))
threads.append(t2)#将这个子线程添加到线程列表中
t3 = Thread(target=main_function,args=(2,))
threads.append(t3)#将这个子线程添加到线程列表中
t4 = Thread(target=main_function,args=(3,))
threads.append(t4)#将这个子线程添加到线程列表中
t5 = Thread(target=main_function,args=(4,))
threads.append(t5)#将这个子线程添加到线程列表中
t6 = Thread(target=main_function,args=(5,))
threads.append(t6)#将这个子线程添加到线程列表中
t7 = Thread(target=main_function,args=(6,))
threads.append(t7)#将这个子线程添加到线程列表中
t8 = Thread(target=main_function,args=(7,))
threads.append(t8)#将这个子线程添加到线程列表中
t9 = Thread(target=main_function,args=(8,))
threads.append(t9)#将这个子线程添加到线程列表中
t10 = Thread(target=main_function,args=(9,))
threads.append(t10)#将这个子线程添加到线程列表中



for t in threads:
    t.start()
for t in threads:
    t.join()

#按关注人数排序并输出
#m=sorted(m.items(), key=lambda d:d[1], reverse=True)	
#with open('C:/Python34/test.txt', 'wt') as f:
#    print(m, file=f)

#输出到excel
x=list(m.keys())
list_length = len(x)
book = Workbook()
sheet1 = book.add_sheet('sheet1')

for n in range(list_length):
    l = n-1
    z = x[l]
    sheet1.write(n,0,z)
    q = m[z]
    sheet1.write(n,1,q)
book.save('2016-1-4.xls')

#输出爬取失败的股票代码
with open('C:/Python34/failed.txt', 'wt') as f:
    print(e, file=f)


endtime=time.time();#记录程序结束时间
totaltime=endtime-starttime;#计算程序执行耗时

print (totaltime)