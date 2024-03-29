
#xl文件处理库
import xlwt
from lxml import etree

#自动化测试库
from selenium import webdriver

#网络请求库
import requests

#其他库
import traceback
import time
import pandas as pd
import numpy as np
import json
from sqlalchemy import create_engine
from utils import DateAndTimeUtil as datu
from utils import CommonUtil as cu
from utils import StockDataUtil as sdu

# import DateAndTimeUtil as datu
# import CommonUtil as cu
# import StockDataUtil as sdu

# 使用xlwt.Workbook() 初始化
def addHeaders(workbook,board_name,heads):
    sheet = workbook.add_sheet(board_name) 
    for h in range(len(heads)):
        sheet.write(0, h, heads[h]) 
    return sheet


STOCK_TYPE_HSA = '沪深A股'
STOCK_TYPE_SHA = '沪市A股'
STOCK_TYPE_SZA = '深市A股'
STOCK_TYPE_ZXB = '中小板'
STOCK_TYPE_CYB = '创业板'

#根据股票类型获取相关股票列表
def getStockDataByType(dir_path,stock_type=STOCK_TYPE_HSA,save_to_db=False):
    # StockCode：股票代码
    # StockName：股票简称
    # Price：当前股价
    '-------每日浮动数据-------'
    # PriceLimit：涨跌幅
    # QuantityRelativeR：量比
    # TurnoverRate：换手率
    '-------相对固定数据-------'
    # Trade：行业
    # PB：市净率
    # PE_S：静态市盈率
    # PE_D：动态市盈率
    # EarningsPerShare：每股收益
    # NetProfitDes：净利润描述（包括净利润和增长比率）
    # ROE：净资产收益率
    # GrossProfitRate：毛利率
    # NetAssetValuePerShare：每股净资产
    # CapitalStock：股本
    # ScaleShareType：类型（大盘股、小盘股）
    # FinanceAnalize：财务分析
    # MoneyFlowPerShare：每股现金流量
    '-------备用字段-------'
    # Proceeds：营业收入
    # ProceedsYOY：营业收入-同比增长
    # ProceedsQOQ：营业收入-环比增长
    # NetProfit：净利润
    # NetProfitYOY：净利润-同比增长
    # NetProfitQOQ：净利润-环比增长
    # GrossProfitRate：销售毛利率
    heads = ['StockCode','StockName','Price','PriceLimit','QuantityRelativeR','TurnoverRate',\
    'Trade','PB','PE_S','PE_D','EarningsPerShare','NetProfitDes','ROE','GrossProfitRate','NetAssetValuePerShare','CapitalStock',\
     'ScaleShareType','FinanceAnalize','MoneyFlowPerShare']
    
    # 初始化webdriver
    browser = webdriver.Chrome()
    browser.get("http://quote.eastmoney.com/center/gridlist.html#hs_a_board")
    
    parent = browser.find_element_by_xpath('//*[@id="tab"]/ul')
    boards = parent.find_elements_by_xpath('.//*')
    
    # 初始化xl对象
    workbook = xlwt.Workbook()
    
    #连接数据库
    engine = create_engine('mysql+pymysql://root:root@localhost:3306/listed_company')
    
    try:
        for i,board in enumerate(boards):
            board_name =board.text
            if board_name == stock_type:
                sheet = addHeaders(workbook,board_name,heads)
                board.click()
                time.sleep(2)
                print("正在获取："+board_name+"..." + '请勿关闭浏览器')
                page_parent = browser.find_element_by_class_name('paginate_page')
                pages = page_parent.find_elements_by_xpath('.//*')
                total_page = int(pages[len(pages) - 1].text)
                # total_page = 2
                
                #统计每页缺失数据的数量
                missing_count = 0
                row_index = 0
                for page in range(1,total_page+1):
                    for j in range(20):
                        cu.printProgress("第"+str(page)+"页-"+"第"+str(j)+"条...")
                        company = browser.find_element_by_xpath('//*[@id="table_wrapper-table"]/tbody/tr['+str(j+1)+']/td[3]/a')
                        code = browser.find_element_by_xpath('//*[@id="table_wrapper-table"]/tbody/tr['+str(j+1)+']/td[2]/a')
                        price = browser.find_element_by_xpath('//*[@id="table_wrapper-table"]/tbody/tr['+str(j+1)+']/td[5]/span')
                        pl = browser.find_element_by_xpath('//*[@id="table_wrapper-table"]/tbody/tr['+str(j+1)+']/td[6]/span')
                        qrr = browser.find_element_by_xpath('//*[@id="table_wrapper-table"]/tbody/tr['+str(j+1)+']/td[15]')
                        tr = browser.find_element_by_xpath('//*[@id="table_wrapper-table"]/tbody/tr['+str(j+1)+']/td[16]')
                        # 排除无价格、ST和退市的股票
                        if ((price.text == '-') | sdu.isST(company.text) | sdu.isDelist(company.text)):
                            missing_count += 1
                            continue
                        else:
                            row_index += 1
                            if save_to_db:
                                sql = 'INSERT INTO t_tonghua (stock_code,stock_name) VALUES ("{}","{}")'.format(code.text,company.text)
                                engine.execute(sql)
                            else:
                                sheet.write(row_index,0,code.text)
                                sheet.write(row_index,1,company.text)
                                sheet.write(row_index,2,price.text)
                                sheet.write(row_index,3,pl.text)
                                sheet.write(row_index,4,qrr.text)
                                sheet.write(row_index,5,tr.text)
                                getBaseDataFromF10(code.text,row_index,6,sheet)                            
                    
                    page_btn = browser.find_element_by_xpath('//*[@id="main-table_paginate"]/a[2]')
                    if page < total_page:
                        page_btn.click()
                        time.sleep(2)
                
    except Exception:
        traceback.print_exc()
    
    if save_to_db == False:
        #保存为xls文件
        file_path = dir_path + board_name +'.xls'
        workbook.save(file_path)
        print('数据下载完毕，已保存到'+file_path)
    
    browser.close()
    browser.quit()
    
#获取指数成分股
def getIndexStockByType(dir_path,index_type):
    heads = ['StockCode','StockName','Price',\
    'Trade','PB','PE_S','PE_D','EarningsPerShare','NetProfitDes','ROE','GrossProfitRate','NetAssetValuePerShare','CapitalStock',\
     'ScaleShareType','FinanceAnalize']
    
    browser = webdriver.Chrome()
    browser.get('http://data.eastmoney.com/other/index/hs300.html')
    
    parent = browser.find_element_by_id("mk_type")
    
    boards = parent.find_elements_by_xpath('.//*')
    
    workbook = xlwt.Workbook()
    
    for i,board in enumerate(boards):
        board_name =board.text
        if board_name == index_type:
            sheet = addHeaders(workbook,board_name,heads)
            board.click()
            time.sleep(2)
            print("正在获取："+board_name+"...")
            total_page = int(browser.find_element_by_xpath('//*[@id="miniPageNav"]/b[4]/span').text)
            for page in range(1,total_page+1):
                for j in range(50):
                    cu.printProgress("第"+str(page)+"页-"+"第"+str(j)+"条...")
                    company = browser.find_element_by_xpath('//*[@id="dt_1"]/tbody/tr['+str(j+1)+']/td[3]/a')
                    code = browser.find_element_by_xpath('//*[@id="dt_1"]/tbody/tr['+str(j+1)+']/td[2]/a')
                    price = browser.find_element_by_xpath('//*[@id="dt_1"]/tbody/tr['+str(j+1)+']/td[4]/span')
                    
                    row = (page-1)*50+1+j
                    sheet.write(row,0,code.text)
                    sheet.write(row,1,company.text)
                    sheet.write(row,2,price.text)
                    getBaseDataFromF10(code.text,row,3,sheet)
                
                
                page_parent = browser.find_element_by_id("PageCont")
                page_btns = page_parent.find_elements_by_xpath('.//*')
                subscript = 0
                if (len(page_btns) >0)  & (page != total_page+1):
                    for page_btn in page_btns:
                        if page_btn.text == "下一页":
                            subscript = page_btns.index(page_btn)
                    page_btns[subscript].click()
                    time.sleep(2)
                    
            file_path = dir_path + board_name +'.xls'
            workbook.save(file_path)
            
            print('数据下载完毕，已保存到'+file_path)
                    
    browser.close()
    browser.quit()

#从同花顺F10获取个股数据
def getBaseDataFromF10(code,row,index,sheet):
    try:
        #设置用户代理，不然会被网站屏蔽
        headers = {"user-agent":"PostmanRuntime/7.13.0"}
        response_1 = requests.get("http://basic.10jqka.com.cn/"+code+"/company.html",headers = headers)
        response_1.encoding = 'gbk'
        e_1 = etree.HTML(response_1.text)
        # 获取行业
        try:
            trade = e_1.xpath('//*[@id="detail"]/div[2]/table/tbody/tr[2]/td[2]/span')[0]
            sheet.write(row,index,trade.text)
        except IndexError:
            try:
                trade = e_1.xpath('//*[@id="detail"]/div[3]/table/tbody/tr[2]/td[2]/span')[0]
                sheet.write(row,index,trade.text)
            except IndexError:
                sheet.write(row,index,'未知行业')
            
#         #上市时间
#         ttm = e_1.xpath('//*[@id="publish"]/div[2]/table/tbody/tr[2]/td[1]/span')[0]
#         sheet.write(row,index+1,ttm.text)

        # 获取市净率
        response_2 = requests.get("http://basic.10jqka.com.cn/"+code+"/",headers = headers)
        response_2.encoding = 'gbk'
        e_2 = etree.HTML(response_2.text)
        pb = e_2.xpath('//*[@id="sjl"]')[0]
        sheet.write(row,index+1,pb.text)
        # 获取市盈率
        # 静态市盈率
        pe_s = e_2.xpath('//*[@id="jtsyl"]')[0]
        sheet.write(row,index+2,pe_s.text)
        # 动态市盈率
        pe_d = e_2.xpath('//*[@id="dtsyl"]')[0]
        sheet.write(row,index+3,pe_d.text)
        
        #每股收益
        eps = e_2.xpath('//*[@id="profile"]/div[2]/table[2]/tbody/tr[1]/td[2]/span[2]')[0]
        sheet.write(row,index+4,eps.text)
        # 净利润
        np_des = e_2.xpath('//*[@id="profile"]/div[2]/table[2]/tbody/tr[3]/td[2]/span[2]')[0]
        sheet.write(row,index+5,np_des.text)

        # 净资产收益率
        roe = e_2.xpath('//*[@id="profile"]/div[2]/table[2]/tbody/tr[4]/td[3]/span[2]')[0]
        sheet.write(row,index+6,roe.text)
        
        #毛利率
        gpr = e_2.xpath('//*[@id="profile"]/div[2]/table[2]/tbody/tr[4]/td[2]/span[2]')[0]
        sheet.write(row,index+7,gpr.text)
        
        #每股净资产
        nvps = e_2.xpath('//*[@id="profile"]/div[2]/table[2]/tbody/tr[4]/td[1]/span[2]')[0]
        sheet.write(row,index+8,nvps.text)
        #股本
        cs = e_2.xpath('//*[@id="profile"]/div[2]/table[2]/tbody/tr[2]/td[4]/span[2]/text()')[0]
        sheet.write(row,index+9,str(cs))
        #类型
        sst = e_2.xpath('//*[@id="profile"]/div[2]/table[2]/tbody/tr[1]/td[4]/span[2]')[0]
        sheet.write(row,index+10,sst.text)
        
        #财务分析
        fa_text = ''
        fa = e_2.xpath('//*[@id="profile"]/div[2]/table[1]/tbody/tr[2]/td[2]/div[2]')
        if len(fa) > 0:
            for sub_fa in fa[0].getchildren():
                fa_text = fa_text+(sub_fa.text+',')
            sheet.write(row,index+11,fa_text)
        #每股现金流
        mf = e_2.xpath('//*[@id="profile"]/div[2]/table[2]/tbody/tr[3]/td[3]/span[2]')[0]
        sheet.write(row,index+12,mf.text)
            
    except ConnectionResetError:
        print(code)
        
#获取年报
def getAnnualReportByStockCode (stock_code,date):
    #拼接完整股票代码
    symbol= "SZ"+stock_code if stock_code[0] in ["0", "3"] else "SH"+stock_code
    # 报告数量取1，会从制定日期向前获取最近的一个报告
    count = 1

    # https://stock.xueqiu.com/v5/stock/finance/cn/indicator.json?symbol=SH601318&type=all&is_detail=true&count=5&timestamp=1574826156123
    url = "https://stock.xueqiu.com/v5/stock/finance/cn/indicator.json?" \
            "symbol={}&type=all&is_detail=true&count={}&timestamp={}" .\
            format(symbol, count,datu.date2TimeStamp(date))
    
    # print(url)
    # header里面必须加入Cookie，否则会报400错误
    headers = {"user-agent":"PostmanRuntime/7.13.0",
               "Cookie":"Hm_lvt_1db88642e346389874251b5a1eded6e3=1633742515; device_id=148ce45f1c3420f2560f46f55f5929cb; s=bq11i2gfwp; __utmc=1; __utmz=1.1633742720.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); acw_tc=3ccdc15316337500942814179e76a00bfa3c27ced08eeef334a330b259b937; remember=1; xq_a_token=576cf04ba868551f310417bf21651d9866114e88; xqat=576cf04ba868551f310417bf21651d9866114e88; xq_id_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJ1aWQiOjEzMzMwMTA2ODEsImlzcyI6InVjIiwiZXhwIjoxNjM2MzQyMTI4LCJjdG0iOjE2MzM3NTAxMjg4MDQsImNpZCI6ImQ5ZDBuNEFadXAifQ.BfGWPUfSLE361SRrkvj0V5UhOj_D5snwMTG0nXyD-EfeHc6wxJzR9FPBGtS5F9nDK587x1pYBZzUqWs8RhU1ekblDzJCxYjpfwaqSGm0m40LHnP9DFGqQn1D-k4RY9wpwDqnQ2A-vEa9Pv_uCpRrQBWRSgXoXuIIp0rc1twsdzjpZyfxMJ1rx6PS8b6R_suepppvsSmHHwZojhVxj5n8fBtJclf8eB-TauxvwySMvi8cmmrx5cHG82zVr4xbJ7h6zf_mKOfs7VHNVPb6FiJSDgH3_62faj-XNpQPJAH-4tr-0ydKL3xygzIfvgevETIKE_3CmmRfN54_WeN9d95VlA; xq_r_token=a13f2a57855b8169a04086e17d0067c78228c969; xq_is_login=1; u=1333010681; bid=232c888fc9a3a35853a183e5b3261fdc_kuj8n7qw; Hm_lpvt_1db88642e346389874251b5a1eded6e3=1633750128; __utma=1.501209350.1633742720.1633742720.1633750128.2; __utmt=1; __utmb=1.2.10.1633750128"}

    data = requests.get(url, headers = headers)
    try:
        # print(data.text)
        data = pd.read_json(data.text, dtype=False, orient='records')
    except Exception as e:
        print(e)
        return {}
    # title = data['data']['quote_name'] + '({})'.format(symbol)
    try:
        dict_fin = pd.Series(data['data']['list'])[0]
        return dict_fin
    except Exception as e:
        print(e)
        return {}
    
years = {'2013':'2013-12-31','2014':'2014-12-31','2015':'2015-12-31','2016':'2016-12-31','2017':'2017-12-31','2018':'2018-12-31','2019':'2019-12-31','2020':'2020-12-31'} 
def getAnnualReports (dir_path,df,year):
    workbook = xlwt.Workbook()

    # stock_code：股票代码
    # stock_name：股票简称
    '-------财务数据-------'
    # total_revenue：营业收入
    # total_revenue_gr：营业收入增长率
    # net_profit_atsopc：净利润
    
    # basic_eps：每股收益
    # basic_eps_gr：每股收益增长率
    # np_per_share：每股净资产
    # operate_cash_flow_ps：现金流
    
    # avg_roe：净资产收益率
    # asset_liab_ratio：资产负债率
    
    heads = ['stock_code','stock_name','total_revenue','total_revenue_gr','net_profit_atsopc','net_profit_atsopc_gr','basic_eps','basic_eps_gr',
             'np_per_share','np_per_share_gr','operate_cash_flow_ps','operate_cash_flow_ps_gr','avg_roe','asset_liab_ratio']
    #添加表头
    sheet = workbook.add_sheet(year + "年报") 
    for h in range(len(heads)):
        sheet.write(0, h, heads[h]) 
    
    row_index = 0
    for index, row in df.iterrows():
        stock_code = row['StockCode']
        dict_ar = getAnnualReportByStockCode(stock_code,years[year])
        if len(dict_ar) == 0:
            continue
        if year not in dict_ar['report_name']:
            continue
        row_index = row_index + 1
        cu.printProgress('获取第'+str(row_index)+'条')
        
        sheet.write(row_index,0,stock_code)
        sheet.write(row_index,1,row['StockName'])
        sheet.write(row_index,2,dict_ar['total_revenue'][0])
        sheet.write(row_index,3,dict_ar['total_revenue'][1])
        sheet.write(row_index,4,dict_ar['net_profit_atsopc'][0])
        sheet.write(row_index,5,dict_ar['net_profit_atsopc'][1])
        sheet.write(row_index,6,dict_ar['basic_eps'][0])
        sheet.write(row_index,7,dict_ar['basic_eps'][1])
        sheet.write(row_index,8,dict_ar['np_per_share'][0])
        sheet.write(row_index,9,dict_ar['np_per_share'][1])
        sheet.write(row_index,10,dict_ar['operate_cash_flow_ps'][0])
        sheet.write(row_index,11,dict_ar['operate_cash_flow_ps'][1])
        sheet.write(row_index,12,float(0 if dict_ar['avg_roe'][0] is None else dict_ar['avg_roe'][0])/100)
        sheet.write(row_index,13,float(dict_ar['asset_liab_ratio'][0])/100)
        
    #     if index == 10:
    #         break
    file_path = dir_path + year + '年报.xls'
    workbook.save(file_path)
    print('数据下载完毕，已保存到'+file_path)
    
# 获取历史动态市盈率
def getPETTM(stock_code,years = 5):
    url = 'http://www.dashiyetouzi.com/tools/compare/historical_valuation_data.php'
    # 这里必须带上Cookie，否则获取不到数据
    headers = {"user-agent":"PostmanRuntime/7.13.0",
              "Cookie":"PHPSESSID=bcb37fudh7g5l18hg1ath85fd3; Hm_lvt_210e7fd46c913658d1ca5581797c34e3=1633760068,1633760666,1633760674; Hm_lpvt_210e7fd46c913658d1ca5581797c34e3=1633760723"}
    
    from_date = datu.timeStamp2Date(time.time() - (datu.oneDaySecond() * years * 365))
    to_date = datu.timeStamp2Date(time.time())
    params = (('report_type', 'pettm'),('report_stock_id', stock_code),('from_date', from_date),('to_date',to_date ))
    response = requests.post(url, headers = headers,data=params)
    
    return json.loads(response.text)

#计算市盈率中位数
def getPEMedian(stock_code):
    data = getPETTM(stock_code)
    pe_list = []
    try:
        for pair in data['list']:
            pe_list.append(pair[1])
        pe_median = np.median(pe_list)
#        pe_newest = pe_list[len(pe_list) - 1]
    except Exception:
        return 0
    
    return pe_median
        
#-----------------------------------测试代码-------------------------------------------------
# getStockDataByType('E:/wangtao/PythonWorkSpace/DataAnalysis/StockAnalyze/data/')
#getIndexStockByType('E:/wangtao/PythonWorkSpace/SpiderSpace/Stock_Companies/data/','沪深300')
    
# df = pd.read_excel("E:/wangtao/PythonWorkSpace/DataAnalysis/StockAnalyze/data/沪深A股.xls",converters= {u'StockCode':str})
# dir_path = 'E:/wangtao/PythonWorkSpace/DataAnalysis/StockAnalyze/data/'
# getAnnualReports(dir_path,df,'2020')
    
# print(getPEMedian('600167'))