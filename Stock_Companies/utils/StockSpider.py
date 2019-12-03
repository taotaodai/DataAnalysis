import sys
import os

#xl文件处理库
import xlwt
import xlrd
from lxml import etree
from xlutils.copy import copy

#自动化测试库
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

#网络请求库
import requests

#其他库
import traceback
import time
#from utils import CommonUtil as cu
#from utils import StockDataUtil as sdu

import CommonUtil as cu
import StockDataUtil as sdu

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
def getStockDataByType(dir_path,stock_type=STOCK_TYPE_HSA):
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
                #保存为xls文件
                file_path = dir_path + board_name +'.xls'
                workbook.save(file_path)
                
                print('数据下载完毕，已保存到'+file_path)
    except Exception as e:
        traceback.print_exc()
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
        
#getStockDataByType('E:/wangtao/PythonWorkSpace/SpiderSpace/Stock_Companies/data/')
#getIndexStockByType('E:/wangtao/PythonWorkSpace/SpiderSpace/Stock_Companies/data/','沪深300')
        