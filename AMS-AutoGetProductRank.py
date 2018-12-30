# -*- coding:UTF-8 -*-

import sys
import time
import csv
import os
import json
import re
import datetime as dt
from sys import stdin
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

DATA_DIR = os.path.abspath(os.path.dirname(__file__)) + "\\data\\ProductRanks.csv"
NO_RANK_INFO = 0
PAGE_ERROR = 1

def LoadProductData():
    fs = {}
    csvFile = open(DATA_DIR)
    reader = csv.reader(csvFile)
    rows = [row for row in reader]
    asins = rows[0]
    hrefs = rows[1]
    
    if len(asins) != len(hrefs):
        print('信息有误，请检查asin跟链接是否一一对应！')
        return fs
    
    for item in range(len(asins)):
        if item == 0:
            continue
        fs[asins[item]] = hrefs[item]
        
    return fs

def GetProductsRank(driver,data):
    getRanks = []
    
    for index in data:
        print('正在获取%s: %s的产品信息，请稍后。'%(index,data[index]))
        getRank = MatchPagePattern(driver,index,data[index])
        if getRank == PAGE_ERROR:
            print('页面异常，请重点关注前后台情况！')
            getRank="(No Data)"
        elif getRank == NO_RANK_INFO:
            print('未获得 %s 的排名信息！'%index)
            getRank="(No Data)"
        else:
            print('已获得 %s 的排名信息！'%(index)) 
        getRanks.append((index,getRank))
    return getRanks
    
def MatchPagePattern(driver,asin,link):

    #打开产品链接
    driver.get(link)
    
    #获取页面代码
    page = driver.page_source
    
    #检查是否正常打开到产品页面
    re_asin = r'' + asin.strip() + ''
    m_asin = re.findall(re_asin, page, re.S|re.M)
    if len(m_asin) <= 0:
        return PAGE_ERROR
    
    #正则匹配布局模式
    #布局模式1
    re_pattern1 = r'Best Sellers Rank\s*</th>(.*?)</td>'
    #布局模式2
    re_pattern2 = r'<li id="SalesRank">\s*<b>Amazon Bestseller(.*?)</ul>\s*</li>'
    #布局模式3
    re_pattern3 = r'<tr id="SalesRank">(.*?)</tr>'
    
    matches = re.findall(re_pattern1, page, re.S|re.M)
    if len(matches) > 0:
        return GetFristPatternRank(matches[0])
        
    matches = re.findall(re_pattern2, page, re.S|re.M)
    if len(matches) > 0:
        return GetSecondPatternRank(matches[0])  
        
    matches = re.findall(re_pattern3, page, re.S|re.M)
    if len(matches) > 0:
        return GetThirdPatternRank(matches[0])
    else:
        return NO_RANK_INFO
   
   
def GetFristPatternRank(e_match):
    rank1 = ""
    #筛选span标签
    re_span = r'<span(.*?)/span>'
    m_spans = re.findall(re_span,e_match,re.S|re.M)
    rank1 = GetRankText(m_spans)
    return rank1
    
def GetSecondPatternRank(e_match):
    rank2 = ""
    re_tag = r'</b(.*?)<style'
    re_li = r'<li class="zg_hrsr_item">(.*?)</li>'
    
    #提取主目录排名
    m_mainranks = re.findall(re_tag,e_match,re.S|re.M)
    rank2 = GetRankText(m_mainranks) + '\n'
    #筛选li标签        
    m_lis = re.findall(re_li,e_match,re.S|re.M)
    rank2 += GetRankText(m_lis)
            
    return rank2
    
    
def GetThirdPatternRank(e_match):
    rank3 = ""
    re_li = r'<li class="zg_hrsr_item"(.*?)/li>'
    #筛选li标签
    m_lis = re.findall(re_li,e_match,re.S|re.M)
    rank3 = GetRankText(m_lis)
    return rank3

def GetRankText(matches):
    rankText = ""
    re_text = r'>(.*?)<'
    for match in matches:
        texts = re.findall(re_text, match, re.S|re.M)
        for text in texts:
            rankText += text.strip()
        if match != matches[len(matches) - 1]:
            if len(rankText) > 0:
                rankText += '\n'
    #文本处理
    rankText = rankText.replace('&nbsp;',' ')
    rankText = rankText.replace('&gt;',' > ')
    rankText = rankText.replace('&amp;','&')
    
    return rankText
    
def WriteRankData(productRanks):
    ranks = []
    for row in productRanks:
        ranks.append(row[1])
    #encoding='utf-8'打开文件时就声明编码方式为utf-8,防止遇到gbk不支持字符报错
    #"a+"写入csv文件追加方式不多空行
    csvFile = open(DATA_DIR,'a+',encoding='utf-8')
    writer = csv.writer(csvFile)
    rankText = []
    for item in ranks:
        item = item.encode('utf-8').decode('utf-8')
        rankText.append(item)
    writer.writerow(rankText)
    csvFile.close()

        
if __name__=="__main__":
    try:
        driver = None
        
        #打开Chrome
        options = webdriver.ChromeOptions()
        driver = webdriver.Chrome(chrome_options=options)
        
        #从data.csv加载产品列表文件，获取产品名及产品链接
        print("获取产品信息")
        productData = LoadProductData()
        
        if len(productData) > 0:
            print("开始收集产品排名！")
            productRanks = []
            productRanks.append(("Update Time",dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            productRanks += GetProductsRank(driver,productData)
            print("获取排名完成！")
        
            #把排名信息写入文件
            print("排名情况如下：")
            for row in productRanks:
                print('%s:\n%s'%(row[0],row[1]))
                
            WriteRankData(productRanks)
            print("已把排名数据写入ProductRanks.csv中！")
        else:
            print("并未获得产品信息！请检查data.csv文件是否填好产品信息！")
        
        #结束退出
        print("程序结束！")
        
        
    finally:
        if driver is not None:
            driver.close()    
