# This Python file uses the following encoding: utf-8
import sys
import requests
import datetime
import json
import pandas as pd
import time
from bs4 import BeautifulSoup as bs4
from datetime import datetime

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def getGeometricMean(years):
    totalGrowth = 1
    for index, year in enumerate(years):
        if index > 3: # 마지막은 연결산이 아니라 분기결산이라 제외함
            break
        if is_number(year.text):
            yoyGrowth = float(year.text)
        else:
            continue
        totalGrowth = totalGrowth * (1 + yoyGrowth / 100)
    return round(totalGrowth ** (1/4) - 1, 3) * 100

def getArithmeticMean(years):
    total = 0
    for index, year in enumerate(years):
        if index > 3: # 마지막은 연결산이 아니라 분기결산이라 제외함
            break
        if is_number(year.text):
            yoyValue = float(year.text)
        else:
            yoyValue = 0
        total += yoyValue
    return round(total / 4, 2)

if __name__ == '__main__':
    df = pd.read_excel('./data.xlsx')
    columns = ['회사명', '종목코드', 'P/E',  '산업평균 PER', 'P/B', '장기ROE', '장기성장률', '배당수익률', 'PEGR', '당좌비율', '순현금(억)']
    result = pd.DataFrame(columns = columns)
    try:
        for index, row in df.iterrows():
            #if index < 1100:
            #    continue
            stockPage = requests.get(f'http://comp.fnguide.com/SVO2/ASP/SVD_FinanceRatio.asp?pGB=1&gicode=A{row["종목코드"]:06d}')
            stockPage = bs4(stockPage.text, 'html.parser')
            # financialStatement = requests.get(f'http://comp.fnguide.com/SVO2/ASP/SVD_Finance.asp?pGB=1&gicode=A{row["종목코드"]:06d}')
            # financialStatement = bs4(financialStatement.text, 'html.parser') 
            stockName = row["회사명"]
            stockCode = row["종목코드"]
            print(f'Processing {index}:{stockName}:{stockCode}')
            years = stockPage.select("#p_grid1_10 > td")
            averageGrowthRate = getGeometricMean(years)
            roes = stockPage.select("#p_grid1_18 > td")
            averageRoe = getArithmeticMean(roes)
            peratio = stockPage.select("#corp_group2 > dl:nth-child(1) > dd")
            peratioIndustry = stockPage.select("#corp_group2 > dl:nth-child(3) > dd")
            pbratio = stockPage.select("#corp_group2 > dl:nth-child(4) > dd")
            dividend = stockPage.select("#corp_group2 > dl:nth-child(5) > dd")
            quickRatio = stockPage.select("#p_grid1_2 > td.r.cle")

            total_liability = stockPage.select("#compBody > div.section.ul_de > div:nth-child(3) > div.um_table > table > tbody > tr:nth-child(9) > td.r.cle")
            liquid_liability = stockPage.select("#compBody > div.section.ul_de > div:nth-child(3) > div.um_table > table > tbody > tr:nth-child(4) > td.r.cle")
            quickAssets = stockPage.select('#compBody > div.section.ul_de > div:nth-child(3) > div.um_table > table > tbody > tr:nth-child(6) > td.r.cle')

            items = [peratio, peratioIndustry, pbratio, dividend, quickRatio]
            [peratio, peratioIndustry, pbratio, dividend, quickRatio] = [item[0].text if item else '-' for item in items]
            items = [total_liability, liquid_liability, quickAssets]
            [total_liability, liquid_liability, quickAssets] = [item[0].text if item else '-' for item in items]

            if is_number(dividend.replace('%', '')):
                dividend = dividend.replace('%', '')
            else:
                dividend = 0
            if is_number(averageGrowthRate) and is_number(peratio):
                pegr = round((averageGrowthRate + float(dividend)) / float(peratio), 2)
            else:
                pegr = 0
                
            if is_number(total_liability) and is_number(liquid_liability) and is_number(quickAssets):
                netCash = float(quickAssets) - float(total_liability) + float(liquid_liability)
            else:
                netCash = "-"

            result.loc[index] = [stockName, stockCode, peratio, peratioIndustry, pbratio, averageRoe, averageGrowthRate, dividend, pegr, quickRatio, netCash]
    except Exception as exc:
        print(exc)
        result.to_csv('atLeastIHaveThis.csv', mode='a', header=False)
    result.to_csv(f'{datetime.today().strftime("%Y-%m-%d")}.csv', encoding='utf-8-sig', header=False)
        
        
        
            
        