from selenium import webdriver
from selenium.webdriver import ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome import service as fs
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver import Chrome
from subprocess import CREATE_NO_WINDOW

import time
import openpyxl
import pandas as pd
import xlrd
import os
import webbrowser

options = webdriver.ChromeOptions()
#ヘッドレスモードの場合指定
#options.add_argument('headless')
#シークレットモードの場合の指定
options.add_argument('incognito')
options.add_experimental_option("excludeSwitches", ['enable-automation']);
chrome_service = fs.Service(executable_path='./driver/chromedriver.exe')
chrome_service.creationflags = CREATE_NO_WINDOW
serv = Service(ChromeDriverManager().install())#driverの自動更新
driver = webdriver.Chrome(service=serv, options=options)
wait = WebDriverWait(driver, 10)#タイムアウト時間の設定

def xpath_click(driver,xpath):
    driver.find_element(By.XPATH, xpath).click()

   
#######################################
#CSVファイルの読み込み
dirname = os.getcwd()
df = pd.read_excel('./site.xlsx', sheet_name=["会社", "あなたの情報"],header=0,index_col=0)
df1 = df["会社"]

company_df =  df["会社"]
info_df = df["あなたの情報"]

#print(company_df["ホームページ"]["三井ホーム"])
#print(info_df["あなたの情報"]["名前（上）"])

#######################################

#動画で見やすいように遅延処理
time.sleep(1)
driver.get(company_df["アカウント作成"]["三井ホーム"])

time.sleep(1)
xpath_click(driver,'//*[@id="first_access"]')#

time.sleep(1)
xpath_click(driver,'/html/body/div[1]/div[3]/div[2]/div[4]/p[1]/a')#

time.sleep(1)
xpath_write(driver,'/html/body/div[1]/div[3]/div[2]/form/dl/div[2]/dd/span/span[1]/div/div',info_df["あなたの情報"]["名前（上）"])

#終了
time.sleep(5)
exit()
