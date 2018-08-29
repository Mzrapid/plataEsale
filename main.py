# -*- coding:utf-8 -*-
"""
    Headless Chrome Version
    売電実績入力支援ツール plataEsale
    Selenium, Chrome, Python3.6
    since August 14 2018
"""
from logininformation                   import *
from selenium                           import webdriver
from selenium.webdriver.common.keys     import Keys
from selenium.webdriver                 import Chrome, ChromeOptions
from selenium.webdriver.chrome.options  import Options
from selenium.common.exceptions         import NoSuchElementException
import selenium.webdriver               as webdriver
import selenium.webdriver.support.ui    as ui
import time
import traceback
import sys
import subprocess


def main():
    #ブザウザを開き、ログインしてプラント選択画面まで行く
    if Reach_to_TEPCO() == False:
        print("Error: プラント選択画面に到達できません")
        sys.exit(1) #異常終了

    #各ブラントに対してやる
    for n in range(PLANT_NUM):
        if Get_Plant_Information(n) == False:
            print("Error: プラント情報が得られません")
            sys.exit(1) #異常終了

    #無事にすべてが終わった
    driver.quit()
    print("End: main.py")


def Reach_to_TEPCO():
    """
    ■共通処理
        ブラウザを開き、ログインして、プラントを選択するところに行くまで
    """
    notrouble = True
    driver.implicitly_wait(10)  #要素が見つからないとき、10秒まで待つことにする

    try:
        #ブラウザを開く
        driver.get(url)
        print("Success: ビジネスTEPCO ログインページ に到達")

        #ユーザーIDを入力
        elemID = driver.find_element_by_id("loginid")
        elemID.send_keys(UserID)
        print("Success: ユーザーIDを正常に入力しました")

        #パスワードを入力
        elemPwd = driver.find_element_by_id("loginpwd")
        elemPwd.send_keys(Passwd)
        print("Success: パスワードを正常に入力しました")

        #ログインボタンをクリック
        driver.find_element_by_id("login").click()
        print("Success: ログインボタンを正常にクリックしました")


    except Exception as e:
        notrouble = False
        if NoSuchElementException:
            print("Error: 要素がありません")
        else:
            print("Error: 未分類のエラー")

    else:
        notrouble = True
        print("Success: ログイン操作を完了しました")

    finally:
        return notrouble


def Get_Plant_Information(plantID):
    """
    それぞれのプラントの売電情報を得ます
    try-except節による例外処理はしていません
    note: タブ切り替え時の読み込みのタイミングがシビアである。time.sleep()を不用意に外さないこと
    """
    notrouble = True
    row_of_table = 14
    driver.implicitly_wait(4)
    print("Process: %s の売電実績を取得します：" % plantName[plantID])
    time.sleep(1.000)
    driver.find_element_by_xpath(plant[plantID]).click()
    print("Success: ご使用実績 を正常にクリックしました")

    #新しいタブが開かれるため切り替え
    time.sleep(2.000)
    driver.switch_to.window(driver.window_handles[1])
    time.sleep(2.000)

    #最新のデータはいつのものであるか分析
    while row_of_table > 1:
        elems = driver.find_elements_by_xpath("/html/body/form[1]/div/table/tbody/tr/td[1]/table[3]/tbody/tr/td/table[1]/tbody/tr/td/table[1]/tbody/tr[2]/td[2]/table[2]/tbody/tr/td/table/tbody/tr[%s]/td[1]" % str(row_of_table))
        print(row_of_table)
        if len(elems)==0:   #リストが空である
            row_of_table -= 1
        elif elems[0].text==" " or elems[0].text=="&nbsp;":  #空欄である
            row_of_table -= 1
        else:   #データが見つかる
            break

    if row_of_table < 2:   #一番上まで探したけどデータが存在しなかったとき
        notrouble = False
        print("Error: データが見つかりません")

    latest = driver.find_element_by_xpath("/html/body/form[1]/div/table/tbody/tr/td[1]/table[3]/tbody/tr/td/table[1]/tbody/tr/td/table[1]/tbody/tr[2]/td[2]/table[2]/tbody/tr/td/table/tbody/tr[%s]/td[1]" % str(row_of_table))
    print("Message: %s までデータがあります" % latest.text)
    str_latest = latest.text

    #千葉市プラントと白子町プラントではXPathが違うことが判明したため分類
    if plantName[plantID] == "千葉市若葉区小間子町１ー３":
        driver.find_element_by_xpath("/html/body/form[1]/div/table/tbody/tr/td[1]/table[3]/tbody/tr/td/table[1]/tbody/tr/td/table[1]/tbody/tr[2]/td[2]/table[2]/tbody/tr/td/table/tbody/tr[%s]/td[9]" % str(row_of_table)).click()
    if plantName[plantID] == "千葉市若葉区金親町８３":
        driver.find_element_by_xpath("/html/body/form[1]/div/table/tbody/tr/td[1]/table[3]/tbody/tr/td/table[1]/tbody/tr/td/table[1]/tbody/tr[2]/td[2]/table[2]/tbody/tr/td/table/tbody/tr[%s]/td[9]" % str(row_of_table)).click()
    if plantName[plantID] == "長生郡白子町発電所１":
        driver.find_element_by_xpath("/html/body/form[1]/div/table/tbody/tr/td[1]/table[3]/tbody/tr/td/table[1]/tbody/tr/td/table[1]/tbody/tr[2]/td[2]/table[2]/tbody/tr/td/table/tbody/tr[%s]/td[8]" % str(row_of_table)).click()
    if plantName[plantID] == "長生郡白子町発電所２":
        driver.find_element_by_xpath("/html/body/form[1]/div/table/tbody/tr/td[1]/table[3]/tbody/tr/td/table[1]/tbody/tr/td/table[1]/tbody/tr[2]/td[2]/table[2]/tbody/tr/td/table/tbody/tr[%s]/td[8]" % str(row_of_table)).click()
    if plantName[plantID] == "長生郡白子町発電所３":
        driver.find_element_by_xpath("/html/body/form[1]/div/table/tbody/tr/td[1]/table[3]/tbody/tr/td/table[1]/tbody/tr/td/table[1]/tbody/tr[2]/td[2]/table[2]/tbody/tr/td/table/tbody/tr[%s]/td[8]" % str(row_of_table)).click()


    print("Success: 「内訳」ボタンをクリックしました")
    print("Process: 購入電力のお知らせを表示します")
    driver.find_element_by_xpath("/html/body/div/table/tbody/tr/td/table[1]/tbody/tr[1]/td[3]/form").click()
    driver.switch_to.window(driver.window_handles[2])

    #一応時間を設けておきます
    time.sleep(2.000)
    Make_HTML(driver.page_source, str_latest)    #整形したHTMLファイルを生成
    Make_PDF(str_latest)                         #HTMLからPDFを生成
    Move_PDF(plantID, str_latest)                #PDFをDropboxに移動

    #はじめにもどる
    time.sleep(5.000)
    driver.switch_to.window(driver.window_handles[0])
    driver.get("https://www30.tepco.co.jp/dv05/dfw/biztepco/D3BWwwAP/D3BBTUM001G01_loginauth.act?NEXTACT=D3BBTMP001G01_UsaLocationListSV&FW_SCTL=INIT")

    return notrouble


def Make_HTML(source, latest_raw):
    print("Process: 現在のページのHTMLファイルを作成します")
    filename = Era[latest_raw]
    try:
        f = open("%s.html" % filename, "w")
    except IOError as e:
        print("Error: ファイルを開けません")
        sys.exit(1) #異常終了
    
    source_utf8 = source.replace("Windows-31J", "utf-8", 1) #文字コードを修正する
    f.write(source_utf8)
    f.close()


def Make_PDF(latest_raw):
    filename = Era[latest_raw]
    print("Process: HTMLからPDFを生成します")
    cmd = "google-chrome --disable-gpu --headless --print-to-pdf="+filename+".pdf "+filename+".html"
    subprocess.run(cmd, shell=True)


def Move_PDF(plantid, latest_raw):
    filename = Era[latest_raw]
    toPath = plantDirectory[plantid]
    print("Process: PDFをDropboxに移動します")
    cmd = "mv "+filename+".pdf "+toPath
    subprocess.run(cmd, shell=True)


def Init_Selenium():
    options = ChromeOptions()
    options.add_argument('--headless')      #OPTIONAL: Chromeをヘッドレスモードで起動
    options.add_argument('--disable-gpu')   #暫定必須らしい

    #エラーの許容
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-running-insecure-content')
    options.add_argument('--disable-web-security')

    #不要
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-desktop-notifications')

    #便利
    options.add_argument('--start-maximized')

    prefs = {
        "download.default_directory" : DOWNLOAD_PATH,
        "download.prompt_for_download" : False,
        "safebrowsing.enabled" : False,
        "safebrowsing.disable_download_protection" : True
    }
    options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(chrome_options=options)
    return driver



# Python: main関数を呼び出す
if __name__ == '__main__':
    DOWNLOAD_PATH = "./Downloads"
    DRIVER_PATH = "./chromedriver"
    url = "https://www30.tepco.co.jp/dv05s/dfw/biztepco/D3BWwwAP/D3BBTUM00101.act?FW_SCTL=INIT"
    driver = Init_Selenium()

    #プラント辞書
    plant = {
        0 : "//*[@id='myTable']/tbody/tr[1]/td[4]/ul/li[1]/form",
        1 : "//*[@id='myTable']/tbody/tr[2]/td[4]/ul/li[1]/form",
        2 : "//*[@id='myTable']/tbody/tr[3]/td[4]/ul/li[1]/form",
        3 : "//*[@id='myTable']/tbody/tr[4]/td[4]/ul/li[1]/form",
        4 : "//*[@id='myTable']/tbody/tr[5]/td[4]/ul/li[1]/form"
    }
    plantName = {
        0 : "千葉市若葉区小間子町１ー３",
        1 : "千葉市若葉区金親町８３",
        2 : "長生郡白子町発電所１",
        3 : "長生郡白子町発電所２",
        4 : "長生郡白子町発電所３"
    }
    plantDirectory = {
        0 : "~/Dropbox/千葉市　若葉区小間子町１ー３",
        1 : "~/Dropbox/千葉市若葉区金親町８３",
        2 : "~/Dropbox/長生郡白子町発電所１",
        3 : "~/Dropbox/長生郡白子町発電所２",
        4 : "~/Dropbox/長生郡白子町発電所３"
    }
    PLANT_NUM = len(plant)

    #元号
    Era = {
        "平成30年 4月" : "201804",
        "平成30年 5月" : "201805",
        "平成30年 6月" : "201806",
        "平成30年 7月" : "201807",
        "平成30年 8月" : "201808",
        "平成30年 9月" : "201809",
        "平成30年10月" : "201810",
        "平成30年11月" : "201811",
        "平成30年12月" : "201812",
        "平成31年 1月" : "201901",
        "平成31年 2月" : "201902",
        "平成31年 3月" : "201903",
        "平成31年 4月" : "201904"
    }
    main()
