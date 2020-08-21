import requests
from urllib.parse import urljoin
from selenium import webdriver
from tqdm import tqdm
import urllib.request
import time
import os

#画像保存用ファイルを作成し、自社サイトの商品画像を取得して商品番号の名前でファイルに保存する。


url_list = []
title_list = []
race_urls = []

#フォルダ作成
folder= 'SeminarPhoto'
os.makedirs(folder, exist_ok=True)

#executable_path= クロームドライバーの格納場所を入力
driver = webdriver.Chrome(executable_path='E:/chromedriver.exe')
#range(1,X) Xには商品一覧ページの総数+1の値を代入（1から15ページまである場合はrange(1,16）
for i in range(1,7):
    #セミナーの場合：http://hikarulandpark.jp/shopbrand/001/X/page{}/recommend/
    #グッズの場合：http://hikarulandpark.jp/shopbrand/002/X/page{}/recommend/
    URL = 'http://hikarulandpark.jp/shopbrand/001/X/page{}/recommend/'.format(i)
    driver.get(URL)
    time.sleep(5)
    #リストの初期化
    url_list.clear()
    title_list.clear()
    race_urls.clear()
    #商品一覧ページから、各商品のURLを取得
    elems_race_url = driver.find_elements_by_class_name("M_cl_name > a")
    for elem_race_url in elems_race_url:
        race_url = elem_race_url.get_attribute('href')
        race_urls.append(race_url)
    #取得した各URLに移動
    for race_url_list in race_urls:
        driver.get(race_url_list)
        #画像を取得
        try:
            urlFact = driver.find_element_by_css_selector("#M_mainContents > div.M_clearfix > div.itemleft > a > img").get_attribute("src")
        except:
            urlFact = driver.find_element_by_xpath("/html/body/div[3]/table/tbody/tr[1]/td[3]/form[2]/div/div[1]/div[4]/div[1]/div/div[1]/a/img").get_attribute("src")
        url_list.append(urlFact)
        print("画像データ取得...SUCCESS!!")
        print("画像URL「" + urlFact + "」")
        #タイトル(商品番号)を取得
        element_titile = driver.find_element_by_xpath("/html/body/div[3]/table/tbody/tr[1]/td[3]/form[2]/div/div[1]/div[4]/div[2]/table[1]/tbody/tr[3]/td").get_attribute("textContent")
        title_list.append(element_titile)
        print("商品番号「" + element_titile + "」取得...SUCCESS!!\n")
        time.sleep(3)
        driver.back()
        time.sleep(3)

    #画像を商品番号の名前で保存
    for n in tqdm (range(len(url_list))):
    	image_url =  url_list[n]
    	print('画像をダウンロード中 {}...'.format(image_url))
    	res = requests.get(image_url)
    	res.raise_for_status()

    	image_file = open(os.path.join(folder, os.path.basename(title_list[n]+'.jpg')), 'wb')
    	for chunk in res.iter_content(100000):
    		image_file.write(chunk)
    	image_file.close()
