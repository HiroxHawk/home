import pyautogui as pg
import sys
import os
import time
import subprocess
import win32gui
import win32process
import win32api
import pyperclip
from plyer import notification

#分析シークエンス関数定義>>>>>>
def anlSeq():
    #アナライズ選択
    pg.click(90,435)
    time.sleep(2)
    #自動ターゲット選択
    pg.click(443,175)
    time.sleep(3)
    pg.click(443,262)
    time.sleep(3)
    pg.click(443,353)
    time.sleep(3)
    pg.click(443,523)
    time.sleep(3)
    #データベース選択
    pg.click(420,70)
    time.sleep(1)
    pg.click(595,317)
    time.sleep(1)
    pg.click(640,524)
    time.sleep(1)
    #解析開始
    pg.click(330,720)
    time.sleep(60)
    #分析結果マッチング ※これ以降画像解析を使わないとだめかも
    pg.click(585,70)
    time.sleep(1)
    pg.click(770,532)
    time.sleep(10)
    #マッチング有効化
    pg.click(928,140)
    time.sleep(1)
    #送信頻度最適化
    pg.click(815,138)
    time.sleep(1)
    #送信期間設定
    pg.click(400,140)
    time.sleep(1)
    pg.click(513,222)
    time.sleep(1)
    pg.click(513,222)
    time.sleep(1)
    pg.click(496,117)
    time.sleep(1)
    #設定保存
    pg.click(1139,733)
    pg.click(1200,744)
    time.sleep(4)
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

#レベルウェイブ関数定義>>>>>>
def lwSeq():
    #自動アップデート中警告表示
    notification.notify(
        title='レベルウェイブ自動更新',
        message='ただ今、レベルウェイブ更新中です。\nマウスを動かさないでください。\n更新終了までいましばらくお待ちください。',
        app_name='python',
    )

    #最適化モードからの離脱--------------------
    pg.moveTo(63,55,5)
    pg.click(63,55)
    print("最適化モードから離脱しました。")

    #クライアント選択--------------------------
    pg.moveTo(100,160,1)
    pg.click(100,160)
    print("クライアント選択に移りました。")
    pg.moveTo(250,110,1)
    pg.click(250,110)#カテゴリ位置初期化
    #クライアントカテゴリ判定
    try:
        x,y = pg.locateCenterOnScreen("C:/Users/TimeWaver/Desktop/カテゴリ01.png")
        pg.click(x,y)
    except Exception as ex:
        print("対象が見つかりませんでした。")
        print(ex)
    pg.moveTo(500,155,1)
    pg.click(500,155)
    time.sleep(3)

    #レベルウェイブモードへ移行-----------------
    pg.click(95,100)
    time.sleep(3)

    #共鳴中断プロセス--------------------------
    pg.click(90,130)
    pg.click(70,180)
    time.sleep(2)
    #分析シークエンス
    anlSeq()
    #共鳴増幅プロセス--------------------------
    pg.click(90,130)
    pg.click(68,154)
    time.sleep(2)
    #分析シークエンス
    anlSeq()
    #プロフェッショナルモードへの回帰-----------
    pg.click(95,100)
    time.sleep(3)

    #最適化モードへの回帰----------------------
    pg.click(x=88, y=388)

    #自動アップデート作業終了表示
    notification.notify(
        title='レベルウェイブ自動更新終了',
        message='レベルウェイブ更新が終了しました。\nただいまより、操作が可能です。\nご協力ありがとうございました。',
        app_name='python',
    )

#レベルウェイブ実行頻度
for i in range(11):#←実行したい回数-1の値を入力
    lwSeq()
    day = 15#←何日ごと
    interval = 60*60*24*day
    time.sleep(interval)
    #15日ごとに12回、半年間の設定
