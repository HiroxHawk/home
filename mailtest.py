# coding:utf-8
import smtplib
import datetime
import time
import schedule
import sys
import msvcrt
import win32com.client
import os
from pywintypes import com_error
from email.mime.text import MIMEText
from email.utils import formatdate
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from os.path import basename

#指定した曜日・時間に、共有フォルダ内にあるセミナースケジュール表のエクセルデータをPDFに変換し、
#配信用GmailBOTアカウントから、全体や指定した連絡先にメールで定期送信する。

#メールの送信先(複数可)↓
send_list = "〇〇@〇〇.co.jp,〇〇@〇〇.co.jp"
#メールに送付したいPDFファイルのパスを入力　ダブルクオーテーションで囲み、カンマで区切る↓
path_list = ["//〇〇share/〇〇共有/セミナースケジュール/自動送信用蓄積ファイル/★セミナー一覧【申込み数】.pdf","//〇〇/share/〇〇共有/セミナースケジュール/セミナールームスケジュール2020.pdf"]
sendtime = "14:00"#定期送信時間

#定期送信ユニット
def seminar_schedule():
    #エクセルからPDFに変換しファイル生成(1)
    WB_PATH = r'//〇〇/share/〇〇共有/セミナースケジュール/セミナールームスケジュール2020.xlsx'
    PATH_TO_PDF = r'//〇〇/share/〇〇共有/セミナースケジュール/自動送信用蓄積ファイル/セミナールームスケジュール2020.pdf'

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        print('PDFへ変換開始')
        # 開く
        wb = excel.Workbooks.Open(WB_PATH)
        # PDF保存したいシートをインデックスで指定。1が最初（一番左）のシート。
        ws_index_list = [1]
        wb.WorkSheets(ws_index_list).Select()
        # 保存
        wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
    except com_error as e:
        print('セミナールームスケジュール2020→PDF変換失敗しました')
    else:
        print('セミナールームスケジュール2020→PDF変換成功しました')
    finally:
        time.sleep(1)
        wb.Close()
        excel.Quit()
    #エクセルからPDFに変換しファイル生成(2)
    WB_PATH2 = r'//〇〇/share/〇〇共有/★セミナー用/★セミナー一覧【会場事前申込用】控えつき.xlsx'
    PATH_TO_PDF2 = r'//〇〇/share/〇〇共有/セミナースケジュール/自動送信用蓄積ファイル/★セミナー一覧【申込み数】.pdf'

    excel2 = win32com.client.Dispatch("Excel.Application")
    excel2.Visible = False
    try:
        print('PDFへ変換開始')
        # 開く
        wb2 = excel2.Workbooks.Open(WB_PATH2)
        # PDF保存したいシートをインデックスで指定。1が最初（一番左）のシート。
        ws_index_list2 = [2]
        wb2.WorkSheets(ws_index_list2).Select()
        # 保存
        wb2.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF2)
    except com_error as e:
        print('★セミナー一覧【申込み数】→PDF変換失敗しました')
    else:
        print('★セミナー一覧【申込み数】→PDF変換成功しました')
    finally:
        time.sleep(1)
        wb2.Close()
        excel2.Quit()
    #メールログイン
    smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpobj.ehlo()
    smtpobj.starttls()
    smtpobj.ehlo()
    smtpobj.login('hikarulandbot@gmail.com', 'landbot_311')

    #メールタイトル・本文
    date = datetime.datetime.now()
    msg = MIMEMultipart()
    #メールタイトルは送信時の日時を反映
    msg['Subject'] = 'セミナールーム申込み最新情報【'+ str(date.year) + '年' + str(date.month) + '月' + str(date.day) + '日更新】'
    msg['From'] = 'hikarulandbot@gmail.com'
    msg['To'] = send_list
    msg['Date'] = formatdate()
    body = MIMEText('みなさま\n\nお疲れ様です。\n〇〇botです。 \nセミナールームの最新情報と申込み状況を添付いたします。 \nご確認をよろしくお願いします 。\nセルの色分けにつきましてはかっこ内をご参照ください。 \nどうぞよろしくお願いいたします。')
    msg.attach(body)
    #ファイル送付
    for path in path_list:

        with open(path, "rb") as f:
            part = MIMEApplication(
                f.read(),
                Name=basename(path)
            )

        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(path)
        msg.attach(part)

    #送信
    sendToList = send_list.split(',')
    smtpobj.sendmail('〇〇bot@gmail.com',sendToList, msg.as_string())
    smtpobj.close()
    #コンソールに送信ログを表示
    print('【'+ str(date.year) + '年' + str(date.month) + '月' + str(date.day) + '日にお知らせメッセージが送信されました】')
    time.sleep(4)
    os.remove('//〇〇/share/〇〇共有/セミナースケジュール/自動送信用蓄積ファイル/★セミナー一覧【申込み数】.pdf')
    os.remove('//〇〇/share/〇〇共有/セミナースケジュール/自動送信用蓄積ファイル/セミナールームスケジュール2020.pdf')

def keystart():
    #schedule.every().week.at(sendtime).do(seminar_schedule)
    schedule.every().thursday.at(sendtime).do(seminar_schedule)　#←曜日を指定する場合
    print('セミナールーム最新情報お知らせメールの自動定期送信プログラム...実行中')
    today = datetime.datetime.now()
    print(str(today.year) + '年' + str(today.month) + '月' + str(today.day) + '日' + str(sendtime) + 'から1週間ごとに自動でメッセージが送信されます。')
    print("送信を停止するには、キーボードのxボタンを押してください。\n")
    while True:
        schedule.run_pending()
        time.sleep(1)
        if msvcrt.kbhit(): # キーが押されているか
            kb = msvcrt.getch()
            if kb.decode() == 'x' :
                print("xボタンが押され、自動送信が停止されました。")
                print("①→再開するにはsボタンを押してください。")
                print("②→このままプログラムを終了するにはcボタンを押してください\n")
                break

print('セミナールーム最新情報お知らせメールの自動定期送信プログラムへようこそ')
print('定期送信を開始するにはキーボードの「s」ボタンを押してください。\n')

while True:
    if msvcrt.kbhit(): # キーが押されているか
        kb = msvcrt.getch()
        if kb.decode() == 's' :# sキーが押されたらスタート
            keystart()
        elif kb.decode() == 'c' :# cキーが押されたらプログラム終了
            print("\nSee you...!!")
            time.sleep(3)
            sys.exit()
