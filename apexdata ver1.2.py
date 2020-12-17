import openpyxl
import datetime
import time
#from colorama import Fore, Back, Style
#from openpyxl.drawing.image import PIL
#from multiprocessing.resource_tracker import _posixshmem
#from openpyxl.xml import defusedxml
#from openpyxl.xml import lxml
#from openpyxl.compat.numbers import numpy
#from openpyxl.reader.excel import tests     
#from openpyxl.cell.cell import pandas
#from logging.handkers import resourve 
#from cmd import readline
#from logging.handlers import win32evtlog 
#from logging.handlers import win32evtlogutil 

#レジェンド選択関数
def select_legend(lengendname):
  if legendname == "ブラッドハウンド" or legendname == "ブラハ":
    legend = "bloodhound"
  elif legendname == "ジブラルタル" or legendname == "ジブ":
    legend = "gibraltar"
  elif legendname == "ライフライン" or legendname == "ライフラ":
    legend = "LIFELINE"
  elif legendname == "パスファインダー" or legendname == "パスファ":
    legend = "pathfinder"
  elif legendname == "レイス":
    legend = "wraith"
  elif legendname == "バンガロール" or legendname == "バンガ":
    legend = "bangalore"
  elif legendname == "コースティック" or legendname == "ガスおじ" or legendname == "ガス":
    legend = "caustic"
  elif legendname == "ミラージュ":
    legend = "mirage"
  elif legendname == "オクタン":
    legend = "octane"
  elif legendname == "ワットソン" or legendname == "ワット":
    legend = "wattson"
  elif legendname == "クリプト":
    legend = "crypto"
  elif legendname == "レブナント" or legendname == "レブ" or legendname == "デス" or legendname == "レヴナント" or legendname == "レヴ":
    legend = "revenant"
  elif legendname == "ローバ":
    legend = "loba"
  elif legendname == "ランパート" or legendname == "ランパ":
    legend = "rampart"
  elif legendname == "ホライゾン" or legendname == "ホラさん":
    legend = "horizon"
  else:
    print("デフォルトでレイスになります")
    legend = "wraith"

  return legend


print("ようこそ")
#エラー回避用読み込み

#無限ループ（複数回用）フラグ
flag = True

while(flag):

  apex_dataxlsx = openpyxl.load_workbook("apex_data_casual.xlsx")
  apex_data = apex_dataxlsx.active
  filename = "apex_data_casual.xlsx"
  #マッチ選択
  print("マッチを入力 カジュアルならc ランクならrを入力")
  match = input()
  if match == "c": #カジュアル
    print("シートを読み込んでいます")
    time.sleep(1)
  elif match == "r": #ランク
    apex_dataxlsx.close()
    print("シートを読み込んでいます")
    time.sleep(1)
    apex_dataxlsx = openpyxl.load_workbook("apex_data_rank.xlsx")
    apex_data = apex_dataxlsx.active 
    filename = "apex_data_rank.xlsx" 


  #TODO データ読み込み
  #行カウンター
  row = 2
  #もし現在地にデータがあったら改行
  print("行を調整しています")
  time.sleep(1)
  while not apex_data["A"+str(row)].value == None:
    row = row + 1

  #日付時刻の取得
  dateandtime = datetime.datetime.now() 

  year = dateandtime.year
  month = dateandtime.month
  day = dateandtime.day
  hour = dateandtime.hour
  minute = dateandtime.minute

  #日付
  apex_data["A"+str(row)] = datetime.date(year, month, day)
  #時間
  apex_data["B"+str(row)] = datetime.time(hour,minute)

  #マップ
  print("マップの入力 ワールズエッジならw オリンパスならo")
  map_select = input()
  if map_select == "w":
    apex_data["C"+str(row)] = "world's edge"
  elif map_select == "o":
    apex_data["C"+str(row)] = "olympus"
  else:
    print("マップ名の選択ができませんでした")

  #キル数
  print("キル数の入力(整数)")
  apex_data["D"+str(row)] = int(input())

  #与ダメ
  print("与ダメの入力(整数)")
  apex_data["E"+str(row)] = int(input())

  #生存時間
  print("生存時間の入力(分)")
  survivalminute = int(input())
  print("生存時間の入力(秒)")
  survivalsecond = int(input())
  apex_data["F"+str(row)] = datetime.time(0,survivalminute,survivalsecond)

  #レジェンド
  print("レジェンドの入力(カタカナ)　デフォルトでレイスになっています。")
  legendname = input()
  legend = select_legend(legendname)
  print(legend)

  apex_data["G"+str(row)] = legend

  #順位
  print("順位の入力")
  rank = int(input())
  apex_data["H"+str(row)] = rank

  #残り部隊
  if rank == 1:
    print("部隊数入力中")
    apex_data["I"+str(row)] = 1
    time.sleep(2)
  else:
    print("部隊数数の入力(整数)")
    apex_data["I"+str(row)] = int(input())

  #残り人数
  if rank == 1:
    print("残り人数入力中")
    apex_data["J"+str(row)] = 1
    time.sleep(2)
  else:
    print("残り人数の入力(整数)")
    apex_data["J"+str(row)] = int(input())

  #死亡場所


  #TODO 保存
  apex_dataxlsx.save(filename)
  time.sleep(1)
  print("正常に入力が完了しました。")
  time.sleep(1)

  print("まだ入力を続けますか？ yes or no")
  answer = input()
  
  #選択フラグ
  select_flag = True
  while select_flag:
    if answer == "Yes" or answer == "YES" or answer == "yes" or answer == "y":
      flag = True
      select_flag = False
    elif answer == "No" or answer == "NO" or answer == "no" or answer == "n":
      flag = False
      select_flag = False
    else:
      print("yesかnoで答えてください。")


print("終了します。")
time.sleep(2)