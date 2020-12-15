import openpyxl
import datetime
import time

print("ようこそ")
#エラー回避用読み込み
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
while(not apex_data["A"+str(row)].value == None):
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
#キル数
print("キル数の入力(整数)")
apex_data["C"+str(row)] = int(input())

#与ダメ
print("与ダメの入力(整数)")
apex_data["D"+str(row)] = int(input())

#生存時間
print("生存時間の入力(分)")
survivalminute = int(input())
print("生存時間の入力(秒)")
survivalsecond = int(input())
apex_data["E"+str(row)] = datetime.time(0,survivalminute,survivalsecond)

#レジェンド
print("レジェンドの入力(文字列)")
apex_data["F"+str(row)] = input()

#順位
print("順位の入力")
rank = int(input())
apex_data["G"+str(row)] = rank

#残り部隊
if rank == 1:
  print("部隊数入力中")
  apex_data["H"+str(row)] = 1
  time.sleep(2)
else:
  print("部隊数数の入力(整数)")
  apex_data["H"+str(row)] = int(input())

#残り人数
if rank == 1:
  print("残り人数入力中")
  apex_data["I"+str(row)] = 1
  time.sleep(2)
else:
  print("部隊数数の入力(整数)")
  apex_data["I"+str(row)] = int(input())

#マップ
print("マップの入力 ワールズエッジならw オリンパスならo")
map_select = input()
if map_select == "w":
  apex_data["J"+str(row)] = "world's edge"
elif map_select == "o":
  apex_data["J"+str(row)] = "olympus"
else:
  print("マップ名の選択ができませんでした")

#死亡場所


#TODO 保存
apex_dataxlsx.save(filename)
time.sleep(1)
print("正常に入力が完了しました。")