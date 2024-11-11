# %%
print('プログラム開始')

# %%
import pandas as pd
import datetime
import requests
from bs4 import BeautifulSoup
import re
import time
import openpyxl
from io import StringIO
import sys
import tqdm
from urllib.request import urlopen

# %%
today = datetime.datetime.today()
today_year = str(today.year)
today_month = str(today.month)
today_monthz = str(today.month).zfill(2)
today_day = str(today.day)
today_dayz = str(today.day).zfill(2)
today_today = today_year +"/"+ today_month +"/"+ today_day
today_todayz = today_year +"/"+ today_monthz +"/"+ today_dayz
today_today2 = today_year + today_monthz + today_dayz
jra = today_year +'/'+ today_year +"/"+ today_month +"/"+ today_monthz + today_dayz
today_tw = today_month + '/' + today_day
today_nt = today_monthz + today_dayz

# %%
"""
if today.month >= 9 and today.day >= 22:
    raise SystemExit("今日の日付が9/22以降です。強制終了します。")
"""

# %%
input_year = input('取得したいレースの年を記入してください'\
    '(例: 20240101)')

# %%
# %%
place_list = [
    '門別', '盛岡', '水沢', '浦和', '船橋', '大井', '川崎', '金沢', '笠松', '名古屋', '名古屋', '姫路', '高知', '佐賀', '帯広(ば)'
]

# %%
place_dict = {
    '30':'門別',  '35':'盛岡',  '36':'水沢',  '42':'浦和',  '43':'船橋', '44':'大井',
    '45':'川崎',  '46':'金沢',  '47':'笠松',  '48':'名古屋', '50':'園田', '51':'姫路',
    '54':'高知', '55':'佐賀', '65':'帯広(ば)'
}

# %%
def race_id_get_month(race_year, race_month):
    race_id_list = []
    #開催一覧からレースID（レース数を除く）レース日の取得
    url = 'https://nar.netkeiba.com/top/calendar.html?year='+str(race_year)+'&month='+str(race_month)
    html = requests.get(url)
    html.encoding = "EUC-JP"
    soup = BeautifulSoup(html.text, "html.parser")
    race_list_1 = []
    race_list_2 = []
    day_list = []
    for link in soup.find_all('a'):
        if link.get('href').startswith('../top/race_list.html?kaisai_date') == True:
            race_link = link.get('href')[-10:]
            race_day = link.get('href')[34:-21]
            race_list_1.append(race_link[4:6])
            race_list_2.append(race_link)
            day_list.append(race_day)

    #レース数の確認
    race_len_list = []
    for kaijo, day in zip(tqdm.tqdm(race_list_1), day_list):
        url = 'https://db.netkeiba.com/race/sum/'+kaijo+'/'+day+'/'
        html = requests.get(url)
        html.encoding = "EUC-JP"
        soup = BeautifulSoup(html.text, "html.parser")
        len_list = []
        for link in soup.find_all('img'):
            if link.get('src').startswith('/style/netkeiba.ja/image') == True:
                len_list.append(link)
        race_len_list.append(len(len_list))
    #レース数とrace_idをconcatする
    for race, ran in zip(race_list_2,race_len_list):
        for ran_len in range(1,int(ran)+1):
            race_id_list.append(str(race)+str(ran_len).zfill(2))
    
    return race_id_list

# %%
def race_id_get_year(race_year):
    race_id_list = []
    #開催一覧からレースID（レース数を除く）レース日の取得
    for race_month in range(1,13):
        url = 'https://nar.netkeiba.com/top/calendar.html?year='+str(race_year)+'&month='+str(race_month)
        html = requests.get(url)
        html.encoding = "EUC-JP"
        soup = BeautifulSoup(html.text, "html.parser")
        race_list_1 = []
        race_list_2 = []
        day_list = []
        for link in soup.find_all('a'):
            if link.get('href').startswith('../top/race_list.html?kaisai_date') == True:
                race_link = link.get('href')[-10:]
                race_day = link.get('href')[34:-21]
                race_list_1.append(race_link[4:6])
                race_list_2.append(race_link)
                day_list.append(race_day)

        #レース数の確認
        race_len_list = []
        for kaijo, day in zip(tqdm.tqdm(race_list_1), day_list):
            url = 'https://db.netkeiba.com/race/sum/'+kaijo+'/'+day+'/'
            html = requests.get(url)
            html.encoding = "EUC-JP"
            soup = BeautifulSoup(html.text, "html.parser")
            len_list = []
            for link in soup.find_all('img'):
                if link.get('src').startswith('/style/netkeiba.ja/image') == True:
                    len_list.append(link)
            race_len_list.append(len(len_list))
        #レース数とrace_idをconcatする
        for race, ran in zip(race_list_2,race_len_list):
            for ran_len in range(1,int(ran)+1):
                race_id_list.append(str(race)+str(ran_len).zfill(2))
    
    return race_id_list

# %%
print('レースID取得中')
race_id_list = race_id_get_month(int(input_year[:4]), int(input_year[4:6]))


# %%
only_race_id = []
for data in race_id_list:
    if data[:4] + data[6:8] + data[8:10] == input_year:
        only_race_id.append(data)

# %%
def scrape_results_processore(df, race_id):
    columns_list= ['開催日時',  '開催場所', 'レース']
    data_list = [str(race_id[:4])+'/'+str(race_id[6:8])+'/'+str(race_id[8:10]), place_dict[race_id[4:6]], int(race_id[-2:])]
    for a in range(len(df)):
        if len(df[2][a].split(' ')) >= 2:
            if df[0][a] == '単勝' or df[0][a] == '複勝':
                for data in range(len(df[2][a].split(' '))): #同着の場合の処理
                    columns_list.append(str(df[0][a])+'_'+str(data+1)) #券種の追加
                    columns_list.append(str(df[0][a])+'_'+str(data+1)+'_払い戻し')
                    data_list.append(int(df[1][a].split(' ')[data])) #馬番の追加
                    data_list.append(int(df[2][a].split(' ')[data].replace('円', '').replace(',', ''))/100) #払い戻しの追加
            elif df[0][a] == '枠連' or df[0][a] == '馬連' or df[0][a] == '馬単' or df[0][a] == 'ワイド':
                for data in range(0,len(df[2][a].split(' '))): #同着の場合の処理
                    if data == 0:
                        columns_list.append(str(df[0][a])+'_'+str(data+1)) #券種の追加
                        columns_list.append(str(df[0][a])+'_'+str(data+1)+'_払い戻し')
                        data_list.append(df[1][a].split(' ')[data]+'-'+df[1][a].split(' ')[data+1]) #馬番の追加
                        data_list.append(int(df[2][a].split(' ')[data].replace('円', '').replace(',', ''))/100) #払い戻しの追加
                    elif data == 1:
                        columns_list.append(str(df[0][a])+'_'+str(data+1)+'_払い戻し')
                        columns_list.append(str(df[0][a])+'_'+str(data+1)) #券種の追加
                        data_list.append(int(df[2][a].split(' ')[data].replace('円', '').replace(',', ''))/100) #払い戻しの追加
                        data += 1
                        data_list.append(df[1][a].split(' ')[data]+'-'+df[1][a].split(' ')[data+1]) #馬番の追加
                    else:
                        columns_list.append(str(df[0][a])+'_'+str(data+1)+'_払い戻し')
                        columns_list.append(str(df[0][a])+'_'+str(data+1)) #券種の追加
                        data_list.append(int(df[2][a].split(' ')[data].replace('円', '').replace(',', ''))/100) #払い戻しの追加
                        data += 2
                        data_list.append(df[1][a].split(' ')[data]+'-'+df[1][a].split(' ')[data+1]) #馬番の追加
                    
            elif df[0][a] == '3連複' or df[0][a] == '3連単':
                for data in range(0,len(df[2][a].split(' ')),3):
                    columns_list.append(str(df[0][a])+'_'+str(data+1)) #券種の追加
                    columns_list.append(str(df[0][a])+'_'+str(data+1)+'_払い戻し')
                    data_list.append(df[1][a].split(' ')[data]+'-'+df[1][a].split(' ')[data+1]+'-'+df[1][a].split(' ')[data+2]) #馬番の追加
                    data_list.append(int(df[2][a].split(' ')[data].replace('円', '').replace(',', ''))/100) #払い戻しの追加

        else:
            if df[0][a] == '単勝' or df[0][a] == '複勝':
                columns_list.append(str(df[0][a])) #券種の追加
                columns_list.append(str(df[0][a])+'_払い戻し') #払い戻しの追加
                data_list.append(int(df[1][a])) #馬番の追加
                data_list.append(int(df[2][a].replace('円', '').replace(',', ''))/100) #払い戻しの追加
            elif df[0][a] == '枠連' or df[0][a] == '馬連' or df[0][a] == '馬単':
                columns_list.append(str(df[0][a])) #券種の追加
                columns_list.append(str(df[0][a])+'_払い戻し') #払い戻しの追加
                data_list.append(df[1][a].replace(' ', '-')) #馬番の追加
                data_list.append(int(df[2][a].replace('円', '').replace(',', ''))/100) #払い戻しの追加
            elif df[0][a] == '3連複' or df[0][a] == '3連単':
                columns_list.append(str(df[0][a])) #券種の追加
                columns_list.append(str(df[0][a])+'_払い戻し') #払い戻しの追加
                data_list.append(df[1][a].replace(' ', '-')) #馬番の追加
                data_list.append(int(df[2][a].replace('円', '').replace(',', ''))/100) #払い戻しの追加
    return columns_list, data_list

# %%
def scrape_results(race_id):
    try:
        url = "https://nar.netkeiba.com/race/result.html?race_id=" + race_id
        #メインとなるテーブルデータを取得
        response = requests.get(url)
        soup = BeautifulSoup(response.content, "html.parser")
        tables = soup.find_all('table')

        # StringIOオブジェクトにHTMLをラップ
        html_io = StringIO(str(tables))

        # pandasのDataFrameに変換
        dfs = pd.read_html(html_io, encoding=response.encoding)

        # 最初のテーブルを取得
        df_1 = dfs[1]
        df_2 = dfs[2]
        df_concat = pd.concat([df_1, df_2], axis=0).reset_index(drop=True)
        
        columns_list, data_list = scrape_results_processore(df_concat, race_id)
        df = pd.DataFrame([data_list], columns=columns_list)
    except ImportError:
        df = pd.DataFrame()
    return df

# %%
print('レース結果取得中')
list = []
for race_id in tqdm.tqdm(only_race_id):
    list.append(scrape_results(race_id))

# %%
columns = ['開催日時', '開催場所', 'レース', 
        '単勝', '単勝_払い戻し', '単勝_1', '単勝_1_払い戻し', '単勝_2','単勝_2_払い戻し', 
        '複勝_1', '複勝_1_払い戻し', '複勝_2', '複勝_2_払い戻し', '複勝_3', '複勝_3_払い戻し', '複勝_4', '複勝_4_払い戻し',
        '枠連', '枠連_払い戻し', '枠連_1', '枠連_1_払い戻し', '枠連_2', '枠連_2_払い戻し',
        '馬連', '馬連_払い戻し', '馬連_1', '馬連_1_払い戻し', '馬連_2', '馬連_2_払い戻し',
        'ワイド_1', 'ワイド_1_払い戻し', 'ワイド_2', 'ワイド_2_払い戻し', 'ワイド_3', 'ワイド_3_払い戻し', 'ワイド_4', 'ワイド_4_払い戻し', 'ワイド_5', 'ワイド_5_払い戻し',
        '馬単', '馬単_払い戻し', '馬単_1', '馬単_1_払い戻し', '馬単_2', '馬単_2_払い戻し',
        '3連複', '3連複_払い戻し','3連複_1', '3連複_1_払い戻し',
        '3連単', '3連単_払い戻し','3連単_1', '3連単_1_払い戻し',]
df_columns = pd.DataFrame(columns=columns)

# %%
df_data = pd.concat(list, axis=0)

# %%
df_cocnat = pd.concat([df_columns, df_data], axis=0)

# %%
df_cocnat.to_excel('results_'+input_year+'.xlsx', index=False)


