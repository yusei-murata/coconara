# %%
print('プログラム開始')

# %%
input_year = input('取得したいレースの年月日を記入してください'\
    '(例: 20240101)')

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
# Check if the current date is after March 22
#if datetime.date.today() > datetime.date(2024, 4, 8):
#    sys.exit("Program terminated. Date is after April 8, 2024.")

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
def scrape_results(race_id):
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
    df = dfs[0]

    # データの整形
    df = df.T.reset_index(level=0, drop=True).T
    df = df[[2,10]].rename(columns={2: '馬番', 10: 'オッズ',})
    df['race_id'] = [race_id]*len(df)
    df = df[['race_id', '馬番', 'オッズ']]
    df_merge = pd.merge(df[:1].rename(columns={'馬番':'1着_馬番', 'オッズ':'1着_オッズ'}),
                        df[1:2].rename(columns={'馬番':'2着_馬番', 'オッズ':'2着_オッズ'}), on='race_id')
    df_merge[0] = ['枠連']*len(df_merge)
    return df_merge

# %%
def umaren(race_id):
    url = 'https://nar.netkeiba.com/odds/index.html?type=b4&race_id='+race_id+'&housiki=c99'
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")
    tables = soup.find_all('table')

    # StringIOオブジェクトにHTMLをラップ
    html_io = StringIO(str(tables))

    # pandasのDataFrameに変換
    dfs = pd.read_html(html_io, encoding=response.encoding)

    # 最初のテーブルを取得
    df = dfs[0]
    df = df[['組み合わせ','オッズ']]
    df.columns = ['組み合わせ', 'オッズ']
    df_0 = df.copy()
    df_0['race_id'] = [race_id]*len(df_0)

    df_2 = df_0[1:2].copy() #馬連の2番人気のみを抽出
    df_2.loc[:, '馬番_1'] = df_2['組み合わせ'].str.split(' ').str[0]
    df_2.loc[:, '馬番_2'] = df_2['組み合わせ'].str.split(' ').str[1]
    df_3 = df_2[['race_id', '馬番_1', '馬番_2', 'オッズ']].rename(columns={'馬番_1':'2番人気馬連_馬番_1', '馬番_2':'2番人気馬連_馬番_2', 'オッズ':'2番人気馬連_オッズ'})
    df_4 = df_0[2:3].copy() #馬連の3番人気のみを抽出
    df_4.loc[:, '馬番_1'] = df_4['組み合わせ'].str.split(' ').str[0]
    df_4.loc[:, '馬番_2'] = df_4['組み合わせ'].str.split(' ').str[1]
    df_5 = df_4[['race_id', '馬番_1', '馬番_2', 'オッズ']].rename(columns={'馬番_1':'3番人気馬連_馬番_1', '馬番_2':'3番人気馬連_馬番_2', 'オッズ':'3番人気馬連_オッズ'})
    df_merge = pd.merge(df_3, df_5, on=['race_id'])
    df_4 = df_0[3:4].copy() #馬連の4番人気のみを抽出
    df_4.loc[:, '馬番_1'] = df_4['組み合わせ'].str.split(' ').str[0]
    df_4.loc[:, '馬番_2'] = df_4['組み合わせ'].str.split(' ').str[1]
    df_5 = df_4[['race_id', '馬番_1', '馬番_2', 'オッズ']].rename(columns={'馬番_1':'4番人気馬連_馬番_1', '馬番_2':'4番人気馬連_馬番_2', 'オッズ':'4番人気馬連_オッズ'})
    df_merge = pd.merge(df_merge, df_5, on=['race_id'])
    df_4 = df_0[4:5].copy() #馬連の5番人気のみを抽出
    df_4.loc[:, '馬番_1'] = df_4['組み合わせ'].str.split(' ').str[0]
    df_4.loc[:, '馬番_2'] = df_4['組み合わせ'].str.split(' ').str[1]
    df_5 = df_4[['race_id', '馬番_1', '馬番_2', 'オッズ']].rename(columns={'馬番_1':'5番人気馬連_馬番_1', '馬番_2':'5番人気馬連_馬番_2', 'オッズ':'5番人気馬連_オッズ'})
    df_merge = pd.merge(df_merge, df_5, on=['race_id'])
    
    return df_merge


# %%
def scrape(race_id_list):
    return_tables = {}
    for race_id in tqdm.tqdm(race_id_list):
        time.sleep(1)
        try:
            url = "https://nar.netkeiba.com/race/result.html?race_id=" + race_id
            #普通にスクレイピングすると複勝やワイドなどが区切られないで繋がってしまう。
            #そのため、改行コードを文字列brに変換して後でsplitする
            f = urlopen(url)
            html = f.read()
            html = html.replace(b'<br />', b'br')
            dfs = pd.read_html(html)
            #dfsの1番目に単勝〜馬連、2番目にワイド〜三連単がある
            df = pd.concat([dfs[1], dfs[2]])
            df['race_id'] = [race_id] * len(df)
            df = df[['race_id', 0, 2]].rename(columns={2: '枠連_配当'})
            #結果を取得
            results_data = scrape_results(race_id)
            #馬連を取得
            umaren_df = umaren(race_id)
            #race_idと枠連でマージ
            df_merge = pd.merge(results_data,df,on=['race_id',0],how='left')
            df_merge = pd.merge(df_merge,umaren_df,on=['race_id'],how='left')
            df_merge['開催場所'] = place_dict[race_id[4:6]]
            df_merge['レース'] = df_merge['race_id'].str[-2:]
            df_merge['開催日時'] = str(race_id[:4])+'/'+str(race_id[6:8])+'/'+str(race_id[8:10])
            df_merge.drop([0,'race_id'], axis=1, inplace=True)
            df_merge = df_merge[[ '開催日時',  '開催場所', 'レース','1着_馬番', '1着_オッズ', '2着_馬番', '2着_オッズ', '枠連_配当', 
                        '2番人気馬連_馬番_1','2番人気馬連_馬番_2', '2番人気馬連_オッズ', 
                        '3番人気馬連_馬番_1', '3番人気馬連_馬番_2', '3番人気馬連_オッズ',
                        '4番人気馬連_馬番_1', '4番人気馬連_馬番_2', '4番人気馬連_オッズ',
                        '5番人気馬連_馬番_1', '5番人気馬連_馬番_2', '5番人気馬連_オッズ']]
            return_tables[race_id] = df_merge
            
        except IndexError:
            continue
        except AttributeError: #存在しないrace_idでAttributeErrorになるページもあるので追加
            continue
        except Exception as e:
            print(e)
            break
        except:
            break
        
    #pd.DataFrame型にして一つのデータにまとめる
    return_tables_df = pd.concat([return_tables[key] for key in return_tables])
    return return_tables_df

# %%
print('データ取得中')
df = scrape(only_race_id)

# %%
df['枠連_配当'] = df['枠連_配当'].str.replace('円','').str.replace(',','')

# %%
df['枠連_配当'] = pd.to_numeric(df['枠連_配当'], errors='coerce')

# %%
list = ['レース',
        '2番人気馬連_馬番_1', '2番人気馬連_馬番_2',
        '3番人気馬連_馬番_1', '3番人気馬連_馬番_2',
        '4番人気馬連_馬番_1', '4番人気馬連_馬番_2',
        '5番人気馬連_馬番_1', '5番人気馬連_馬番_2']

# %%
for data in list:
    df[data] = df[data].astype(int)

# %%
input_year

# %%
df.to_excel(input_year+'_data.xlsx', index=False)


