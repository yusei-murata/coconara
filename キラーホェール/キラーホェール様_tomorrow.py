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

# %%
today = datetime.datetime.today()+datetime.timedelta(days=1)
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
kaijou_id_list = ['30', '35', '36', '42',  '43', '44', '45', '46', '47', '48', '50', '51', '54', '55']

# %%
race_id_list = []
for kaijou_id in kaijou_id_list:
    race_url = 'https://nar.netkeiba.com/race/shutuba.html?race_id='+today_year+kaijou_id+today_nt+'01'
    response = requests.get(race_url)
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
    if len(df)>1:
        for a in range(1,13,1):
            race_id_list.append(today_year+kaijou_id+today_nt+str(a).zfill(2))
        

# %%
place_dict = {
    '30':'門別',  '35':'盛岡',  '36':'水沢',  '42':'浦和',  '43':'船橋', '44':'大井',
    '45':'川崎',  '46':'金沢',  '47':'笠松',  '48':'名古屋', '50':'園田', '51':'姫路', '54':'高知', '55':'佐賀'
}

# %%
def race_data_scrape(race_id_list, date):
    df_list = []
    for race_id in race_id_list:
        time.sleep(1)
        url = 'https://nar.netkeiba.com/race/shutuba.html?race_id=' + race_id
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
        if len(df) != 0:
            html = requests.get(url)
            html.encoding = "EUC-JP"
            soup = BeautifulSoup(html.text, "html.parser")

            texts = soup.find('div', attrs={'class': 'RaceData01'}).text
            texts = re.findall(r'\w+', texts)
            for text in texts:
                #if 'm' in text:
                #    df['course_len'] = [int(re.findall(r'\d+', text)[-1])] * len(df) #20211212：[0]→[-1]に修正
                if text in ["曇", "晴", "雨", "小雨", "小雪", "雪"]:
                    df["weather"] = [text] * len(df)
                if text in ["良", "稍重", "重"]:
                    df["ground_state"] = [text] * len(df)
                if '不' in text:
                    df["ground_state"] = ['不良'] * len(df)
                # 2020/12/13追加
                if '稍' in text:
                    df["ground_state"] = ['稍重'] * len(df)
                if '芝' in text:
                    df['race_type'] = ['芝'] * len(df)
                if '障' in text:
                    df['race_type'] = ['障害'] * len(df)
                if 'ダ' in text:
                    df['race_type'] = ['ダート'] * len(df)
            df['date'] = [date] * len(df)

            # horse_id
            horse_id_list = []
            horse_td_list = soup.find_all("td", attrs={'class': 'HorseInfo'})
            for td in horse_td_list:
                horse_id = re.findall(r'\d+', td.find('a')['href'])[0]
                horse_id_list.append(horse_id)

            corse_around = []
            table = soup.find('div',class_ ='RaceData01')
            #for a in table:
            #    b = a.text
            #    if '(' in b :
            #        c = b.split(')')[0]
            #        d = c[2]
            #        corse_around.append(d)


            df["horse_id"] = horse_id_list
            #df['corse_around'] = corse_around* len(df)
            #インデックスをrace_idにする 
            df.index = [race_id] * len(df)
            df['競馬場'] = [place_dict[race_id[4:6]]] * len(df)
            df['レース'] = [race_id[10:]] * len(df)
            df.drop(['印','登録','メモ'],axis=1,inplace=True)
            df['厩舎'] = df['厩舎'].str.split(' ', expand=True)[1]
            df = df.rename(columns={'厩舎':'調教師'})
            df_list.append(df)
    try:
        df_concat = pd.concat(df_list)
    except:
        df_concat = pd.DataFrame()
    return df_concat

# %%
def horse_data_scrape(horse_id_list):

    #horse_idをkeyにしてDataFrame型を格納
    horse_results = []
    for horse_id in horse_id_list:
        time.sleep(1)
        try:
            session = requests.Session()
            url = 'https://db.netkeiba.com/horse/' + horse_id
            res = session.get(url, timeout=(3.0, 7.5))
            df = pd.read_html(res.content)[1].T
            df.columns = df.iloc[0,:].values
            df = df[1:].reset_index(drop=True)
            session = requests.Session()
            url = 'https://db.netkeiba.com/horse/' + horse_id
            res = session.get(url, timeout=(3.0, 7.5))
            df_1 = pd.merge(pd.read_html(res.content)[2][:1], pd.read_html(res.content)[2][1:2], on=0)
            df_1.columns = ['父', '父_父', '父_母']
            df_2 = pd.merge(pd.read_html(res.content)[2][2:3], pd.read_html(res.content)[2][3:4], on=0)
            df_2.columns = ['母', '母_父', '母_母']
            df_concat = pd.concat([df, df_1, df_2], axis=1)
            df_concat['horse_id'] = horse_id
            df_concat = df_concat.drop('調教師',axis=1)
            horse_results.append(df_concat)
        except IndexError:
            continue
        except Exception as e:
            print(e)
            break
        except:
            break

    #pd.DataFrame型にして一つのデータにまとめる        
    horse_results_df = pd.concat(horse_results)

    return horse_results_df

# %%
def data_merge(race_data, horse_data):
    df_merge = pd.merge(race_data, horse_data, on='horse_id', how='left')
    try:
        df_merge = df_merge.drop(['募集情報', 'horse_id'],axis=1)
    except:
        pass
    df_merge = df_merge[['date', '競馬場',  'レース', '馬 番', '生年月日', '馬名', '騎手', '調教師', '馬主']]
    return df_merge

# %%
book = openpyxl.Workbook()

# %%
race_id_only_list = []
for race_id in race_id_list:
    if race_id[4:6] not in race_id_only_list:
        race_id_only_list.append((race_id[4:6]))

# %%
race_id_only_list

# %%
df_merge_list = []
for keibajo in race_id_only_list:
    only_race_data_list = []
    for race_id in race_id_list:
        if keibajo == race_id[4:6]:
            only_race_data_list.append(race_id)
    race_data = race_data_scrape(only_race_data_list, today_todayz)
    if len(race_data) != 0:
        horse_data = horse_data_scrape(race_data['horse_id'].unique())
        df_merge = data_merge(race_data, horse_data)
        df_merge_list.append(df_merge)
    else:
        print('レースデータがありません。')

# %%
with pd.ExcelWriter(today_today2+'.xlsx') as writer:
    for i, data in enumerate(df_merge_list):
        sheet_name = data['競馬場'][0]
        data.to_excel(writer, sheet_name=sheet_name, index=False)


