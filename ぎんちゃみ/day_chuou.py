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
from msedge.selenium_tools import Edge, EdgeOptions
import getpass
from selenium.webdriver.common.by import By
from selenium import webdriver
user_name = getpass.getuser()
# Edgeのオプション
options = EdgeOptions()
options.use_chromium = True
options.add_argument('--log-level=3')  # コンソールログのレベルを設定
options.add_experimental_option('excludeSwitches', ['enable-logging'])  # ログの無効化
options.add_argument('--headless')

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
jra = input_year[:4]+"/"+input_year[:4]+"/"+str(int(input_year[4:6])).zfill(1)+"/"+input_year[4:6]+input_year[6:8]
today_tw = today_month + '/' + today_day
today_nt = today_monthz + today_dayz

# %%
print('レースID取得開始')
#JRAからrace_idを取得してスクレイピングする
#race_idがなければプログラムを終える
jra_url = 'https://www.jra.go.jp/keiba/calendar'+ jra +'.html'
res = requests.get(jra_url)
res.encoding = res.apparent_encoding
soup = BeautifulSoup(res.text,'html.parser')

list=[]
list_place=[]
list_replace=[]
list_kai=[]
list_place_kai = []
list_hm=[]
list_hmc=[]
list_place_baken=[]
race_time_list = []
subscribers_time = soup.find_all('td',attrs = {'class':'time'})
subscribers = soup.find_all('div',attrs = {'class':'main'})
subscribers7 = subscribers[7:]

for a in subscribers7:
    b = a.text.split('回')
    list_place_baken.append(b[-1])
    
for a in subscribers7:
    b = a.text.split('回')
    list_kai.append(b[0])
    list_place.append(b[-1])

#レース時間
for a in subscribers_time:
    b = a.text
    c = b.replace('時',':').replace('分','')
    d = today_today + ' ' +  c
    e = datetime.datetime.strptime(d,'%Y/%m/%d %H:%M')
    race_time_list.append(e)
    f = e - datetime.timedelta(minutes=20)
    list_hmc.append(f)


if len(list_hmc) == 0:
    print('該当なし')
    sys.exit()
    
list_rekai = [a.replace('1','01')\
    .replace('2','02')\
    .replace('3','03')\
    .replace('4','04')\
    .replace('5','05')\
    .replace('6','06')\
    .replace('7','07')\
    .replace('8','08')\
    .replace('9','09')\
    .replace('10','10')\
    .replace('11','11')\
    .replace('12','12')for a in list_kai]

for a in list_place:
    if a.endswith('11日') == True or a.endswith('12日') ==True:
        b = a.replace('札幌','01')\
                .replace('函館','02')\
                .replace('福島','03')\
                .replace('新潟','04')\
                .replace('東京','05')\
                .replace('中山','06')\
                .replace('中京','07')\
                .replace('京都','08')\
                .replace('阪神','09')\
                .replace('小倉','10')\
                .replace('3日','03')\
                .replace('4日','04')\
                .replace('5日','05')\
                .replace('6日','06')\
                .replace('7日','07')\
                .replace('8日','08')\
                .replace('9日','09')\
                .replace('10日','10')\
                .replace('11日','11')\
                .replace('12日','12')
        list_replace.append(b)
    else:
        b = a.replace('札幌','01')\
            .replace('函館','02')\
            .replace('福島','03')\
            .replace('新潟','04')\
            .replace('東京','05')\
            .replace('中山','06')\
            .replace('中京','07')\
            .replace('京都','08')\
            .replace('阪神','09')\
            .replace('小倉','10')\
            .replace('1日','01')\
            .replace('2日','02')\
            .replace('3日','03')\
            .replace('4日','04')\
            .replace('5日','05')\
            .replace('6日','06')\
            .replace('7日','07')\
            .replace('8日','08')\
            .replace('9日','09')\
            .replace('10日','10')
        list_replace.append(b)

list_place_kai = [a[:2]+b+a[2:] for a,b in zip(list_replace,list_rekai)]

race_id_list = []
for b in range(0,4): 
    if int(b) != len(list_place_kai) and int(b) < len(list_place_kai):
        for r in range(1,13,1):
            race_id = today_year + list_place_kai[int(b)] + str(r).zfill(2)
            race_id_list.append(race_id)
    else:
        pass
print('レースID取得完了')

# %%
# Check if the current date is after March 22
#if datetime.date.today() > datetime.date(2024, 4, 8):
#    sys.exit("Program terminated. Date is after April 8, 2024.")

# %%
place_list = [
    '札幌', '函館', '福島', '新潟', '東京', '中山', '中京', '京都', '阪神', '小倉'
]

# %%
place_dict = {
    '01':'札幌',  '02':'函館',  '03':'福島',  '04':'新潟',  '05':'東京',
    '06':'中山',  '07':'中京',  '08':'京都',  '09':'阪神',  '10':'小倉',
}

# %%
def race_id_get_month(race_year, race_month):
    race_id_list = []
    race_day_list = []
    url = 'https://race.netkeiba.com/top/calendar.html?year='+str(race_year)+'&month='+str(race_month)
    html = requests.get(url)
    html.encoding = "EUC-JP"
    soup = BeautifulSoup(html.text, "html.parser")
    kaisaiday_list = []
    for link in soup.find_all('a'):
        if link.get('href').startswith('../top/race_list.html?kaisai_date') == True:
            race_link = link.get('href')[-8:]
            kaisaiday_list.append(race_link)
    for kaisai in kaisaiday_list:
        url = 'https://db.netkeiba.com/race/list/'+kaisai
        html = requests.get(url)
        html.encoding = "EUC-JP"
        soup = BeautifulSoup(html.text, "html.parser")
        kaisai_day = kaisai[:4]+'/'+kaisai[4:6]+'/'+kaisai[6:8]
        for link in soup.find_all('a'):
            if link.get('href').startswith('/race/movie') == True:
                race_id = link.get('href')[-12:]
                race_day_list.append(kaisai_day)
                race_id_list.append(race_id)
    return race_id_list, race_day_list

# %%
def race_id_get(race_year):
    race_id_list = []
    for race_month in range(1,13):
        url = 'https://race.netkeiba.com/top/calendar.html?year='+str(race_year)+'&month='+str(race_month)
        html = requests.get(url)
        html.encoding = "EUC-JP"
        soup = BeautifulSoup(html.text, "html.parser")
        kaisaiday_list = []
        for link in soup.find_all('a'):
            if link.get('href').startswith('../top/race_list.html?kaisai_date') == True:
                race_link = link.get('href')[-8:]
                kaisaiday_list.append(race_link)
        race_day_list = []
        for kaisai in kaisaiday_list:
            url = 'https://db.netkeiba.com/race/list/'+kaisai
            html = requests.get(url)
            html.encoding = "EUC-JP"
            soup = BeautifulSoup(html.text, "html.parser")
            kaisai_day = kaisai[:4]+'/'+kaisai[4:6]+'/'+kaisai[6:8]
            for link in soup.find_all('a'):
                if link.get('href').startswith('/race/movie') == True:
                    race_id = link.get('href')[-12:]
                    race_day_list.append(kaisai_day)
                    race_id_list.append(race_id)
    return race_id_list

# %%
def scrape_results(race_id):
    url = "https://race.netkeiba.com/race/result.html?race_id=" + race_id
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
browser = Edge(executable_path='C:\\Users\\'+str(user_name)+'\\Downloads\\msedgedriver\\msedgedriver.exe',options=options)
browser.implicitly_wait(11)  # タイムアウト値を整数または浮動小数点数に更新

# %%
def umaren(race_id):
    url = 'https://race.netkeiba.com/odds/index.html?type=b4&race_id='+race_id+'&housiki=c99'
    browser.get(url)
    umaban1_data = browser.find_elements(By.ID, 'umaban_1')[:5]
    umaban2_data = browser.find_elements(By.ID, 'umaban_2')[:5]
    kaime_list = []
    for a,b in zip(umaban1_data,umaban2_data):
        kaime_list.append(a.text+'-'+b.text)
    odds_data = browser.find_elements(By.CLASS_NAME, 'Name_Odds')[:5]
    odds_list = []
    for a in odds_data:
        odds_list.append(a.text)
    # 最初のテーブルを取得
    df = pd.DataFrame({'組み合わせ':kaime_list, 'オッズ':odds_list})
    df = df[['組み合わせ','オッズ']]
    df.columns = ['組み合わせ', 'オッズ']
    df_0 = df.copy()
    df_0['race_id'] = [race_id]*len(df_0)

    df_2 = df_0[1:2].copy() #馬連の2番人気のみを抽出
    df_2.loc[:, '馬番_1'] = df_2['組み合わせ'].str.split('-').str[0]
    df_2.loc[:, '馬番_2'] = df_2['組み合わせ'].str.split('-').str[1]
    df_3 = df_2[['race_id', '馬番_1', '馬番_2', 'オッズ']].rename(columns={'馬番_1':'2番人気馬連_馬番_1', '馬番_2':'2番人気馬連_馬番_2', 'オッズ':'2番人気馬連_オッズ'})
    df_4 = df_0[2:3].copy() #馬連の3番人気のみを抽出
    df_4.loc[:, '馬番_1'] = df_4['組み合わせ'].str.split('-').str[0]
    df_4.loc[:, '馬番_2'] = df_4['組み合わせ'].str.split('-').str[1]
    df_5 = df_4[['race_id', '馬番_1', '馬番_2', 'オッズ']].rename(columns={'馬番_1':'3番人気馬連_馬番_1', '馬番_2':'3番人気馬連_馬番_2', 'オッズ':'3番人気馬連_オッズ'})
    df_merge = pd.merge(df_3, df_5, on=['race_id'])
    df_4 = df_0[3:4].copy() #馬連の4番人気のみを抽出
    df_4.loc[:, '馬番_1'] = df_4['組み合わせ'].str.split('-').str[0]
    df_4.loc[:, '馬番_2'] = df_4['組み合わせ'].str.split('-').str[1]
    df_5 = df_4[['race_id', '馬番_1', '馬番_2', 'オッズ']].rename(columns={'馬番_1':'4番人気馬連_馬番_1', '馬番_2':'4番人気馬連_馬番_2', 'オッズ':'4番人気馬連_オッズ'})
    df_merge = pd.merge(df_merge, df_5, on=['race_id'])
    df_4 = df_0[4:5].copy() #馬連の5番人気のみを抽出
    df_4.loc[:, '馬番_1'] = df_4['組み合わせ'].str.split('-').str[0]
    df_4.loc[:, '馬番_2'] = df_4['組み合わせ'].str.split('-').str[1]
    df_5 = df_4[['race_id', '馬番_1', '馬番_2', 'オッズ']].rename(columns={'馬番_1':'5番人気馬連_馬番_1', '馬番_2':'5番人気馬連_馬番_2', 'オッズ':'5番人気馬連_オッズ'})
    df_merge = pd.merge(df_merge, df_5, on=['race_id'])
    
    return df_merge


# %%
def scrape(race_id_list):
    return_tables = {}
    for race_id in tqdm.tqdm(race_id_list):
        time.sleep(1)
        try:
            url = "https://race.netkeiba.com/race/result.html?race_id=" + race_id
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
            df_merge['開催日時'] = str(input_year[:4])+'/'+str(input_year[4:6])+'/'+str(input_year[6:8])
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
df = scrape(race_id_list)

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
list_odds = ['2番人気馬連_オッズ',
            '3番人気馬連_オッズ',
            '4番人気馬連_オッズ',
            '5番人気馬連_オッズ']

# %%
for data in list:
    df[data] = df[data].astype(int)

# %%
for data in list_odds:
    df[data] = df[data].astype(float)

# %%
df.to_excel(input_year+'_data.xlsx', index=False)


