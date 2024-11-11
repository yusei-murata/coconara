# %%
import os
import io
import zipfile
import requests
import win32api
import getpass
user_name = getpass.getuser()
# フォルダーからMicrosoft Edgeのバージョンを取得する
edge_link = 'C:\Program Files (x86)\Microsoft\Edge\Application'
new_version = os.listdir(edge_link)[0]
# Microsoft Edge WebDriverのダウンロードURLを生成する
url = f"https://msedgedriver.azureedge.net/{new_version}/edgedriver_win64.zip"
# ダウンロードしたファイルを保存する場所を指定する
driver_path = "C:\\Users\\"+str(user_name)+"\\Downloads\\msedgedriver\\msedgedriver.exe" # ココを自分のedgedriverがあるpathに変更する
# ダウンロードしたedgedriverのバージョン情報を取得する
info = win32api.GetFileVersionInfo(driver_path, "\\")
download_version = f"{info['FileVersionMS'] >> 16}.{info['FileVersionMS'] & 0xffff}.{info['FileVersionLS'] >> 16}.{info['FileVersionLS'] & 0xffff}"
if new_version != download_version:
    # ファイルがすでに存在する場合は削除する
    if os.path.exists(driver_path):
        os.remove(driver_path)
    # Microsoft Edge WebDriverをダウンロードして解凍する
    response = requests.get(url, verify=False)
    z = zipfile.ZipFile(io.BytesIO(response.content))
    z.extractall("C:\\Users\\"+str(user_name)+"\\Downloads\\msedgedriver") # ココを自分のedgedriver.exeがあるpathに変更する
    # ダウンロードしたファイルに実行権限を与える
    os.chmod(driver_path, 0o755)
    print("更新完了")
else:
    print('更新不要')

# %%



