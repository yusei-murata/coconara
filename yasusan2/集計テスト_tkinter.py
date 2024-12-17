# %%
# ライブラリのインポート
import tkinter as tk

# tkオブジェクトの作成
root = tk.Tk()
root.title("回収率計算")  # ウィンドウのタイトルを設定
# ウィンドウを最大化
root.state('zoomed')

# キャンバスとスクロールバーを作成
canvas = tk.Canvas(root)
scrollbar_y = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
scrollbar_x = tk.Scrollbar(root, orient="horizontal", command=canvas.xview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

# スクロールバーをキャンバスに配置
scrollbar_y.pack(side="right", fill="y")
scrollbar_x.pack(side="bottom", fill="x")
canvas.pack(side="left", fill="both", expand=True)

# マウスホイールイベントをバインド
def on_mouse_wheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

canvas.bind_all("<MouseWheel>", on_mouse_wheel)


"""
ここまでがウィジェットの設定
"""


# %%

# ラジオボタンをクリアする関数
def clear_radio_buttons_and_labels():
    kenshu_var.set(None)
    kaikata_var.set(None)
    column_var_1.set(None)
    column_var_2.set(None)
    column_var_3.set(None)
    up_down_var_1.set(None)
    up_down_var_2.set(None)
    up_down_var_3.set(None)
    lower_head_var.set(None)
    maximum_head_var.set(None)
    umaban_1_vars.clear()
    umaban_2_vars.clear()
    umaban_3_vars.clear()
    race_vars.clear()
    #clear_frame(frame_umaban_1)
    #clear_frame(frame_umaban_2)
    #clear_frame(frame_umaban_3)
    #clear_frame(frame_column)
    #clear_frame(frame_up_down)
    #clear_frame(frame_lower_heads)
    #clear_frame(frame_maximum_heads)
    #clear_frame(frame_race)

# フレーム内のウィジェットをクリアする関数
def clear_frame(frame):
    for widget in frame.winfo_children():
        widget.destroy()

def show_selection():
    
    # 新しいウィンドウを作成
    new_window = tk.Toplevel(root)
    new_window.title("選択内容の確認")
    
    # 結果を表示するラベルを作成して配置
    kenshu_label = tk.Label(new_window, text="")
    kenshu_selection = f"券種 : {kenshu_var.get()}"
    kenshu_label.config(text=kenshu_selection)
    kenshu_label.grid(row=0, column=10, columnspan=2, padx=5, pady=10)
    # 馬1の表示
    umaban_label_1 = tk.Label(new_window, text="")
    umaban_selection_1 = "馬1 : " + ", ".join([str(umaban) for umaban, var in umaban_1_vars.items() if var.get()])
    umaban_label_1.config(text=umaban_selection_1)
    umaban_label_1.grid(row=1, column=10, columnspan=2, padx=5, pady=10)
    
    
    
    column_label = tk.Label(new_window, text="")
    column_label.grid(row=1, column=10, columnspan=2, padx=5, pady=10)
    up_down_label = tk.Label(new_window, text="")
    up_down_label.grid(row=2, column=10, columnspan=2, padx=5, pady=10)
    umaban_label_1 = tk.Label(new_window, text="")
    umaban_label_2 = tk.Label(new_window, text="")
    umaban_label_2.grid(row=4, column=10, columnspan=2, padx=5, pady=10)
    umaban_label_3 = tk.Label(new_window, text="")
    umaban_label_3.grid(row=5, column=10, columnspan=2, padx=5, pady=10)
    lower_heads_label = tk.Label(new_window, text="")
    lower_heads_label.grid(row=6, column=10, columnspan=2, padx=5, pady=10)
    maximum_heads_label = tk.Label(new_window, text="")
    maximum_heads_label.grid(row=7, column=10, columnspan=2, padx=5, pady=10)
    race_label = tk.Label(new_window, text="")
    race_label.grid(row=8, column=10, columnspan=2, padx=5, pady=10)
    umaban_selection_1 = "馬1 : " + ", ".join([str(umaban) for umaban, var in umaban_1_vars.items() if var.get()])
    umaban_selection_2 = "馬2 : " + ", ".join([str(umaban) for umaban, var in umaban_2_vars.items() if var.get()])
    umaban_selection_3 = "馬3 : " + ", ".join([str(umaban) for umaban, var in umaban_3_vars.items() if var.get()])
    column_selection_1 = f"馬1_列 : {column_var_1.get()}"
    column_selection_2 = f"馬2_列 : {column_var_2.get()}"
    column_selection_3 = f"馬3_列 : {column_var_3.get()}"
    up_down_selection_1 = f"馬1_昇降順 : {up_down_var_1.get()}"
    up_down_selection_2 = f"馬2_昇降順 : {up_down_var_2.get()}"
    up_down_selection_3 = f"馬3_昇降順 : {up_down_var_3.get()}"
    lower_heads_selection = f"下限頭数 : {lower_head_var.get()}"
    maximum_heads_selection = f"上限頭数 : {maximum_head_var.get()}"
    race_selection = "対象レース : " + ", ".join([str(race) for race, var in race_vars.items() if var.get()])
    
    umaban_label_2.config(text=umaban_selection_2)
    umaban_label_3.config(text=umaban_selection_3)
    column_label_1.config(text=column_selection_1)
    column_label_2.config(text=column_selection_2)
    column_label_3.config(text=column_selection_3)
    up_down_label_1.config(text=up_down_selection_1)
    up_down_label_2.config(text=up_down_selection_2)
    up_down_label_3.config(text=up_down_selection_3)
    lower_heads_label.config(text=lower_heads_selection)
    maximum_heads_label.config(text=maximum_heads_selection)
    race_label.config(text=race_selection)


# %%

#kenshu_list = ["馬単", "三連複"]
kenshu_list = ["三連複"]
kenshu_var = tk.StringVar(value=" ")  # 初期値を設定しない

# ウィジェットの配置や、イベント処理などを記述
kenshu_label = tk.Label(scrollable_frame, text="券種を選んでください")
kenshu_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

# メインラジオボタンを横に並べるためのフレーム
main_frame = tk.Frame(scrollable_frame)
main_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W)
for index, kenshu in enumerate(kenshu_list):
    radio = tk.Radiobutton(main_frame, text=kenshu, variable=kenshu_var, value=kenshu)
    radio.grid(row=0, column=index, padx=5, pady=5)

# %% [markdown]
# # 馬1

# %%
# 馬番の選択肢を更新
umaban_label_1 = tk.Label(scrollable_frame, text="馬1を選んでください")
umaban_label_1.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)  # sticky=tk.Wを追加
umaban_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
umaban_1_vars = {}  # 馬番の1着を選択するチェックボックス用の変数を保持する辞書
frame_umaban_1 = tk.Frame(scrollable_frame)
frame_umaban_1.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=10)
# サブ選択肢フレームのクリア
for widget in frame_umaban_1.winfo_children():
    widget.destroy()
# 馬番の選択肢を追加
for i, umaban in enumerate(umaban_list):
    var = tk.BooleanVar()
    umaban_1_vars[umaban] = var
    check = tk.Checkbutton(frame_umaban_1, text=umaban, variable=var)
    if umaban < 10:
        check.grid(row=0, column=umaban-1)  # 1〜9は1行目
    else:
        check.grid(row=1, column=umaban-10)  # 10〜18は2行目

"""
列の選択
"""
# 馬1の列を選択
column_label_1 = tk.Label(scrollable_frame, text="馬1の列を選んでください")
column_label_1.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)  # sticky=tk.Wを追加
column_list = ['着順', '馬番', '馬齢', 'オッズ', '斤量', '脚質', '総合値',
               'SP値', 'AG値', 'SA値', '馬連率', '戦数', '賞金平均', 'KI値']
# 買い方の選択肢の追加
column_var_1 = tk.StringVar(value=" ")  # 初期値を設定しない
frame_column_1 = tk.Frame(scrollable_frame)
frame_column_1.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=10)
# サブ選択肢フレームのクリア
for widget in frame_column_1.winfo_children():
    widget.destroy()
for i, option in enumerate(column_list):
    radio = tk.Radiobutton(frame_column_1, text=option, variable=column_var_1, value=option)
    if i < 7:
        radio.grid(row=0, column=i, padx=5, pady=5)
    else:
        radio.grid(row=1, column=i-7, padx=5, pady=5)

"""
昇降順の選択
"""
up_down_label_1 = tk.Label(scrollable_frame, text="馬1の列の昇降順を選んでください")
up_down_label_1.grid(row=6, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)  # sticky=tk.Wを追加
up_down_list = ['昇順', '降順']
up_down_var_1 = tk.StringVar(value=" ")  # 初期値を設定しない
frame_up_down_1 = tk.Frame(scrollable_frame)
frame_up_down_1.grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=10)
# サブ選択肢フレームのクリア
for widget in frame_up_down_1.winfo_children():
    widget.destroy()
for i, option in enumerate(up_down_list):
    radio = tk.Radiobutton(frame_up_down_1, text=option, variable=up_down_var_1, value=option)
    radio.grid(row=0, column=i, padx=5, pady=5)


end_label = tk.Label(scrollable_frame, text="-"*100)
end_label.grid(row=8, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)  # sticky=tk.Wを追加

# %% [markdown]
# # 馬2

# %%
# 馬番の選択肢を更新
umaban_label_2 = tk.Label(scrollable_frame, text="馬2を選んでください")
# row記入
umaban_label_2.grid(row=9, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
umaban_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
umaban_2_vars = {}  # 馬番の1着を選択するチェックボックス用の変数を保持する辞書
frame_umaban_2 = tk.Frame(scrollable_frame)
#row記入
frame_umaban_2.grid(row=10, column=0, columnspan=2, sticky=tk.W, pady=10)
# サブ選択肢フレームのクリア
for widget in frame_umaban_2.winfo_children():
    widget.destroy()
# 馬番の選択肢を追加
for i, umaban in enumerate(umaban_list):
    var = tk.BooleanVar()
    umaban_2_vars[umaban] = var
    check = tk.Checkbutton(frame_umaban_2, text=umaban, variable=var)
    if umaban < 10:
        check.grid(row=0, column=umaban-1)  # 1〜9は1行目
    else:
        check.grid(row=1, column=umaban-10)  # 10〜18は2行目
"""
列の選択
"""
#馬1の列を選択
column_label_2 = tk.Label(scrollable_frame, text="馬2の列を選んでください")
#row記入
column_label_2.grid(row=11, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
column_list = ['着順', '馬番', '馬齢', 'オッズ', '斤量', '脚質', '総合値',
                    'SP値', 'AG値', 'SA値', '馬連率', '戦数', '賞金平均', 'KI値']
# 買い方の選択肢の追加
column_var_2 = tk.StringVar(value=" ") # 初期値を設定しない
frame_column_2 = tk.Frame(scrollable_frame)
# row記入
frame_column_2.grid(row=12, column=0, columnspan=2, sticky=tk.W, pady=10)
# サブ選択肢フレームのクリア
for widget in frame_column_2.winfo_children():
    widget.destroy()
for i, option in enumerate(column_list):
    radio = tk.Radiobutton(frame_column_2, text=option, variable=column_var_2, value=option)
    if i < 7:
        radio.grid(row=0, column=i, padx=5, pady=5)
    else:
        radio.grid(row=1, column=i-7, padx=5, pady=5)
"""
昇降順の選択
"""
up_down_label_2 = tk.Label(scrollable_frame, text="馬2の列の昇降順を選んでください")
#row記入
up_down_label_2.grid(row=13, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
up_down_list = ['昇順', '降順']
up_down_var_2 = tk.StringVar(value=" ") # 初期値を設定しない
frame_up_down_2 = tk.Frame(scrollable_frame)
#row記入
frame_up_down_2.grid(row=14, column=0, columnspan=2, sticky=tk.W, pady=10)
# サブ選択肢フレームのクリア
for widget in frame_up_down_2.winfo_children():
    widget.destroy()
for i, option in enumerate(up_down_list):
    radio = tk.Radiobutton(frame_up_down_2, text=option, variable=up_down_var_2, value=option)
    radio.grid(row=0, column=i, padx=5, pady=5)

end_label = tk.Label(scrollable_frame, text="-"*100)
end_label.grid(row=15, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)  # sticky=tk.Wを追加

# %% [markdown]
# # 馬3

# %%
# 馬番の選択肢を更新
umaban_label_3 = tk.Label(scrollable_frame, text="馬3を選んでください")
# row記入
umaban_label_3.grid(row=16, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
umaban_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
umaban_3_vars = {}  # 馬番の1着を選択するチェックボックス用の変数を保持する辞書
frame_umaban_3 = tk.Frame(scrollable_frame)
#row記入
frame_umaban_3.grid(row=17, column=0, columnspan=2, sticky=tk.W, pady=10)
# サブ選択肢フレームのクリア
for widget in frame_umaban_3.winfo_children():
    widget.destroy()
# 馬番の選択肢を追加
for i, umaban in enumerate(umaban_list):
    var = tk.BooleanVar()
    umaban_3_vars[umaban] = var
    check = tk.Checkbutton(frame_umaban_3, text=umaban, variable=var)
    if umaban < 10:
        check.grid(row=0, column=umaban-1)  # 1〜9は1行目
    else:
        check.grid(row=1, column=umaban-10)  # 10〜18は2行目
"""
列の選択
"""
#馬1の列を選択
column_label_3 = tk.Label(scrollable_frame, text="馬3の列を選んでください")
#row記入
column_label_3.grid(row=18, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
column_list = ['着順', '馬番', '馬齢', 'オッズ', '斤量', '脚質', '総合値',
                    'SP値', 'AG値', 'SA値', '馬連率', '戦数', '賞金平均', 'KI値']
# 買い方の選択肢の追加
column_var_3 = tk.StringVar(value=" ") # 初期値を設定しない
frame_column_3 = tk.Frame(scrollable_frame)
# row記入
frame_column_3.grid(row=19, column=0, columnspan=2, sticky=tk.W, pady=10)
# サブ選択肢フレームのクリア
for widget in frame_column_3.winfo_children():
    widget.destroy()
for i, option in enumerate(column_list):
    radio = tk.Radiobutton(frame_column_3, text=option, variable=column_var_3, value=option)
    if i < 7:
        radio.grid(row=0, column=i, padx=5, pady=5)
    else:
        radio.grid(row=1, column=i-7, padx=5, pady=5)
"""
昇降順の選択
"""
up_down_label_3 = tk.Label(scrollable_frame, text="馬3の列の昇降順を選んでください")
#row記入
up_down_label_3.grid(row=20, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
up_down_list = ['昇順', '降順']
up_down_var_3 = tk.StringVar(value=" ") # 初期値を設定しない
frame_up_down_3 = tk.Frame(scrollable_frame)
#row記入
frame_up_down_3.grid(row=21, column=0, columnspan=2, sticky=tk.W, pady=10)
# サブ選択肢フレームのクリア
for widget in frame_up_down_3.winfo_children():
    widget.destroy()
for i, option in enumerate(up_down_list):
    radio = tk.Radiobutton(frame_up_down_3, text=option, variable=up_down_var_3, value=option)
    radio.grid(row=0, column=i, padx=5, pady=5)

end_label = tk.Label(scrollable_frame, text="-"*100)
end_label.grid(row=22, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)  # sticky=tk.Wを追加

# %% [markdown]
# # 頭数制限

# %%
#下限頭数
lower_heads_label = tk.Label(scrollable_frame, text="下限頭数を選んでください")
#row記入
lower_heads_label.grid(row=23, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
# 馬番の選択肢を更新
umaban_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
frame_lower_heads = tk.Frame(scrollable_frame)
#row記入
frame_lower_heads.grid(row=24, column=0, columnspan=2, sticky=tk.W, pady=10)
# サブ選択肢フレームのクリア
for widget in frame_lower_heads.winfo_children():
    widget.destroy()
lower_head_var = tk.StringVar(value=" ") # 初期値を設定しない
# 買い方の選択肢の追加
for i, option in enumerate(umaban_list):
    radio = tk.Radiobutton(frame_lower_heads, text=option, variable=lower_head_var, value=option)
    if i < 9:
        radio.grid(row=0, column=i, padx=5, pady=5)
    else:
        radio.grid(row=1, column=i-9, padx=5, pady=5)

#上限頭数
maximum_heads_label = tk.Label(scrollable_frame, text="上限頭数を選んでください")
#row記入
maximum_heads_label.grid(row=25, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
# 馬番の選択肢を更新
umaban_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
frame_maximum_heads = tk.Frame(scrollable_frame)
#row記入
frame_maximum_heads.grid(row=26, column=0, columnspan=2, sticky=tk.W, pady=10)
# サブ選択肢フレームのクリア
for widget in frame_maximum_heads.winfo_children():
    widget.destroy()
maximum_head_var = tk.StringVar(value=" ") # 初期値を設定しない
# 買い方の選択肢の追加
for i, option in enumerate(umaban_list):
    radio = tk.Radiobutton(frame_maximum_heads, text=option, variable=maximum_head_var, value=option)
    if i < 9:
        radio.grid(row=0, column=i, padx=5, pady=5)
    else:
        radio.grid(row=1, column=i-9, padx=5, pady=5)
        
end_label = tk.Label(scrollable_frame, text="-"*100)
end_label.grid(row=27, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)  # sticky=tk.Wを追加

# %% [markdown]
# # レース選択

# %%
race_label = tk.Label(scrollable_frame, text="集計するレースを選んでください")
#row記入
race_label.grid(row=28, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)


# レースの選択肢を更新
race_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
#レースの選択
frame_race = tk.Frame(scrollable_frame)
#row記入
frame_race.grid(row=29, column=0, columnspan=2, sticky=tk.W, pady=10)

# サブ選択肢フレームのクリア
for widget in frame_race.winfo_children():
    widget.destroy()
race_vars = {} # 初期値を設定しない
# 買い方の選択肢の追加
for i, race in enumerate(race_list):
    var = tk.BooleanVar()
    race_vars[race] = var
    check = tk.Checkbutton(frame_race, text=race, variable=var)
    if race < 7:
        check.grid(row=0, column=race-1)  # 1〜6は1行目
    else:
        check.grid(row=1, column=race-7)  # 6〜12は2行目


# %%
"""
ここから別
"""
# 「クリック」ボタンをframe_umabanの下に配置
#このrowの数字はレースの選択のrowより3下にする
submit_button = tk.Button(scrollable_frame, text="確認", command=show_selection)
submit_button.grid(row=31, column=0, columnspan=2, pady=10)
# クリアボタンを作成して、ラジオボタンの選択をクリアする
clear_button = tk.Button(scrollable_frame, text="クリア", command=clear_radio_buttons_and_labels)
clear_button.grid(row=31, column=3, padx=5, pady=5)
# メインループの実行
root.mainloop()

# %%



