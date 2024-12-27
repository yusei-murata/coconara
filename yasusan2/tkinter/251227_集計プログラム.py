# %%
import tkinter as tk
from tkinter import ttk
import pandas as pd
from tqdm import tqdm
import time
import glob
import threading
import datetime

# Suppress pandas warnings
pd.options.mode.chained_assignment = None

# %%
files = glob.glob('*.xlsx')

# %%
today = datetime.datetime.today()

# %%
sheet_name_list = [
            '結果_1', '結果_2', '結果_3', '結果_4', '結果_5', '結果_6',
            '結果_7', '結果_8', '結果_9', '結果_10', '結果_11', '結果_12',
            ]

# %%
def aggregate():
    df_data = pd.read_csv('data_concat.csv', encoding='cp932')
    df_resutls = pd.read_csv('result_concat.csv', encoding='cp932')
    
    # Tkinterウィンドウの作成
    root = tk.Tk()
    root.title("データ集計・計算")
    root.geometry("600x600")

    # プログレスバーの作成
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.pack(padx=10, pady=10)

    # 経過時間と残り時間ラベルの作成
    time_label = tk.Label(root, text="経過時間: 0.00秒 残り時間: 0.00秒")
    time_label.pack(pady=10)

    def run_aggregation():
        # レース数による集計
        race_list = [int(race) for race, var in race_vars.items() if var.get()]
        df_data_race = df_data[df_data['race'].isin(race_list)]

        # 頭数による集計
        df_data_heads = df_data_race[(int(lower_head_var.get())<= df_data_race['頭数']) & (df_data_race['頭数']<=int(maximum_head_var.get()))]

        # columnによる集計
        df_data_column_1 = df_data_heads.dropna(subset=[column_var_1.get()])
        df_data_column_2 = df_data_column_1.dropna(subset=[column_var_2.get()])
        df_data_column_3 = df_data_column_2.dropna(subset=[column_var_3.get()])

        # 馬1の馬番を取得
        uma1_list = [int(uma1)-1 for uma1, var in umaban_1_vars.items() if var.get()]
        # 馬2の馬番を取得
        uma2_list = [int(uma2)-1 for uma2, var in umaban_2_vars.items() if var.get()]
        # 馬3の馬番を取得
        uma3_list = [int(uma3)-1 for uma3, var in umaban_3_vars.items() if var.get()]

        total_files = len(df_data_column_3['id'].unique())
        progress_step = 100 / total_files  # プログレスバーのステップサイズ
        start_time = time.time()  # 処理開始時間

        kaime_list_all = []
        id_list_all = []
        kingaku_list_all = []
        # レースIDのみにする
        for file_index, id_data in enumerate(df_data_column_3['id'].unique()):
            df_data_id = df_data_column_3[df_data_column_3['id'] == id_data]
            # 馬1の集計
            if column_var_1.get() != '馬番':
                df_data_1 = df_data_id[['馬番',column_var_1.get()]]
            else:
                df_data_1 = df_data_id[['馬番']]
            if up_down_var_1.get() == '昇順':
                df_data_1_sort = df_data_1.sort_values(column_var_1.get()).reset_index(drop=True)
            else:
                df_data_1_sort = df_data_1.sort_values(column_var_1.get(),ascending=False).reset_index(drop=True)
            # 馬2の集計
            if column_var_2.get() != '馬番':
                df_data_2 = df_data_id[['馬番',column_var_2.get()]]
            else:
                df_data_2 = df_data_id[['馬番']]
            if up_down_var_2.get() == '昇順':
                df_data_2_sort = df_data_2.sort_values(column_var_2.get()).reset_index(drop=True)
            else:
                df_data_2_sort = df_data_2.sort_values(column_var_2.get(),ascending=False).reset_index(drop=True)
            # 馬3の集計
            if column_var_3.get() != '馬番':
                df_data_3 = df_data_id[['馬番',column_var_3.get()]]
            else:
                df_data_3 = df_data_id[['馬番']]
            if up_down_var_3.get() == '昇順':
                df_data_3_sort = df_data_3.sort_values(column_var_3.get()).reset_index(drop=True)
            else:
                df_data_3_sort = df_data_3.sort_values(column_var_3.get(),ascending=False).reset_index(drop=True)
            # 馬1の馬番を取得
            uma1_list_data = [str(df_data_1_sort['馬番'][uma1]) for uma1 in uma1_list]
            # 馬2の馬番を取得
            uma2_list_data = [str(df_data_2_sort['馬番'][uma2]) for uma2 in uma2_list]
            # 馬3の馬番を取得
            uma3_list_data = [str(df_data_3_sort['馬番'][uma3]) for uma3 in uma3_list]
            # 買い目を作成
            for uma1 in uma1_list_data:
                for uma2 in uma2_list_data:
                    for uma3 in uma3_list_data:
                        if uma1 != uma2 and uma2 != uma3 and uma3 != uma1:
                            kaime_list = [
                                    int(uma1),
                                    int(uma2),
                                    int(uma3)
                                    ]
                            kaime_list_sorted = sorted(kaime_list)
                            kaime_list_all.append(str(kaime_list_sorted[0])+'_'+str(kaime_list_sorted[1])+'_'+str(kaime_list_sorted[2]))
                            id_list_all.append(id_data)
                            kingaku_list_all.append(100)
            # プログレスバーを更新
            progress_var.set((file_index + 1) * progress_step)
            progress_bar.update_idletasks()

            # 経過時間と残り時間を更新
            elapsed_time = time.time() - start_time
            remaining_time = (elapsed_time / (file_index + 1)) * (total_files - (file_index + 1))
            time_label.config(text=f"経過時間: {elapsed_time:.2f}秒 残り時間: {remaining_time:.2f}秒")
            time_label.update_idletasks()
        df_kaime = pd.DataFrame({
                        'id':id_list_all,
                        '買い目':kaime_list_all,
                        '購入金額':kingaku_list_all
                        }).drop_duplicates(keep='first')
        df_merge = pd.merge(df_kaime,df_resutls, on=['id', '買い目'], how='left')
        
        kounyu_money = df_merge['購入金額'].sum()
        haraimodoshi = df_merge['odds'].sum()
        kaishu_rate = round(haraimodoshi / kounyu_money*100, 2)
        kounyu_suu = len(df_merge)
        tekityu = df_merge['odds'].count()
        tekityu_rate = round(tekityu / kounyu_suu*100, 2)
        # 集計結果を表示
        result_label_money_1 = tk.Label(root, text=f"購入金額: {kounyu_money}円 ")
        result_label_money_1.pack(pady=10)
        result_label_money_2 = tk.Label(root, text=f"払戻金: {haraimodoshi}円")
        result_label_money_2.pack(pady=10)
        result_label_money_3 = tk.Label(root, text=f"回収率: {kaishu_rate} %")
        result_label_money_3.pack(pady=10)
        hypen_label = tk.Label(root, text="----------------------------------------")
        result_label_1 = tk.Label(root, text=f"購入数: {kounyu_suu}")
        result_label_1.pack(pady=10)
        result_label_2 = tk.Label(root, text=f"的中数: {tekityu}")
        result_label_2.pack(pady=10)
        result_label_3 = tk.Label(root, text=f"的中率: {tekityu_rate}%")
        result_label_3.pack(pady=10)
        df_merge.to_csv(
                str(today).split(' ')[0]+'_aggregation.csv',
                index=False,
                encoding='cp932',)

    # 別スレッドで集計処理を実行
    threading.Thread(target=run_aggregation).start()

    root.mainloop()


# %%
def show_next_window():
    # tkオブジェクトの作成
    root = tk.Tk()
    root.title("条件選択")  # ウィンドウのタイトルを設定
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
    

    def show_selection():
        
        # 新しいウィンドウを作成
        new_window = tk.Toplevel(root)
        new_window.title("選択内容の確認")
        new_window.geometry("500x500")
        
        
        # 結果を表示するラベルを作成して配置
        kenshu_label = tk.Label(new_window, text="")
        kenshu_selection = f"券種 : {kenshu_var.get()}"
        kenshu_label.config(text=kenshu_selection)
        kenshu_label.grid(row=0, column=10, columnspan=2, sticky=tk.W, padx=5, pady=10)
        
        # 馬1の表示
        umaban_label_1 = tk.Label(new_window, text="")
        umaban_selection_1 = "馬1 : " + ", ".join([str(umaban) for umaban, var in umaban_1_vars.items() if var.get()])
        umaban_label_1.config(text=umaban_selection_1)
        umaban_label_1.grid(row=1, column=10, columnspan=2, sticky=tk.W, padx=5, pady=10)
        # 馬1の列表示
        column_label_1 = tk.Label(new_window, text="")
        column_selection_1 = f"馬1_列 : {column_var_1.get()}"
        column_label_1.config(text=column_selection_1)
        column_label_1.grid(row=2, column=10, columnspan=2, sticky=tk.W, padx=5, pady=10)
        #馬1の列昇降
        up_down_label_1 = tk.Label(new_window, text="")
        up_down_selection_1 = f"馬1_昇降順 : {up_down_var_1.get()}"
        up_down_label_1.config(text=up_down_selection_1)
        up_down_label_1.grid(row=3, column=10, columnspan=2, sticky=tk.W, padx=5, pady=10)
        
        # 馬2の表示
        umaban_label_2 = tk.Label(new_window, text="")
        umaban_selection_2 = "馬2 : " + ", ".join([str(umaban) for umaban, var in umaban_2_vars.items() if var.get()])
        umaban_label_2.config(text=umaban_selection_2)
        umaban_label_2.grid(row=1, column=15, columnspan=2, sticky=tk.W, padx=5, pady=10)
        # 馬2の列表示
        column_label_2 = tk.Label(new_window, text="")
        column_selection_2 = f"馬2_列 : {column_var_2.get()}"
        column_label_2.config(text=column_selection_2)
        column_label_2.grid(row=2, column=15, columnspan=2, sticky=tk.W, padx=5, pady=10)
        #馬2の列昇降
        up_down_label_2 = tk.Label(new_window, text="")
        up_down_selection_2 = f"馬2_昇降順 : {up_down_var_2.get()}"
        up_down_label_2.config(text=up_down_selection_2)
        up_down_label_2.grid(row=3, column=15, columnspan=2, sticky=tk.W, padx=5, pady=10)
        
        # 馬3の表示
        umaban_label_3 = tk.Label(new_window, text="")
        umaban_selection_3 = "馬3 : " + ", ".join([str(umaban) for umaban, var in umaban_3_vars.items() if var.get()])
        umaban_label_3.config(text=umaban_selection_3)
        umaban_label_3.grid(row=1, column=20, columnspan=2, sticky=tk.W, padx=5, pady=10)
        # 馬3の列表示
        column_label_3 = tk.Label(new_window, text="")
        column_selection_3 = f"馬3_列 : {column_var_3.get()}"
        column_label_3.config(text=column_selection_3)
        column_label_3.grid(row=2, column=20, columnspan=2, sticky=tk.W, padx=5, pady=10)
        # 馬3の列昇降
        up_down_label_3 = tk.Label(new_window, text="")
        up_down_selection_3 = f"馬3_昇降順 : {up_down_var_3.get()}"
        up_down_label_3.config(text=up_down_selection_3)
        up_down_label_3.grid(row=3, column=20, columnspan=2, sticky=tk.W, padx=5, pady=10)
        # 下限頭数
        lower_head_label = tk.Label(new_window, text="")
        lower_head_selection = f"下限頭数 : {lower_head_var.get()}"
        lower_head_label.config(text=lower_head_selection)
        lower_head_label.grid(row=4, column=10, columnspan=2, sticky=tk.W, padx=5, pady=10)
        
        # 上限頭数
        maximum_head_label = tk.Label(new_window, text="")
        maximum_head_selection = f"上限頭数 : {maximum_head_var.get()}"
        maximum_head_label.config(text=maximum_head_selection)
        maximum_head_label.grid(row=4, column=15, columnspan=2, sticky=tk.W, padx=5, pady=10)
        
        # 選択レース
        race_label = tk.Label(new_window, text="")
        race_selection = "レース : " + ", ".join([str(race) for race, var in race_vars.items() if var.get()])
        race_label.config(text=race_selection)
        race_label.grid(row=5, column=20, columnspan=2, sticky=tk.W, padx=5, pady=10)
        
        
        #集計に進むボタン
        submit_button = tk.Button(new_window, text="集計に進む", command=aggregate)
        submit_button.grid(row=6, column=5, columnspan=2, sticky=tk.W, pady=10)
        
        def new_window_close():
            new_window.destroy()
            clear_radio_buttons_and_labels()
        #選択し直すボタン
        chancele_button = tk.Button(new_window, text="選択に戻る", command=new_window_close)
        chancele_button.grid(row=6, column=15, columnspan=2, pady=10)
    
    # グローバル変数の定義
    global kenshu_var, column_var_1, column_var_2, column_var_3
    global up_down_var_1, up_down_var_2, up_down_var_3
    global umaban_1_vars, umaban_2_vars, umaban_3_vars
    global lower_head_var, maximum_head_var, race_vars
    
    
    kenshu_var = tk.StringVar(value=" ")
    umaban_1_vars = {}
    column_var_1 = tk.StringVar(value=" ")  # 初期値を設定しない
    up_down_var_1 = tk.StringVar(value=" ")  # 初期値を設定しない
    umaban_2_vars = {}
    column_var_2 = tk.StringVar(value=" ")  # 初期値を設定しない
    up_down_var_2 = tk.StringVar(value=" ")  # 初期値を設定しない
    umaban_3_vars = {}
    column_var_3 = tk.StringVar(value=" ")  # 初期値を設定しない
    up_down_var_3 = tk.StringVar(value=" ")  # 初期値を設定しない
    lower_head_var = tk.StringVar(value=" ") # 下限頭数
    maximum_head_var = tk.StringVar(value=" ") # 上限頭数
    race_vars = {} # 初期値を設定しない
    
    #kenshu_list = ["馬単", "三連複"]
    kenshu_list = ["三連複"]

    # ウィジェットの配置や、イベント処理などを記述
    kenshu_label = tk.Label(scrollable_frame, text="券種を選んでください")
    kenshu_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

    # ラジオボタンを横に並べるためのフレーム
    main_frame = tk.Frame(scrollable_frame)
    main_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W)
    for index, kenshu in enumerate(kenshu_list):
        radio = tk.Radiobutton(main_frame, text=kenshu, variable=kenshu_var, value=kenshu)
        radio.grid(row=0, column=index, padx=5, pady=5)
    # 馬1
    # 馬番の選択肢を更新
    umaban_label_1 = tk.Label(scrollable_frame, text="馬1を選んでください")
    umaban_label_1.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)  # sticky=tk.Wを追加
    umaban_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
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
    
    # 馬2
    # 馬番の選択肢を更新
    umaban_label_2 = tk.Label(scrollable_frame, text="馬2を選んでください")
    # row記入
    umaban_label_2.grid(row=9, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
    umaban_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
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
    
    # 馬3
    # 馬番の選択肢を更新
    umaban_label_3 = tk.Label(scrollable_frame, text="馬3を選んでください")
    # row記入
    umaban_label_3.grid(row=16, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
    umaban_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
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
    
    # 頭数制限
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
    # 買い方の選択肢の追加
    for i, option in enumerate(umaban_list):
        radio = tk.Radiobutton(frame_maximum_heads, text=option, variable=maximum_head_var, value=option)
        if i < 9:
            radio.grid(row=0, column=i, padx=5, pady=5)
        else:
            radio.grid(row=1, column=i-9, padx=5, pady=5)
            
    end_label = tk.Label(scrollable_frame, text="-"*100)
    end_label.grid(row=27, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)  # sticky=tk.Wを追加
    
    # レース選択
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
    # 買い方の選択肢の追加
    for i, race in enumerate(race_list):
        var = tk.BooleanVar()
        race_vars[race] = var
        check = tk.Checkbutton(frame_race, text=race, variable=var)
        if race < 7:
            check.grid(row=0, column=race-1)  # 1〜6は1行目
        else:
            check.grid(row=1, column=race-7)  # 6〜12は2行目

    """
    ここから別
    """
    # 「クリック」ボタンをframe_umabanの下に配置
    #このrowの数字はレースの選択のrowより3下にする
    submit_button = tk.Button(scrollable_frame, text="確認", command=show_selection)
    submit_button.grid(row=31, column=0, columnspan=2, pady=10)
    
    def clear_radio_buttons_and_labels():
        # すべての変数をリセット
        kenshu_var.set(" ")
        column_var_1.set(" ")
        up_down_var_1.set(" ")
        column_var_2.set(" ")
        up_down_var_2.set(" ")
        column_var_3.set(" ")
        up_down_var_3.set(" ")
        lower_head_var.set(" ")
        maximum_head_var.set(" ")
        for var in umaban_1_vars.values():
            var.set(False)
        for var in umaban_2_vars.values():
            var.set(False)
        for var in umaban_3_vars.values():
            var.set(False)
        for var in race_vars.values():
            var.set(False)
    
    # クリアボタンを作成して、ラジオボタンの選択をクリアする
    clear_button = tk.Button(scrollable_frame, text="クリア", command=clear_radio_buttons_and_labels)
    clear_button.grid(row=31, column=3, padx=5, pady=5)
    # メインループの実行
    root.mainloop()

# %%
def data_concat(files, sheet_name_list, progress_var, progress_bar, time_label):
    df_list = []
    result_list = []
    total_files = len(files)
    progress_step = 100 / total_files  # プログレスバーのステップサイズ
    start_time = time.time()  # 処理開始時間

    for file_index, file in enumerate(files):
        try:
            for sheet_name in sheet_name_list:
                id = file.split('.')[0] + '_' + sheet_name
                df = pd.read_excel(file, header=1, sheet_name=sheet_name)
                if '総合 値' in df.columns:
                    df = df.rename(columns={'総合 値': '総合値'})
                if 'AG 値' in df.columns:
                    df = df.rename(columns={'AG 値': 'AG値'})
                if 'SP 値' in df.columns:
                    df = df.rename(columns={'SP 値': 'SP値'})
                if 'SA 値' in df.columns:
                    df = df.rename(columns={'SA 値': 'SA値'})
                if 'KI 値' in df.columns:
                    df = df.rename(columns={'KI 値': 'KI値'})
                if 'Unnamed: 10' in df.columns:
                    df = df.drop(['Unnamed: 10'], axis=1)
                df_drop = df.dropna(subset=['着順'])
                try:
                    df_drop = df_drop.drop(['馬名', '騎手名', '前回騎乗', '調教師'], axis=1)
                except KeyError:
                    df_drop = df_drop.drop(['馬名', '騎手名', '調教師'], axis=1)
                result = df.query(str(df.columns[5]) + ' == "3連複"')[[str(df.columns[5]), str(df.columns[6]), str(df.columns[7])]]
                result = result.rename(columns={str(df.columns[5]): '券種', str(df.columns[6]): '買い目', str(df.columns[7]): 'odds'})
                df_drop['id'] = id
                df_drop['race'] = id.split('_')[-1]
                #頭数追加
                df_drop['頭数'] = len(df_drop)
                result['id'] = id
                result['買い目'] = result['買い目'].str.replace('-', '_')
                df_column = df_drop
                result_list.append(result)
                df_list.append(df_column)
        except:
            pass
        # プログレスバーを更新
        progress_var.set((file_index + 1) * progress_step)
        progress_bar.update_idletasks()

        # 経過時間と残り時間を更新
        elapsed_time = time.time() - start_time
        remaining_time = (elapsed_time / (file_index + 1)) * (total_files - (file_index + 1))
        time_label.config(text=f"経過時間: {elapsed_time:.2f}秒 残り時間: {remaining_time:.2f}秒")
        time_label.update_idletasks()

    # データフレームを結合してCSVファイルに保存
    df_data = pd.concat(df_list).drop_duplicates(keep='first')
    result_data = pd.concat(result_list).drop_duplicates(keep='first')
    df_data.to_csv('data_concat.csv', index=False, encoding='cp932')
    result_data.to_csv('result_concat.csv', index=False, encoding='cp932')
    # 最初のウィンドウを閉じて次のウィンドウを表示
    root.destroy()
    show_next_window()

def start_processing():
    files = glob.glob('*.xlsx')  # 処理するファイルのリスト
    sheet_name_list = [
                    '結果_1', '結果_2', '結果_3', '結果_4', '結果_5', '結果_6',
                    '結果_7', '結果_8', '結果_9', '結果_10', '結果_11', '結果_12',
                    ]  # 処理するシートのリスト
    data_concat(files, sheet_name_list, progress_var, progress_bar, time_label)

def skip_processiong():
    # 最初のウィンドウを閉じて次のウィンドウを表示
    root.destroy()
    show_next_window()

# Tkinterウィンドウの作成
root = tk.Tk()
root.title("データ処理")
root.geometry("300x200")

# プログレスバーの作成
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.pack(padx=10, pady=10)

# 経過時間と残り時間ラベルの作成
time_label = tk.Label(root, text="経過時間: 0.00秒 残り時間: 0.00秒")
time_label.pack(pady=10)

# 開始ボタンの作成
start_button = tk.Button(root, text="開始", command=start_processing)
start_button.pack(pady=10)

# スキップボタンの作成
skip_button = tk.Button(root, text="スキップ", command=skip_processiong)
skip_button.pack(pady=10)

# メインループの実行
root.mainloop()

# %%


# %%


# %%



