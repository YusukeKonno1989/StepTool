import tkinter
from tkinter import ttk
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import scrolledtext
from ttkwidgets import CheckboxTreeview
from chardet import detect
import os
import csv
import datetime
import json
import threading
import pandas as pd

# フォルダ指定の関数
def dirdialog_clicked():
    # ツリー初期化
    tree.delete(*tree.get_children())

    # 参照欄のパス取得
    iDir = os.path.abspath(os.path.dirname(__file__))
    iDirPath = filedialog.askdirectory(initialdir = iDir)
    entry1.set(iDirPath)

    if not iDirPath:
        return
    dname = os.path.basename(iDirPath)
    tree.heading("#0", text=dname, anchor=W)
    exclusion = read_json('exclusion.json')
    process_directory("", iDirPath, exclusion)

    tree.pack()

def read_json(filepath):
    # jsonファイルを開く
    json_file = open(filepath, 'r', encoding="utf-8")
    return json.load(json_file)

#ファイル・フォルダ名を取得してinsertする関数
def process_directory(parent, path, exclusion):
    
    exclusion_list = exclusion['exclusion_list']
    exclusion_folder_list = exclusion['exclusion_folder_list']
    delete_dir = os.path.dirname(entry1.get())
    
    for p in os.listdir(path):
        abspath = os.path.join(path, p)
        dir = abspath.replace(delete_dir, '')
        child = tree.insert(
            parent,
            "end",
            iid=dir,
            text=p,
            )
        if dir in exclusion_list:
            tree.change_state(child, 'checked')
        if dir in exclusion_folder_list:
            tree.change_state(child, 'checked')

        if os.path.isdir(abspath):
            #子要素がある場合は再帰呼び出し
            process_directory(child, abspath, exclusion)

#重複したリストを削除する関数
def get_unique_list(seq):
    seen = []
    return [x for x in seq if x not in seen and not seen.append(x)]

#親のチェックボックス状態を取得しリストに格納する関数
def search_parent_checked(id, exclusion_folder_list):
    parent = tree.parent(id)
    if not parent:
        return
    
    if len(tree.item(parent, 'tags')) > 0 and tree.item(parent, 'tags')[0] == 'checked':
        exclusion_folder_list.append(parent)

    if os.path.isdir(os.path.dirname(entry1.get()) + parent):
        search_parent_checked(parent, exclusion_folder_list)

# 除外チェック関数
def check_exclusion(exclusion_list, count_target):
    result = True
    for exclusion in exclusion_list:
        if exclusion in count_target:
            result = False
            break

    return result

#ステップ数カウント関数
def step_count():
    iDirPath = entry1.get()
    csv_path = os.path.join(os.getcwd(), str(datetime.datetime.now().strftime('%Y%m%d-%H%M%S')) + "_StepCounter.csv")

    if not iDirPath:
        return
    
    # configファイルを開く
    config = read_json("config.json")
    extension_list = config["extension_list"]
    exclusion_extension_list = config["exclusion_extension_list"]
    exclusion_folder = config["exclusion_folder"]

    exclusion_list = []
    exclusion_folder_list = []
    if len(tree.get_checked()) > 0:
        for id in tree.get_checked():
            if len(tree.item(tree.parent(id), 'tags')) > 0 and tree.item(tree.parent(id), 'tags')[0] == 'checked':
                search_parent_checked(tree.parent(id), exclusion_folder_list)
                exclusion_folder_list.append(tree.parent(id))
            else:
                exclusion_list.append(id)
    exclusion_folder_list = get_unique_list(exclusion_folder_list)

    count_list = []
    for current_dir, sub_dirs, files_list in os.walk(iDirPath): 
        for file_name in files_list:
            count_list.append(os.path.join(current_dir,file_name))

    p.config(maximum=len(count_list))
    st.configure(state='normal')
    write_data = []    
    for count_target in count_list:
        try:
            # 除外フォルダのチェック
            if not check_exclusion(exclusion_folder_list, count_target):
                pbval.set(pbval.get()+1)
                continue

            # 除外フォルダ(config側指定)のチェック
            if not check_exclusion(exclusion_folder, count_target):
                pbval.set(pbval.get()+1)
                continue
            
            # 除外ファイルのチェック
            if not check_exclusion(exclusion_list, count_target):
                pbval.set(pbval.get()+1)
                continue

            # 拡張子のチェック
            root, ext = os.path.splitext(count_target)
            if not ext in extension_list or ext in exclusion_extension_list:
                pbval.set(pbval.get()+1)
                continue

            if os.path.isdir(count_target):
                # ディレクトリの場合はスキップ
                pbval.set(pbval.get()+1)
                continue
            else:
                count = 0

                # ファイルのエンコード確認
                with open(count_target, 'rb') as f:  # バイナリファイルとしてファイルをオープン
                    b = f.read()
                enc = detect(b)

                with open(count_target, encoding=enc['encoding']) as file:
                    for line in file:
                        count += 1
                    st.insert('end', count_target + ' ' + str(count) + '行' + '\n')
                    st.see('end')
                    write_data.append([count_target, count, ext])
                    pbval.set(pbval.get()+1)
        except:
            continue

    # CSVに結果出力
    with open(csv_path, 'w') as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(["ファイルパス", "ステップ数", "拡張子"])
        writer.writerows(write_data)
        writer.writerow("")
        writer.writerow(["拡張子", "合計ステップ数"])

        # 結果出力
        df = pd.DataFrame(write_data,columns =['path','count','ext'])
        df_sum = df[['ext','count']].groupby('ext').sum()
        st.insert('end', '※※※※※※※※※※※※※※※※※※※※※※※※※' + '\n')
        for sum in df_sum.iterrows():
            writer.writerow([sum[0], str(sum[1][0])])
            st.insert('end', '拡張子：' + sum[0] + '  合計ステップ数：' + str(sum[1][0]) + '\n')
        st.insert('end', '※※※※※※※※※※※※※※※※※※※※※※※※※' + '\n')
        st.see('end')

    # configファイル書き込み
    set_data = {'extension_list':extension_list ,'exclusion_extension_list':exclusion_extension_list, 'exclusion_folder':exclusion_folder}
    with open('config.json', 'w', encoding="utf-8") as fs:
        json.dump(set_data, fs, indent=2, ensure_ascii=False)
    
    # exclusionファイル書き込み
    set_data = {'exclusion_list':exclusion_list, 'exclusion_folder_list':exclusion_folder_list}
    with open('exclusion.json', 'w', encoding="utf-8") as fs:
        json.dump(set_data, fs, indent=2, ensure_ascii=False)
    
    st.configure(state='disabled')
    messagebox.showinfo('Info', 'ステップカウント終了')  

#実行ボタン押下時のイベント関数
def exec_step_count():
  
    if not entry1.get():
        messagebox.showinfo('Info', '参照するフォルダが選択されていません') 
        return
    
    st.delete("1.0","end")
    pbval.set(0)

    thread = threading.Thread(target=step_count)
    thread.start()

if __name__ == "__main__":

    root = tkinter.Tk()
    root.geometry("780x780")
    root.resizable(width=False, height=False)
    root.title("ステップカウンタ")

    # Frame1の作成
    frame1 = ttk.Frame(root, padding=10)
    frame1.grid(row=0, column=1, sticky=E)

    # 「フォルダ参照」ラベルの作成
    IDirLabel = ttk.Label(frame1, text="フォルダ参照＞＞", padding=(5, 2))
    IDirLabel.pack(side=LEFT)

    # 「フォルダ参照」エントリーの作成
    entry1 = StringVar()
    IDirEntry = ttk.Entry(frame1, textvariable=entry1, width=90)
    IDirEntry.pack(side=LEFT)

    # 「フォルダ参照」ボタンの作成
    IDirButton = ttk.Button(frame1, text="参照", command=dirdialog_clicked)
    IDirButton.pack(side=LEFT)

    # Frame2の作成
    frame2 = ttk.Frame(root, padding=10)
    frame2.grid(row=1, column=1, sticky=E)

    # 「フォルダ参照」ラベルの作成
    IDirLabel = ttk.Label(frame2, text="除外ファイル選択＞＞", padding=(5, 2))
    
    #Treeview
    tree = CheckboxTreeview(frame2, height=20)
    tree.bind("<<TreeviewSelect>>")
    tree.column("#0", minwidth=0, width=500, stretch=NO)

    # スクロールバーの追加
    scrollbar = ttk.Scrollbar(frame2, orient=VERTICAL, command=tree.yview)
    tree.configure(yscroll=scrollbar.set)

    # ウィジェットの配置
    scrollbar.pack(side=RIGHT, fill=Y)
    tree.pack(side=RIGHT)
    IDirLabel.pack(side=RIGHT)

    # Frame3の作成
    frame3 = ttk.Frame(root, padding=10)
    frame3.grid(row=2,column=1,sticky=E)

    #ScrolledTextウィジェットを作成
    st= scrolledtext.ScrolledText(
        frame3, 
        width=105,
        state='disabled',
        height=13)
    st.pack()
    
    # Frame4の作成
    frame4 = ttk.Frame(root, padding=10)
    frame4.grid(row=3,column=1,sticky=E)

    # 実行ボタンの設置
    btnExec = ttk.Button(frame4, text="実行", command=exec_step_count, padding=(0,0))
    btnExec.pack(fill="x", padx=30)

    # Frame5の作成
    frame5 = ttk.Frame(root, padding=10)
    frame5.grid(row=4,column=1,sticky=E)

    #プログレスバー
    pbval = IntVar()
    p = ttk.Progressbar(
        frame5,
        orient=HORIZONTAL,
        maximum=10000,
        variable=pbval,
        length=250,
        mode="determinate",
        )
    p.pack()

    root.mainloop()