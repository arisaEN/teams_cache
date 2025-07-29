import os
import shutil
import subprocess
import tkinter as tk
from tkinter import messagebox, Toplevel, Label
from tkinter import ttk
import threading

classic_path = os.path.expandvars(r"%APPDATA%\Microsoft\Teams")
new_path = os.path.expandvars(r"%LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe")

# Teamsが動作中か確認（tasklist使用）
def is_teams_running():
    result = subprocess.run("tasklist", capture_output=True, text=True, shell=True) #windowsのタスクリスト確認
    return "Teams.exe" in result.stdout or "ms-teams.exe" in result.stdout #新旧どちらかのtesamsが動いていればtrueを返す

###Teams終了
def kill_teams():
    subprocess.call("taskkill /F /IM Teams.exe", shell=True) #Teams.exeというプロセスをキル
    subprocess.call("taskkill /F /IM ms-teams.exe", shell=True) #ms-Teams.exeというプロセスをキル (新版)

###キャッシュ削除
def clear_cache():
    if is_teams_running():
        if not messagebox.askokcancel("確認", "Teamsが実行中です。終了してキャッシュを削除しますか？"):
            return
        kill_teams()

    loading = Toplevel(root) #ローディング画面を表示
    loading.title("処理中") #ローディング画面に表示するタイトル
    loading.geometry("250x100") #ローディング画面のサイズ
    Label(loading, text="キャッシュを削除中...\nしばらくお待ちください", font=("Arial", 12)).pack(expand=True) #ローディング画面に表示するテキスト
    loading.update() #ローディング画面を更新

    def worker(): #キャッシュ削除ロジック本体
        deleted = [] #削除したファイルのリスト
        #昔Ver全部消しちゃう
        # if os.path.exists(classic_path): #Teamsフォルダが存在するか確認
        #     shutil.rmtree(classic_path, ignore_errors=True) #Teamsフォルダごと削除
        #     deleted.append("Classic Teams") #削除したファイルのリストに追加 昔のteams削除したよってことを通知するため。
        
        #昔Ver 特定のファイルだけ消す。。
        # for item in os.listdir(classic_path):
        #     if item in ["Cache", "tmp", "Service Worker"]:  # 消したいフォルダ名だけ指定
        #         item_path = os.path.join(classic_path, item)
        #         shutil.rmtree(item_path, ignore_errors=True)
                
        #昔Ver フォルダの中身消す。
        if os.path.exists(classic_path):
            for item in os.listdir(classic_path):
                item_path = os.path.join(classic_path, item)
                if os.path.isdir(item_path):
                    shutil.rmtree(item_path, ignore_errors=True)
                else:
                    os.remove(item_path)
            deleted.append("Classic Teams")
        else:
            print("Classic Teamsフォルダは存在しません")
                #新Ver キャッシュフォルダだけ消す   
        for sub in ["LocalCache", "LocalState", "TempState"]: #新しいteamsのキャッシュフォルダ subの中にフォルダ名いれる。
            path = os.path.join(new_path, sub) #新しいteamsのキャッシュフォルダのパス subはフォルダ名のリスト ファイルパス+フォルダ名
            if os.path.exists(path): #新しいteamsのキャッシュフォルダが存在するか確認
                shutil.rmtree(path, ignore_errors=True) #新しいteamsのキャッシュフォルダごと削除
                deleted.append(f"New Teams - {sub}") #削除したファイルのリストに追加 新しいteams削除したよってことを通知するため。

        loading.destroy() #ローディング画面を削除
        if deleted: #空かどうか判定
            messagebox.showinfo("完了", "キャッシュを削除しました:\n" + "\n".join(deleted))
        else:
            messagebox.showinfo("情報", "削除対象のキャッシュは見つかりませんでした。")
    threading.Thread(target=worker).start() #キャッシュ削除 #スレッド処理なのでステップ実行できない
    #worker() #デバック用

root = tk.Tk()
root.title("Teams キャッシュクリア")
root.geometry("300x150")

btn = ttk.Button(root, text="Clear Teams Cache", command=clear_cache)
btn.pack(expand=True, ipadx=20, ipady=10)

btn.pack(expand=True)

root.mainloop()
