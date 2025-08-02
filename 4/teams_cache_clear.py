import os
import shutil
import subprocess
import tkinter as tk
from tkinter import ttk
import threading
from dataclasses import dataclass, field
from typing import List

# ===============================
# Teams のキャッシュフォルダのパス
# ===============================
# 古い Teams のキャッシュフォルダ
classic_path = os.path.expandvars(r"%APPDATA%\Microsoft\Teams")
# 新しい Teams のキャッシュフォルダ
new_path = os.path.expandvars(r"%LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe")

# ===============================
# データモデル：削除対象1件の情報
# ===============================
@dataclass
class DeletedItem:
    name: str   # 対象の名前（Classic Teams / New Teams - LocalCache など）
    path: str   # 削除対象のパス
    status: str # ステータス（待機中 / 削除中 / success / failed など）

# ===============================
# データモデル：削除結果のリスト
# ===============================
@dataclass
class DeletedItems:
    items: "List[DeletedItem]" = field(default_factory=list)
    def add(self, name: str, path: str, status: str):
        # 削除結果をリストに追加する
        self.items.append(DeletedItem(name=name, path=path, status=status))

# ===============================
# Teams の実行状態を確認
# ===============================
def get_running_teams():
    # Windows の tasklist コマンドで Teams のプロセスを調べる
    # 古い Teams は Teams.exe、新しい Teams は ms-teams.exe
    result = subprocess.run("tasklist", capture_output=True, text=True, shell=True)
    running = []
    if "Teams.exe" in result.stdout:
        running.append("classic")
    if "ms-teams.exe" in result.stdout:
        running.append("new")
    return running

# ===============================
# Teams のプロセスを終了
# ===============================
def kill_teams():
    # Teams のプロセスを強制終了する（古い Teams / 新しい Teams 両方）
    subprocess.call("taskkill /F /IM Teams.exe", shell=True)
    subprocess.call("taskkill /F /IM ms-teams.exe", shell=True)

# ===============================
# Teams を再起動
# ===============================
def restart_teams(running_list):
    # 起動していた Teams を再度起動する。
    # running_list に classic/new の情報が入っている場合のみ起動。
    if "classic" in running_list:
        classic_exe = os.path.expandvars(r"%LocalAppData%\Microsoft\Teams\current\Teams.exe")
        if os.path.exists(classic_exe):
            subprocess.Popen([classic_exe], shell=True)
    if "new" in running_list:
        new_teams_exe = os.path.expandvars(r"%LocalAppData%\Microsoft\WindowsApps\ms-teams.exe")
        if os.path.exists(new_teams_exe):
            subprocess.Popen([new_teams_exe], shell=True)

# ===============================
# 実際の削除処理
# ===============================
def delete_path(path: str):
    # 指定パスを削除。フォルダなら rmtree、ファイルなら remove
    if os.path.isdir(path):
        shutil.rmtree(path, ignore_errors=True)
    else:
        os.remove(path)

def safe_delete(name: str, path: str, result: DeletedItems, timeout: int = 10):
    # 削除処理をタイムアウト付きで実行。
    # 10秒以内に終わらなければ timeout 扱いにしてスキップ。
    update_tree_status(path, "削除中")
    import concurrent.futures
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(delete_path, path)
        try:
            future.result(timeout=timeout)
            result.add(name, path, status="success")
            update_tree_status(path, "success")
        except concurrent.futures.TimeoutError:
            result.add(name, path, status="timeout")
            update_tree_status(path, "timeout")
        except Exception as e:
            result.add(name, path, status=f"failed: {e}")
            update_tree_status(path, f"failed: {e}")

# ===============================
# Treeview のステータス更新
# ===============================
def update_tree_status(path, status):
    # Treeview の表示を更新する。
    # 同じ path の行を探してステータスを書き換える。
    for row in tree.get_children():
        values = tree.item(row, "values")
        if values[2] == path:
            tree.item(row, values=(values[0], status, values[2]))
            break

# ===============================
# リトライ（レコード単位）
# ===============================
def retry_item(path, name):
    # ダブルクリックされたレコードをリトライする。
    # Treeview のステータスを「リトライ中」にして再削除。
    update_tree_status(path, "リトライ中")
    threading.Thread(target=lambda: safe_delete(name, path, DeletedItems())).start()

# ===============================
# 親フォルダをエクスプローラで開く
# ===============================
def open_folder(path):
    # 削除済みのフォルダは存在しないので、1つ上の親フォルダを開く。
    parent = os.path.dirname(path)
    if os.path.exists(parent):
        subprocess.Popen(f'explorer "{parent}"')

# ===============================
# キャッシュ削除のメイン処理
# ===============================
def clear_cache():
    # Teams のキャッシュを削除するメイン関数。
    # - 起動中の Teams を終了
    # - 古い Teams と新しい Teams のキャッシュフォルダを削除
    # - 起動していた Teams を再起動
    result = DeletedItems() # 削除結果を格納するデータモデル

    # 起動していた Teams を確認
    running_before = get_running_teams() # 起動中の Teams 変数に入れとく
    if running_before:
        kill_teams() # 起動中の Teams を終了

    # Treeview をクリア（毎回新しいリストを表示）
    for row in tree.get_children(): # GUIの行をクリア
        tree.delete(row)

    # 古い Teams のキャッシュ削除
    if os.path.exists(classic_path):
        for item in os.listdir(classic_path):
            item_path = os.path.join(classic_path, item)
            tree.insert("", tk.END, values=("Classic Teams", "待機中", item_path)) #GUIに削除状況を追加
            safe_delete("Classic Teams", item_path, result) #削除した後に、GUIに削除状況を追加
    else:
        tree.insert("", tk.END, values=("Classic Teams", "not found", classic_path)) #フォルダがない場合

    # 新しい Teams のキャッシュ削除
    for sub in ["LocalCache", "LocalState", "TempState"]:
        path = os.path.join(new_path, sub)
        if os.path.exists(path):
            tree.insert("", tk.END, values=(f"New Teams - {sub}", "待機中", path)) #GUIに削除状況を追加
            safe_delete(f"New Teams - {sub}", path, result) #削除した後に、GUIに削除状況を追加
        else:
            tree.insert("", tk.END, values=(f"New Teams - {sub}", "not found", path)) #フォルダがない場合

    # Teams を再起動（起動していた場合のみ）
    if running_before: 
        restart_teams(running_before)

# ===============================
# GUI 本体
# ===============================
root = tk.Tk()
root.title("Teams キャッシュクリア")
root.geometry("600x250")

# アプリのアイコン設定
ico_path = os.path.join(os.path.dirname(__file__), "teams_cache.ico")
if os.path.exists(ico_path):
    root.iconbitmap(ico_path)

# キャッシュ削除ボタン
btn = ttk.Button(root, text="Clear Teams Cache", command=lambda: threading.Thread(target=clear_cache).start())
btn.pack(pady=5)

# ==== ボタンにツールチップ ====
btn_tooltip = tk.Label(root, text="", bg="yellow", relief="solid", bd=1)
btn_tooltip.place_forget()

def on_btn_enter(event):
    btn_tooltip.config(text="起動中のTeamsが終了し、キャッシュが削除されます")
    x = event.x_root - root.winfo_rootx() + 20
    y = event.y_root - root.winfo_rooty() + 20
    btn_tooltip.place(x=x, y=y)

def on_btn_leave(event):
    btn_tooltip.place_forget()

btn.bind("<Enter>", on_btn_enter)
btn.bind("<Leave>", on_btn_leave)

# Treeview（削除結果を表形式で表示）
columns = ("名前", "ステータス", "パス")
tree = ttk.Treeview(root, columns=columns, show="headings", height=10)
tree.heading("名前", text="名前")
tree.heading("ステータス", text="ステータス")
tree.heading("パス", text="パス")

# 列幅調整
tree.column("名前", width=150)
tree.column("ステータス", width=80)
tree.column("パス", width=350)

tree.pack(padx=10, pady=10, fill="both", expand=True)

# ダブルクリックイベント（列で動作を分ける）
def on_double_click(event):
    selected = tree.selection()
    if not selected:
        return
    column = tree.identify_column(event.x)
    item = tree.item(selected[0])
    name, _, path = item["values"]

    if column == "#1":  # 名前列 → リトライ
        retry_item(path, name)
    elif column == "#3":  # パス列 → 親フォルダを開く
        open_folder(path)

tree.bind("<Double-1>", on_double_click)

# ツールチップ（名前列とパス列でメッセージを出し分け）
tooltip = tk.Label(root, text="", bg="yellow", relief="solid", bd=1)
tooltip.place_forget()

def on_motion(event):
    region = tree.identify_region(event.x, event.y)
    column = tree.identify_column(event.x)
    if region == "cell" and column == "#1":
        tooltip.config(text="ダブルクリックでリトライ")
        tooltip.place(x=event.x_root - root.winfo_rootx() + 20,
                      y=event.y_root - root.winfo_rooty() + 20)
    elif region == "cell" and column == "#3":
        tooltip.config(text="ダブルクリックで親フォルダを開く")
        tooltip.place(x=event.x_root - root.winfo_rootx() + 20,
                      y=event.y_root - root.winfo_rooty() + 20)
    else:
        tooltip.place_forget()

tree.bind("<Motion>", on_motion)

root.mainloop()
