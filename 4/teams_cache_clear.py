import os
import shutil
import subprocess
import tkinter as tk
from tkinter import ttk
import threading
from dataclasses import dataclass, field
from typing import List

classic_path = os.path.expandvars(r"%APPDATA%\Microsoft\Teams")
new_path = os.path.expandvars(r"%LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe")

@dataclass
class DeletedItem:
    name: str
    path: str
    status: str

@dataclass
class DeletedItems:
    items: "List[DeletedItem]" = field(default_factory=list)
    def add(self, name: str, path: str, status: str):
        self.items.append(DeletedItem(name=name, path=path, status=status))

def get_running_teams():
    result = subprocess.run("tasklist", capture_output=True, text=True, shell=True)
    running = []
    if "Teams.exe" in result.stdout:
        running.append("classic")
    if "ms-teams.exe" in result.stdout:
        running.append("new")
    return running

def kill_teams():
    subprocess.call("taskkill /F /IM Teams.exe", shell=True)
    subprocess.call("taskkill /F /IM ms-teams.exe", shell=True)

def restart_teams(running_list):
    if "classic" in running_list:
        classic_exe = os.path.expandvars(r"%LocalAppData%\Microsoft\Teams\current\Teams.exe")
        if os.path.exists(classic_exe):
            subprocess.Popen([classic_exe], shell=True)
    if "new" in running_list:
        new_teams_exe = os.path.expandvars(r"%LocalAppData%\Microsoft\WindowsApps\ms-teams.exe")
        if os.path.exists(new_teams_exe):
            subprocess.Popen([new_teams_exe], shell=True)

def delete_path(path: str):
    if os.path.isdir(path):
        shutil.rmtree(path, ignore_errors=True)
    else:
        os.remove(path)

def safe_delete(name: str, path: str, result: DeletedItems, timeout: int = 10):
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

def update_tree_status(path, status):
    for row in tree.get_children():
        values = tree.item(row, "values")
        if values[2] == path:
            tree.item(row, values=(values[0], status, values[2]))
            break

def retry_item(path, name):
    update_tree_status(path, "リトライ中")
    threading.Thread(target=lambda: safe_delete(name, path, DeletedItems())).start()

def open_folder(path):
    parent = os.path.dirname(path)
    if os.path.exists(parent):
        subprocess.Popen(f'explorer "{parent}"')

def clear_cache():
    result = DeletedItems()
    running_before = get_running_teams()
    if running_before:
        kill_teams()
        result.add("Teams", "プロセス", status="killed")

    for row in tree.get_children():
        tree.delete(row)

    if os.path.exists(classic_path):
        for item in os.listdir(classic_path):
            item_path = os.path.join(classic_path, item)
            tree.insert("", tk.END, values=("Classic Teams", "待機中", item_path))
            safe_delete("Classic Teams", item_path, result)
    else:
        tree.insert("", tk.END, values=("Classic Teams", "not found", classic_path))
        result.add("Classic Teams", classic_path, status="not found")

    for sub in ["LocalCache", "LocalState", "TempState"]:
        path = os.path.join(new_path, sub)
        if os.path.exists(path):
            tree.insert("", tk.END, values=(f"New Teams - {sub}", "待機中", path))
            safe_delete(f"New Teams - {sub}", path, result)
        else:
            tree.insert("", tk.END, values=(f"New Teams - {sub}", "not found", path))
            result.add(f"New Teams - {sub}", path, status="not found")

    if running_before:
        restart_teams(running_before)
        tree.insert("", tk.END, values=("Teams", "started", "再起動"))
        result.add("Teams", "再起動", status="started")

root = tk.Tk()
root.title("Teams キャッシュクリア")
root.geometry("600x250")

ico_path = os.path.join(os.path.dirname(__file__), "teams_cache.ico")
if os.path.exists(ico_path):
    root.iconbitmap(ico_path)

btn = ttk.Button(root, text="Clear Teams Cache", command=lambda: threading.Thread(target=clear_cache).start())
btn.pack(pady=5)

# ==== ボタンにツールチップ ====
btn_tooltip = tk.Label(root, text="", bg="yellow", relief="solid", bd=1)
btn_tooltip.place_forget()

def on_btn_enter(event):
    btn_tooltip.config(text="起動中のTeamsが終了→キャッシュ削除")
    x = event.x_root - root.winfo_rootx() + 20
    y = event.y_root - root.winfo_rooty() + 20
    btn_tooltip.place(x=x, y=y)

def on_btn_leave(event):
    btn_tooltip.place_forget()

btn.bind("<Enter>", on_btn_enter)
btn.bind("<Leave>", on_btn_leave)

columns = ("名前", "ステータス", "パス")
tree = ttk.Treeview(root, columns=columns, show="headings", height=15)
tree.heading("名前", text="名前")
tree.heading("ステータス", text="ステータス")
tree.heading("パス", text="パス")

tree.column("名前", width=150)
tree.column("ステータス", width=80)
tree.column("パス", width=350)

tree.pack(padx=10, pady=10, fill="both", expand=True)

def on_double_click(event):
    selected = tree.selection()
    if not selected:
        return
    column = tree.identify_column(event.x)
    item = tree.item(selected[0])
    name, _, path = item["values"]

    if column == "#1":  # 名前列
        retry_item(path, name)
    elif column == "#3":  # パス列
        open_folder(path)

tree.bind("<Double-1>", on_double_click)

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
