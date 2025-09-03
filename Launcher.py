# main_gui_app.py の内容

import tkinter as tk
from tkinter import messagebox
from isms_report import HrmosAutomation  # isms_report.py からインポート
from pms_report import HrmosAutomationp  # pms_report.py からインポート

def run_isms_automation():
    """ISMS報告処理を実行する関数"""
    try:
        automation = HrmosAutomation()
        automation.run()
        messagebox.showinfo("完了", "ISMS報告処理が完了しました。")
    except Exception as e:
        messagebox.showerror("エラー", f"ISMS報告処理中にエラーが発生しました:\n{e}")

def run_pms_automation():
    """PMS報告処理を実行する関数"""
    try:
        automationp = HrmosAutomationp()
        automationp.runp()
        messagebox.showinfo("完了", "PMS報告処理が完了しました。")
    except Exception as e:
        messagebox.showerror("エラー", f"PMS報告処理中にエラーが発生しました:\n{e}")

def show_about():
    """このツールについての情報を表示する関数"""
    messagebox.showinfo("このツールについて", "ISMS自動登録ツール v1.0\n作成者：あなた")

def exit_app():
    """アプリケーションを終了する関数"""
    root.destroy()

# フォント設定 (Windows環境向けにMeiryoを設定)
yu_gothic_font = ("Meiryo", 20)

# GUI画面の定義
root = tk.Tk()
root.title("ISMS自動登録ツール")
root.geometry("600x400")
root.attributes("-topmost", True) # 常に最前面に表示
root.configure(bg="#FFE4E1") # 背景色をピンクに設定 (ミストローズ)

# ラベル（背景色もピンクに設定）
tk.Label(root, text="実行したい処理を選んでください", font=yu_gothic_font, bg="#FFE4E1").pack(pady=10)

# ボタン（背景色は落ち着いた色に設定）
tk.Button(root, text="ISMS報告処理を開始", command=run_isms_automation, font=yu_gothic_font, bg="#D2B48C").pack(pady=5)
tk.Button(root, text="PMS報告処理を開始", command=run_pms_automation, font=yu_gothic_font, bg="#D2B48C").pack(pady=5)
tk.Button(root, text="このツールについて", command=show_about, font=yu_gothic_font, bg="#D2B48C").pack(pady=5)
tk.Button(root, text="終了", command=exit_app, font=yu_gothic_font, bg="#D2B48C").pack(pady=5)

# GUIイベントループの開始
root.mainloop()
