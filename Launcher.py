import tkinter as tk
from tkinter import messagebox
import sys
import os

#---< own class　フォルダ名を頭につける。
  # このファイルの親フォルダ（kikakuportal）をパスに追加
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from isms.isms_report import HrmosAutomation  # isms_report.py からインポート
from pms.pms_report import HrmosAutomationp # pms_report.py からインポート
from Bild import BildProc  # pms_report.py からインポート
from mailfuriwake.mailfuriwake import PpapProcessor  #mailfuriwake.py からインポート
from doxjidoushutoku.doxjidoushutoku import DoxGetProc #doxjidoushutoku.py からインポート
#-----------------------------------------------------------------------
class Launcher:
#-----------------------------------------------------------------------
#  ツール画面
#-----------------------------------------------------------------------
    def __init__(self):
       wk_pgm_start = "Launcher start"
#-----------------------------------------------------------------------
#  isms申請
#-----------------------------------------------------------------------
    def  run_isms_automation(self):
        """ISMS報告処理を実行する関数"""
        try:
            automation = HrmosAutomation()
            automation.run()
            messagebox.showinfo("完了", "ISMS報告処理が完了しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"ISMS報告処理中にエラーが発生しました:\n{e}")
#----------------------------------------------------------------------
#  pms申請
#-----------------------------------------------------------------------
    def run_pms_automation(self):
        """PMS報告処理を実行する関数"""
        try:
            automationp = HrmosAutomationp()
            automationp.runp()
            messagebox.showinfo("完了", "PMS報告処理が完了しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"PMS報告処理中にエラーが発生しました:\n{e}")
#----------------------------------------------------------------------
#  Eーメールの日付範囲抽出
#-----------------------------------------------------------------------
    def run_mailfuriwake(self):
        print("作る予定outlook")

        try:
            malefuriwake = PpapProcessor()
            malefuriwake.run()
            messagebox.showinfo("完了", "メール抽出が処理が完了しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")


#----------------------------------------------------------------------
#  Doxより抽出
#-----------------------------------------------------------------------
    def run_doxjidoushutoku(self):
        print("作る予定dox")

        try:
            Own_DoxGetProc = DoxGetProc()
            Own_DoxGetProc.run()
            messagebox.showinfo("完了", "ｄｏｘ抽出が処理が完了しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")
#-----------------------------------------------------------------------
#  暇ならやる
#-----------------------------------------------------------------------
    def show_about(self):
        """このツールについての情報を表示する関数"""
        messagebox.showinfo("このツールについて", "ISMS自動登録ツール v1.0\n作成者：あなた")
#-----------------------------------------------------------------------
#  画面を閉じる　今はどこも閉じる行為はしない
#-----------------------------------------------------------------------
    def exit_app(self):
        """アプリケーションを終了する関数"""
        self.root.destroy()

# --- メイン処理の開始 ---
if __name__ == '__main__':
  launcher = Launcher()
  bild = BildProc()
  launcher.root = bild.GamenBild(launcher)  # launcherを渡して、class BildProc()でボタンを押されたらlauncherへ戻る
  launcher.root.mainloop()
    
