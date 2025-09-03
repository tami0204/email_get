import os
from typing import Self
import win32com.client
import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta

class GetDataFromToProc:
     def __init__(self):

   #----< 画面作成
        self.root = tk.Tk()
        self.root.title("抽出期間の入力")
        self.root.geometry("500x250")
        font_large = ("Meiryo", 14)
        tk.Label(self.root, text="").grid(row=1)
        
        #----<開始日定義
        tk.Label(self.root, text="").grid(row=1)
        self.label_kaishibi = tk.Label(self.root, text="開始日（YYYYMMDD）", font=font_large)
        self.label_kaishibi.grid(row=2, column=1)
        self.entry_kaishibi = tk.Entry(self.root, font=font_large, width=20)
        self.entry_kaishibi.grid(row=2, column=2)

        tk.Label(self.root, text="").grid(row=3)
        #----<終了日定義
        self.label_endbi = tk.Label(self.root, text="終了日（YYYYMMDD）", font=font_large)
        self.label_endbi.grid(row=4, column=1)
        self.entry_endbi = tk.Entry(self.root, font=font_large, width=20)
        self.entry_endbi.grid(row=4, column=2)
        
        tk.Label(self.root, text="").grid(row=5)
        tk.Label(self.root, text="").grid(row=6)
      #----< 実行ボタン
        self.button_execute = tk.Button(self.root, text="実行", font=font_large, command=self._submit)
        self.button_execute.grid(row=7, column=1, columnspan=2, pady=40)
        #tk.Button(self.root, text="実行", font=font_large, command=submit).pack(pady=20)

     def _submit(self):
        try:
            self.start = datetime.strptime(self.entry_kaishibi.get(), "%Y%m%d")
            self.end = datetime.strptime(self.entry_endbi.get(), "%Y%m%d") + timedelta(days=1) - timedelta(seconds=1)
            #---<時分秒をつける
            self.date_from = self.start
            self.date_to = self.end

            self.start_str = self.date_from.strftime('%m/%d/%Y %H:%M %p')
            self.end_str = self.date_to.strftime('%m/%d/%Y %H:%M %p')
            self.root.destroy()
        except ValueError:
            messagebox.showerror("入力エラー", "日付は YYYYMMDD 形式で入力してください。")

    #======================================================================================#
    #  対象日付を入力
    #======================================================================================#
     def get_dates_from_to(self):
        """
        Tkinterを使って抽出期間を入力するフォームを表示する。
        """
        self.root.mainloop()
        return [f"[ReceivedTime] >= '{self.start_str}' AND [ReceivedTime] <= '{self.end_str}'",
                self.start_str,self.end_str]