import tkinter as tk
from tkinter import messagebox
#----------------------------------------------------------------------
class BildProc:
#-----------------------------------------------------------------------
#  ツール画面
#-----------------------------------------------------------------------

    def __init__(self):
        wk_pgm_start = "BildProc"# フォルダパスの設定
#-----------------------------------------------------------------------
#  主処理
#-----------------------------------------------------------------------
    def GamenBild(self,Launcher):
        # フォント設定 (Windows環境向けにMeiryoを設定)
        yu_gothic_font = ("Meiryo", 20)
    
        # GUI画面の定義
        root = tk.Tk()
        root.title("開発企画課ポータル")
        root.geometry("1200x600")   # yoko tate
        root.attributes("-topmost", True) # 常に最前面に表示
        root.configure(bg="#FFE4E1") # 背景色をピンクに設定 (ミストローズ)
        
        # ラベル（背景色もピンクに設定）
        tk.Label(root, text="実行したい処理を選んでください", font=yu_gothic_font, bg="#FFE4E1").pack(pady=10)
        # ボタン配置用フレーム
        button_frame = tk.Frame(root, bg="#FFE4E1")
        button_frame.pack()

    # ボタン定義（テキストと対応する関数）
        buttons = [
         ("ISMS報告書作成", Launcher.run_isms_automation),
         ("PMS報告書作成", Launcher.run_pms_automation),
         ("mail振り分け", Launcher.run_mailfuriwake),
         ("DOX自動取得", Launcher.run_doxjidoushutoku),
         ("このツールについて", Launcher.show_about),
         ("終了", Launcher.exit_app)
         ]

    # 2列でボタンを配置
        for i, (text, command) in enumerate(buttons):
            row = i // 2    #割り算
            col = i % 2     #割り算の余り
            # tk.Button(button_frame, text=text, command=command, font=yu_gothic_font, bg="#D2B48C", width=25).grid(row=row, column=col, padx=20, pady=10)

            btn = tk.Button(
                button_frame,
                text=text,
                command=lambda cmd=command: self.run_and_minimize(Launcher, cmd),
                font=yu_gothic_font,
                bg="#D2B48C",
                width=25
            )
            btn.grid(row=row, column=col, padx=20, pady=10)

        return root
    def run_and_minimize(self, Launcher, func):
        Launcher.root.iconify()
        func()

        #################################################################################################

        
        # # ボタン（背景色は落ち着いた色に設定）
        # tk.Button(root, text="ISMS報告書作成", command=Launcher.run_isms_automation, font=yu_gothic_font, bg="#D2B48C").pack(pady=5)
        # tk.Button(root, text="PMS報告書作成", command=Launcher.run_pms_automation, font=yu_gothic_font, bg="#D2B48C").pack(pady=5)
        # tk.Button(root, text="このツールについて", command=Launcher.show_about, font=yu_gothic_font, bg="#D2B48C").pack(pady=5)
        # tk.Button(root, text="終了", command=Launcher.exit_app, font=yu_gothic_font, bg="#D2B48C").pack(pady=5)
        
        # return root
