# -*- coding: cp932 -*-
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from datetime import datetime, timedelta
import time
import os
import zipfile
import shutil
import glob
import chardet
from pathlib import Path

#-----------------------------------------------------------------------
class DoxGetProc:
#-----------------------------------------------------------------------
#  コンストラクタ
#-----------------------------------------------------------------------
    def __init__(self):
    #固定値情報
        self.username = "t-tamura"
        self.password = "merumo0526!!"
        self.target_extensions = ('.pdf', '.xlsx', '.docx', '.zip')

    #格納先フォルダ作成
        base_path = os.path.join(Path.home(), "Desktop", "ai_ocr", "メール","dox")
        current_time_str = datetime.now().strftime("%Y%m%d_%H%M")
        dated_path = os.path.join(base_path, current_time_str)
        os.makedirs(dated_path, exist_ok=True)
        save_folder=dated_path   
    # フォルダの指定
        self.save_folder = save_folder
        self.download_dir = os.path.join(Path.home(), "Downloads")
        self.final_destination_dir = os.path.join(save_folder, "00_合体フォルダ")
        self.temp_extract_dir = os.path.join(self.final_destination_dir, "temp_unzip")
    #フォルダの存在チェックと作成(ない時だけ作る)
    #exist_ok=True ->すでにそのディレクトリが存在していてもエラーにならない
        os.makedirs(self.download_dir, exist_ok=True)
        os.makedirs(self.final_destination_dir, exist_ok=True)
        os.makedirs(self.temp_extract_dir, exist_ok=True)
    # EdgeDriverのパス
        driver_path = r"msedgedriver.exe"

        service = Service(driver_path)
        #Edge  オプション設定
        options = webdriver.EdgeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--remote-allow-origins=*")
        # ダウンロード設定を追加
        prefs = {
        "download.default_directory": self.download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
        }
        # WebDriver起動
        self.driver = webdriver.Edge(service=service, options=options)
        # Web_open ｱﾄﾞﾚｽｾｯﾄ
        self.login_url = "https://ur-systems.dg.dox.jp/w/project"
    #======================================================================================#
    #  処理基幹部
    #======================================================================================#
    def run(self):
        # --- メインの実行処理 ---
        try:
            # 1. ログインページへアクセス
            self.driver.get(self.login_url)

            # 2. ログイン情報を入力
            WebDriverWait(self.driver, 2).until(
                EC.presence_of_element_located((By.ID, "username"))
            ).send_keys(self.username)
            self.driver.find_element(By.NAME, "j_password").send_keys(self.password)
            self.driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()
    
            # 3. ログイン後のページに遷移するまで待機
            WebDriverWait(self.driver, 5).until(EC.url_changes(self.login_url))
            print("ログインに成功しました。")

            # 4. 「kyouryoku_seikyu」プロジェクトのリンクをクリック
            project_link_id = "detail:projectlist_1:_id58"
            WebDriverWait(self.driver, 2).until(
                EC.element_to_be_clickable((By.ID, project_link_id))
            ).click()
            print("「kyouryoku_seikyu」プロジェクトに移動しました。")

            # 5. ファイル探索の再帰関数を呼び出す
            self.traverse_folders()

        except Exception as e:
            print(f"致命的なエラーが発生しました: {e}")

        finally:
            # 7. 処理が完了したら、このメッセージを表示する
            print("ファイル取得処理完了")
    
            # ブラウザを閉じる（デバッグ中はコメントアウト）
            # driver.quit()
            pass

#-----------------------------------------------------------------------
#  解凍用フォルダへ対象ファイルをＣｏｐｙ
#-----------------------------------------------------------------------
    def unzip_file_and_move(self,zip_path, target_dir):
    #---指定されたZIPファイルを解凍し、中のファイルを指定ディレクトリに移動します。
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                for member in zip_ref.infolist():
                    raw_filename = member.filename.encode('cp437')
                    result = chardet.detect(raw_filename)
                    charset = result['encoding']
                
                    try:
                        corrected_filename = raw_filename.decode(charset)
                    except:
                        corrected_filename = raw_filename.decode('cp932', 'ignore')

                    member_path = os.path.join(self.temp_extract_dir, corrected_filename)
                
                    if member.is_dir():
                        os.makedirs(member_path, exist_ok=True)
                    else:
                        os.makedirs(os.path.dirname(member_path), exist_ok=True)
                        with open(member_path, "wb") as f:
                            f.write(zip_ref.read(member))

            print(f"'{zip_path}' を一時フォルダに解凍しました。")

            for root, _, files in os.walk(self.temp_extract_dir):
                for file in files:
                    source_path = os.path.join(root, file)
                    destination_path = os.path.join(target_dir, file)
                    shutil.move(source_path, destination_path)
                    print(f"ファイル '{file}' を '{target_dir}' に移動しました。")

            shutil.rmtree(self.temp_extract_dir)
            print("一時解凍フォルダを削除しました。")
            return True
        except zipfile.BadZipFile:
            print(f"警告: '{zip_path}' は有効なZIPファイルではありません。")
            return False
        except Exception as e:
            print(f"ZIPファイルの解凍またはファイル移動中にエラーが発生しました: {e}")
            return False
#-----------------------------------------------------------------------
#  再帰探索しファイルをダウンロード
#-----------------------------------------------------------------------
    def traverse_folders(self):
    
        #現在のページ内のフォルダとファイルを再帰的に探索し、
        #特定の種類のファイルをダウンロードします。
    
        try:
            WebDriverWait(self.driver, 2).until(
                EC.presence_of_element_located((By.ID, "detail:nodelist"))
            )
            self.List_select()
            self.folders_select()
            self.files_select()
        except TimeoutException:
            print("フォルダ内に要素がありませんでした。")
            return

#-----------------------------------------------------------------------
#  リスト構造体を1件づつ、リストへ
#-----------------------------------------------------------------------
    def List_select(self):    
     
        self.folders = []
        self.files = []
    #---<detail:nodelist" のテーブル内で、class に list を含む行
        nodes = self.driver.find_elements(By.XPATH, "//table[@id='detail:nodelist']//tr[contains(@class, 'list')]//a")
    #---<list class 1ｹﾝﾁｭｳｼｭﾂ
        for node in nodes:
            node_text = node.text.strip()
            link_href = node.get_attribute('href')
            if "icon_folder_small.gif" in node.get_attribute("innerHTML") and node_text:
                self.folders.append({'text': node_text, 'href': link_href})
            elif node_text:
                self.files.append({'text': node_text, 'href': link_href})
#-----------------------------------------------------------------------
#  リスト構造体フォルダ探索
#-----------------------------------------------------------------------
    def folders_select(self):    
        for folder_info in self.folders:
            folder_name = folder_info['text']
            folder_href = folder_info['href']
        
            print(f"フォルダ '{folder_name}' を見つけました。探索します。")
        
            self.driver.get(folder_href)
        
            self.traverse_folders()
        
            self.driver.back()
            time.sleep(1)
            print(f"フォルダ '{folder_name}' から戻りました。")
    
#-----------------------------------------------------------------------
#  リスト構造体フォルダ探索
#-----------------------------------------------------------------------
    def files_select(self):       
        
    #-----<ファイル選択>
        self.found_target_file = False
        for file_info in self.files:
            self.file_name = file_info['text']
            self.file_href = file_info['href']
        
            if self.file_name.lower().endswith(self.target_extensions):
                print(f"ダウンロード対象のファイルです: {self.file_name}")
                self.found_target_file = True
            
                # ファイルが既に存在する場合でも上書きしてダウンロードする
                self.driver.get(self.file_href) #ｸﾘｯｸしてダウンロード
                print(f"'{self.file_name}' のダウンロードを開始しました。")
                time.sleep(1)
                self.copy_files_to_final_destination()
            
                self.driver.back()
                time.sleep(1)

        if not self.folders and not self.found_target_file:
            print("ファイルは存在しませんでした。")
#-----------------------------------------------------------------------
#  解凍用フォルダへ対象ファイルをＣｏｐｙ
#-----------------------------------------------------------------------
    def copy_files_to_final_destination(self,):
        #ダウンロードフォルダ内の対象ファイルを最終的な保存先にコピーします。
        try:
            copied_files_count = 0
            source_path = os.path.join(self.download_dir, self.file_name)
            destination_path = os.path.join(self.final_destination_dir, self.file_name)
            # ファイルが既に存在する場合でも上書きしてコピーする
            shutil.copy2(source_path, destination_path)
            copied_files_count += 1
        except Exception as e:
            print(f"ファイルのコピー中にエラーが発生しました: {e}")
