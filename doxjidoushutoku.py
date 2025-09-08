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
#  �R���X�g���N�^
#-----------------------------------------------------------------------
    def __init__(self):
    #�Œ�l���
        self.username = "t-tamura"
        self.password = "merumo0526!!"
        self.target_extensions = ('.pdf', '.xlsx', '.docx', '.zip')

    #�i�[��t�H���_�쐬
        base_path = os.path.join(Path.home(), "Desktop", "ai_ocr", "���[��","dox")
        current_time_str = datetime.now().strftime("%Y%m%d_%H%M")
        dated_path = os.path.join(base_path, current_time_str)
        os.makedirs(dated_path, exist_ok=True)
        save_folder=dated_path   
    # �t�H���_�̎w��
        self.save_folder = save_folder
        self.download_dir = os.path.join(Path.home(), "Downloads")
        self.final_destination_dir = os.path.join(save_folder, "00_���̃t�H���_")
        self.temp_extract_dir = os.path.join(self.final_destination_dir, "temp_unzip")
    #�t�H���_�̑��݃`�F�b�N�ƍ쐬(�Ȃ����������)
    #exist_ok=True ->���łɂ��̃f�B���N�g�������݂��Ă��Ă��G���[�ɂȂ�Ȃ�
        os.makedirs(self.download_dir, exist_ok=True)
        os.makedirs(self.final_destination_dir, exist_ok=True)
        os.makedirs(self.temp_extract_dir, exist_ok=True)
    # EdgeDriver�̃p�X
        driver_path = r"msedgedriver.exe"

        service = Service(driver_path)
        #Edge  �I�v�V�����ݒ�
        options = webdriver.EdgeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--remote-allow-origins=*")
        # �_�E�����[�h�ݒ��ǉ�
        prefs = {
        "download.default_directory": self.download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
        }
        # WebDriver�N��
        self.driver = webdriver.Edge(service=service, options=options)
        # Web_open ���ڽ���
        self.login_url = "https://ur-systems.dg.dox.jp/w/project"
    #======================================================================================#
    #  �������
    #======================================================================================#
    def run(self):
        # --- ���C���̎��s���� ---
        try:
            # 1. ���O�C���y�[�W�փA�N�Z�X
            self.driver.get(self.login_url)

            # 2. ���O�C���������
            WebDriverWait(self.driver, 2).until(
                EC.presence_of_element_located((By.ID, "username"))
            ).send_keys(self.username)
            self.driver.find_element(By.NAME, "j_password").send_keys(self.password)
            self.driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()
    
            # 3. ���O�C����̃y�[�W�ɑJ�ڂ���܂őҋ@
            WebDriverWait(self.driver, 5).until(EC.url_changes(self.login_url))
            print("���O�C���ɐ������܂����B")

            # 4. �ukyouryoku_seikyu�v�v���W�F�N�g�̃����N���N���b�N
            project_link_id = "detail:projectlist_1:_id58"
            WebDriverWait(self.driver, 2).until(
                EC.element_to_be_clickable((By.ID, project_link_id))
            ).click()
            print("�ukyouryoku_seikyu�v�v���W�F�N�g�Ɉړ����܂����B")

            # 5. �t�@�C���T���̍ċA�֐����Ăяo��
            self.traverse_folders()

        except Exception as e:
            print(f"�v���I�ȃG���[���������܂���: {e}")

        finally:
            # 7. ����������������A���̃��b�Z�[�W��\������
            print("�t�@�C���擾��������")
    
            # �u���E�U�����i�f�o�b�O���̓R�����g�A�E�g�j
            # driver.quit()
            pass

#-----------------------------------------------------------------------
#  �𓀗p�t�H���_�֑Ώۃt�@�C�����b������
#-----------------------------------------------------------------------
    def unzip_file_and_move(self,zip_path, target_dir):
    #---�w�肳�ꂽZIP�t�@�C�����𓀂��A���̃t�@�C�����w��f�B���N�g���Ɉړ����܂��B
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

            print(f"'{zip_path}' ���ꎞ�t�H���_�ɉ𓀂��܂����B")

            for root, _, files in os.walk(self.temp_extract_dir):
                for file in files:
                    source_path = os.path.join(root, file)
                    destination_path = os.path.join(target_dir, file)
                    shutil.move(source_path, destination_path)
                    print(f"�t�@�C�� '{file}' �� '{target_dir}' �Ɉړ����܂����B")

            shutil.rmtree(self.temp_extract_dir)
            print("�ꎞ�𓀃t�H���_���폜���܂����B")
            return True
        except zipfile.BadZipFile:
            print(f"�x��: '{zip_path}' �͗L����ZIP�t�@�C���ł͂���܂���B")
            return False
        except Exception as e:
            print(f"ZIP�t�@�C���̉𓀂܂��̓t�@�C���ړ����ɃG���[���������܂���: {e}")
            return False
#-----------------------------------------------------------------------
#  �ċA�T�����t�@�C�����_�E�����[�h
#-----------------------------------------------------------------------
    def traverse_folders(self):
    
        #���݂̃y�[�W���̃t�H���_�ƃt�@�C�����ċA�I�ɒT�����A
        #����̎�ނ̃t�@�C�����_�E�����[�h���܂��B
    
        try:
            WebDriverWait(self.driver, 2).until(
                EC.presence_of_element_located((By.ID, "detail:nodelist"))
            )
            self.List_select()
            self.folders_select()
            self.files_select()
        except TimeoutException:
            print("�t�H���_���ɗv�f������܂���ł����B")
            return

#-----------------------------------------------------------------------
#  ���X�g�\���̂�1���ÂA���X�g��
#-----------------------------------------------------------------------
    def List_select(self):    
     
        self.folders = []
        self.files = []
    #---<detail:nodelist" �̃e�[�u�����ŁAclass �� list ���܂ލs
        nodes = self.driver.find_elements(By.XPATH, "//table[@id='detail:nodelist']//tr[contains(@class, 'list')]//a")
    #---<list class 1��������
        for node in nodes:
            node_text = node.text.strip()
            link_href = node.get_attribute('href')
            if "icon_folder_small.gif" in node.get_attribute("innerHTML") and node_text:
                self.folders.append({'text': node_text, 'href': link_href})
            elif node_text:
                self.files.append({'text': node_text, 'href': link_href})
#-----------------------------------------------------------------------
#  ���X�g�\���̃t�H���_�T��
#-----------------------------------------------------------------------
    def folders_select(self):    
        for folder_info in self.folders:
            folder_name = folder_info['text']
            folder_href = folder_info['href']
        
            print(f"�t�H���_ '{folder_name}' �������܂����B�T�����܂��B")
        
            self.driver.get(folder_href)
        
            self.traverse_folders()
        
            self.driver.back()
            time.sleep(1)
            print(f"�t�H���_ '{folder_name}' ����߂�܂����B")
    
#-----------------------------------------------------------------------
#  ���X�g�\���̃t�H���_�T��
#-----------------------------------------------------------------------
    def files_select(self):       
        
    #-----<�t�@�C���I��>
        self.found_target_file = False
        for file_info in self.files:
            self.file_name = file_info['text']
            self.file_href = file_info['href']
        
            if self.file_name.lower().endswith(self.target_extensions):
                print(f"�_�E�����[�h�Ώۂ̃t�@�C���ł�: {self.file_name}")
                self.found_target_file = True
            
                # �t�@�C�������ɑ��݂���ꍇ�ł��㏑�����ă_�E�����[�h����
                self.driver.get(self.file_href) #�د����ă_�E�����[�h
                print(f"'{self.file_name}' �̃_�E�����[�h���J�n���܂����B")
                time.sleep(1)
                self.copy_files_to_final_destination()
            
                self.driver.back()
                time.sleep(1)

        if not self.folders and not self.found_target_file:
            print("�t�@�C���͑��݂��܂���ł����B")
#-----------------------------------------------------------------------
#  �𓀗p�t�H���_�֑Ώۃt�@�C�����b������
#-----------------------------------------------------------------------
    def copy_files_to_final_destination(self,):
        #�_�E�����[�h�t�H���_���̑Ώۃt�@�C�����ŏI�I�ȕۑ���ɃR�s�[���܂��B
        try:
            copied_files_count = 0
            source_path = os.path.join(self.download_dir, self.file_name)
            destination_path = os.path.join(self.final_destination_dir, self.file_name)
            # �t�@�C�������ɑ��݂���ꍇ�ł��㏑�����ăR�s�[����
            shutil.copy2(source_path, destination_path)
            copied_files_count += 1
        except Exception as e:
            print(f"�t�@�C���̃R�s�[���ɃG���[���������܂���: {e}")
