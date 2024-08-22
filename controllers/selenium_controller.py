import time
import os
import sys
import logging
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium import webdriver

# Seleniumのログレベルを設定
logging.getLogger('selenium').setLevel(logging.CRITICAL)

class SeleniumController:
    def __init__(self, driver_path, csv_save_path):
        if getattr(sys, 'frozen', False):
            # PyInstallerでパッケージされた場合
            self.driver_path = os.path.join(sys._MEIPASS, driver_path)
            self.csv_save_path = os.path.join(sys._MEIPASS, csv_save_path)
        else:
            # 通常の実行の場合
            self.driver_path = driver_path
            self.csv_save_path = csv_save_path
            #print(f"Current working directory: {os.getcwd()}")
            #print(f"Driver path exists: {os.path.exists(driver_path)}")

    def download_file(self,order_dict,password_dict):

        # EdgeDriverが起動できないときは終了
        if not self.driver_path:
            print("Edge Driverのバージョンが異なります.管理者に確認してください")
            return

        #ブラウザアクセス時のオプション設定
        options = EdgeOptions() 
        options.add_experimental_option("excludeSwitches", ["enable-automation"]) #自動化メッセージの非表示
        options.add_experimental_option('useAutomationExtension', False) #ブラウザが自動化されていることを隠す(これがないとEdgeにブロックされる)
        #options.add_argument(f"user-agent={user_agent}") # ユーザーエージェントの指定
        options.add_argument("--log-level=3")  # ログレベルをERRORに設定
        options.add_argument("--disable-logging")  # 追加のログメッセージの抑制
        options.add_argument("--disable-extensions")  # 拡張機能を無効化してログを減らす
        options.add_argument("--disable-default-apps")  # デフォルトアプリを無効化してログを減らす
        # options.add_argument("--no-sandbox")  # サンドボックスを無効化
        # options.add_argument("--disable-gpu")  # GPUの無効化
        # options.add_argument("--headless")  # ヘッドレスモードで実行
        options.add_experimental_option("prefs", {
            "download.default_directory": self.csv_save_path,  # ファイルの保存先のパス
            "download.prompt_for_download": False,        # ダウンロード確認ダイアログを表示しない
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
            })

        download_zip_dict = {}

        for key, url in order_dict.items():
            if key in password_dict:

                edge_service = EdgeService(executable_path=self.driver_path)
                driver = webdriver.Edge(service=edge_service,options=options)
                # 指定したURLにアクセス
                driver.get(url)
                # パスワード入力欄を探し、パスワードを入力する
                password_field = driver.find_element(By.CSS_SELECTOR, '#dlPassword')
                password_field.send_keys(password_dict[key])
                # #ダウンロードボタンを探し押下する
                download_button = driver.find_element(By.CSS_SELECTOR, "#submitbutton")
                download_button.click()

                #ダウンロード確認ページから保存ファイルを選択
                download_link = driver.find_element(By.CSS_SELECTOR, "#shadow > table > tbody > tr > td > form > div.box2 > table > tbody > tr > td:nth-child(1) > div:nth-child(1) > table > tbody > tr:nth-child(2) > td:nth-child(1) > div > div > a")
                download_link.click()

                #ダウンロード完了まえ待機
                time.sleep(8)
                driver.quit()

                #ダウンロードファイルを取得
                downloaded_files = set(os.listdir(self.csv_save_path))
                latest_time = 0
                latest_file = None
                # ファイルのダウンロード時間をもとに最も新しいダウンロードファイルを見つける
                for file in downloaded_files:
                    if file.endswith(".zip"):
                        file_path = os.path.join(self.csv_save_path, file)
                        download_time = os.path.getctime(file_path)
                        if download_time > latest_time:
                            latest_time = download_time
                            latest_file = file_path

                if latest_file:
                    download_zip_dict[key] = latest_file

        return download_zip_dict

