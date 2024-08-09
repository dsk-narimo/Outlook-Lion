#import json5
import os
import sys

# OutlookProjectディレクトリをsys.pathに追加
if getattr(sys, 'frozen', False):
    # PyInstallerでパッケージされた場合、実行ファイルのパスを取得
    project_root = os.path.join(sys._MEIPASS)
else:
    # 通常の実行の場合、カレントディレクトリを取得
    project_root = os.path.abspath(os.path.dirname(__file__))

sys.path.insert(0, project_root)

#デバッグ用
# print("Project root:", project_root)
# print("sys.path:", sys.path)
# print("Controllers directory exists:", os.path.exists(os.path.join(project_root, 'controllers')))
# print("Models directory exists:", os.path.exists(os.path.join(project_root, 'models')))
# print("Config file exists:", os.path.exists(os.path.join(project_root, 'config.json5')))
# print("Edge driver exists:", os.path.exists(os.path.join(project_root, 'msedgedriver.exe')))

from controllers.outlook_controller import OutlookController
from controllers.outlook_processor import OutlookProcessor
from controllers.selenium_controller import SeleniumController
from models.exists_checker import FolderExistsCheck,AddressExistsCheck


#メインスクリプト
def main():
    try:
        # # 設定を直接定義
        # settings = {
        #     "受信者アドレス": "dsk_gyoumu@daishinkogyo.co.jp",
        #     "送信者アドレス": "lion_order@lion-jimuki.co.jp",
        #     "受信フォルダ名": "LION（FTS）",
        #     "PDF保存先パス": "\\\\192.168.175.4\\fax\\本社\\受信FAX\\本社_大熊\\LION(FTS)\\注文書出力分",
        #     "CSV保存先パス": "\\\\192.168.175.4\\業務課\\OKUMA\\LIONETサービス(新）",
        #     "PDFパスワード件名": "[パスワードの通知] ご注文書",
        #     "PDF注文書件名": "ご注文書",
        #     "CSVパスワード件名": "パスワード通知",
        #     "CSV注文書件名": "ご注文データの送付",
        #     "処理後移動先フォルダ名": "LION（FTS）処理済",
        #     "Edgeドライバー": "msedgedriver.exe"
        # }
        
        #テスト用
        settings = {
            "受信者アドレス": "tatsuya_narimo@daishinkogyo.co.jp",
            "送信者アドレス": "lion_order@lion-jimuki.co.jp",
            "受信フォルダ名": "テスト",
            "PDF保存先パス": "\\\\192.168.175.4\\SourceTree\\pdf3",
            "CSV保存先パス": "\\\\192.168.175.4\\SourceTree\\csv3",
            "PDFパスワード件名": "[パスワードの通知] ご注文書",
            "PDF注文書件名": "ご注文書",
            "CSVパスワード件名": "パスワード通知",
            "CSV注文書件名": "ご注文データの送付",
            "処理後移動先フォルダ名": "削除済みアイテム",
            "Edgeドライバー": "msedgedriver.exe"
        }

        receive_address = settings['受信者アドレス']
        sender_address = settings['送信者アドレス']
        folder_name = settings['受信フォルダ名']
        pdf_save_path = settings['PDF保存先パス']
        csv_save_path = settings['CSV保存先パス']
        pdf_password_subject = settings['PDFパスワード件名']
        pdf_order_subject = settings['PDF注文書件名']
        csv_password_subject = settings['CSVパスワード件名']
        csv_order_subject = settings['CSV注文書件名']
        remove_folder_name = settings['処理後移動先フォルダ名']
        driver_path = settings['Edgeドライバー']
        
        #保存フォルダの存在チェック
        if not FolderExistsCheck.check_folder_exists(pdf_save_path):
            print(f"PDF保存先フォルダが存在しません。フォルダ構成変更してないか確認してください。: {pdf_save_path}")
            return

        if not FolderExistsCheck.check_folder_exists(csv_save_path):
            print(f"CSV保存先フォルダが存在しません。フォルダ構成変更してないか確認してください。: {csv_save_path}")
            return

        
        pdf_password_dict ={
            "帳票番号": r"帳票番号\s*：\s*(\d+)",
            "パスワード": r"パスワード\s*：\s*([^\s]+)"
        }
        
        pdf_order_dict = {
            "帳票番号": r"帳票番号\s*：\s*(\d+)"
        }
        
        csv_order_dict = {
            "帳票番号": r'(\S+_注文データ_\S+\.csv)',
            "ダウンロードURL": r'ダウンロードURL\s*:\s*(https?://\S+)'
        }

        csv_password_dict = {
            "帳票番号": r'(\S+_注文データ_\S+\.csv)',
            "ダウンロードパスワード":  r"ダウンロードパスワード\s*：\s*([^\s]+)",
            "パスワード":  r"ZIPファイル解凍のパスワード\s*：\s*([^\s]+)",
        }
 
        #インスタンス生成
        outlook_processor = OutlookProcessor(folder_name,
                                            pdf_save_path,
                                            csv_save_path,
                                            sender_address,
                                            receive_address,
                                            remove_folder_name,
                                            pdf_password_subject,
                                            pdf_order_subject,
                                            csv_password_subject,
                                            csv_order_subject,
                                            driver_path)

        #インスタンス化
        outlook_controller = OutlookController()
            
        selenium_processor = SeleniumController(driver_path,csv_save_path)
        
        if not FolderExistsCheck.check_file_exists(driver_path):
            print(f"edgeドライバーが存在しません。管理者に問い合わせてください。: {driver_path}")
            return
        
        # 受信フォルダが存在するかチェックして指定
        account = None
        for acc in outlook_controller.outlook.Folders:
            if acc.Name == receive_address:
                account = acc
                break

        if account is None:
            print(f"受信アドレス {receive_address} が見つかりません")
            return

        receive_folder_exists = AddressExistsCheck.folder_exists(account, folder_name)
        if not receive_folder_exists:
            print(f"受信フォルダ {folder_name} が見つかりません")
            return
        
        remove_folder_exists = AddressExistsCheck.folder_exists(account, remove_folder_name)
        if not remove_folder_exists:
            print(f"処理済み格納フォルダ {remove_folder_name} が見つかりません")
            return

        # フォルダを取得
        folder = None
        for f in account.Folders:
            if f.Name == folder_name:
                folder = f
                break

        # 送信者アドレスが存在するかチェック
        sender_exists = AddressExistsCheck.check_sender_exists(folder, sender_address)
        if not sender_exists:
            print(f"送信者アドレス {sender_address} からのメールが見つかりません")
            return
        
        print("zipファイルの保存・解凍処理を開始します。")
        print('\n')
        
        #処理済メールの移動先を指定
        recipient_account = outlook_controller.outlook.Folders(receive_address)
        deleted_items_folder = recipient_account.Folders(remove_folder_name)
        
        #対象メールをインポート
        target_mails = outlook_controller.import_target_mail(folder,sender_address,receive_address)

        print("LION事務器ダウンロード用URLにアクセスし解凍処理を開始します。")
        print('------ダウンロード中です(※以下メッセージは無視してください)------')
        #パスワードメールからDLパスワード、ZIP解凍パスワード、CSV名を取得
        csv_password_dict,url_password_dict,finished_cvs_mails = outlook_processor.get_password_info(csv_password_subject,csv_password_dict,target_mails)
        
        #注文書メールからダウンロードURLを抽出
        csv_order_dict,finished_url_mails = outlook_processor.get_csv_info(csv_order_subject,csv_order_dict,target_mails)

        #ダウンロードURLからZIPファイルをダウンロード
        download_zip_dict = selenium_processor.download_file(csv_order_dict, url_password_dict)
        
        print('\n')
        print('------ダウンロード終了-------')
        print('\n')
        print('------解凍処理開始-------')
        #帳票番号との紐づけからZIPファイルの解凍
        finished_keys = []
        finished_keys = outlook_processor.extract_and_save_zip_files(download_zip_dict, csv_password_dict,csv_save_path)

        #処理済メールの移動処理
        moved_csv_mail_count = outlook_controller.move_to_folder(finished_keys,finished_cvs_mails,finished_url_mails,deleted_items_folder)
        print((f"注文データCSV保存件数：{moved_csv_mail_count}件"))
        print('\n')
        #パスワードメールから帳票番号とパスワード抽出
        pdf_password_dict,null_dict,finished_pass_mails = outlook_processor.get_password_info(pdf_password_subject,pdf_password_dict,target_mails)

        #注文メールから帳票番号を抽出と添付ファイルの保存
        pdf_order_dict,finished_pdf_mails = outlook_processor.get_pdf_info(pdf_order_subject,pdf_order_dict,target_mails,pdf_save_path)

        #帳票番号との紐づけからZIPファイルの解凍
        finished_keys = []
        finished_keys = outlook_processor.extract_and_save_zip_files(pdf_order_dict, pdf_password_dict,pdf_save_path)
        
        #処理済メールの移動処理
        moved_pdf_mail_count = outlook_controller.move_to_folder(finished_keys,finished_pdf_mails,finished_pass_mails,deleted_items_folder)

        print(f"注文書PDF保存件数：{moved_pdf_mail_count}件")
        
    except Exception as e:
        print("Error: " + str(e) + "\n")
    finally:
        print('\n')
        input("処理終了。Enterを押して閉じてOKです。")


# def load_config():
#     """
#     設定ファイル読み込み
#     :return: 設定の辞書
#     """
#     if getattr(sys, 'frozen', False):
#         # PyInstallerでパッケージされた場合
#         config_path = os.path.join(sys._MEIPASS, 'config.json5')
#     else:
#         # 通常の実行の場合
#         config_path = os.path.join(os.path.dirname(__file__), 'config.json5')
        
#     if not os.path.exists(config_path):
#         raise FileNotFoundError(f"設定ファイルが見つかりません: {config_path}")
#     with open(config_path, 'r', encoding='utf-8') as file:
#         config = json5.load(file)
#     return config['Settings']


if __name__ == "__main__":
    main()