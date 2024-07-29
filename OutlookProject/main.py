import json5
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


#メインスクリプト
def main():
    try:
        
        settings = load_config()
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

        print("zipファイルの保存・解凍処理を開始します。")
        print('\n')

        # デバッグ時とパッケージ化時のパスを設定
        if getattr(sys, 'frozen', False):
            driver_path = os.path.basename(driver_path)  # PyInstallerの場合はファイル名のみ
        else:
            driver_path = os.path.join(os.path.dirname(__file__), driver_path)
        
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
        
        #対象メールをインポート
        target_mails = outlook_controller.import_target_mail(folder_name,sender_address,receive_address)

        print("LION事務器ダウンロード用URLにアクセスし解凍処理を開始します。")
        print('------ダウンロード中です(※以下メッセージは無視してください)------')
        #パスワードメールからDLパスワード、ZIP解凍パスワード、CSV名を取得
        csv_password_dict,url_password_dict = outlook_processor.get_password_info(csv_password_subject,csv_password_dict,target_mails)
        
        #注文書メールからダウンロードURLを抽出
        csv_order_dict = outlook_processor.get_csv_info(csv_order_subject,csv_order_dict,target_mails)

        #ダウンロードURLからZIPファイルをダウンロード
        download_zip_dict = selenium_processor.download_file(csv_order_dict, url_password_dict)
        print('\n')
        print('------ダウンロード終了-------')
        print('\n')
        print('------解凍処理開始-------')
        #帳票番号との紐づけからZIPファイルの解凍
        outlook_processor.extract_and_save_zip_files(download_zip_dict, csv_password_dict,csv_save_path)
        
        #パスワードメールから帳票番号とパスワード抽出
        pdf_password_dict,null_dict = outlook_processor.get_password_info(pdf_password_subject,pdf_password_dict,target_mails)

        #注文メールから帳票番号を抽出と添付ファイルの保存
        pdf_order_dict = outlook_processor.get_pdf_info(pdf_order_subject,pdf_order_dict,target_mails,pdf_save_path)

        #帳票番号との紐づけからZIPファイルの解凍
        outlook_processor.extract_and_save_zip_files(pdf_order_dict, pdf_password_dict,pdf_save_path)

    except Exception as e:
        with open("error_log.txt", "w") as f:
            f.write("Error: " + str(e) + "\n")
        print("Error: " + str(e) + "\n")
    finally:
        print('\n')
        input("処理終了。Enterを押して閉じてOKです")


def load_config():
    """
    設定ファイル読み込み
    :return: 設定の辞書
    """
    if getattr(sys, 'frozen', False):
        # PyInstallerでパッケージされた場合
        config_path = os.path.join(sys._MEIPASS, 'config.json5')
    else:
        # 通常の実行の場合
        config_path = os.path.join(os.path.dirname(__file__), 'config.json5')
        
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"設定ファイルが見つかりません: {config_path}")
    with open(config_path, 'r', encoding='utf-8') as file:
        config = json5.load(file)
    return config['Settings']


if __name__ == "__main__":
    main()