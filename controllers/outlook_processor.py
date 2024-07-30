import os
from controllers.outlook_controller import OutlookController
from models.zip_handler import ZipFileHandler

class OutlookProcessor:
    def __init__(self,
                folder_name,
                pdf_save_path,
                csv_save_path,
                sender_address,
                receive_address,
                remove_folder_name,
                pdf_password_subject,
                pdf_order_subject,
                csv_password_subject,
                csv_order_subject,
                driver_path):

        
        self.folder_name = folder_name
        self.pdf_save_path = pdf_save_path
        self.csv_save_path = csv_save_path
        self.sender_address = sender_address
        self.receive_address = receive_address
        self.remove_folder_name = remove_folder_name
        self.pdf_password_subject = pdf_password_subject
        self.pdf_order_subject = pdf_order_subject
        self.csv_password_subject = csv_password_subject
        self.csv_order_subject = csv_order_subject
        self.driver_path = driver_path
        #インスタンス化
        self.outlook_controller = OutlookController()


    def get_password_info(self,subject,patterns,emails):
        """
        パスワード通知のメール処理
        :param subject:件名
        :param patterns: 抽出パターン
        :return:
        """

        zip_password_dict = {}
        url_password_dict = {}
        recipient_account = self.outlook_controller.outlook.Folders(self.receive_address)
        deleted_items_folder = recipient_account.Folders(self.remove_folder_name)
        for email in emails:
            if subject == email.subject_name:
                info = {}
                for key, pattern in patterns.items():
                    #キー：”帳票番号:xxx”、値：パスワード:xxx
                    info[key] = self.outlook_controller.extract_info(email.contents, pattern)
                
                if "帳票番号" in info:
                    zip_password_dict[info["帳票番号"]] = info['パスワード'].strip('"')
                    if "ダウンロードパスワード" in info:
                        url_password_dict[info["帳票番号"]] = info['ダウンロードパスワード'].strip('"')

                    try:
                        # メールを削除済みアイテムフォルダに移動
                        email.message.Move(deleted_items_folder)
                    except Exception as e:
                        # エラーが発生した場合はそのメールの処理をスキップ
                        print(f"メール移動失敗: {e}")
                        continue

        return zip_password_dict, url_password_dict


    def get_pdf_info(self, subject, patterns,emails,pdf_save_path):
        """
        注文書のメール処理
        :param subject: 件名
        :param patterns: 正規表現パターン
        :return:
        """
        order_file_dict = {}
        recipient_account = self.outlook_controller.outlook.Folders(self.receive_address)
        deleted_items_folder = recipient_account.Folders(self.remove_folder_name)
        for email in emails:
            if subject == email.subject_name:
                info = {}
                for key, pattern in patterns.items():
                    #キー：”帳票番号:xxx”、値：zipファイルパス
                    info[key] = self.outlook_controller.extract_info(email.contents, pattern)
                #添付されたzipファイルをダウンロード
                self.outlook_controller.save_attached_file([email],pdf_save_path)
                if "帳票番号" in info:
                    for attachment in email.attached_files:
                        #zipファイルの保存パスを辞書登録
                        order_file_dict[info["帳票番号"]] = os.path.join(pdf_save_path, attachment.FileName)
                
                    try:
                        # メールを削除済みアイテムフォルダに移動
                        email.message.Move(deleted_items_folder)
                    except Exception as e:
                        # エラーが発生した場合はそのメールの処理をスキップ
                        print(f"メール移動失敗: {e}")
                        continue

        return order_file_dict
    
    def get_csv_info(self, subject, patterns,emails):
        """
        注文書のメール処理
        :param subject: 件名のフィルター
        :param patterns: 正規表現パターン
        :return:
        """

        order_file_dict = {}
        recipient_account = self.outlook_controller.outlook.Folders(self.receive_address)
        deleted_items_folder = recipient_account.Folders(self.remove_folder_name)
        for email in emails:
            if subject in email.subject_name:
                info = {}
                for key, pattern in patterns.items():
                    #キー：”帳票番号:xxx”、値：zipファイルパス
                    info[key] = self.outlook_controller.extract_info(email.contents, pattern)

                if "帳票番号" in info:
                    #zipファイルの保存パスを辞書登録
                    order_file_dict[info["帳票番号"]] = info['ダウンロードURL']

                    try:
                        # メールを削除済みアイテムフォルダに移動
                        email.message.Move(deleted_items_folder)
                    except Exception as e:
                        # エラーが発生した場合はそのメールの処理をスキップ
                        print(f"メール移動失敗: {e}")
                        continue

        return order_file_dict


    def extract_and_save_zip_files(self, order_file_dict, password_dict,save_path):
        """
        ZIPファイルの解凍処理
        :param order_file_dict
        :param password_dict
        :return:
        """
        for key in order_file_dict:

            target_password = None
            # 対象パスワードを見つける
            if key in password_dict:
                target_password = password_dict[key]
                zip_path = order_file_dict[key]
                extract_to = save_path  # 解凍先を保存ディレクトリに変更
                ZipFileHandler.extract_zip(zip_path, target_password, extract_to)