import os #ファイル操作ライブラリ
import re #正規表現ライブラリ
import win32com.client
from models.outlook_model import Outlook
from models.exists_checker import AddressExistsCheck

class OutlookController():
    def __init__(self):
        """
        コンストラクタ
        :receive_address: 受信する自身のアドレス
        """
        #outlookアプリケーションにアクセス
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")


    def import_target_mail(self, folder, sender_address, receive_address):

        #受信フォルダの中から対象のメールアドレスから来たメールを取得する
        inbox_mails = folder.Items
        target_mails = []
        for message in inbox_mails:
            if sender_address in message.SenderEmailAddress:
                attachments = []

                for i in range(1,message.Attachments.Count + 1):
                    attachment = message.Attachments.Item(i)
                    attachments.append(attachment)

                email = Outlook(message.Subject, message.SenderName, message.Body, attachments,message)
                target_mails.append(email)

        return target_mails


    def extract_info(self, contents, pattern):
        """
        メールの本文から特定の情報を抽出する
        :contents: メールの本文
        :pattern: 情報を抽出するパターン
        """
        match = re.search(pattern, contents)
        if match:
            return match.group(1)
        return None
    

    def save_attached_file(self,emails,save_path):
        """
        メールの添付ファイルを保存
        """
        for email in emails:
            email.save_file(save_path)
            
    def move_to_folder(self, finished_keys,finished_pdf_mails,finished_pass_mails,deleted_items_folder):
        """
        処理済みメールを特定のフォルダに移動
        """
        moved_mail_count = 0
        for finished_key in finished_keys:
            for finished_pdf_mail in finished_pdf_mails:
                if finished_key in finished_pdf_mail.contents:
                    target_remove_mail1 = finished_pdf_mail
                    break
                
            for finished_pass_mail in finished_pass_mails:
                if finished_key in finished_pass_mail.contents:
                    target_remove_mail2 = finished_pass_mail
                    break

            # 両方のメールが見つかった場合に移動
            if target_remove_mail1 and target_remove_mail2:
                try:
                    # メールを削除済みアイテムフォルダに移動
                    target_remove_mail1.message.Move(deleted_items_folder)
                    target_remove_mail2.message.Move(deleted_items_folder)
                    moved_mail_count += 1
                except Exception as e:
                    # エラーが発生した場合はそのメールの処理をスキップ
                    print(f"メール移動失敗: {e}")
                    continue
        
        return moved_mail_count