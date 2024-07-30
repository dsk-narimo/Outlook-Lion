import os #ファイル操作ライブラリ
import re #正規表現ライブラリ
import win32com.client
from models.outlook_model import Outlook


class OutlookController():
    def __init__(self):
        """
        コンストラクタ
        :receive_address: 受信する自身のアドレス
        """
        #outlookアプリケーションにアクセス
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")


    def import_target_mail(self, folder_name, sender_address, receive_address):
        """
        メール取得
        :folder_name: 取得対象のフォルダ名
        :sender_address: 送信者アドレス
        :receive_address: 受信するアドレス
        :return: 取得したメール
        """
        #受信アドレス(自分のアドレス)を指定
        account = None
        for acc in self.outlook.Folders:
            if acc.Name == receive_address:
                account = acc
                break

        if account is None:
            raise ValueError(f"メールアドレス {receive_address} が見つかりません")
        
        #受信フォルダを指定
        folder = account.Folders[folder_name]

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