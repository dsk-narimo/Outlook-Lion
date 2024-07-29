#メールデータの管理
import os


class Outlook():
    def __init__(self,subject_name,sender_address,contents,attached_files,message):
        """
        コンストラクタ
        :param subject_name: 件名
        :param sender_address: 送信者アドレス
        :param contents: メールの中身
        :param attached_files: 添付ファイル(複数)
        """
        self.subject_name = subject_name
        self.sender_address = sender_address
        self.contents = contents
        self.attached_files = attached_files
        self.message = message


    def save_file(self, save_path):
        """
        添付ファイル保存
        :param save_path: 保存先パス
        """
        if not os.path.exists(save_path):
            os.makedirs(save_path)
        for attachment in self.attached_files:
            attachment.SaveAsFile(f"{save_path}/{attachment.FileName}")



