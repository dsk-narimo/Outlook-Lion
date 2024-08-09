import os
import sys

class FolderExistsCheck:
    @staticmethod
    def check_folder_exists(folder_path):
        return os.path.exists(folder_path)

    @staticmethod
    def check_file_exists(file_path):
        if getattr(sys, 'frozen', False):
            #driver_path = os.path.basename(driver_path)  # PyInstallerの場合はファイル名のみ
            file_path = os.path.join(sys._MEIPASS, os.path.basename(file_path))
        return os.path.exists(file_path)


class AddressExistsCheck:
    @staticmethod
    def check_receive_exists(outlook_object, receive_address):
        """Check if the receive address exists in the Outlook object."""
        for account in outlook_object.Folders:
            if account.Name == receive_address:
                return True
        return False

    @staticmethod
    def check_sender_exists(folder, sender_address):
        """Check if the sender address exists in the specified folder."""
        for message in folder.Items:
            if message.Class == 43:  # 43 is the class code for MailItem
                if sender_address in message.SenderEmailAddress:
                    return True
        return False

    @staticmethod
    def folder_exists(account, folder_name):
        """Check if the specified folder exists within the given account."""
        for folder in account.Folders:
            if folder.Name == folder_name:
                return True
        return False