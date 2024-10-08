import zipfile
import os
from PyPDF2 import PdfFileWriter, PdfFileReader


class ZipFileHandler:
    @staticmethod
    def extract_zip(zip_path, password, extract_to):
        """
        ZIPファイルを解凍する
        :param zip_path: ZIPファイルのパス
        :param password: 解凍パスワード
        :param extract_to: 解凍先のパス
        """
        # ZIPファイルでない場合、処理をスキップ
        if not zipfile.is_zipfile(zip_path):
            print(f"破損したファイルが保存された可能性があります。処理終了後に確認してください。{zip_path} ")
            return False
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            for file_info in zip_ref.infolist():
                # ファイル名のエンコーディングを修正
                try:
                    file_name = file_info.filename.encode('cp437').decode('shift_jis')
                except UnicodeDecodeError:
                    file_name = file_info.filename.encode('utf-8').decode('utf-8')

                # ファイルをパスワードで解凍
                extract_result = False
                try:
                    with zip_ref.open(file_info, pwd=bytes(password, 'utf-8')) as source:
                        with open(os.path.join(extract_to, file_name), 'wb') as target:
                            target.write(source.read())
                    print(f"解凍完了: {file_name}")
                    extract_result = True
                except RuntimeError as e:
                    print(f"エラー: {file_name} を解凍できませんでした。パスワードが正しくない可能性があります: {e}")
                    continue
                except zipfile.BadZipFile as e:
                    print(f"エラー: ZIPファイルが破損しています: {e}")
                    continue
                except Exception as e:
                    print(f"エラー: {e}")
                    continue
        
        #保存したzipファイルを削除
        ZipFileHandler.delete_zip_file(zip_path)
        return extract_result

    @staticmethod
    def list_zip_files(directory):
        """
        ディレクトリ内のZIPファイルをリスト化する
        :param directory: ディレクトリのパス
        :return: ZIPファイルのリスト
        """
        return [f for f in os.listdir(directory) if f.endswith('.zip')]
    
    @staticmethod
    def delete_zip_file(zip_path):
        """
        解凍し終わったZIPファイルを削除する
        """
        try:
            os.remove(zip_path)
            #print(f"ZIPファイルを削除しました: {zip_path}")
        except Exception as e:
            print(f"ZIPファイルの削除に失敗しました: {e}")