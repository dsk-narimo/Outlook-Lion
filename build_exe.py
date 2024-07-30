import subprocess

# requirements.txtからモジュール名を読み取る
with open('requirements.txt') as f:
    modules = f.read().splitlines()

# hidden-importオプションを生成（バージョン番号を除外）
hidden_imports = ' '.join([f'--hidden-import {module.split("==")[0]}' for module in modules if module.strip()])

# PyInstallerコマンドを一行で組み立て
pyinstaller_command = (
    f'pyinstaller --onefile --add-data "controllers;controllers" '
    f'--add-data "models;models" --add-data "msedgedriver.exe;." '
    f'--add-data "requirements.txt;." '
    f'--paths venv\\Lib\\site-packages {hidden_imports} main.py'
)

# デバッグ用にコマンドを表示
print(pyinstaller_command)

# PyInstallerを実行
subprocess.run(pyinstaller_command, shell=True, check=True)
