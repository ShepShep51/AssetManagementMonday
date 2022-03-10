import PyInstaller.__main__
import os
import shutil

PyInstaller.__main__.run(['Main.py', '--onefile', '--console'])
pull_path = r'C:\Users\dshepard\PycharmProjects\AssManUpload\venv\dist\Main.exe'
end_path = r'L:\7. Departmental Shortcuts\7.3 Asset Management\AssetManagementUpload'
shutil.copy(src=pull_path,dst=end_path)
