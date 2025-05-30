pip install pyinstaller

pyinstaller --onefile --windowed --version-file=version.txt --icon=icon.ico --hidden-import win32timezone NPD_GUI.py