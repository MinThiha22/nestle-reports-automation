pip install pyinstaller

pyinstaller --onefile --windowed --add-data "NPD.py;." --version-file=version.txt --icon=icon.ico NPD_GUI.py
