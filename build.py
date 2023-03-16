import os
os.system('pip uninstall pathlib')
os.system('pip install pandas pyqt6 pyqt-tools numpy pyinstaller')
os.system('pyinstaller --noconsole main.py')
