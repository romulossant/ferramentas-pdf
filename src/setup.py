# setup.py
import PyInstaller.__main__

PyInstaller.__main__.run([
    '--name=PDFs',
    '--onefile',
    '--windowed',
    '--icon=src/icon.ico',
    '--add-data=src/icon.ico;src',
    'src/app_gui.py'
])