# setup.py
import PyInstaller.__main__

PyInstaller.__main__.run([
    '--name=CONVERSÃO E CONSOLIDAÇÃO PDFs',
    '--onefile',
    '--console',
    '--icon=src/icon.ico',
    'src/app_gui.py'
])