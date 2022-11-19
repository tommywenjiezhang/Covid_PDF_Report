import PyInstaller.__main__
import os

PyInstaller.__main__.run([
    '--name=%s' % "pdfReport",
    '--onefile',
    '--windowed',
    '--add-data=%s' % "./wkhtmltopdf.exe;.",
    "main.py"
])
