import PyInstaller.__main__
import os

# PyInstaller.__main__.run([
#     '--name=%s' % "pdfReport",
#     '--onefile',
#     '--windowed',
#     '--add-data=%s' % "./wkhtmltopdf.exe;.",
#     "main.py"
# ])

# PyInstaller.__main__.run([
#     '--name=%s' % "update_test",
#     '--onefile',
#     '--windowed',
#     "queryRunner.py"
# ])

# PyInstaller.__main__.run([
#     '--name=%s' % "most_common_visitor",
#     '--onefile',
#     '--windowed',
#     "db.py"
# ])


PyInstaller.__main__.run([
    '--name=%s' % "custom_reports",
    '--onefile',
    '--windowed',
    '--add-data=%s' % "./wkhtmltopdf.exe;.",
    "CustomReport.py"
])
