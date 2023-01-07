import PyInstaller.__main__
import os

# PyInstaller.__main__.run([
#     '--name=%s' % "pdfReport",
#     '--onefile',
#     '--windowed',
#     '--add-data=%s' % "./wkhtmltopdf.exe;.",
#     "main.py"
# ])

PyInstaller.__main__.run([
    '--name=%s' % "update_resident_test",
    '--onefile',
    '--windowed',
    "update_resident.py"
])

# PyInstaller.__main__.run([
#     '--name=%s' % "most_common_visitor",
#     '--onefile',
#     '--windowed',
#     "db.py"
# ])


# PyInstaller.__main__.run([
#     '--name=%s' % "custom_reports",
#     '--onefile',
#     '--windowed',
#     '--add-data=%s' % "./wkhtmltopdf.exe;.",
#     "CustomReport.py"
# ])


# PyInstaller.__main__.run([
#     '--name=%s' % "auto_email_responder",
#     '--onefile',
#     '--windowed',
#     '--add-data=%s' % "./wkhtmltopdf.exe;.",
#     "AutoReplyRunner.py"
# ])


# PyInstaller.__main__.run([
#     '--name=%s' % "search_visitor",
#     '--windowed',
#     "search_visitor.py"
# ])
