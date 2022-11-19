import pdfkit
from db import Testingdb
import os, sys
from datetime import datetime



def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def convert_pdf(out_path, body):
    path_wkhtmltopdf = resource_path('./wkhtmltopdf.exe')
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
    pdfkit.from_string(body, out_path, configuration=config)
    
if __name__ == "__main__":
    today_date = datetime.now().strftime("%Y_%m_%d")
    convert_pdf("{} Testing Report.pdf".format(today_date))
    

