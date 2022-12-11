import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
import os, sys
import re
import shutil
import requests
from convert_pdf import convert_pdf
from sendEmail import send_email
import functools
import time


def highlight_pos_rows(x):
  if x["result"] != "N":
    return ["background-color:#e76f51;border-collapse: collapse;font-weight: bold;"] * len(x)
  else:
    return ["background-color:#ffffff;border-collapse: collapse;"] * len(x)


def highlight_rows(x):
  if x["no test"] != 0:
    return ["background-color:#e76f51"] * len(x)
  elif x["Total Tests"] < 2 and  x["Total Tests"] > 0:
    return ["background-color:#e9c46a"] * len(x)
  else:
    return ["background-color:#ffffff"] * len(x)

def copyActiveTestingToCurr(path=""):
  home_dir = os.environ['USERPROFILE']
  testingFolderPath = os.path.join(home_dir,"Desktop/test_history")
  basepath = "active_testing.xlsx"
  combin_path = os.path.join(home_dir, testingFolderPath,basepath)
  shutil.copy(path, combin_path)


def main_dir():
  home_dir = os.environ['USERPROFILE']
  testingFolderPath = os.path.join(home_dir,"Desktop/test_history")
  return  os.path.join(home_dir, testingFolderPath)

def format_whatsapp_report(df):
    table = pd.pivot_table(df, values="timeTested",index=['Category'], columns=['typeOfTest'],aggfunc="count")
    table.fillna(0, inplace=True)
    report_str = ""
    if table.empty:
      report_str = "No testing today"
    else:
      emps = {}
      emps["PCR"] =table.loc["EMPLOYEE","PCR"]
      emps["RAPID"]= table.loc["EMPLOYEE","RAPID"]
      visitor = {}
      visitor["PCR"] = table.loc["VISITOR","PCR"]
      visitor["RAPID"] = table.loc["VISITOR","RAPID"]
      report_str = "Employee Total:\n" + iterate_dict(emps) + "\nVisitor Total:\n" + iterate_dict(visitor)
    return report_str

def iterate_dict(d):
    result_str = ""
    for key in d.keys():
        result_str += key + ":" + str(d[key]) + "\n"
    return result_str

def formatMessages(df):
    message = ""
    if df.empty:
      return message
    else:
      for row in df.itertuples(index=False):
        msg_idx = list(df.columns).index("message")
        created_at_idx = list(df.columns).index("created_at")
        message += "<p><span><strong>{}</strong></span>: {}</p>".format(row[created_at_idx],row[msg_idx])
    return message

def is_email_exists(email,email_dict):
    for d in email_dict.keys():
      if d in email:
        return d
    else:
      return None

def check_internet_connection():
  timeout = 1
  try:
    requests.head("http://www.google.com/", timeout=timeout)
    return True
  except:
    return False

def makeFolder():
    home_dir = os.environ['USERPROFILE']
    testingFolderPath = os.path.join(home_dir,"Desktop/test_history")
    todayDate = datetime.now().strftime("%B_%d_%Y")
    combined_path = os.path.join(testingFolderPath,todayDate)
    os.makedirs(combined_path,exist_ok=True)
    return combined_path


def read_email_list(email_list_path):
    email_list = []
    with open(email_list_path) as f:
      raw_list = f.read()
      email_list = re.findall('\S+@\S+', raw_list)
    return email_list



def retries(num_times, delay=-1):
    def decorator_repeat(func):
        @functools.wraps(func)
        def wrapper_repeat(*args, **kwargs):
          try:
            func(*args, **kwargs)
          except:
            for _ in range(num_times):
                if delay > 0:
                    time.sleep(delay)
                func(*args, **kwargs)
        return wrapper_repeat
    return decorator_repeat


if __name__ == "__main__":
  pass
  # table = t.getMissingTests(datetime.strptime("10/01/2021", "%m/%d/%Y"), datetime.strptime("10/10/2021", "%m/%d/%Y"))
  # memo = t.uploadActiveTesting("active_testing.xlsx")
  # rf = MissingReportFormater(table, datetime.today().strftime("%Y%m%d"))
  # memo = rf.combine_memo(memo)
  # memo.to_csv("memo.csv")
  # table = t.getEmpData(datetime.strptime("11/01/2021", "%m/%d/%Y"), datetime.strptime("11/05/2022", "%m/%d/%Y"),"J1")
  # ef = EmployeReportFormatter(table)
  # empData = ef.summary()
  # with open("output.html", mode="w") as f:
  #   f.write(empData)