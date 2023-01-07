import argparse
from db import Testingdb
from datetime import datetime
import logging
import os
import pyautogui
import time


def parse_args():
  """
  Parse input arguments
  """
  parser = argparse.ArgumentParser(description="Enter Visitor")
  parser.add_argument('-v', '--visitor', help='visitor input', type=str)
  args = parser.parse_args()
  return args

user_dir = os.environ['USERPROFILE']
onedrive = os.path.join(user_dir, "OneDrive", "Documents", "log")
if not os.path.exists(onedrive):
  os.makedirs(onedrive)

logging.basicConfig(format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', filename=os.path.join(onedrive,'search_visitor.log'), encoding='utf-8', level=logging.DEBUG)
if __name__ == "__main__":
  args = parse_args()
  if args.visitor:
    try:

        tdb = Testingdb()
        visitor_info = tdb.search_visitor(args.visitor.strip().upper())
        visitor_name = visitor_info[2]
        visitor_birthday = visitor_info[3]
        pyautogui.keyDown('shift')
        pyautogui.press('right')
        pyautogui.keyUp('shift')
        time.sleep(2)
        pyautogui.typewrite(visitor_name)
        time.sleep(1)
        pyautogui.press('tab')
        pyautogui.typewrite(visitor_birthday)
    except:
        logging.error("Visitor Name {} fail".format(args.visitor))
        pass