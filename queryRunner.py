import argparse
from db import Testingdb
from datetime import datetime
import logging
import os


def parse_args():
  """
  Parse input arguments
  """
  parser = argparse.ArgumentParser(description='enter the start and end date')
  parser.add_argument('--update', action='store_true')
  parser.add_argument('-l', '--list', help='delimited list input', type=str)
  parser.add_argument('-v', '--visitor', help='visitor input', type=str)
  args = parser.parse_args()
  return args


user_dir = os.environ['USERPROFILE']
onedrive = os.path.join(user_dir, "OneDrive", "Documents", "log")
if not os.path.exists(onedrive):
  os.makedirs(onedrive)

logging.basicConfig(format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', filename=os.path.join(onedrive,'update_testing.log'), encoding='utf-8', level=logging.DEBUG)
if __name__ == "__main__":
  args = parse_args()
  if args.list and len(args.list) > 0:
    pos = [str(item)for item in args.list.split(',')]
  else:
    pos= []
  if args.visitor and len(args.visitor) > 0:
    vist_pos = [str(item)for item in args.visitor.split('|')]
  else:
    vist_pos= []
  print(vist_pos)
  logging.debug("update {}".format(vist_pos))
  tdb = Testingdb()
  todaysDate = datetime.now().strftime("%m/%d/%Y")
  tdb.updateTesting(todaysDate,pos)
  tdb.updateVisitorTesting(vist_pos,todaysDate)