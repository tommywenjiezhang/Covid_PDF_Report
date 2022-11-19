import argparse
def parse_args():
  """
  Parse input arguments
  """
  parser = argparse.ArgumentParser(description='enter the start and end date')
  parser.add_argument("--start", help="start date", type=str)
  parser.add_argument("--end", help="end date", type=str)
  parser.add_argument("--empID", help="end date", type=str)
  parser.add_argument('-csv', action='store_true')
  parser.add_argument('--missing', action='store_true')
  args = parser.parse_args()
  return args