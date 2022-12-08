import pandas as pd
from datetime import datetime
import logging, argparse
from helper import main_dir, makeFolder
import os
from ReportFomatter import SpecialReportFormatter
import sys

class CustomReportBase():
    def __init__(self,df):
        self.df = df
    def pivot(self, rows, columns):
        self.df = self.df.reset_index()
        table = pd.pivot_table(self.df, index=rows, values = 'index',columns=columns,aggfunc= ['count'], fill_value="", \
                          margins = True, margins_name='Total')
        return table
    def _process(self):
        pass
    def export(self, output_path):
        pass

class TestsByDays(CustomReportBase):
    def __init__(self, df):
        super().__init__(df)
        self.output = None
    def _process(self):
        return self.pivot(["empID","empName","DOB"],["timeTested","typeOfTest","result"])
    def get_data(self):
        return self._process()
    
class TestsByDepartment(CustomReportBase):
    def _process(self):
        self.df = self.pivot(["empID","empName"],["timeTested","typeOfTest","result"])
        return self.df
    def get_data(self):
        return self._process()
    def export(self, output_path):
        table = self._process()
        table.to_excel(output_path)

class CustomPivot(CustomReportBase):
    def setColumns(self, cols):
        self.columns = cols
    def setRows(self, rows):
        self.rows = rows
    def _process(self):
        self.df = self.pivot(self.rows,self.columns)
        return self.df
    def get_data(self):
        return self._process()
    def export(self, output_path):
        table = self._process()
        table.to_excel(output_path)

def parse_args():
  """
  Parse input arguments
  """
  parser = argparse.ArgumentParser(description='enter the start and end date')
  parser.add_argument('-d', '--date', help='delimited date input', type=str)
  parser.add_argument('-r', '--report_type', help='delimited date input', type=str)
  parser.add_argument('--rows',  help='row', type=str)
  parser.add_argument('--columns',  help='column', type=str)
  args = parser.parse_args()
  return args




logging.basicConfig(format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', filename='db.log', encoding='utf-8', level=logging.DEBUG)
if __name__ == "__main__":
  args = parse_args()
  if args.date and len(args.date) > 0:
    output_dir = os.path.join(main_dir(),makeFolder())
    from db import Testingdb
    tdb = Testingdb()
    day_range = [str(item).strip() for item in args.date.split(',')]

    day_range = [datetime.strptime(c, "%m/%d/%Y") for c in day_range]
    df = tdb.getCustomDayRange(day_range)
    if args.report_type:
        if args.report_type == "EMP_BY_DAY":
            day_condense_str = "_".join([d.strftime("%Y_%m_%d") for d in day_range])
            export_path = os.path.join(output_dir, "COVID TESTING {}.xlsx".format(day_condense_str))
            td = TestsByDays(df)
            df = td.get_data()
            sFormatter = SpecialReportFormatter(df,day_condense_str)
            sFormatter.set_heading("BY EMPLOYEE")
            email_list_path = "emailList.txt"
            if getattr(sys, 'frozen', False):
                application_path = os.path.dirname(sys.executable)
            elif __file__:
                application_path = os.path.dirname(__file__)
            email_path = os.path.join(application_path,email_list_path)
            sFormatter.to_pdf(os.path.join(output_dir, "COVID TESTING {}.pdf".format(day_condense_str))).send_email(email_path, subject ="COVID TESTING for{}".format(day_condense_str) )
            sFormatter.to_csv(export_path)
        elif args.report_type == "CUSTOM":
            if args.rows and args.columns:
                rows = [str(item)for item in args.rows.split(',')]
                columns = [str(item)for item in args.columns.split(',')]
                day_condense_str = "_".join([d.strftime("%Y_%m_%d") for d in day_range])
                export_path = os.path.join(output_dir, "COVID TESTING {}.xlsx".format(day_condense_str))
                td = CustomPivot(df)
                td.setColumns(columns)
                td.setRows(rows)
                df = td.get_data()
                sFormatter = SpecialReportFormatter(df)
                sFormatter.set_heading("Custom")
                email_list_path = "emailList.txt"
                if getattr(sys, 'frozen', False):
                    application_path = os.path.dirname(sys.executable)
                elif __file__:
                    application_path = os.path.dirname(__file__)
                email_path = os.path.join(application_path,email_list_path)
                sFormatter.to_pdf(os.path.join(output_dir, "COVID TESTING {}.pdf".format(day_condense_str))).send_email(email_path, subject ="COVID TESTING for{}".format(day_condense_str) )
                sFormatter.to_csv(export_path)
        elif args.report_type == "DATA":
            day_condense_str = "_".join([d.strftime("%Y_%m_%d") for d in day_range])
            export_path = os.path.join(output_dir, "COVID TESTING {}.xlsx".format(day_condense_str))
            with pd.ExcelWriter(export_path, engine="xlsxwriter") as writer:
                df.to_excel(writer, sheet_name="data", index=False)
                os.startfile(export_path)

        

