import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
import os, sys
import re
import shutil
import requests
from convert_pdf import convert_pdf
from sendEmail import send_email
from helper import check_internet_connection, read_email_list, highlight_pos_rows, highlight_rows, main_dir


def split_emp_vist(df):
    emp_col = ['empID', 'empName',"Time Tested","DOB","Date Tested", 'symptom', 'typeOfTest',"result"]
    vist_col = ["visitorName","visitorDOB", "Time Tested","Date Tested", "symptom", "typeOfTest", "result"]
    emp_df = df[df["Category"]=="EMPLOYEE"][emp_col]
    emp_df.index += 1
    vist_df = df[df["Category"]=="VISITOR"][vist_col]
    vist_df.index += 1
    vist_df.rename(columns={'visitorName':'Visitor Name', 'visitorDOB':'DOB'},inplace=True)
    emp_df= emp_df.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
    emp_df = emp_df.apply(highlight_pos_rows, axis= 1)
    emp_html = emp_df.to_html()
    vist_html = vist_df.to_html()
    combine_html = "<h3>Employee Testing:</h3></br>{}<h3>Visitor Testing:</h3></br>{}".format(emp_html,vist_html)
    return combine_html


def split_emp_vist_csv(df):
    emp_col = ['empID', 'empName',"Time Tested","DOB","Date Tested", 'symptom', 'typeOfTest',"result"]
    vist_col = ["visitorName","visitorDOB", "Time Tested","Date Tested", "symptom", "typeOfTest", "result"]
    emp_df = df[df["Category"]=="EMPLOYEE"][emp_col]
    emp_df.index += 1
    vist_df = df[df["Category"]=="VISITOR"][vist_col]
    vist_df.index += 1
    vist_df.rename(columns={'visitorName':'Visitor Name', 'visitorDOB':'DOB'},inplace=True)
    emp_df= emp_df.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
    emp_df = emp_df.apply(highlight_pos_rows, axis= 1)
    return (emp_df,vist_df)


def validate_df(df):
    df.loc[df["symptom"]==False, "symptom"] = "None"
    df['DOB']= pd.to_datetime(df['DOB'])
    df['timeTested']= pd.to_datetime(df['timeTested'])
    df["Date Tested"]= df.timeTested.dt.strftime("%Y-%m-%d")
    df['DOB'] = df['DOB'].dt.strftime("%Y-%m-%d")
    df["Time Tested"] = df.timeTested.dt.strftime('%H:%M %p')
    return df


class BaseFormatter():
    def __init__(self, df, report_date = datetime.now().strftime("%Y-%m-%d")):
        self.df = df
        self.report_date = report_date
        self.heading = "<h1 style='text-align:center;'> Covid Testing Report for {} </h1><h3>Test Summary:</h3>".format(self.report_date)
        self.body = ""
        self.pdf_path = ""
    def pivot(self):
        pass

    def set_heading(self,type_of_report):
        self.heading = "<h1 style='text-align:center;'> {} Covid Testing Report for {} </h1><h3>Test Summary:</h3>".format(type_of_report,self.report_date)

    def _format_table(self,table_html):
        html = ""
        css= "<style>{}</style>".format("th, td {padding: 15px;}")
        soup = BeautifulSoup(table_html,'html.parser')
        tables = soup.find_all("table")
        for table in tables:
            table["style"] = "border-collapse: collapse;\
                                margin: 5px 0;\
                                font-size: 1em;\
                                font-family: sans-serif;\
                                min-width: 400px;\
                                width:100%;\
                                "
        ths = soup.find_all("tr")
        for tr in ths:
            if tr.parent.name == 'thead':
                tr["style"] = "background-color: #2b2d42;color: #ffffff;text-align: left;"
        html += soup.prettify()
        body = """<html>{}{}</html>""".format(css,html)
        return body

    def summary(self):
        return ""

    def detail_records(self):
        return ""
    
    def combine_html(self):
        return self.heading + self.summary() + self.detail_records() + "<br/><p> Auto-generated testing report for {} </p>".format(datetime.now().strftime("%m/%d/%Y %H:%M:%S"))

    def to_csv(self, output_path):
        pass


    def to_pdf(self,pdf_path):
        self.pdf_path = pdf_path
        html = self.combine_html()
        convert_pdf(pdf_path, html)
        os.startfile(pdf_path)
        return self

    def send_email(self, email_list_path, subject):
        try:
            for recipient in read_email_list(email_list_path):
                if self.pdf_path != "":
                    send_email(recipient,self.pdf_path,subject,self.summary())
                else:
                    pass
        except:
            if not check_internet_connection():
                from filedialog import showMessage
                showMessage("internet")

class DailyReportFormatter(BaseFormatter):
    def pivot(self):
        table = pd.pivot_table(self.df, values="timeTested",index=['Category',"result"], columns=['typeOfTest'],aggfunc= ['count'], \
                            margins = True, margins_name='Total')
        table.fillna(0, inplace=True)
        table.columns.name = ""
        return table

    def summary(self):
        if self.df.empty:
            self.body += """<h1>No Testing Today</h1>"""
            return self.body
        else:
            table = pd.pivot_table(self.df, values="timeTested",index=['Category',"result"], columns=['typeOfTest'],aggfunc= ['count'], \
                            margins = True, margins_name='Total')
            table.fillna(0, inplace=True)
            table.columns.name = ""
            pivot_table_html = table.to_html()
            pivot_table_html = self._format_table(pivot_table_html)
            self.body += pivot_table_html
            return pivot_table_html

    def detail_records(self):
        if self.df.empty:
            self.body += """<h1>No Testing Today</h1>"""
            return self.body
        else:
            self.df = validate_df(self.df)
            record_table_html =  split_emp_vist(self.df)
            positive_html = self._format_table(self.positive_test())
            record_table_html = positive_html + self._format_table(record_table_html)
            self.body += record_table_html
            return record_table_html

    def to_csv(self, output_path):
        if not self.df.empty:
            emp_df, vist_df = split_emp_vist_csv(self.df)
            summary = self.pivot()
            with pd.ExcelWriter(output_path,engine='xlsxwriter') as writer:
                summary.to_excel(writer, sheet_name= "SUMMARY")
                emp_df.to_excel(writer,sheet_name = "EMP_TESTING")
                vist_df.to_excel(writer,sheet_name = "VISITOR_TESTING")

        

    def positive_test(self):
        if self.df.empty:
            return ""
        else:
            positive_df = self.df.loc[self.df["result"].astype("str") =="P"]
            if positive_df.empty:
                return ""
            else:
                positive_df.index += 1
                emp_col = ['empID', 'empName',"Time Tested","DOB","Date Tested", 'symptom', 'typeOfTest',"result"]
                positive_df = positive_df[emp_col]
                positive_df= positive_df.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
                positive_df = positive_df.apply(highlight_pos_rows, axis= 1)
                return "</br><h3>Positive Employees:</h3></br>" + positive_df.to_html()
        
            
class SpecialReportFormatter(BaseFormatter):

    def summary(self):
        return ""
    def detail_records(self):
        if self.df.empty:
            self.body += """<h1>No Testing Records</h1>"""
            return self.body
        else:
            self.df.columns.name = ""
            table_html = self.df.to_html()
            table_html = self._format_table(table_html)
            return table_html
            
    def to_csv(self, output_path):
        if not self.df.empty:
            with pd.ExcelWriter(output_path,engine='xlsxwriter') as writer:
                self.df.to_excel(writer, sheet_name= "SUMMARY")
        
    

class WeeklyReportFormatter(BaseFormatter):
    def pivot(self):
        table = pd.pivot_table(self.df, values='timeTested',index=['Category',"TestByDate", "result"], columns=['typeOfTest'],aggfunc= ['count'], \
                            margins = True, margins_name='Total')
        table.fillna(0, inplace=True)
        table.columns.name = ""
        return table

    def summary(self):
        if self.df.empty:
            self.body += """<h1>No Testing Records found</h1>"""
            return self.body
        else:
            self.df["TestByDate"]  = pd.to_datetime(self.df['timeTested']).dt.strftime('%Y-%m-%d')
            table = pd.pivot_table(self.df, values='timeTested',index=['Category',"TestByDate", "result"], columns=['typeOfTest'],aggfunc= ['count'], \
                            margins = True, margins_name='Total')
            table.fillna(0, inplace=True)
            table.columns.name = ""
            pivot_table_html = table.to_html()
            pivot_table_html = self._format_table(pivot_table_html)
            self.body += pivot_table_html
            return pivot_table_html

    def detail_records(self):
        if self.df.empty:
            self.body += """<h1>No Testing Today</h1>"""
            return self.body
        else:
            self.df = validate_df(self.df)
            record_table_html = split_emp_vist(self.df)
            positve_html = self._format_table(self.positive_test())
            record_table_html = positve_html + self._format_table(record_table_html)
            self.body += record_table_html
            return record_table_html
    
    def to_csv(self, output_path):
        if not self.df.empty:
            emp_df, vist_df = split_emp_vist_csv(self.df)
            summary = self.pivot()
            with pd.ExcelWriter(output_path,engine='xlsxwriter') as writer:
                summary.to_excel(writer, sheet_name= "SUMMARY")
                emp_df.to_excel(writer,sheet_name = "EMP_TESTING")
                vist_df.to_excel(writer,sheet_name = "VISITOR_TESTING")
    
    def positive_test(self):
        if self.df.empty:
            return ""
        else:
            positive_df = self.df.loc[self.df["result"].astype("str") =="P"]
            if positive_df.empty:
                return ""
            else:
                positive_df.index += 1
                emp_col = ['empID', 'empName',"Time Tested","DOB","Date Tested", 'symptom', 'typeOfTest',"result"]
                positive_df = positive_df[emp_col]
                positive_df= positive_df.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
                positive_df = positive_df.apply(highlight_pos_rows, axis= 1)
                return "</br><h3>Positive Employees:</h3></br>" + positive_df.to_html()

class MissingReportFormatter(BaseFormatter):
    def __init__(self, df, report_date=datetime.now().strftime("%Y-%m-%d")):
        super().__init__(df, report_date)
        self.memo_html = ""

    def set_heading(self,type_of_report):
        self.heading = "<h1 style='text-align:center;'> {} Covid Testing Audit for {} </h1><h3>Test Summary:</h3>".format(type_of_report, self.report_date)
    
    def memo_body(self, active_testing_df):
        if self.df.empty:
            self.body += """<h1>No Testing Today</h1>"""
            return self
        elif active_testing_df is None:
            df = self.df.loc[self.df["no test"] != 0].reset_index()["empName"]
            df.index += 1
            df = df.to_frame()
            df = df[df["empName"] != "Total Tests"]
            record_table_html = df.to_html()
            record_table_html = self._format_table(record_table_html)
            self.memo_html = "The following employees missed covid tests for the period {}".format(self.report_date)  + record_table_html
            return self
        else:
            df = self.df.loc[self.df["no test"] != 0].reset_index()["empName"]
            df.index += 1
            df = df.to_frame()
            df = df[df["empName"] != "Total Tests"]
            combin_df = df.merge(active_testing_df, how="inner",on="empName")
            combin_df = combin_df.rename({"TITLE":"Title"}, axis= 1)
            memo = combin_df[["empID", "empName","Title", "MEMO"]]
            memo.index += 1
            memo_html= memo.to_html()
            self.memo_html =  "The following employees missed covid tests for the period {}".format(self.report_date)  +  self._format_table(memo_html)
            return self

    def pivot(self):
        return self.df
    
    def to_csv(self, output_path):
        with pd.ExcelWriter(output_path,engine='xlsxwriter') as writer:
            self.df.to_excel(writer, sheet_name= "Missing")

    def summary(self):
        return self.memo_html

    def detail_records(self):
        if self.df.empty:
            self.body += """<h1>No Testing Record Found</h1>"""
            return self.body
        else:
            table = self.df
            table = table.fillna(0)
            table = table.astype(int)
            table = table.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
            table = table.apply(highlight_rows, axis= 1)
            table_html = table.to_html()
            table_html = self._format_table(table_html)
            return table_html

    def combine_html(self):
        return self.heading + self.detail_records() + "<br/><p> Auto-generated testing report for {} </p>".format(datetime.now().strftime("%m/%d/%Y %H:%M:%S"))

        
class EmployeeReportFormatter(BaseFormatter):
    def __init__(self, df, report_date=datetime.now().strftime("%Y-%m-%d")):
        super().__init__(df, report_date)
        if self.df.empty:
            empName = ""
        else:
            empName = self.df['empName'].iat[0]
        self.heading = "<h1 style='text-align:center;'> {} Covid Testing Report for {}  for {} </h1><h3>Test Summary:</h3>".format("Employee", empName,self.report_date)
    
    def pivot(self):
        self.df["TestByDate"]  = pd.to_datetime(self.df['timeTested']).dt.strftime('%Y-%m-%d')
        table = pd.pivot_table(self.df, values='timeTested',index=["TestByDate", "result"], columns=['typeOfTest'],aggfunc= ['count'], \
                        margins = True, margins_name='Total')
        table.fillna(0, inplace=True)
        table.columns.name = ""
        return table

    def summary(self):
        if self.df.empty:
            self.body += """<h1>No Testing found</h1>"""
            return self.body
        else:
            self.df["TestByDate"]  = pd.to_datetime(self.df['timeTested']).dt.strftime('%Y-%m-%d')
            table = pd.pivot_table(self.df, values='timeTested',index=["TestByDate", "result"], columns=['typeOfTest'],aggfunc= ['count'], \
                            margins = True, margins_name='Total')
            table.fillna(0, inplace=True)
            table.columns.name = ""
            pivot_table_html = table.to_html()
            pivot_table_html = self._format_table(pivot_table_html)
            self.body += pivot_table_html
            return pivot_table_html
        
    def to_csv(self, output_path):
        emp_df = self.pivot()
        with pd.ExcelWriter(output_path,engine='xlsxwriter') as writer:
            emp_df.to_excel(writer, sheet_name= "Testing Records")
    
    def detail_records(self):
        return ""
        
 
if __name__ == "__main__":
  from db import Testingdb
  t = Testingdb()

  table = t.getMissingTests(datetime.strptime("10/01/2021", "%m/%d/%Y"), datetime.strptime("10/14/2021", "%m/%d/%Y"))
  rf = MissingReportFormatter(table, datetime.today().strftime("%Y%m%d"))
  rf.set_heading("Missing" )
  active_file_path = os.path.join(main_dir(),"active_testing.xlsx")
  memo_df = t.uploadActiveTesting(active_file_path)
  rf.memo_body(memo_df).summary()
  html = rf.combine_html()
 

  html = rf.combine_html()

  with open("output.html", mode="w") as f:
    f.write(html)
  # memo = t.uploadActiveTesting("active_testing.xlsx")
 
  # memo = rf.combine_memo(memo)
  # memo.to_csv("memo.csv")
#   table = t.getEmpData(datetime.strptime("11/01/2022", "%m/%d/%Y"), datetime.strptime("11/11/2022", "%m/%d/%Y"),"J2")
#   ef = EmployeeReportFormatter(table,datetime.strptime("11/01/2022", "%m/%d/%Y"))

#   empData = ef.summary()
#   ef.set_heading("Employee","11/01/2022")
#   empHtml = ef.combine_html()


