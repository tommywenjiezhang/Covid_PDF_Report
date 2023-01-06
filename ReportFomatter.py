import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
import os, sys
import re
import time
import shutil
import requests
from convert_pdf import convert_pdf
from sendEmail import send_email
from helper import check_internet_connection, read_email_list, highlight_pos_rows, highlight_rows, main_dir, retries
from wifi import connect_to_wifi
import logging
from load_html import load_context_to_template

def split_emp_vist(df):
    emp_col = ['empID', 'empName',"Time Tested","DOB","Date Tested", 'symptom', 'typeOfTest',"result"]
    vist_col = ["visitorName","visitorDOB", "Time Tested","Date Tested", "symptom", "typeOfTest", "result"]
    emp_df = df[df["Category"]=="EMPLOYEE"][emp_col]
    emp_df.index += 1
    vist_df = df[df["Category"]=="VISITOR"][vist_col]
    vist_df.index += 1
    vist_df.rename(columns={'visitorName':'Visitor Name', 'visitorDOB':'DOB'},inplace=True)
    emp_df= emp_df.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
    vist_df= vist_df.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
    emp_df = emp_df.apply(highlight_pos_rows, axis= 1)
    vist_df = vist_df.apply(highlight_pos_rows, axis= 1)
    emp_html = emp_df.to_html()
    vist_html = vist_df.to_html()
    combine_html = "<h3>Employee Testing:</h3></br>{}<h3>Visitor Testing:</h3></br>{}".format(emp_html,vist_html)
    return combine_html

def split_resident_by_wing(df):
    output =[]
    cols = ["residentName","DOB","Time Tested", "Date Tested", "roomNum",\
                 "symptom", "typeOfTest","result",	"lotNumber","expirationDate"]
    wings = df.wings.value_counts().index
    for wing in wings:
        df_bywing = df[df.wings == wing]
        df_bywing = df_bywing[cols]
        df_bywing = df_bywing.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
        df_bywing = df_bywing.apply(highlight_pos_rows, axis= 1)
        wing_html = df_bywing.to_html()
        combine_html = "<h3>{} Testing:</h3></br>".format(wing) + wing_html
        output.append(combine_html)
    return "</br>".join(output)



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
    df['visitorDOB'] = pd.to_datetime(df['visitorDOB'] )
    df['timeTested']= pd.to_datetime(df['timeTested'])
    df["Date Tested"]= df.timeTested.dt.strftime("%Y-%m-%d")
    df['DOB'] = df['DOB'].dt.strftime("%Y-%m-%d")
    df['visitorDOB'] = df['visitorDOB'].dt.strftime("%Y-%m-%d")
    df["Time Tested"] = df.timeTested.dt.strftime('%H:%M %p')
    return df


def validate_resident_df(df):
    df.loc[df["symptom"]==False, "symptom"] = "None"
    df['DOB']= pd.to_datetime(df['DOB'])
    df['timeTested']= pd.to_datetime(df['timeTested'])
    df["expirationDate"] = pd.to_datetime(df["expirationDate"])
    df["Date Tested"]= df.timeTested.dt.strftime("%Y-%m-%d")
    df['DOB'] = df['DOB'].dt.strftime("%Y-%m-%d")
    df["expirationDate"] = df["expirationDate"].dt.strftime("%Y-%m-%d")
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
        css= "<style>{}</style>".format("th, td {padding: 10px;}")
        soup = BeautifulSoup(table_html,'html.parser')
        tables = soup.find_all("table")
        for table in tables:
            table["style"] = "border-collapse: collapse;\
                                margin: 5px 0;\
                                font-size: 12px;\
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
    
    @retries(num_times=2)
    def send_email(self, email_list_path, subject):
        try:
            for recipient in read_email_list(email_list_path):
                if self.pdf_path != "":
                    send_email(recipient,self.pdf_path,subject,self.summary())
                else:
                    pass
        except Exception as ex:
            logging.debug("Email was not send {}".format(ex))
            connect_to_wifi()
            if not check_internet_connection():
                from filedialog import showMessage
                showMessage("internet")
            raise ex
            

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
            table = table.droplevel(0, axis=1)
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
                postive_emp_str  = ""
                postive_vist_str = ""
                emp_col = ['empID', 'empName',"Time Tested","DOB","Date Tested", 'symptom', 'typeOfTest',"result"]
                vist_col = ["visitorName","visitorDOB", "Time Tested","Date Tested", "symptom", "typeOfTest", "result"]
                positive_df.index += 1
                emp_positive_df = positive_df[positive_df['Category'].astype("str") == "EMPLOYEE"]
                visitor_positive_df = positive_df[positive_df['Category'].astype("str") == "VISITOR"]
                if not emp_positive_df.empty:
                    emp_positive_df = emp_positive_df[emp_col]
                    emp_positive_df= emp_positive_df.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
                    emp_positive_df = emp_positive_df.apply(highlight_pos_rows, axis= 1)
                    postive_emp_str =  "<h3>Positive Employees:</h3></br>"  + emp_positive_df.to_html()
                else:
                    postive_emp_str = ""
                if not visitor_positive_df.empty:
                    visitor_positive_df  = visitor_positive_df[vist_col]
                    visitor_positive_df = visitor_positive_df.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
                    visitor_positive_df = visitor_positive_df.apply(highlight_pos_rows, axis= 1)
                    postive_vist_str =  "<h3>Positive Visitor:</h3></br>"  + visitor_positive_df.to_html()
                else:
                    postive_vist_str = ""
                return "</br>{}{}".format(postive_emp_str,postive_vist_str)
        
            
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
            table = table.droplevel(0, axis=1)
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
                postive_emp_str  = ""
                postive_vist_str = ""
                emp_col = ['empID', 'empName',"Time Tested","DOB","Date Tested", 'symptom', 'typeOfTest',"result"]
                vist_col = ["visitorName","visitorDOB", "Time Tested","Date Tested", "symptom", "typeOfTest", "result"]
                positive_df.index += 1
                emp_positive_df = positive_df[positive_df['Category'].astype("str") == "EMPLOYEE"]
                visitor_positive_df = positive_df[positive_df['Category'].astype("str") == "VISITOR"]
                if not emp_positive_df.empty:
                    emp_positive_df = emp_positive_df[emp_col]
                    emp_positive_df= emp_positive_df.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
                    emp_positive_df = emp_positive_df.apply(highlight_pos_rows, axis= 1)
                    postive_emp_str =  "<h3>Positive Employees:</h3></br>"  + emp_positive_df.to_html()
                else:
                    postive_emp_str = ""
                if not visitor_positive_df.empty:
                    visitor_positive_df  = visitor_positive_df[vist_col]
                    visitor_positive_df = visitor_positive_df.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
                    visitor_positive_df = visitor_positive_df.apply(highlight_pos_rows, axis= 1)
                    postive_vist_str =  "<h3>Positive Visitor:</h3></br>"  + visitor_positive_df.to_html()
                else:
                    postive_vist_str = ""
                return "</br>{}{}".format(postive_emp_str,postive_vist_str)
        

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
            table = table.droplevel(0, axis=1)
            table.index.names = ["Test Date", "Result"]
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

class ResidentFormatter(BaseFormatter):
    def summary(self):
        if self.df.empty:
            self.body += """<h1>No Testing Records found</h1>"""
            return self.body
        else:
            self.df["TestByDate"]  = pd.to_datetime(self.df['timeTested']).dt.strftime('%Y-%m-%d')
            table = pd.pivot_table(self.df, values='timeTested',index=['wings',"TestByDate", "result"], columns=['typeOfTest'],aggfunc= ['count'], \
                            margins = True, margins_name='Total')
            table.fillna(0, inplace=True)
            table = table.droplevel(0, axis=1)
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
            self.df = validate_resident_df(self.df)
            record_table_html = split_resident_by_wing(df)
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
                positive_df= positive_df.style.set_table_styles([{'selector': 'tr,td', 'props': [('font-size', '12pt'),('border-style','solid'),('border-width','1px')]}])
                positive_df = positive_df.apply(highlight_pos_rows, axis= 1)
                postive_resident_str =  "<h3>Positive Resident:</h3></br>"  + positive_df.to_html()
                return "</br>{}".format(postive_resident_str)
    
    def batch_html(self, pdf_path):
        if self.df.empty:
            return
        else:
            template_df = self.df.rename(columns={"residentName":"name","TestByDate":"testDate"})
            residents = template_df.to_dict(orient="records")
            if not os.path.exists(pdf_path):
                os.makedirs(pdf_path)
            for resident in residents:
                name = resident.get("name")
                testDate = resident.get("testDate")
                filename = name + "test_export" + testDate + ".pdf"
                fullpath = os.path.join(pdf_path, filename)
                context = {"resident":resident, "whitespaces": " "*20}
                html = load_context_to_template(context,"ResidentTestingTemplate.html")
                convert_pdf(fullpath, html)
            os.startfile(pdf_path)


 
if __name__ == "__main__":
    from db import ResidentDB
    rdb = ResidentDB("ResidentDb.accdb")
    start_date = datetime.strptime("12/01/2022", "%m/%d/%Y")
    end_date = datetime.strptime("01/06/2023", "%m/%d/%Y")
    df = rdb.getWeeklyResidentTesting(start_date, end_date)
    formattter = ResidentFormatter(df)
    formattter.to_pdf("Resident Testing.pdf")
    formattter.batch_html("Resident_Testing_Batch")
  # memo = t.uploadActiveTesting("active_testing.xlsx")
 
  # memo = rf.combine_memo(memo)
  # memo.to_csv("memo.csv")
#   table = t.getEmpData(datetime.strptime("11/01/2022", "%m/%d/%Y"), datetime.strptime("11/11/2022", "%m/%d/%Y"),"J2")
#   ef = EmployeeReportFormatter(table,datetime.strptime("11/01/2022", "%m/%d/%Y"))

#   empData = ef.summary()
#   ef.set_heading("Employee","11/01/2022")
#   empHtml = ef.combine_html()


