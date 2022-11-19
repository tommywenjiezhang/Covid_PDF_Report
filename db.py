import pyodbc
import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import numpy as np
from nameparser import HumanName
import difflib
# from sendEmail import send_email


def fuzzy_match(x, vect):
    if len(difflib.get_close_matches(x, vect, cutoff=0.8)) >0:
        return difflib.get_close_matches(x, vect, cutoff=0.8)[0]
    else:
        return None


class Testingdb():
    def __init__(self):
        access_driver = [d for d in pyodbc.drivers() if "Access" in d]
        # print(access_driver)
        home_dir = os.environ['USERPROFILE']
        self.dbPath = os.path.join(home_dir,"Desktop\Testingdb.accdb")
        print(self.dbPath)
        connection_str = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + self.dbPath
        print(connection_str)
        self.conn = pyodbc.connect(connection_str)

    def getTodayStatsData(self):
        today_date = "#{}#".format(datetime.now().strftime("%Y-%m-%d"))
        emp_qry = "Select empList.empID, empList.empName, empList.DOB, \
                    Testing.timeTested, Testing.typeOfTest, Testing.result from empList left JOIN Testing \
                    ON empList.empID = Testing.empID where timeTested > %s order by Testing.timeTested asc" % (today_date)
        emp_df = pd.read_sql(emp_qry, self.conn)
        emp_df["Category"] = "EMPLOYEE"
        vistor_qry = "select * from visitorTesting where timeTested > %s" % (today_date)
        visitor_df = pd.read_sql(vistor_qry,self.conn)
        visitor_df["Category"] = "VISITOR"
        combined_df = pd.concat([emp_df,visitor_df])
        combined_df["symptom"].fillna("None",inplace=True)
        combined_df["result"].replace("", None, inplace=True)
        combined_df["result"].fillna("N", inplace=True)
        combined_df["timeTested"]= pd.to_datetime(combined_df["timeTested"])
        return combined_df

    def uploadActiveTesting(self,workbook_path):
        emp_qry = "SELECT empID, empName from empList"
        emp_df = pd.read_sql(emp_qry, self.conn)
        active_testing  = pd.read_excel(workbook_path, sheet_name=0, skiprows=1)
        active_testing = active_testing[["NAME","TITLE", "MEMO"]]
        active_testing = active_testing.loc[pd.notnull(active_testing.MEMO)]
        emp_df["match_key"] = emp_df['empName'].apply(lambda x: fuzzy_match(x, active_testing["NAME"]))
        combin_df = emp_df.merge(active_testing, how="inner",left_on="match_key", right_on="NAME")
        combin_df = combin_df[["empID","empName","TITLE","MEMO"]]
        return combin_df


    def get_duplicated_employee(self):
        emp_qry = "select empID, empName from empList"
        df = pd.read_sql(emp_qry,self.conn)
        df["first_name"] = df["empName"].apply(lambda x: HumanName(x).first).str.upper()
        df["last_name"] = df["empName"].apply(lambda x: HumanName(x).last).str.upper()
        df["duplicated"] = df.duplicated(subset=["first_name", "last_name"],keep=False).map({True:'Yes', False:'No'})
        df = df[df["duplicated"] == "Yes"].sort_values(by=["last_name","first_name"])
        return df


    def get_most_common_visitor(self):
        start_date = datetime.now() - timedelta(weeks=100)
        end_date = datetime.now()
        qry_start_date = "#{}#".format(start_date.strftime("%Y-%m-%d %H:%M:%S"))
        qry_end_date = "#{}#".format(end_date.strftime("%Y-%m-%d %H:%M:%S"))
        vistor_qry = "select visitorName, visitorDOB from visitorTesting"
        visitor_df = pd.read_sql(vistor_qry,self.conn)
        visitor_df["first_name"] = visitor_df["visitorName"].apply(lambda x: HumanName(x).first).str.upper()
        visitor_df["last_name"] = visitor_df["visitorName"].apply(lambda x: HumanName(x).last).str.upper()
        # visitor_df = visitor_df.drop_duplicates(subset=["first_name", "last_name"])
        visitor_df["visitorName"] = visitor_df["last_name"].astype(str) + "," + visitor_df["first_name"].astype(str)
        most_common = visitor_df["visitorName"].value_counts().to_frame(name="Freq").reset_index().rename(columns={"index":"visitorName"}).sort_values(by="Freq", ascending=False)
        most_common = most_common[most_common["Freq"] >=3]
        most_common = most_common.drop_duplicates(subset="visitorName")
        combin_df = most_common.merge(visitor_df, how="inner", on="visitorName")
        combin_df =  combin_df[["last_name","first_name","visitorName", "visitorDOB"]].drop_duplicates(subset="visitorName")
        return combin_df



    def getMissingTests(self, start_date:datetime, end_date:datetime):
        qry_start_date = "#{}#".format(start_date.strftime("%Y-%m-%d %H:%M:%S"))
        qry_end_date = "#{}#".format(end_date.strftime("%Y-%m-%d %H:%M:%S"))
        emp_qry = "Select empList.empID, empList.empName, empList.DOB,t.timeTested,t.typeOfTest, t.result from empList left join (Select Testing.empID, Testing.timeTested,Testing.typeOfTest, Testing.result FROM Testing where Testing.timeTested >= {} and Testing.timeTested <= {}) AS t ON empList.empID = t.empID order by t.timeTested asc".format(qry_start_date, qry_end_date)
        emp_df = pd.read_sql(emp_qry, self.conn)
        week_day_category = ["Sunday","Monday", "Tuesday"\
                            , "Wednesday", "Thursday",\
                                "Friday","Saturday","no test"]
        emp_df["timeTested"] = pd.to_datetime(emp_df["timeTested"])
        emp_df["Weekday"] = emp_df["timeTested"].dt.day_name()
        emp_df["timeTested"] = emp_df["timeTested"].fillna(0)
        emp_df["Weekday"] = emp_df["Weekday"].fillna("no test")
        emp_df["Weekday"] = pd.Categorical(emp_df["Weekday"], categories = week_day_category)
        table = pd.pivot_table(emp_df, values="timeTested",index=['empName'], columns=["Weekday"],aggfunc= ["count"], margins = True, margins_name='Total Tests')
        table  = table.droplevel(0, axis=1)
        table.loc[table["no test"] != 0,'Total Tests'] = 0
        return table

    def getEmpData(self, start_date:datetime, end_date:datetime, empID:str):
        qry_start_date = "#{}#".format(start_date.strftime("%Y-%m-%d %H:%M:%S"))
        qry_end_date = "#{}#".format(end_date.strftime("%Y-%m-%d %H:%M:%S"))
        emp_qry = "Select empList.empID, empList.empName, empList.DOB, Testing.timeTested, Testing.typeOfTest, Testing.result from  empList \
                  left JOIN Testing ON empList.empID = Testing.empID where empList.empID = '{}' and Testing.timeTested >= {} and Testing.timeTested <= {} order by Testing.timeTested asc".format(empID, qry_start_date, qry_end_date)
        emp_df = pd.read_sql(emp_qry, self.conn)
        emp_df["result"].replace("", None, inplace=True)
        emp_df["result"].fillna("N", inplace=True)
        emp_df["timeTested"]= pd.to_datetime(emp_df["timeTested"])
        return emp_df


    def getWeeklyStatsData(self, start_date:datetime, end_date:datetime):
        qry_start_date = "#{}#".format(start_date.strftime("%Y-%m-%d %H:%M:%S"))
        qry_end_date = "#{}#".format(end_date.strftime("%Y-%m-%d %H:%M:%S"))
        emp_qry = "Select empList.empID, empList.empName, empList.DOB, Testing.timeTested, Testing.typeOfTest, Testing.result from  empList \
                  left JOIN Testing ON empList.empID = Testing.empID where Testing.timeTested >= {} and Testing.timeTested <= {} order by Testing.timeTested asc".format(qry_start_date, qry_end_date)
        vistor_qry = "select * from visitorTesting where timeTested >= {} and timeTested <= {}".format(qry_start_date,qry_end_date)
        print(emp_qry)
        emp_df = pd.read_sql(emp_qry, self.conn)
        emp_df["Category"] = "EMPLOYEE"
        visitor_df = pd.read_sql(vistor_qry,self.conn)
        visitor_df["Category"] = "VISITOR"
        combined_df = pd.concat([emp_df,visitor_df])
        combined_df["symptom"].fillna("None",inplace=True)
        combined_df["result"].replace("", None, inplace=True)
        combined_df["result"].fillna("N", inplace=True)
        combined_df["timeTested"]= pd.to_datetime(combined_df["timeTested"])
        return combined_df

class Emaildb():
    def __init__(self):
        home_dir = os.environ['USERPROFILE']
        self.dbPath = os.path.join(home_dir,"Desktop\Emaildb.accdb")
        access_driver = [d for d in pyodbc.drivers() if "Access" in d]
        print(access_driver)
        print(self.dbPath)
        connection_str = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + self.dbPath
        print(connection_str)
        self.conn = pyodbc.connect(connection_str)

    def get_subscribers(self):
        qry = "SELECT * FROM subscribers where subscribed <> 0"
        sub_df = pd.read_sql(qry, self.conn)
        subscribers_lst = list(sub_df[["subscriber_email","is_phone_number"]].itertuples(index= False))
        return subscribers_lst

    def update_subscriber_email(self,old_email, new_email):
        qry = """Update subscribers set subscriber_email = '{}' where subscriber_email = '{}'""".format(new_email, old_email)
        cursor = self.conn.cursor()
        print(qry)
        cursor.execute(qry)
        self.conn.commit()

class MessageDB():
    def __init__(self):
        home_dir = os.environ['USERPROFILE']
        self.dbPath = os.path.join(home_dir,"Desktop\Emaildb.accdb")
        access_driver = [d for d in pyodbc.drivers() if "Access" in d]
        print(self.dbPath)
        connection_str = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + self.dbPath
        self.conn = pyodbc.connect(connection_str)
    def get_messages(self):
        today_date = "#{}#".format(datetime.now().strftime("%Y-%m-%d"))
        qry = "SELECT * FROM messages where created_at >= %s" % (today_date)
        messages = pd.read_sql(qry, self.conn)
        return messages


if __name__ == "__main__":
    # t = Testingdb()
    # df  = t.getTodayStatsData()
    s = Testingdb()
    df = s.get_duplicated_employee()
    # print(df.loc[df["no test"] != 0].reset_index()["empName"].to_frame())
    df.to_excel("duplicated_employee.xlsx", index=None)


    
    