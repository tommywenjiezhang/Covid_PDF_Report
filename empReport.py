import threading
from db import Testingdb
from datetime import datetime
import pandas as pd



if __name__ == "__main__":
    tdb = Testingdb()
    empList = tdb.getEmpList()
    start_date = datetime.strptime("11/24/2022", "%m/%d/%Y")
    end_date = datetime.strptime("11/29/2022", "%m/%d/%Y")
    weeklyTesting  = tdb.getWeeklyStatsData(start_date, end_date)
    weeklyTesting["timeTested"] = weeklyTesting["timeTested"].dt.strftime("%m/%d/%Y")
    weeklyTesting = weeklyTesting.loc[weeklyTesting["timeTested"].isin(["11/24/2022", "11/26/2022", "11/28/2022"])]
    weeklyTesting = weeklyTesting[["empID", "empName", "timeTested","typeOfTest", "result"]]
    merge_df = empList.merge(weeklyTesting, how = "left", on = ["empID", "empName"])
    merge_df = merge_df.reset_index()
    pvt_df =  pd.pivot_table(merge_df, index=["empID", "empName", "DOB", "result"], values= "index", columns=['timeTested', "typeOfTest"],aggfunc= ['count'], \
                          margins = True, margins_name='Total')
    pvt_df.reset_index().to_csv("empTestingqry.csv", index = False)
