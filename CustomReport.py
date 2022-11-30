import pandas as pd
from datetime import datetime


class CustomReportBase():
    def __init__(self,df):
        self.df = df
    def pivot(self, rows, columns):
        self.df = self.df.reset_index()
        table = pd.pivot_table(self.df, index=rows, values = 'index',columns=columns,aggfunc= ['count'], \
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

    def export(self, output_path):
        table = self._process()
        table.to_excel(output_path)
class TestsByDepartment(CustomReportBase)


if __name__ == "__main__":
    from db import Testingdb
    tdb = Testingdb()
    day_range = ["11/24/2022", "11/26/2022", "11/28/2022"]
    day_range = [datetime.strptime(c, "%m/%d/%Y") for c in day_range]
    df = tdb.getCustomDayRange(day_range)
    TestsByDays(df).pivot().export("TestsByDays.xlsx")
