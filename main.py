from datetime import datetime, timedelta
from db import Testingdb
from helper import makeFolder, copyActiveTestingToCurr, main_dir
from ReportFomatter import EmployeeReportFormatter, MissingReportFormatter, DailyReportFormatter, WeeklyReportFormatter
from parse_input import parse_args
import os, sys
import warnings
import traceback

import logging


def get_pdf_path(folder_path,report_type, day_range):
    pdf_name = "{} {}Testing Report".format(day_range, report_type)
    pdf_path = os.path.join(folder_path,pdf_name) + ".pdf"
    return pdf_path

def get_xlsx_path(folder_path,report_type, day_range):
    xlsx_name = "{} {}Testing Report".format(day_range, report_type)
    xlsx_path = os.path.join(folder_path,xlsx_name) + ".xlsx"
    return xlsx_path


# Create a custom logger
logging.basicConfig(format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', filename='file.log', encoding='utf-8', level=logging.DEBUG)

if __name__ == "__main__":
    warnings.filterwarnings("ignore")
    email_list_path = "emailList.txt"
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    email_path = os.path.join(application_path,email_list_path)
    # email_list = read_email_list(email_path)
    args = parse_args()
    tdb = Testingdb()
    folder_path = makeFolder()
    summary = ""
    if args.start and args.start != datetime.now().strftime("%m/%d/%Y"):
        start_date = datetime.strptime(args.start,"%m/%d/%Y")
        if args.end:
            end_date = datetime.strptime(args.end,"%m/%d/%Y") + timedelta(hours=23)
        else:
            end_date = datetime.strptime(args.start,"%m/%d/%Y") + timedelta(hours=23)
        if args.missing:
            logging.info("Missing Report Ran {} - {}".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")))
            try:
                from filedialog import open_files, showMessage
                if not os.path.exists(os.path.join(main_dir(),"active_testing.xlsx")):
                    active_file_path = open_files()
                    copyActiveTestingToCurr(active_file_path)
                    showMessage("info")
                else:
                    active_file_path = os.path.join(main_dir(),"active_testing.xlsx")
                missing_df = tdb.getMissingTests(start_date, end_date)
                day_range = "{} - {}".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
                pdf_path = get_pdf_path(folder_path, "Missing", day_range)
                xlsx_path = get_xlsx_path(folder_path, "Missing",day_range)
                memo_df = tdb.uploadActiveTesting(active_file_path)
                missingft = MissingReportFormatter(missing_df, day_range)
                missingft.set_heading("Missing")
                subject = "{}-{} Missing Testing Report Audit".format(start_date.strftime("%Y_%m_%d"), end_date.strftime("%Y_%m_%d"))
                missingft.memo_body(memo_df).to_pdf(pdf_path).send_email(email_path,subject=subject)
                missingft.to_csv(xlsx_path)
                # convert_pdf(pdf_path, report_html)
            except Exception as e:
                logging.error("Missing Report No active testing error {}".format(e))
                traceback.print_exc()
                showMessage("error")
                df = tdb.getMissingTests(start_date, end_date)
                day_range = "{} - {}".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
                pdf_path = get_pdf_path(folder_path,"Missing",day_range)
                xlsx_path = get_xlsx_path(folder_path, "Missing",day_range)
                missingrf = MissingReportFormatter(df,day_range)
                missingrf.set_heading("Missing")
                subject = "{}-{} Missing Testing Report Audit".format(start_date.strftime("%Y_%m_%d"), end_date.strftime("%Y_%m_%d"))
                missingrf.memo_body(None).to_pdf(pdf_path).send_email(email_path,subject=subject)
                missingrf.to_csv(xlsx_path )
                # convert_pdf(pdf_path, report_html)
        elif args.empID:
                logging.info("Employee Report Ran {} - {}".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")))
                subject = "{}-{} Employee Testing Report".format(start_date.strftime("%Y_%m_%d"), end_date.strftime("%Y_%m_%d"))
                day_range = "{} - {}".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
                pdf_path = get_pdf_path(folder_path, "Employee", day_range)
                xlsx_path = get_xlsx_path(folder_path, "Employee " + args.empID ,day_range)
                emp_df = tdb.getEmpData(start_date, end_date,args.empID.strip())
                ef = EmployeeReportFormatter(emp_df, day_range)
                ef.to_pdf(pdf_path)
                ef.to_csv(xlsx_path)
        else:
            logging.debug("Employee Report Ran {} - {}".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")))
            subject = "{}-{} Testing Report".format(start_date.strftime("%Y_%m_%d"), end_date.strftime("%Y_%m_%d"))
            day_range = "{} - {}".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
            pdf_path = get_pdf_path(folder_path, "Weekly", day_range)
            xlsx_path = get_xlsx_path(folder_path, "Daily",day_range)
            df = tdb.getWeeklyStatsData(start_date, end_date)
            rformater = WeeklyReportFormatter(df, day_range)
            rformater.to_pdf(pdf_path).send_email(email_path, subject)
            rformater.to_csv(xlsx_path)
            # if args.csv:
            #     csv_path = os.path.join(folder_path,pdf_name)
            #     rformater.split_emp_vist_csv(csv_path)
    else:
        today_date = datetime.now().strftime("%Y-%m-%d")
        subject = "{} Testing Report".format(today_date)
        pdf_path = get_pdf_path(folder_path, "Daily",today_date)
        xlsx_path = get_xlsx_path(folder_path, "Daily",today_date)
        df = tdb.getTodayStatsData()
        rformater = DailyReportFormatter(df, today_date)
        rformater.to_pdf(pdf_path).send_email(email_path, subject)
        rformater.to_csv(xlsx_path)
