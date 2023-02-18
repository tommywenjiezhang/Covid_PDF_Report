from db import Testingdb, Emaildb
from auto_responder.check_inbox import  MailBox, get_reply_email
from auto_responder.bot import check_if_message_is_different, get_dates_from_msgbody, check_report_type, get_empID_from_msgbody
import os, sys
from configparser import ConfigParser
import logging
import re
from helper import makeFolder, retries
from ReportFomatter import EmployeeReportFormatter, MissingReportFormatter, WeeklyReportFormatter
from sendEmail import send_email
import time
import traceback
from helper import main_dir

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)
email_config = os.path.join(application_path,"config.ini")
config_object = ConfigParser()
config_object.read(email_config)
user_obj = config_object["USEREMAIL"]
imap_username = user_obj["email"]
imap_password = user_obj["password"]

folder_path = makeFolder()

logging.basicConfig(filename=os.path.join(folder_path,"auto_responder.log"),
                            filemode='a',
                            format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                            datefmt='%H:%M:%S',
                            level=logging.DEBUG)
logger=logging.getLogger()



def get_pdf_path(folder_path,report_type, day_range):
    pdf_name = "{} {}Testing Report".format(day_range, report_type)
    pdf_path = os.path.join(folder_path,pdf_name) + ".pdf"
    return pdf_path

def get_xlsx_path(folder_path,report_type, day_range):
    xlsx_name = "{} {}Testing Report".format(day_range, report_type)
    xlsx_path = os.path.join(folder_path,xlsx_name) + ".xlsx"
    return xlsx_path


def is_user_subscribed(userEmail):
    email_db = Emaildb("Emaildb.accdb")
    subscriber_lst = email_db.get_subscribers()
    if userEmail in subscriber_lst[0]:
        return True
    else:
        return False



@retries(num_times=2,delay=3)
def auto_run():
    try:
        # t = Testingdb()
        subject = ""
        with MailBox(imap_username, imap_password) as mbox:
            msgs = mbox.get_lastest_email()
            From,Subject,Email_body = mbox.parse_msg()
            From = get_reply_email(From)
            if check_if_message_is_different(msgs, folder_path) and is_user_subscribed(From):
                logging.debug("subscriber {}".format(From))
                start_date, end_date = get_dates_from_msgbody(Email_body)
                logging.debug("get data from {}-{}".format(start_date, end_date))
                from db import Testingdb
                tdb = Testingdb()
                report_type = check_report_type(Email_body)
                if report_type == "WEEKLY":
                    df = tdb.getWeeklyStatsData(start_date, end_date)
                    subject = "{}-{} Testing Report".format(start_date.strftime("%Y_%m_%d"), end_date.strftime("%Y_%m_%d"))
                    day_range = "{} - {}".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
                    pdf_path = get_pdf_path(folder_path, "Weekly", day_range)
                    xlsx_path = get_xlsx_path(folder_path, "Weekly",day_range)
                    rformater = WeeklyReportFormatter(df, day_range)
                    rformater.to_pdf(pdf_path)
                    rformater.to_csv(xlsx_path)
                    send_email(From, pdf_path,subject,"")
                    mbox.delete_message(msgs)
                    logging.debug("Weekly testing Email sent to {}".format(From))
                elif report_type == "MISSING":
                    if os.path.exists(os.path.join(main_dir(),"active_testing.xlsx")):
                        active_file_path = os.path.join(main_dir(),"active_testing.xlsx")
                        missing_df = tdb.getMissingTests(start_date, end_date)
                        day_range = "{} - {}".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
                        pdf_path = get_pdf_path(folder_path, "Missing", day_range)
                        xlsx_path = get_xlsx_path(folder_path, "Missing",day_range)
                        memo_df = tdb.uploadActiveTesting(active_file_path)
                        missingft = MissingReportFormatter(missing_df, day_range)
                        missingft.set_heading("Missing")
                        subject = "{}-{} Missing Testing Report Audit".format(start_date.strftime("%Y_%m_%d"), end_date.strftime("%Y_%m_%d"))
                        missingft.memo_body(memo_df).to_pdf(pdf_path)
                        send_email(From, pdf_path,subject,"")
                        logging.debug("MISSING testing Email sent to {}".format(From))
                    else:
                        logging.debug("ERROR APPENDING THE Active Testing")
                elif report_type == "EMP_ID":
                    logging.info("Employee Report Ran {} - {}".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")))
                    subject = "{}-{} Employee Testing Report".format(start_date.strftime("%Y_%m_%d"), end_date.strftime("%Y_%m_%d"))
                    day_range = "{} - {}".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
                    pdf_path = get_pdf_path(folder_path, "Employee", day_range)
                    emp_id =  get_empID_from_msgbody(Email_body)
                    print(emp_id)
                    xlsx_path = get_xlsx_path(folder_path, "Employee " + emp_id ,day_range)
                    emp_df = tdb.getEmpData(start_date, end_date, emp_id.strip())
                    ef = EmployeeReportFormatter(emp_df, day_range)
                    ef.to_pdf(pdf_path)
                    send_email(From, pdf_path,subject,"")
                    logging.debug("Employee testing Email sent to {}".format(From))
            else:
                logging.info("email has been checked")
    except Exception as ex:
        logging.debug(ex)
        raise ex

if __name__ == "__main__":
    auto_run()