from db import Testingdb, Emaildb
from auto_responder.check_inbox import  MailBox, get_reply_email
from auto_responder.bot import check_if_message_is_different, get_dates_from_msgbody
import os, sys
from configparser import ConfigParser
import logging
import re
from helper import makeFolder
from ReportFomatter import WeeklyReportFormatter
from sendEmail import send_email
import time
import traceback

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


logging.basicConfig(filename="auto_responder.log",
                            filemode='a',
                            format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                            datefmt='%H:%M:%S',
                            level=logging.DEBUG)
logger=logging.getLogger()

folder_path = makeFolder()

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
    print(subscriber_lst)
    if userEmail in subscriber_lst[0]:
        return True
    else:
        return False




def auto_run():
    try:
        # t = Testingdb()
        subject = ""
        with MailBox(imap_username, imap_password) as mbox:
            msgs = mbox.get_lastest_email()
            From,Subject,Email_body = mbox.parse_msg()
            From = get_reply_email(From)
            if check_if_message_is_different(msgs) and is_user_subscribed(From):
                logging.debug("subscriber {}".format(From))
                start_date, end_date = get_dates_from_msgbody(Email_body)
                logging.debug("get data from {}-{}".format(start_date, end_date))
                from db import Testingdb
                tdb = Testingdb()
                df = tdb.getWeeklyStatsData(start_date, end_date)
                subject = "{}-{} Testing Report".format(start_date.strftime("%Y_%m_%d"), end_date.strftime("%Y_%m_%d"))
                day_range = "{} - {}".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
                pdf_path = get_pdf_path(folder_path, "Weekly", day_range)
                xlsx_path = get_xlsx_path(folder_path, "Weekly",day_range)
                rformater = WeeklyReportFormatter(df, day_range)
                rformater.to_pdf(pdf_path)
                rformater.to_csv(xlsx_path)
                send_email(From, pdf_path,subject,"")
                logging.debug("Email sent to {}".format(From))
            else:
                print("old email")
            # email_dict = refresh_email_list("DICT")
            # if is_email_exists(From,email_dict):
            #     old_email = is_email_exists(From,email_dict)
            #     email_db.update_subscriber_email(old_email,From)
            #     if check_if_message_is_different():
            #         is_phone_number = refresh_email_list("DICT")[From]
            #         start_date, end_date = respond_from_email_input(Email_body)
            #         subject = "Testing Stats Report for {} to {}".format(start_date.strftime("%m-%d-%Y"),end_date.strftime("%m-%d-%Y"))
            #         df = t.getWeeklyStatsData(start_date,end_date)
            #         if is_phone_number:
            #             subject += "\n\n"
            #             body = format_whatsapp_report(df) + "\n\n"
            #             sendText(subject,body,From)
            #         else:
            #             body = format_daily_stats(df)
            #             send_email(From, None, subject, body=body)
            #             print(From)
            #             logger.info("sending daily report email to {}".format(From) )
            # else:
            #     print("Not in the email subscriber list")
            #     return
    except Exception as ex:
        logging.debug(ex)
        traceback.print_exc()

if __name__ == "__main__":
    # while True:
    #     start_time = time.time()
    auto_run()
        # et = time.time()
        # elapsed_time = et - start_time
        # print('Execution time:', elapsed_time, 'seconds')