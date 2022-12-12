import os
from  auto_responder.hashText import hash_message,save_hash_to_file,load_hash_to_file
from auto_responder.check_inbox import MailBox
import re
from datetime import datetime, timedelta
from auto_responder.dehumanize import naturaltime



def check_if_message_is_different(msgs,folder_path):
    hashed_str = hash_message(str(msgs))
    if not os.path.exists(os.path.join(folder_path, "hash_msg")):
        save_hash_to_file(hashed_str,os.path.join(folder_path, "hash_msg"))
        return True
    o = load_hash_to_file(os.path.join(folder_path, "hash_msg"))
    if o["last_hash_messsage"] != hashed_str:
        save_hash_to_file(hashed_str,os.path.join(folder_path, "hash_msg"))
        return True
    else:
        return False

def check_report_type(msg):
    if re.search("missing", msg, re.IGNORECASE):
        return "MISSING"
    elif re.search("empid", msg, re.IGNORECASE):
        return "EMP_ID"
    else:
        return "WEEKLY"

def get_empID_from_msgbody(msgbody):
    empID_reg = re.search("[A-Z][0-9]{1,2}",   msgbody)
    if empID_reg:
        return empID_reg.group()

def get_dates_from_msgbody(msgbody):
    try:
        if msgbody:
            date_regex = re.compile(r"([0-9]{2})[.\-/]([0-9]{2})[.\-/]([0-9]{4})")
            if len(date_regex.findall(msgbody)) >= 1:
                dates_rng_list = date_regex.findall(msgbody)
                dates_rng_list =[datetime(month=int(c[0]),day=int(c[1]),year =int(c[2])) for c in dates_rng_list]
                min_date = min(dates_rng_list)
                max_date = max(dates_rng_list)
                max_date = max_date + timedelta(hours=23)
                start_date = min_date
                end_date = max_date
                return start_date, end_date
            else:
                pattern = re.escape("Testing")
                pattern = pattern + "\s[a-zA-Z.]*\s?:\s?([a-zA-Z\s\-0-9\/]+)?"
                test_stats_regex = re.compile(pattern,flags=re.IGNORECASE)
                msgbody = msgbody.strip()
                if test_stats_regex.search(msgbody):
                    msgbody = test_stats_regex.search(msgbody).group(1)
                start_date = naturaltime(msgbody.strip())
                start_date = datetime.strptime(start_date.strftime("%m/%d/%Y"),"%m/%d/%Y")
                if start_date.strftime("%m/%d/%Y") == datetime.now().strftime("%m/%d/%Y"):
                    end_date = start_date + timedelta(hours=23)
                else:
                    end_date = datetime.today() + timedelta(hours=23)
                return start_date, end_date
        else:
            return None
    except Exception as ex:
        print(ex)
        return None


if __name__ == "__main__":
    print(get_empID_from_msgbody("Testing: 11/22/2022 11/13/2022 J2"))
