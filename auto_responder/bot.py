import os
from  auto_responder.hashText import hash_message,save_hash_to_file,load_hash_to_file
from auto_responder.check_inbox import MailBox
import re
from datetime import datetime, timedelta
from auto_responder.dehumanize import naturaltime



def check_if_message_is_different(msgs):
    hashed_str = hash_message(str(msgs))
    if not os.path.exists("hash_msg"):
        save_hash_to_file(hashed_str,"hash_msg")
        return True
    o = load_hash_to_file("hash_msg")
    if o["last_hash_messsage"] != hashed_str:
        save_hash_to_file(hashed_str,"hash_msg")
        return True
    else:
        return False

def get_dates_from_msgbody(msgbody):
    try:
        if msgbody:
            date_regex = re.compile(r"([0-9]{2})[.\-/]([0-9]{2})[.\-/]([0-9]{4})")
            if len(date_regex.findall(msgbody)) >= 1:
                dates_rng_list = date_regex.findall(msgbody)
                dates_rng_list =[datetime(month=int(c[0]),day=int(c[1]),year =int(c[2])) for c in dates_rng_list]
                min_date = min(dates_rng_list)
                max_date = max(dates_rng_list)
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
                end_date = datetime.now() + timedelta(hours=23)
                return start_date, end_date
        else:
            return None
    except Exception as ex:
        print(ex)
        return None


if __name__ == "__main__":
    with open("email_body.txt") as f:
        data = get_dates_from_msgbody(f.read())
        print(data)
