import os
from email.header import decode_header
import imaplib
import time
import uuid
import email
from configparser import ConfigParser
import sys
import  re






"""MailBox class for processing IMAP email.
(To use with Gmail: enable IMAP access in your Google account settings)
usage with GMail:
    import mailbox
    with mailbox.MailBox(gmail_username, gmail_password) as mbox:
        print mbox.get_count()
        print mbox.print_msgs()
for other IMAP servers, adjust settings as necessary.    
"""
IMAP_SERVER = "smtp.office365.com" 
IMAP_PORT = '993'
IMAP_USE_SSL = True

def get_reply_email(From):
    receiver_email = re.search(r"[a-zA-Z0-9]+(?:\.[a-zA-Z0-9]+)*@[a-zA-Z0-9]+(?:\.[a-zA-Z0-9]+)*", From, re.IGNORECASE)
    if receiver_email:
        return receiver_email.group()

class MailBox(object):
    
    def __init__(self, user, password):
        self.user = user
        self.password = password
        self.msg = None
        if IMAP_USE_SSL:
            self.imap = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        else:
            self.imap = imaplib.IMAP4(IMAP_SERVER, IMAP_PORT)

    def __enter__(self):
        self.imap.login(self.user, self.password)
        return self

    def __exit__(self, type, value, traceback):
        self.imap.close()
        self.imap.logout()

    def get_count(self):
        self.imap.select('Inbox')
        status, data = self.imap.search(None, 'ALL')
        return sum(1 for num in data[0].split())

    def fetch_message(self, num):
        self.imap.select('Inbox')
        status, data = self.imap.fetch(str(num), '(RFC822)')
        email_msg = email.message_from_bytes(data[0][1])
        self.msg = email_msg
        return email_msg

    def delete_message(self, num):
        self.imap.select('Inbox')
        self.imap.store(num, '+FLAGS', r'\Deleted')
        self.imap.expunge()

    def delete_all(self):
        self.imap.select('Inbox')
        status, data = self.imap.search(None, 'ALL')
        for num in data[0].split():
            self.imap.store(num, '+FLAGS', r'\Deleted')
        self.imap.expunge()

    def get_lastest_email(self):
        self.imap.select('Inbox')
        status, data = self.imap.search(None, 'ALL')
        messages = data[0].split()
        for num in reversed(messages[-1:]):
            status, data = self.imap.fetch(num, '(RFC822)')
            email_msg = email.message_from_bytes(data[0][1])
            self.msg = email_msg
            return email_msg

    def get_latest_email_sent_to(self, email_address, timeout=300, poll=1):
        start_time = time.time()
        while ((time.time() - start_time) < timeout):
            # It's no use continuing until we've successfully selected
            # the inbox. And if we don't select it on each iteration
            # before searching, we get intermittent failures.
            status, data = self.imap.select('Inbox')
            if status != 'OK':
                time.sleep(poll)
                continue
            status, data = self.imap.search(None, 'FROM', email_address)
            data = [d for d in data if d is not None]
            if status == 'OK' and data:
                for num in reversed(data[0].split()):
                    status, data = self.imap.fetch(num, '(RFC822)')
                    self.msg = email.message_from_bytes(data[0][1])
                    email_msg = email.message_from_bytes(data[0][1])
                    return email_msg
            time.sleep(poll)
        raise AssertionError("No email sent to '%s' found in inbox "
             "after polling for %s seconds." % (email_address, timeout))

    def delete_msgs_sent_to(self, email_address):
        self.imap.select('Inbox')
        status, data = self.imap.search(None, 'TO', email_address)
        if status == 'OK':
            for num in reversed(data[0].split()):
                status, data = self.imap.fetch(num, '(RFC822)')
                self.imap.store(num, '+FLAGS', r'\Deleted')
        self.imap.expunge()

    def parse_msg(self):
        From = self.msg.get("From")
        Subject = self.msg["Subject"]
        if self.msg.is_multipart():
            for part in self.msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                try:
                        # get the email body
                    attachment_file_name = part.get_filename()
                    if bool(attachment_file_name):
                        filePath = os.path.join(os.getcwd(), attachment_file_name)
                        print(filePath)
                        if not os.path.isfile(filePath) :
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()
                    body = part.get_payload(decode=True).decode()
                    # print text/plain emails and skip attachments
                    return (From, Subject,body)
                except:
                    pass

if __name__ == '__main__':
    # example:

    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    email_config = os.path.join(application_path,"config.ini")
    config_object = ConfigParser()
    config_object.read(email_config)
    user_obj = config_object["USEREMAIL"]
    username = user_obj["email"]
    password = user_obj["password"]
    imap_username = username
    imap_password = password
    with MailBox(imap_username, imap_password) as mbox:
        print (mbox.get_count())
        mbox.get_lastest_email()
        print(mbox.parse_msg())
        # msgs = mbox.get_latest_email_sent_to("+18482478883@tmomail.net")
        # hashed_str = hash_message(str(msgs))
        # # save_hash_to_file(hashed_str,"hash_msg")
        # o = load_hash_to_file("hash_msg")
        # if o["last_hash_messsage"] != hashed_str:
        #     print("different message")

