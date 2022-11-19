from email.mime.base import MIMEBase
import smtplib
from email import encoders 
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import traceback
import os, sys
from configparser import ConfigParser

#Read config.ini file




def send_email(receiver, attchment_path, subject, body):
    try:
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        elif __file__:
            application_path = os.path.dirname(__file__)
        email_config = os.path.join(application_path,"config.ini")
        config_object = ConfigParser()
        config_object.read(email_config)
        user_obj = config_object["USEREMAIL"]
        email = user_obj["email"]
        password = user_obj["password"]

        # The server we use to send emails in our case it will be gmail but every email provider has a different smtp 
        # and port is also provided by the email provider.
        smtp = "smtp.office365.com" 
        port = 587
        # This will start our email server
        server = smtplib.SMTP(smtp,port)
        # Starting the server
        server.starttls()
        # Now we need to login
        server.login(email,password)

        # Now we use the MIME module to structure our message.
        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = receiver
        # Make sure you add a new line in the subject
        msg['Subject'] = subject
        # Make sure you also add new lines to your body
        # and then attach that body furthermore you can also send html content.
        
        msg.attach(MIMEText(body, 'html'))

        if attchment_path:
            with open(attchment_path,"rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            attchment_path = os.path.basename(attchment_path)
            encoders.encode_base64(part)

            attchment_path = os.path.basename(attchment_path)

            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {attchment_path}",
            )

            msg.attach(part)

        text = msg.as_string()

        server.sendmail(email,receiver,text)

        # lastly quit the server
        server.quit()
    except Exception as ex:
        traceback.print_exc()

if __name__ == "__main__":
    config_object = ConfigParser()
    config_object.read("config.ini")
    user_obj = config_object["USEREMAIL"]
    email = user_obj["email"]
    password = user_obj["password"]
    print(email, password)
    