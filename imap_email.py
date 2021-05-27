import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import email

from setup import Setup

class Send_Mail():
    def __init__(self,server_str):
        self.setup = Setup()
        server_map = {'Gmail' : {'Server':'smtp.gmail.com','Port':587}, 'Outlook' : {'Server':'smtp-mail.outlook.com','Port':587}}
        server = server_str.get()
        host = server_map[server]['Server']
        port = server_map[server]['Port']

        mailbox = None
        while(True):
            try:
                mailbox = smtplib.SMTP(host,port)
                mailbox.starttls()
                self.SMTP_MAILBOX = mailbox
                break
            except Exception:
                pass

    def login(self,username=None,password=None):
        self.SMTP_MAILBOX.login(username,password)

    def setupHTML(self):
        html = f"""<p>{body}\n\n{self.setup.setup_signature()}</p>"""
