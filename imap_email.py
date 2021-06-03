import os
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import email
import traceback

from setup import Setup
import logging
logger = logging.getLogger(f"MailMerge.{os.path.basename(__file__)}")

class Send_Mail():
    def __init__(self,server_str,sender_email_str,receipent_str,cc_str,subject_str,body_str,data,status_str=None,progressbar=None):
        self.setup = Setup()

        self.status_str = status_str
        self.progressbar = progressbar
        
        server_map = {'Gmail' : {'Server':'smtp.gmail.com','Port':587}, 'Outlook' : {'Server':'smtp-mail.outlook.com','Port':587}}
        self.server = server = server_str.get()
        self.sender_details = sender_email_str.get()
        self.SENDER = self.sender_details.split("<")[1].split(">")[0].strip()
        self.cc_emails = cc_str.get().split(",")
        self.subject = subject_str.get()
        self.body_str = body_str.get("1.0",'end-1c')
        self.data = data
        host = server_map[server]['Server']
        port = server_map[server]['Port']

        mailbox = None
        while(True):
            try:
                self.updateUI(f"Connecting to {self.server}...",10)
                mailbox = smtplib.SMTP(host,port)
                mailbox.starttls()
                self.SMTP_MAILBOX = mailbox
                break
            except Exception:
                pass

    def updateUI(self,message=None,cnt=0):
        logger.info(message)
        if self.status_str:
            self.status_str.set(message)
            self.progressbar['value'] += cnt
        

    def login(self,password,username=None):
        if not username:
            username = self.SENDER
        self.password = password
        self.updateUI(f"Logging in user : {username}",10)
        self.SMTP_MAILBOX.login(username,password)

    def reconnect(self):
        logger.info("Reconnecting to SMTP mailbox")
        self.SMTP_MAILBOX.logout()
        self.login(self.password)

    def send_email(self,row,cnt,idx,total):
        msg = self.createMessage(row)
        text = msg.as_bytes().decode()
        rcpt = [row['Email']] + self.cc_emails
        # sending the mail
        self.updateUI(f"Sending Email to {row['Email']}......({idx} / {total})")
        self.SMTP_MAILBOX.sendmail(self.SENDER, rcpt, text)
        self.updateUI(f"Sent Email to {row['Email']}",cnt)

    def send_emails(self):
        total = len(self.data.index)
        cnt = 70 / total
        idx = 1
        try:
            for row in self.data.to_dict(orient='records'):
                try:
                    self.send_email(row,cnt,idx,total)
                except SMTPRecipientsRefused:
                    self.reconnect()
                    self.send_email(row,cnt,idx,total)
                idx += 1
            self.updateUI("Sent Emails",100)
            return True
        except Exception as e:
            traceback.print_exc
            logger.exception(e)
            return False

    def setupBody(self,row):
        name = row['Name']
        body = f"Dear {name},\n\n{self.body_str}"
        self.body = body
        return body

    def setupHTML(self):
        html = f"""<p>{self.body}\n\n{self.setup.setup_signature()}</p>"""
        self.html = html = html.replace("\n","<br/>")
        return html

    def createMessage(self,row):
        toAddr = row['Email']
        
        msg = MIMEMultipart('mixed')
        msg['From'] = self.sender_details
        msg['To'] = toAddr
        msg['Cc'] = ",".join(self.cc_emails)
        msg['Subject'] = self.subject

        body = self.setupBody(row)
        html = self.setupHTML()
        org_logo = self.setup.configData['org_logo']

        self.updateUI("Creating message")
        msgHtml = MIMEText(html, 'html')
        related = MIMEMultipart('related')
        message = MIMEMultipart('alternative')
        message.attach(MIMEText(body,'plain'))
        message.attach(msgHtml)
        related.attach(message)

        self.addOrgImage(org_logo,related)
        self.attachFile(row,msg)

        msg.attach(related)

        return msg

    def addOrgImage(self,org_logo,related):
        self.updateUI("Attaching image")
        msgImage = None
        try:
            with open(org_logo,"rb") as fp:
                msgImage = MIMEImage(fp.read(),"png")
                fp.close()
            msgImage.add_header('Content-ID', '<image1>')
            msgImage.add_header('Content-Disposition', 'inline; filename="logo.png"')
            related.attach(msgImage)
        except Exception:
            traceback.print_exc()
            pass

    def attachFile(self,row,msg):
        self.updateUI("Attaching PDF")
        filename = f"{row['Email']}.pdf"
        filepath = os.path.join(self.setup.PDF_PATH,filename)
        filename_to_attach = f"{row['Name']}.pdf"
        attachment = open(filepath, "rb")
        # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')
        # To change the payload into encoded form
        p.set_payload((attachment).read())
        # encode into base64
        encoders.encode_base64(p)
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename_to_attach)
        # attach the instance 'p' to instance 'msg'
        msg.attach(p)
