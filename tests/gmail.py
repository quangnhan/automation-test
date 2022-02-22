import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

class Gmail:
    def __init__(self):
        self.fromaddr = "quangnhan145@gmail.com"
        self.password = "yhsafziemqtjjmim"
        self.msg = MIMEMultipart()

    def send_email(self, toaddr, messages):
        created_at = datetime.now().strftime("%m/%d/%Y %H:%M:%S")
        self.msg['From'] = self.fromaddr
        self.msg['To'] = toaddr
        self.msg['Subject'] = f"Automation log - {created_at}"
        self.msg.attach(MIMEText(messages, 'plain'))

        # creates SMTP session and send email
        s = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        s.login(self.fromaddr, self.password)
        text = self.msg.as_string()
        s.sendmail(self.fromaddr, toaddr.split(","), text)
        s.quit()

if __name__ == "__main__":
    toaddr = "nhan.vo@menlosecurity.com"
    messages = "Hello Nathan"
    gmail = Gmail()
    gmail.send_email(toaddr, messages)