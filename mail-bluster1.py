# -*- coding: utf-8 -*-
"""
Created on Fri Oct  4 18:52:42 2019

@author: taspo
"""
import email.message , ssl
import os
import pandas as pd
import smtplib
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.message import EmailMessage

body = ""


msg = MIMEMultipart('alternative')
msg = EmailMessage()
msg['Subject'] = ""
msg['From'] = ""
e = pd.read_excel("Emails.xlsx")
emails = e['Emails'].values
print(os.getcwd())
print(emails)
msg.set_content("")
msg.add_alternative("""\'""")
part1 = MIMEText( "plain",'utf-8')
part2 = MIMEText('',"html", 'utf-8')
encoders.encode_base64(part1)
encoders.encode_base64(part2)
msg.attach(part1)
msg.attach(part2)
part = MIMEBase("application", "octet-stream")


server = smtplib.SMTP("smtp.gmail.com",587)
server.ehlo()
server.starttls()
server.ehlo()
server.login("********@gmail.com","z------a")



for emailz in emails:
    server.sendmail("*****@gmail.com",emailz,msg.as_string())
server.quit()
