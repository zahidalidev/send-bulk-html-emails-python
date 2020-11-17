from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from pathlib import Path
from string import Template
import smtplib
import openpyxl

# 1st allow less secure app access or create app Passwords to send email
# https://support.google.com/accounts/answer/185833?hl=en
# https://myaccount.google.com/u/1/security

filename = './assets/test.xlsx'  # file having emails
xlsx = openpyxl.load_workbook(filename)
sheet = xlsx.active
data = sheet.rows

emails = []
for row in data:
    l = list(row)
    for i in range(len(l)):
        emails.append(str(l[i].value))  # single row

template = Template(Path("template.html").read_text())  # adding template

for row in emails:
    message = MIMEMultipart()
    message["from"] = "Codex"
    message["to"] = row
    message["subject"] = "Hire Virtual Assistant"
    # attaching html template
    message.attach(MIMEText(template.substitute(), "html"))

    with smtplib.SMTP(host="smtp.gmail.com", port=587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login("email", "login_password")  # sender details
        smtp.send_message(message)
        print("sent to " + row)
