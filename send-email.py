import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText

import pandas as pd
import codecs

my_email = "khangcao.skill@gmail.com"
password = "yxhwusqmgzbcfqay"

with open("Attendance sheet.csv") as file:
    df = pd.read_csv(file, encoding="utf8")

for row in df.iterrows():
    file_name = f"0{row[1]['sđt']}"
    file_content = row[1]['attendance']
    name = row[1]["họ và tên"]
    class_ = row[1]["lớp"]
    email = row[1]["email"]

    attachment = f"final/{file_name}.pdf"
    print(name)

    with codecs.open("email-template.txt", 'r', encoding='utf8') as file:
        body = file.read()
        body = body.replace("[NAME]", name)
        body = body.replace("[CLASS_]", class_)

    msg = MIMEMultipart()
    msg['Subject'] = "Vé đăng ký chương trình Cận Lâm Sàng Hè 2024"
    b = MIMEText(body)
    msg.attach(b)
    msg['From'] = my_email
    msg['To'] = email

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(attachment, "rb").read())
    encoders.encode_base64(part)

    part.add_header('Content-Disposition', 'attachment', filename=attachment)

    msg.attach(part)

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(my_email, password)
        server.send_message(msg)