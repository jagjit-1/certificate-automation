import cv2
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import Workbook, load_workbook
from password import pswd


def send_emails(email_to, filepath, filename):
    msg = MIMEMultipart()
    msg['From'] = email_from
    msg['To'] = email_to
    msg['SUbject'] = subject

    filename = 'D:\github\python automation\certificate automation\generated\PRASHI.jpg'

    attachment = open(filepath, 'rb')

    # Encode as base 64
    attachment_package = MIMEBase('application', 'octet-stream')
    attachment_package.set_payload((attachment).read())
    encoders.encode_base64(attachment_package)
    attachment_package.add_header(
        'Content-Disposition', 'attachment; filename= ' + filename)
    msg.attach(attachment_package)
    text = msg.as_string()

    TIE_server.sendmail(email_from, email_to, text)
    print("sending mail to " + email_to)


# ------------------------------------------------------
smtp_port = 587
smtp_server = 'smtp.gmail.com'
email_from = 'fakenamemailnew@gmail.com'

subject = 'Your certificate is ready'

print("connecting")
TIE_server = smtplib.SMTP(smtp_server, smtp_port)
TIE_server.starttls()
TIE_server.login(email_from, pswd)
print("connected")

# reading workbooks and generating photos

wb = load_workbook('names.xlsx')
ws = wb.active
print(tuple(ws.rows))

for i in ws.rows:
    name = i[0].value
    template = cv2.imread('template.jpg')
    if name != 'Names' and name != None and len(name) > 0:
        name = name.upper()
        (width, height), baseline = cv2.getTextSize(
            name, cv2.FONT_HERSHEY_COMPLEX, 4, 4)
        cv2.putText(template, name, (1000 - width//2, 700),
                    cv2.FONT_HERSHEY_SIMPLEX, 4, (0, 0, 0), 4, cv2.LINE_AA)
        path = f"D:\github\python automation\certificate automation\generated\{name}.jpg"
        cv2.imwrite(path, template)
        send_emails(i[1].value, path, name)


TIE_server.quit()
