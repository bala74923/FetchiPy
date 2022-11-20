from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename
from email.mime.application import MIMEApplication
import smtplib
message=MIMEMultipart()
message["from"]="GOWREESH A M"
message["to"]="balamurugan.kcse2020@sece.ac.in"
message["subject"]="SMTP checking.."
message.attach(MIMEText("file attchment trail",'plain'))
filename=r'D:\Leetcode progress(Techdose training).xlsx'
with open(filename,'rb') as f:
    attachment=MIMEApplication(f.read(), Name=basename(filename))
    attachment['Content-Disposition']='attachment;filename="{}"'.format(basename(filename))
    message.attach(attachment)
with smtplib.SMTP(host="smtp.gmail.com",port=587) as smtp:
    smtp.ehlo()
    smtp.starttls()
    smtp.login("manekshaw70@gmail.com","ojlg vwmn agyn wklq")
    smtp.send_message(message)
    print("sent..")