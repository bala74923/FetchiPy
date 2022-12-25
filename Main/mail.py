from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename
from email.mime.application import MIMEApplication
import smtplib
def send(file_list,CONTEST_NAME,to_address):
    message=MIMEMultipart()
    message["from"]= "FetchiPy"
    # to addresss here
    message["to"]= to_address

    message["subject"]="SMTP checking.."
    # body of the message
    message.attach(MIMEText(CONTEST_NAME+" DETAILS",'plain'))
    for filename in file_list:
    #filename=r'D:\Leetcode progress(Techdose training).xlsx'
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