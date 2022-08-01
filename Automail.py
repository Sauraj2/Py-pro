import pandas as pd
import smtplib as s
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

data=pd.read_excel('Book1.xlsx')
emailcol = data.get("Email")
li=list(emailcol)
try:
    server=s.SMTP("smtp.gmail.com",587)
    server.starttls()
    server.login('saurabh34rajpurohit34@gmail.com','')
    f='saurabh34rajpurohit34@gmail.com'
    to_=li
    message=MIMEMultipart('alternative')
    message['Subject']="Message to myself"
    message['from']='saurabh34rajpurohit34@gmail.com'

    html='''
    <html>
    <head>
    </head>
    <body>
        <h1>This a automation project</h1>
        <p>Testing messing</p>
    </body>
    <html>
    '''
    text=MIMEText(html,'html')
    message.attach(text)
    server.sendmail(f,to_,message.as_string())
    print("Email sent")


except Exception as e:
    print(e)