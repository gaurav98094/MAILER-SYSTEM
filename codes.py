import socket
import os

import pandas as pd

import smtplib

# import the corresponding modules
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

df=pd.read_excel('Test1.xlsx')
listed_df=df.values



def send_mail(filename,to_email,mail_subject,mail_body,from_email="ABC IT Solutions <xxxxxxxx@gmail.com>"):  
    smtp_server = "smtp.mailtrap.io"
    login = "xxxxxxxxxxxxxx@gmail.com"
    password="xxxxxxxxxxxxxxx"
    subject = mail_subject
    sender_email =  from_email

    message = MIMEMultipart()
    message["From"] = from_email
    message["To"] = to_email
    message["Subject"] = mail_subject

    # Add body to email
    body =mail_body
    message.attach(MIMEText(body, "plain"))

    filename = filename
    
    file_send_name="Offer_Letter.pdf"
    
    with open(filename, "rb") as attachment:
        # The content type "application/octet-stream" means that a MIME attachment is a binary file
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode to base64
    encoders.encode_base64(part)

    # Add header 
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {file_send_name}",
    )


    # Add attachment to your message and convert it to string
    message.attach(part)
    text = message.as_string()

    # send your email
    server = smtplib.SMTP(host='smtp.gmail.com', port=587)
    server.ehlo()
    server.starttls()
    server.login(login, password)
    server.sendmail(sender_email, to_email, text)
    server.quit()




def mail_the_letter(pdffile,whom_to_send):
    pdfnames=[]

    for i,details in enumerate(listed_df):
        name=details[0]
        college=details[1]
        email=details[3]


        if whom_to_send=="everyone":
            message_body= "Congratulaions! "+name+", You selected for the internship with CodeGnan IT Solutions. Below attached is the Offer letter for your Internship"

            send_mail(filename=pdffile,
                    to_email=email,mail_subject="Offer Letter",mail_body=message_body)


        if whom_to_send==name:
            
            message_body= "Congratulaions! "+name+", You selected for the internship with CodeGnan IT Solutions. Below attached is the Offer letter for your Internship"

            send_mail(filename=pdffile,
                    to_email=email,mail_subject="Offer Letter",mail_body=message_body)