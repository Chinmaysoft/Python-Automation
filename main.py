# importing the pandas and all dependency

import pandas as p
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



# Read the data from excel sheet
data = p.read_excel(r"C:\EmailSender.xlsx")



# print(type(data))
mail_col=data.get("Email")
mail_list = list(mail_col)

# print(mail_list)

try:
    # adding ports and data
    server = sm.SMTP("smtp.gmail.com",587)
    server.starttls()
    # creating the server with user id and password

    server.login("example@email.com","example@1234")

    #sending format

    from_ = "example@email.com"
    to_ = mail_list

    # creating MIMEmultipart object with the msg

    message = MIMEMultipart("alternative")
    message['Subject']="This is just testing msg from Python"
    message['from']="example@email.com"

    # main body of the mail

    html='''
    <html>
    <head> This is just for Python testing
    </head>
    <body>
        <h1>This is my testing Python program</h1>    
    
    </body>
    
    </html>
    '''
    part2=MIMEText(html,"html")
    message.attach(part2)

    #Sending the mail to the server

    server.sendmail(from_,to_,message.as_string())
    print("Massage has been send to all the list mails.")

except Exception as e:
    print(e)