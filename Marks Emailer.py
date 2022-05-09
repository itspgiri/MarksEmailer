import smtplib
import ssl
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart  # New line
from email.mime.base import MIMEBase  # New line
from email import encoders  # New line
#Imports Required Packages from PIL - To manipulate image
from PIL import Image, ImageDraw, ImageFont

#We use pandas for our excel sheet
import pandas as pd

#Import the excel file that contains all the details
data = pd.read_excel("List.xlsx")

#Import 'Name' List from the imported .xlsx file
name_list = data['Name'].to_list()

#Determining the length of the list
max_no = len(name_list)


sender_email = "your email"
sender_name = "your name"
password = "your project"

receiver_emails = data['Email'].to_list()
receiver_names = data['Name'].to_list()
reciever_q1m=data['One'].to_list()
reciever_q2m=data['Two'].to_list()
reciever_q3m=data['Three'].to_list()
reciever_q4m=data['Four'].to_list()
reciever_q5m=data['Five'].to_list()
reciever_T = data['Total'].to_list()

server = smtplib.SMTP('smtp.gmail.com', 587)
context = ssl.create_default_context()
server.starttls(context=context)
server.login(sender_email, password)
print("Reached checkpoint 1")
for receiver_email, receiver_name, reciever_q1m,reciever_q2m, reciever_q3m, reciever_q4m, reciever_q5m, reciever_T in zip(receiver_emails, receiver_names, reciever_q1m, reciever_q2m, reciever_q3m, reciever_q4m, reciever_q5m, reciever_T):


    msg = MIMEMultipart()
    msg['Subject'] = 'MidTerm Score | ' + receiver_name + ' | Winter Semester 2022'
    msg['From'] = formataddr((sender_name, sender_email))
    msg['To'] = formataddr((receiver_name, receiver_email))
    string1 =str(reciever_q1m)
    string2 =str(reciever_q2m)
    string3=str(reciever_q3m)
    string4=str(reciever_q4m)
    string5=str(reciever_q5m)
    stringT = str(reciever_T)
    msg.attach(MIMEText("You're marks are attached below<br>"
                        """
                        <br>Marks for Question 1: """ + string1 +
                        """
                        <br>Marks for Question 2: """ + string2 +
                        """
                        <br>Marks for Question 3: """ + string3 +
                        """
                        <br>Marks for Question 4: """ + string4 +
                        """
                        <br>Marks for Question 5: """ + string5 +
                        """
                        <br>Total: """ + stringT + """


<br><br>Regards,

<p>Mr ABC
<br>Title
<br>Company</p>
                            """, 'html'))

    try:

        server.sendmail(sender_email, receiver_email, msg.as_string())
        print("Email sent to " + receiver_email)
    except Exception as e:
        print(f'Error!n{e}')
        break
    finally:
        print('Completed.')

server.quit()