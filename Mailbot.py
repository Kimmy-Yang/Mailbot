from getpass import getpass

import pandas as pd
import smtplib
# from stdiomask import getpass
from email.message import EmailMessage
from stack import Stack
import json


#df_test = pd.read_excel(r'emails.xlsx', sheet_name= 'test')

#df_meh = pd.read_excel(r'emails.xlsx', sheet_name= 'MEHILÄINEN')


df_kal = pd.read_excel(r'emails.xlsx', sheet_name= 'kalmar') # df_kal = pd.DataFrame(data_kal, columns=['Email','Contact'])

# Data

mails = df_kal.get('Email')
name = df_kal.get('First name')
full_name = df_kal.get('Full name').values.tolist()
team = df_kal.get('Team')
Zoom_link = 'https://aalto.zoom.us/j/65995612558'

sender = "jiaqi.yang@aaltoes.com" #emails
password = "mypassword" #getpass(""mypassword")  #This is used to sign into google smtp server. You need enable lower encryption access temporarily for this account.
# count = 0
# server = smtplib.SMTP('smtp.gmail.com', 465)
# server.starttls()
# server.login(sender,password)
n=0

list=["a", "b", "c", "d"]
mail=["1", "2", "3", "4"]

# for i in range(len(df_test)):
#     df_mates = full_name[i:i + 4]
#     Teammate = df_mates.values.tolist()
#     Teammmates = Teammate.sep=", "
#     print(Teammmates)

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(sender, password)

    for i in range(len(df_kal)):
        link = Zoom_link
        receiver = mails[i]
        first_name = name[i]
        Team = team[i]


        if i %4 ==0:
            list[0:4] = full_name[i:i+4]
            mail[0:4] = mails.values.tolist()[i:i+4]

            #df_mates = full_name[i:i+4]
            #Teammates_list = df_mates.values.tolist()

        Zip = zip(list, mail)
        dic = dict(Zip)
        name_and_mail = json.dumps(dic)

        msg = EmailMessage()
        msg['subject'] = 'Welcome to Dash pre-workshop' # not done yet, email title
        msg['From'] = sender
        msg['To'] = receiver

        msg.set_content("Hey {},\n\nCongratulations! You are in Team {}, for the Kalmar challenge.You have {} in your team.\n\nAre you excited about the Dash Hackathon? But before the main event, we are organizing a pre-worship to speed you up about DASH. In this workshop, you can meet with your teammates, know more about the practicalities of the event, and much much more.\n\nTo help you learn about design thinking concepts that could help you for solving your challenge, we have onboard Håkan Mitts, a professor at Aalto University, who will talk in detail about design thinking. As a fun pre-task for his session, think of one product or service that you have recently used and would like to improve.\n\nWe highly recommend that you join us for this workshop. Below you will find the details about this pre-workshop.\n \nDate: This Sunday, 3rd October 2021 \nTime: 10:00 - 13:00\nWhere: Online\nZoom Link: {}\n\nThe Zoom link may look weird in your Aalto mail. But it is safe. Thank you and Looking forward to seeing you.\n\nBR\nJiaqi Yang\nHead of Program Dash 2021".format(first_name, Team, name_and_mail, link))
        smtp.send_message(msg)
        n += 1
        print(n)
#

#
#
# "Hey %s,\n ".format(first_name)
#                         "Congratulations! You are in %s, for Mehiläinen challenge.".format(Team)
#                         "You have %s in your team".format(Teammates)
#                         "Are you excited to the Dash Hackathon? But before the main event, we are organizing a pre-worship to speed you up about DASH. In this workshop, you can meet with your teammates,know more about practicalities for the event and much much more.\n"
#                         "To help you learn about design thinking concepts that could help you for solving your challenge, we have on board Håkan Mitts, a professor at Aalto university, who will talk in detail about design thinking. "
#                         "As a fun pre-task for his session, think of one product or service that you have recently used and would like to improve.\n"
#                         "We highly recommend that you join us for this workshop. Below you will find the details about this pre-workshop.\n"
#                         "Date: This Sunday, 3rd October 2021\n"
#                         "Time: 10:00 - 13:00\n"
#                         "Where: Online\n"
#                         "Zoom Link: https://aalto.zoom.us/j/65995612558\n"
#                         "Thank you and Looking forward to seeing you.\n"
#                         "BR\n"
#                         "Jiaqi Yang\n"
#                         "Head of Program Dash 2021"
# with open('', 'rb') as pdf:  #Attachment comment out if something
#         pdf_1 = pdf.read()
#         pdf_1_name = pdf.name
#
# with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
#     smtp.login(sender, password)
#
#     for row in emailList:
#         reciever = row[0]
#         name = row[1]
#         msg = EmailMessage()
#         msg['Subject'] = '' #Title of the email
#         msg['From'] = sender
#         msg['To'] = reciever
#         msg.add_attachment(pdf_1, maintype='application', subtype = 'octet-stream', filename=pdf_1_name)
#         msg.set_content(
#     """
#     """.format()
#
#         smtp.send_message(msg)
#         count += 1
#         print(count)
