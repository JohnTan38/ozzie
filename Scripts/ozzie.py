import streamlit as st
import pandas as pd
#import polars as pl
import win32com.client 
import numpy as np
import openpyxl, os, re
import datetime as dt
from datetime import timedelta
import pythoncom
import warnings
warnings.filterwarnings("ignore")

import smtplib, email, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import glob
from datetime import datetime

st.set_page_config('OOCL', page_icon="🏛️", layout='wide')
def title(url):
     st.markdown(f'<p style="color:#2f0d86;font-size:22px;border-radius:2%;"><br><br><br>{url}</p>', unsafe_allow_html=True)
def title_main(url):
     st.markdown(f'<h1 style="color:#230c6e;font-size:42px;border-radius:2%;"><br>{url}</h1>', unsafe_allow_html=True)

def success_df(html_str):
    html_str = f"""
        <p style='background-color:#baffc9;
        color: #313131;
        font-size: 15px;
        border-radius:5px;
        padding-left: 12px;
        padding-top: 10px;
        padding-bottom: 12px;
        line-height: 18px;
        border-color: #03396c;
        text-align: left;'>
        {html_str}</style>
        <br></p>"""
    st.markdown(html_str, unsafe_allow_html=True)

title_main('OOCL Container Status')
date_time = dt.datetime.now()
lastTwoDaysDateTime = dt.datetime.now() - dt.timedelta(days=2) #set window to last 2 days
newDay = dt.datetime.now().strftime('%Y%m%d')

pythoncom.CoInitialize()

usr_name = st.multiselect('Select your username', ['john.tan', 'vieming'], placeholder='Select 1', 
                          max_selections=2)
def user_email(usr_name):
    usr_email = usr_name[0] + '@sh-cogent.com.sg'
    return usr_email
global lastTwoDaysMessagesOOCLMNR
def lastTwoDaysMessagesOOCLMNR(subFldr):
    usr_email = user_email(usr_name)
    date_time = dt.datetime.now()
    lastTwoDaysDateTime = dt.datetime.now() - dt.timedelta(days=2)
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    #inbox = outlook.GetDefaultFolder(6)
    inbox_0 = outlook.Folders[user_email(usr_name)]
    inbox = inbox_0.Folders["Inbox"]
    sub_folder_oocl = inbox.Folders['OOCL']
    sub_folder_ooclmnr = sub_folder_oocl.Folders[subFldr]
    messages_ooclmnr = sub_folder_ooclmnr.Items
    lastTwoDaysMessagesOOCLMNR = messages_ooclmnr.Restrict("[ReceivedTime] >= '" + (lastTwoDaysDateTime.strftime('%d/%m/%Y %H:%M %p')) + "'")
    st.write(len(lastTwoDaysMessagesOOCLMNR))
    return lastTwoDaysMessagesOOCLMNR

def todaysMessagesOOCLMNR(subFldr):
    usr_email = user_email(usr_name)
    date_time = dt.datetime.now()
    lastTwoDaysDateTime = dt.datetime.now() - dt.timedelta(days=2)
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox_0 = outlook.Folders[user_email(usr_name)]
    inbox = inbox_0.Folders["Inbox"]
    sub_folder_oocl = inbox.Folders['OOCL']
    sub_folder_ooclmnr = sub_folder_oocl.Folders[subFldr]
    messages_ooclmnr = sub_folder_ooclmnr.Items
    startOfDayDateTime = dt.datetime(date_time.year, date_time.month, date_time.day)
    todaysMessagesBlanco = messages_blanco.Restrict("[ReceivedTime] >= '" + startOfDayDateTime.strftime('%d/%m/%Y %H:%M %p') + "'")
    #st.write(len(lastTwoDaysMessagesOOCLMNR))
    return todaysMessagesOOCLMNR

def user_email(usr_name):
    usr_email = usr_name[0] + '@sh-cogent.com.sg'
    return usr_email

if st.button('Confirm Username'):
    if usr_name is not None:
            usr_email = usr_name[0]+ '@sh-cogent.com.sg' #your outlook email address
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox_0 = outlook.Folders[usr_email]
            inbox = inbox_0.Folders["Inbox"]

            sub_folder_oocl = inbox.Folders['OOCL'] #initialize directory paths, sub_folder
            sub_folder_ooclmnr = sub_folder_oocl.Folders['OOCLMNR']
            sub_folder_blanco = sub_folder_oocl.Folders['BLANCO']
            sub_folder_tiara = sub_folder_oocl.Folders['REJECTED']

            messages_ooclmnr = sub_folder_ooclmnr.Items #get all messages in folder #messages = inbox.Items
            messages_blanco = sub_folder_blanco.Items
            messages_tiara = sub_folder_tiara.Items
            lastTwoDaysMessagesOOCLMNR = messages_ooclmnr.Restrict("[ReceivedTime] >= '" +lastTwoDaysDateTime.strftime('%d/%m/%Y %H:%M %p')+"'") #AND "urn:schemas:httpmail:subject"Demurrage" & "'Bill of Lading'")
            lastTwoDaysMessagesBlanco = messages_blanco.Restrict("[ReceivedTime] >= '" +lastTwoDaysDateTime.strftime('%d/%m/%Y %H:%M %p')+"'")
            lastTwoDaysMessagesTiara = messages_tiara.Restrict("[ReceivedTime] >= '" +lastTwoDaysDateTime.strftime('%d/%m/%Y %H:%M %p')+"'")

            startOfDayDateTime = dt.datetime(date_time.year, date_time.month, date_time.day)
            todaysMessagesBlanco = messages_blanco.Restrict("[ReceivedTime] >= '" + startOfDayDateTime.strftime('%d/%m/%Y %H:%M %p') + "'")
            
    else:
            st.write('Please select your username')

global df_ooclmnr
lst_ooclmnr =[]
lst_blanco =[]
lst_blanco_today =[]
lst_tiara =[]
def extract_info(str_oocl):
    container = re.search(r'Container No.\s*([^\n]*)', str_oocl)
    amt = re.findall(r'SGD\s*([\d\.]+)', str_oocl)
    remark = re.search(r'Remark:\s*([^\n]*)', str_oocl)

    container = container.group(1).strip() if container else None
    container = container.replace('-', '')
    amt = amt[4].strip() if amt else None
    remark = remark.group(1).strip() if remark else None

    return [container, amt, remark]

def extract_info_tiara(str_tiara):
    lines = str_tiara.split('\n')
    second_line = lines[1].strip()  # Extract the second line

    container = re.search(r'Container No.\s*([^\n]*)', str_tiara)
    owner_amt = re.findall(r'SGD\s*([\d\.]+)', str_tiara)
    remark = re.search(r'Remark:\s*([^\n]*)', str_tiara)

    container = container.group(1).strip() if container else None
    container = container.replace('-', '')
    owner_amt = owner_amt[4].strip() if owner_amt else None
    remark = remark.group(1).strip() if remark else None
   
    return [second_line, container, owner_amt, remark]

def send_email_oocl(df,usr_email,subj_email):
    email_receiver = usr_email
    #email_receiver = st.multiselect('Select one email', ['john.tan@sh-cogent.com.sg', 'vieming@yahoo.com'])
    email_sender = "john.tan@sh-cogent.com.sg"
    email_password = "Realmadrid8983@" #st.secrets["password"]

    body = """
            <html>
            <head>
            <title>Dear User</title>
            </head>
            <body>
            <p style="color: blue;font-size:25px;">DMS Inventory updated.</strong><br></p>

            </body>
            </html>

            """+ df.reset_index(drop=True).to_html() +"""
        
            <br>This message is computer generated. """+ datetime.now().strftime("%Y%m%d %H:%M:%S")

    mailserver = smtplib.SMTP('smtp.office365.com',587)
    mailserver.ehlo()
    mailserver.starttls()
    mailserver.login(email_sender, email_password)
       
    try:
            if email_receiver is not None:
                try:
                    rgx = r'^([^@]+)@[^@]+$'
                    matchObj = re.search(rgx, email_receiver)
                    if not matchObj is None:
                        usr = matchObj.group(1)
                    
                except:
                    pass

            msg = MIMEMultipart()
            msg['From'] = email_sender
            msg['To'] = email_receiver
            msg['Subject'] = 'OOCL Inventory Status_'+ subj_email+' '+ datetime.today().strftime("%Y%m%d %H:%M:%S")
            msg['Cc'] = 'vieming@yahoo.com'
        
            msg.attach(MIMEText(body, 'html'))
            text = msg.as_string()

            with smtplib.SMTP("smtp.office365.com", 587) as server:
                server.ehlo()
                server.starttls()
                server.login(email_sender, email_password)
                server.sendmail(email_sender, email_receiver, text)
                server.quit()
            st.success(f"Email sent to {email_receiver} 💌 🚀")
    except Exception as e:
            st.error(f"Email not sent: {e}")     

def process_df(df):
    df = df.sort_values(by=['CONTAINER'])
    df.drop_duplicates(subset=['CONTAINER'], keep='first', inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df

if st.button('Scan Email OOCLMNR'):
    lastTwoDaysMessagesOOCLMNR = lastTwoDaysMessagesOOCLMNR('OOCLMNR')
    for message in lastTwoDaysMessagesOOCLMNR:
        #if messages_SG06.Restrict("[Subject] = 'Upcoming Shipments Report SG06/07/55 by SH Cogent - As of 21 Apr 2023'"):
        if ((("repair estimate from SIN13 accepted") or ("repair estimate from SIN13 accepted")) in str(message.Subject)) and (("ooclmnr@oocl.com") in str(message.SenderEmailAddress)):
            body_oocl = str(message.body) #get body of email
            #print(extract_info(body_oocl))
            lst_ooclmnr.append(extract_info(body_oocl))
    
    df_ooclmnr = pd.DataFrame(lst_ooclmnr, columns=['CONTAINER', 'AMT', 'REMARK'])
    process_df(df_ooclmnr)
    st.write('### Email from ooclmnr@oocl.com. Data as of '+newDay)
    st.dataframe(df_ooclmnr, use_container_width=True)
    
    subj_email = 'from ooclmnr'
    send_email_oocl(df_ooclmnr, user_email(usr_name), subj_email)
    st.divider()
  
if st.button('Scan Email BLANCO'):
    lastTwoDaysMessagesBlanco = lastTwoDaysMessagesOOCLMNR('BLANCO')
    for message in lastTwoDaysMessagesBlanco:
        if ((("repair estimate from SIN13 accepted") or ("repair estimate from SIN13 rejected")) in str(message.Subject)) and (("blanco.zhang@equippool.com") in str(message.SenderEmailAddress)):
            body_blanco = str(message.body) #get body of email            
            lst_blanco.append(extract_info(body_blanco))
    
    todaysMessagesBlanco = lastTwoDaysMessagesOOCLMNR('BLANCO')
    for message in todaysMessagesBlanco:
        try:
            msg = email.message_from_string(message.body)
           
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == 'text/plain':
                        payload = part.get_payload(decode=True)
                        str_msg = payload.decode('utf-8')
                        lst_blanco_today.append(extract_info(str_msg))
            else:
                str_msg = msg.get_payload(decode=True).decode('utf-8')
                lst_blanco_today.append(extract_info(str_msg))
        except Exception as e:
            print(e)

    for message in todaysMessagesBlanco:
            if ((("repair estimate from SIN13 accepted") or ("repair estimate from SIN13 rejected")) in str(message.Subject)) and (("blanco.zhang@equippool.com") in str(message.SenderEmailAddress)):
                body_blanco = str(message.body) #get body of email
                lst_blanco_today.append(extract_info(body_blanco)) 
    
    lst_blanco_all = lst_blanco + lst_blanco_today
    df_blanco = pd.DataFrame(lst_blanco_all, columns=['CONTAINER', 'AMT', 'REMARK'])
    process_df(df_blanco)
    st.write('### Email from blanco.zhang@equippool.com. Data as of '+newDay)
    st.dataframe(df_blanco, use_container_width=True)
    
    subj_email = 'from Blanco Zhang'
    send_email_oocl(df_blanco,user_email(usr_name),subj_email)
    st.divider()

if st.button('Scan Email Tiara'):
    lastTwoDaysMessagesTiara = lastTwoDaysMessagesOOCLMNR('REJECTED')
    for message in lastTwoDaysMessagesTiara:
        if ("repair estimate from SIN13 rejected" in str(message.Subject)) and ("tiara@sh-cogent.com.sg" in str(message.SenderEmailAddress)):
            body_tiara = str(message.body) #get body of email
            lst_tiara.append(extract_info_tiara(body_tiara))
    df_tiara = pd.DataFrame(lst_tiara, columns=['REMARK_1', 'CONTAINER', 'OWNER_AMT', 'REMARK_2'])
    process_df(df_tiara)
    st.write('### Email from tiara@sh-cogent.com.sg. Data as of '+newDay)
    st.dataframe(df_tiara, use_container_width=True)
    
    subj_email = 'from Tiara'
    send_email_oocl(df_tiara,user_email(usr_name),subj_email)
    st.divider()

footer_html = """
    <div class="footer">
    <style>
        .footer {
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            background-color: #f0f2f6;
            padding: 10px 20px;
            text-align: center;
        }
        .footer a {
            color: #4a4a4a;
            text-decoration: none;
        }
        .footer a:hover {
            color: #3d3d3d;
            text-decoration: underline;
        }
    </style>
        All rights reserved @2024. Cogent Holdings IT Solutions.      
    </div>
"""
st.markdown(footer_html,unsafe_allow_html=True)
