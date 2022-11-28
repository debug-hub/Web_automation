from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
################
import psycopg2
import pandas as pd
import numpy as np
####################
from email.mime.application import MIMEApplication
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

#################

path =r"C:\Program Files (x86)\chromedriver.exe"


driver= webdriver.Chrome(path)
driver.get("https://morth.nic.in/")
driver.maximize_window()
print(driver.title)
time.sleep(5)
print("windows maximixed")

road_path="/html/body/div[1]/header/div[3]/div/div/div/div/div[1]/div[1]/div/div/ul/li[4]/h2/a"
road= driver.find_element_by_xpath(road_path)
print("Road path opened...")
road.click()
time.sleep(5)

motor_path="/html/body/div[1]/div/div/div/div[2]/div[1]/ul/div/div/div/div/ul/li[7]/a"
motor=driver.find_element_by_xpath(motor_path)
print("motor_path_opened..")
motor.click()
time.sleep(5)
print("counting rows and columns...")
row_count = len(driver.find_elements_by_xpath("//*[@id='block-system-main']/div/div/div[2]/table/tbody/tr"))
col_count = len(driver.find_elements_by_xpath("//*[@id='block-system-main']/div/div/div[2]/table/tbody/tr[2]/td"))
print("rows and columns counted , extracting text....")
data=[]
for i in range(1,row_count+1):
    data1=[]

    for j in range(1,col_count+1):
        path="/html/body/div[1]/div/div/div/div[2]/div[2]/div/div/div[1]/div/div/div[2]/table/tbody/tr[{}]/td[{}]".format(i,j)
        a=driver.find_element_by_xpath(path)
        data1.append(a.text)
    data.append(tuple(data1))

print("Text extracted.")
#print(data)        
    

time.sleep(5)

driver.quit()
############################
print("creating dataframe.....")
df = pd.DataFrame(data, columns =['s_no','Subject', 'status', 'date'])
print("dataframe created.")
print(df.head(5))
print("Saving df as excel...")
df.to_excel("D:\Selenium_python\output.xlsx",index=False)
print("excel file saved.")      
###########################################
print("connecting databse.....")
conn= psycopg2.connect(database="sql_demo", user='postgres',password='albert123',host='localhost',port='5432')
cursor=conn.cursor()
print("Database connected")
query="insert into motor(s_no,subject,status,date_1) values(%s,%s,%s,%s)"
print("inserting query.....")
cursor.executemany(query,data)
print("Data inserted.")
conn.commit()
print("Database changes commited.")
print("Records inserted.")
print("No of records inserted:",cursor.rowcount)
###########################
print("email process started.....")


path = r"D:\Selenium_python"
os.chdir(path)

mail_content = '''Hello,
This is a test mail.
This mail contains an excel file consists of data scrapped from morth.nic.in.
Thank You
'''
#The mail addresses and password
sender_address = 
sender_pass = 
receiver_address = 
#Setup the MIME
message = MIMEMultipart()
message['From'] = sender_address
message['To'] = receiver_address
message['Subject'] = 'A test mail sent by Sharyansh using python.'
#The subject line
#The body and the attachments for the mail
message.attach(MIMEText(mail_content, 'plain'))
attach_file_name = 'output.xlsx'

dirs = os.listdir(path)
for files in dirs:
    if files.endswith('.xlsx'):
        attachment = open(files, 'rb')
        file_name = os.path.basename(files)
        part = MIMEApplication(attachment.read(), _subtype='xlsx')
        part.add_header('Content-Disposition', 'attachment', filename=file_name)
        message.attach(part)

#Create SMTP session for sending the mail
session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
session.starttls() #enable security
session.login(sender_address, sender_pass) #login with mail_id and password
text = message.as_string()
session.sendmail(sender_address, receiver_address, text)
session.quit()
print('Mail Sent')      
