from datetime import date, timedelta, datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime as dt
import smtplib
import pyodbc

CONNECTION_STRING = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER=XX;DATABASE=XX;Trusted_Connection=Yes;Integrated Security=SSPI'#DB connection details 
SENDER = 'XXXX.com' #put your email here
RECEIVER =  'XXXX.com' #put your email here

#Insert the status into the Datable table
def sql_insert(date1,status1):
        conn = pyodbc.connect(CONNECTION_STRING)
        cursor = conn.cursor()
        cursor.execute('''INSERT INTO Tablename ([DataDate],[Status])
        VALUES
        (?,?)''',(date1,status1))
        conn.commit()

#Sending email
def send_email(subject: str, body: str):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = SENDER
    msg['To'] = RECEIVER
    mime_messge = MIMEText(body, 'html')
    msg.attach(mime_messge)
    with smtplib.SMTP(SMTP_SERVER) as smtp_context:
        smtp_context.send_message(msg)

#Getting the workday - day as the report is for T+1
today = date.today()
mondayCheck = today.weekday()
if mondayCheck == 0:
    today2 = date.today()+timedelta(days=-3)
else:
    today2 = date.today()+timedelta(days=-1)
today3 = today2.strftime("%d%m%Y")

# Calling the web drive and go through the report submission process
try:
    driver = webdriver.Chrome("C:/Users/dnalaka/chromedriver.exe")
    driver.implicitly_wait(5)
    driver.get("https://asxonline.com/short-sales-reporting")
    elem = driver.find_element(By.ID, "idToken1")
    elem.send_keys("Usernaem")
    elem2 = driver.find_element(By.ID, "idToken2")
    elem2.send_keys("Password")
    elem3 = driver.find_element(By.ID, "loginButton_0")
    elem3.click()
    error1 = "ASX Short Reporting - Login Error -"+today3 + \
        " Report need to be submitted manually before 9AM" #Error to send via email
    error2 = "Error" #Status to be insert into the DB
    uploadFile = driver.find_element(
        By.XPATH, "//*[@id='broker-form']/div[2]/div/label[1]")
    uploadFile.click()
    error1 = "ASX Short Reporting - Upload Error -"+today3 + \
        " Report need to be submitted manually before 9AM" #Error to send via email if failed at the upload stage
    error2 = "Error" #Status to be insert into the DB
    browser = driver.find_element(By.XPATH, "//*[@id='attachments']")
    browser.send_keys('file path'+today3+'.csv') #Upload file path
    submit = driver.find_element(By.XPATH, "//*[@id='submitBrokerForm']")
    submit.click()
    error1 = ("ASX_Short_Reporting - File submitted Sucessfull for -"+today3)  #Message to send via email if the submission is successfull
    error2 = "Sucessfull" #Status to be insert into the DB
except:
    error1 = "ASX Short Reporting - Upload Error -"+today3 + \
        " Report need to be submitted manually before 9AM" # Assigning the error message if the above fails
    error2 = "Error" # Setting message for the DB if above fails
finally:
    BODY = f"""\
    <html>
    <head>
    </head>
    <body>
        {'ASX_Short_Reporting- '+error1:}
    <br>
    
    <br>
    
    <a href="SOP">ASX Short Reporting SOP</a>

    </body>
    </html>""" # Email body 
    send_email(error1, error1) # calling the email class to send the email based on the error message 
    sql_insert((datetime.strptime(today3,"%d%m%Y")),error2) # calling the class to enter data into the DB
    driver.close() # close the web drive. 
