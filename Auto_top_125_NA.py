import decimal
import pyodbc
import time
import csv
import datetime
import os
import sys
import subprocess
import numpy
import email
import smtplib
import shutil

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
from datetime import datetime as dt

os.chdir('C:/Users/tianxu/Documents/top125/Excel')
pyodbc.pooling = False

#def main():
Login_info = open('C:/Work/LogInMozart_ts.txt', 'r')
server_name = Login_info.readline()
server_name = server_name[:server_name.index(';')+1]
UID = Login_info.readline()
UID = UID[:UID.index(';') + 1]
PWD = Login_info.readline()
PWD = PWD[:PWD.index(';') + 1]
Login_info.close()
#today_dt = datetime.date.today()
print 'Connecting Server to determine date info at: ' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + '.'
conn = pyodbc.connect('DRIVER={Teradata};DBCNAME='+ server_name +'UID=' + UID + 'PWD=' + PWD)
curs = conn.cursor()
curs.execute('''
    SELECT
        MAX(CK_TRANS_DT)
    FROM PRS_RESTRICTED_V.MH_IM_CORE_FAM2_FACT
    WHERE CK_TRANS_DT >= CURRENT_DATE - 10
''')
end_dt1 = curs.fetchall()[0][0]

curs.execute('''
    SELECT
        MAX(TRANS_DT)
    FROM prs_ams_v.AMS_PBLSHR_ERNG
    WHERE TRANS_DT >= CURRENT_DATE - 10
''')
end_dt2 = curs.fetchall()[0][0]

curs.execute('''
    SELECT
        MAX(CLICK_DT)
    FROM PRS_AMS_V.AMS_CLICK
    WHERE CLICK_DT >= CURRENT_DATE - 10
''')
end_dt3 = curs.fetchall()[0][0]

#text_file = open("date.txt", "r")
#record_time = datetime.datetime.strptime(text_file.readlines()[0], '%Y-%m-%d').date()
current_date = datetime.now().date()
Day = timedelta(2)

print 'Table date: ' + end_dt1.strftime("%Y-%m-%d")
#print 'Record date: ' + record_time.strftime("%Y-%m-%d")
#text_file.close()
if (end_dt1 < current_date - Day and end_dt2 < current_date - Day and end_dt3 < current_date - Day):
    print 'Data is not fully ready.'
    print 'Send eMail'
    execfile('EmailSender_Traffic.py')
    conn.close()
    exit(1)
else:
    print 'Updating Excel file'
    p = subprocess.Popen('\"C:/Program Files (x86)/Microsoft Office/Office16/excel.exe\" \"C:/Users/tianxu/Documents/top125/Excel/New_driver.xlsm\"', shell=True, stdout = subprocess.PIPE)
    stdout, stderr = p.communicate()
    print stdout
    print 'Update Finished: NA'
    time.sleep(10)
    # print 'Ending time: ' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    print 'Send eMail'
    execfile('EmailSender_Traffic1_NA.py')
    conn.close()
    exit(0)