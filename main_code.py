import datetime
import pandas as pd
import win32com.client
from bs4 import BeautifulSoup
import pyodbc
import numpy as np

conn = pyodbc.connect(r'Driver={SQL Server};'
                      r'Server=SERVERNAME;'
                      r'Database=DB_NAME;'
                      r'UID=USERNAME;'
                      r'PWD=PASSWORD'
                      )

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6).Folders['Main']

messages = inbox.items

current_time = datetime.datetime.now()

for message in messages:
    if message.SenderEmailAddress == 'EMAILADDRESS':
        msg_date = message.SentOn.strftime("%d-%m-%y")
        yesterday = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%d-%m-%y")
        if msg_date == yesterday:
            words = message.body.split()
            for word in words:
                if word.startswith("Backup") and words[words.index(word) + 1] == "job:":
                    try:
                        soup = BeautifulSoup(message.HTMLbody, 'lxml')
                        table = soup.find_all('table')
                        data = pd.read_html(str(table), skiprows=5, header=0)
                        df = (data[5:11])[0]
                        df = df.loc[:, ["Name", "Status", "Start time", "End time", "Read", "Transferred", "Details"]]
                        df = df.rename(
                            columns={"Name": "Server_Name", "Status": "Current_Status", "Start time": "Start_time",
                                     "End time": "End_time", "Read": "Storage_Found",
                                     "Transferred": "Storage_Transferred"})
                        df['Date_Sent'] = message.SentOn.strftime("%m-%d-%y")
                        df['Date_imported'] = current_time
                        df = df.replace(np.nan, '', regex=True)
                        insert_into = f"INSERT INTO TABLENAME (Server_Name, Server_Status, Start_time, " \
                                      f"End_time, Storage_Found, Storage_Transferred, Details, Date_Sent, " \
                                      f"Date_imported) " \
                                      f"VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)"
                        cursor = conn.cursor()
                        cursor.fast_executemany = True
                        cursor.executemany(insert_into, df.values.tolist())
                        conn.commit()
                        break
                    finally:
                        break
