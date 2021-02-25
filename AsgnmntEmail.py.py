#!/usr/bin/env python
# coding: utf-8

# In[1]:


import win32com.client as win32
import pandas as pd
import numpy as np
import os
import getpass

Login= getpass.getuser()

path = r"C:\Users\Login\Desktop"
path = path.replace('Login',Login)

os.chdir(path)

path1 = r'C:\Users\Login\Desktop\AssgnmtExcel.xlsx'
path1 = path1.replace('Login',Login)

df = pd.read_excel(path1)

df= df[df["Invoice Completed"] == 'Y']
df=pd.DataFrame(df)

Email = df['Email Address']
Name = df['Vendor Name']
Amount = df['Invoiced Amount']

outlook = win32.Dispatch('outlook.application')

for index, row in df.iterrows():
    
    Emailval= row['Email Address']
    Nameval= row['Vendor Name']
    Amountval= row['Invoiced Amount']
    html_template = '''<html> 
    <font style= font-family: Calibri; font style = font-size:10pt>Hello %s,
    <br><br>Hope you are doing great! <br><br> Your amount for the use of funny Services for last month is $ %.2f
    <br><br>For any queries related to the invoiced amount, please write to us on hello@funtime.com. <br> <br>Thanks & Regards,<br>Azam
    </html>''' %  (Nameval,Amountval)

    mail = outlook.CreateItem(0)
    mail.To = '<>%s' %Emailval
    mail.Subject = 'Testing'
    mail.HTMLBody = html_template
    # To attach a file to the email (optional):
#     attachment  = "Path to the attachment"
    # mail.Attachments.Add(attachment)
    mail.display()
    # mail.Send()




# In[52]:


for index, row in df.iterrows():
   print (row['Email Address'], row['Vendor Name'], row['Invoiced Amount'])


# In[ ]:





# In[ ]:




