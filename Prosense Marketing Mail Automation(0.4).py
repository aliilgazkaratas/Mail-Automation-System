#!/usr/bin/env python
# coding: utf-8

# In[2]:


#pip install openpyxl


# In[36]:



import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
data = pd.read_excel (r'C:\Users\ilgaz\OneDrive\Masaüstü\Prosense Marketing Mail Automation\marketingautomation.xlsx')
Mails = data['Mails'].tolist()
Names = data['Name'].tolist()
Subjects = data["Subjects"].tolist()
TimeRange = data["TimeRange"].tolist()
Sender = data["Sender"].tolist()
Password = data["Password"].tolist()
"""msgdraft= data['Draft'].tolist()"""
print(Mails,Names,Subjects,TimeRange,Sender,Password)


# In[37]:


# modules
import time
import smtplib
from email.message import EmailMessage
import random
import codecs
import ssl
 
file1 = codecs.open(r"C:\Users\ilgaz\OneDrive\Masaüstü\Prosense Marketing Mail Automation\Drafts1.html", "r", "utf-8")
file2= codecs.open(r"C:\Users\ilgaz\OneDrive\Masaüstü\Prosense Marketing Mail Automation\Drafts2.html", "r", "utf-8")
file3 = codecs.open(r"C:\Users\ilgaz\OneDrive\Masaüstü\Prosense Marketing Mail Automation\Drafts3.html", "r", "utf-8")
file4 = codecs.open(r"C:\Users\ilgaz\OneDrive\Masaüstü\Prosense Marketing Mail Automation\Drafts4.html", "r", "utf-8")
list_of_drafts=[]
list_of_drafts.append(file1.read())
list_of_drafts.append(file2.read())
list_of_drafts.append(file3.read())
list_of_drafts.append(file4.read())
x=0

#adress=["""<html><body><p><font size="4" color="black" face="Arial">Hi</font>
#</p>""","""<html><body><p><font size="4" color="black" face="Arial">Hello</font>
#</p>""","""<html><body><p><font size="4" color="black" face="Arial">Dear</font>
#</p>"""] 
#<html><body><p><font size="4" color="black" face="Arial">Hi ,</font></p> (alt satırdaydı)


# In[38]:



# content
sender = Sender[0]
while x!=len(Mails):
    
    receiver = Mails[x]
    password = Password[0]
    Names[x]="""<html><body><font size="4">{}</body></html>""".format(Names[x])
    msg_body = ("{} {} ").format(Names[x],list_of_drafts[x%len(list_of_drafts)])
    
# action
    msg = EmailMessage()
    msg['subject'] = random.choice(Subjects)
    x=x+1
    msg['from'] = sender
    msg['to'] = receiver
    msg.add_header('Content-Type','text/html')
    msg.set_content(msg_body,subtype='html')
    list_of_drafts="""{}""".format(list_of_drafts[x%len(list_of_drafts)])
    #with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:    Bu gmail için
    Random_number_for_sleep=random.randint(TimeRange[0],TimeRange[1])
    time.sleep(Random_number_for_sleep)

    context = ssl.SSLContext(ssl.PROTOCOL_TLS)
    with smtplib.SMTP('smtp.office365.com', 587) as smtp:  #Bu outlook için
        smtp.ehlo()
        smtp.starttls(context=context)
        smtp.ehlo()# Gmail için yapınca hata verdi
        smtp.login(sender,password)
        #smtp_server.sendmail(sender, receiver, msg_body)   Gmailde gerek oldu
        #   smtp-mail.outlook.com
        try:
            smtp.send_message(msg)
            print("Sent!")
        except:
            break
            print("There is a problem")
            
        smtp.close()
        


# In[ ]:




