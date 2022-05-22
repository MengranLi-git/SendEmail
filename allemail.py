# -*- coding: utf-8 -*-
"""
Created on Sun May 22 20:33:24 2022

@author: MengranLi
"""

import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilename
import shutil 
import os
import email
import email.encoders
import email.mime.text
import email.mime.base
import smtplib
from email.mime.multipart import MIMEMultipart


window = tk.Tk()

def List_open_file():
    """Open a file for editing."""
    filepath = askopenfilename(
        filetypes=[("Text Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if not filepath:
        return
    List_entry["text"] = f"{filepath}"
    
def Attach_open_file():
    """Open a file for editing."""
    filepath = askopenfilename(
        filetypes=[("Text Files", "*.pdf"), ("All Files", "*.*")]
    )
    if not filepath:
        return
    Attach_entry["text"] = f"{filepath}"

def fun():
    print(2)

Accout_label = tk.Label(text="Email Accout")
Accout_entry = tk.Entry()
Pass_label = tk.Label(text="Password")
Pass_entry = tk.Entry()
List_label = tk.Label(text="List")
List_entry = tk.Label(text = "F:/python/发邮件.xlsx")
Attach_label = tk.Label(text="Attach")
Attach_entry = tk.Label(text = "F:/ZhangYue/WeChat/CV/CV_Mengran Li.pdf")
List_open = tk.Button(text="Open", command = List_open_file)
Attach_open = tk.Button(text="Open", command = Attach_open_file)



def allsend():
    my_email = Accout_entry.get()
    my_passw = Pass_entry.get()
    df = pd.read_excel(List_entry["text"])
    recipients = df.recipients.values.tolist()
    subjects = df.subject.values.tolist()
    message = '简历已在附件中，请查收。'
    file_name = Attach_entry["text"]
    def sendemail(recipient, subject, file):
        # build the message
        msg = MIMEMultipart()
        msg['From'] = my_email
        msg['To'] = recipient
        msg['Date'] = email.utils.formatdate(localtime=True)
        msg['Subject'] = subject
        msg.attach(email.mime.text.MIMEText(message))

        # build the attachment
        att = email.mime.text.MIMEText(open(file, "rb").read(), "base64", "utf-8")
        att["Content-Type"] = "application/octet-stream"
        att.add_header("Content-Disposition", "attachment", filename=("gbk", "", "%s" % os.path.basename(file)))
        msg.attach(att)

        # send the message
        srv = smtplib.SMTP('smtp.outlook.office365.com', 587)
        srv.ehlo()
        srv.starttls()
        srv.login(my_email, my_passw)
        srv.sendmail(my_email, recipients, msg.as_string())
    for i in range(0,len(recipients)):
        attachname = 'F:\\ZhangYue\\WeChat\\CV\\'+subjects[i]+'.pdf'
        shutil.copy(file_name, attachname)
        sendemail(recipients[i], subjects[i], attachname)


Send = tk.Button(text="Send", command = allsend)


Accout_label.grid(row=0, column=0, padx=10)
Accout_entry.grid(row=0, column=1, pady=10)
Pass_label.grid(row=1, column=0, padx=10)
Pass_entry.grid(row=1, column=1, pady=10)
List_label.grid(row=2, column=0, padx=10)
List_entry.grid(row=2, column=1, pady=10)
List_open.grid(row=2, column=2, sticky="ew", padx=5, pady=5)
Attach_label.grid(row=3, column=0, padx=10)
Attach_entry.grid(row=3, column=1, pady=10)
Attach_open.grid(row=3, column=2, sticky="ew", padx=5, pady=5)
Send.grid(row=4, column=1)

window.mainloop()

