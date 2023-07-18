#!/usr/bin/python3

# Student name and No.: Olano James Daniel Cubillas 3035777975
# Development platform: Visual Studio Code (Windows 10)
# Python version: Python 3.8.10

from tkinter import *
from tkinter import ttk
from tkinter import font
from tkinter import messagebox
from tkinter import filedialog
import re
import os
import pathlib
import sys
import base64
import socket

#
# Global variables
#

# Replace this variable with your CS email address
YOUREMAIL = "jdcolano@cs.hku.hk"
# Replace this variable with your student number
MARKER = '3035777975'

# The Email SMTP Server
SERVER = "testmail.cs.hku.hk"   #SMTP Email Server
SPORT = 25                      #SMTP listening port
TIMEOUT = 10                    #Timeout
# For storing the attachment file information
fileobj = None                  #For pointing to the opened file
filename = ''                   #For keeping the filename


#
# For the SMTP communication
#
def do_Send():
  if(nullValueCheck()): #Checks to see if there is any empty inputs/invalid emails
    connection()


def connection():
  client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
  try:
    # set the socket timeout
    client_socket.settimeout(TIMEOUT)
    # server connection
    client_socket.connect((SERVER, SPORT))
    data = client_socket.recv(1024)
    print(data.decode())
    #Greeting
    ehlo_command = "EHLO {}\r\n".format(client_socket.getsockname()[0])
    print(ehlo_command)
    client_socket.send(ehlo_command.encode())
    data = client_socket.recv(1024)
    print(data.decode())
    #From
    mail_from_command = "MAIL FROM: <{}>\r\n".format(YOUREMAIL)
    print(mail_from_command)
    client_socket.send(mail_from_command.encode())
    data = client_socket.recv(1024)
    print(data.decode())
    if("Fail" in data.decode()):  #Check for failure in sending RCPT TO before MAIL
      alertbox(data.decode())
      return
    #To
    for email in to_list:
      rcpt_to_command = "RCPT TO: <{}>\r\n".format(email)
      print(rcpt_to_command)
      client_socket.send(rcpt_to_command.encode())
      data = client_socket.recv(1024)
      print(data.decode())
      if("Fail" in data.decode()):  #Check to see if lookup failed
        alertbox("Fail in sending RCPT TO\n"+data.decode())
        return
    #CC
    if( get_CC() != ""):
      for email in cc_list:
        cc_command = "RCPT TO: <{}>\r\n".format(email)
        print(cc_command)
        client_socket.send(cc_command.encode())
        data = client_socket.recv(1024)
        print(data.decode()) 
        if("denied" in data.decode()):
          alertbox("Fail in sending RCPT TO\n"+data.decode())
          return
    #BCC
    if (get_BCC() != ""):
      for email in bcc_list:
        bcc_command = "RCPT TO: <{}>\r\n".format(email)
        client_socket.send(bcc_command.encode())
        data = client_socket.recv(1024)
        print(data.decode())
        if("denied" in data.decode()):
          alertbox("Fail in sending RCPT TO\n"+data.decode())
          return
    data_command = "DATA\r\n"
    print(data_command)
    client_socket.send(data_command.encode())
    data = client_socket.recv(1024)
    print(data.decode())
    headers = f"From: {YOUREMAIL}\r\nTo: {get_TO()}\r\n"
    if( get_CC() != ""):  #Checks to see if there is a CC
        headers += f"Cc: {get_CC()}\r\n"
    if (get_BCC()!=""): #Checks to see if there is a BCC
      headers += f"Bcc: {get_BCC()}\r\n"
    headers += f"Subject: {get_Subject()}\r\n"
    if filename != "":
        headers += f"MIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary={MARKER}\r\n\r\n--{MARKER}\r\n"
        message_body = f"\r\n{get_Msg()}\r\n\r\n--{MARKER}\r\nContent-Type: application/octet-stream; name=\"{filename}\"\r\nContent-Disposition: attachment; filename=\"{filename}\"\r\nContent-Transfer-Encoding: base64\r\n\r\n"
        with open(filename, "rb") as file:
            message_body += base64.b64encode(file.read()).decode()
        message_body += "\r\n--boundary_string--"
    else:
        message_body = f"\r\n{get_Msg()}\r\n"

    client_socket.send(headers.encode() + message_body.encode()+"\r\n.\r\n".encode())
    data = client_socket.recv(1024)
    print(data.decode())

    quit_command = "QUIT\r\n"
    print(quit_command)
    client_socket.send(quit_command.encode())
    data = client_socket.recv(1024)
    print(data.decode())
    client_socket.close()
    exitbox("Successful!")
    

  except socket.timeout:
      alertbox("SMTP server is not available")
      client_socket.close()

def nullValueCheck():
  global to_list,cc_list,bcc_list
  if (get_TO() == ""):
    alertbox("Must enter the receipient's email")
    return False
  if (get_Subject() == ""):
    alertbox("Must enter subject")
    return False
  if (get_Msg()== "\n"):
    alertbox("Must enter message")
    return False
  to_list = get_TO().strip().replace(" ", "").split(",") #Removes all spaces, trailing/leading spaces
  cc_list = get_CC().strip().replace(" ", "").split(",") if get_CC() != "" else [] #Turns the CC string into a list
  bcc_list = get_BCC().strip().replace(" ", "").split(",") if get_BCC() != "" else [] #Turns the BCC string into a list
  for email in to_list:
      if (not echeck(email)):
        alertbox("Invalid Receipient: Email - "+email)
        return False
  for email in cc_list:
    if (not echeck(email)):
      alertbox("Invalid CC: Email - "+email)
      return False
  for email in bcc_list:
    if (not echeck(email)):
      alertbox("Invalid Receipient: Email - "+email)
      return False
  return True




#
# Utility functions
#

#This set of functions is for getting the user's inputs
def get_TO():
  return tofield.get()

def get_CC():
  return ccfield.get()

def get_BCC():
  return bccfield.get()

def get_Subject():
  return subjfield.get()

def get_Msg():
  return SendMsg.get(1.0, END)

#This function checks whether the input is a valid email
def echeck(email):   
  regex = '^([A-Za-z0-9]+[.\-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+'  
  if(re.fullmatch(regex,email)):   
    return True  
  else:   
    return False

#This function displays an alert box with the provided message
def alertbox(msg):
  messagebox.showwarning(message=msg, icon='warning', title='Alert', parent=win)

def exitbox(msg):
  messagebox.showwarning(message=msg, icon='warning', title='Alert', parent=win)
  sys.exit()

#This function calls the file dialog for selecting the attachment file.
#If successful, it stores the opened file object to the global
#variable fileobj and the filename (without the path) to the global
#variable filename. It displays the filename below the Attach button.
def do_Select():
  global fileobj, filename
  if fileobj:
    fileobj.close()
  fileobj = None
  filename = ''
  filepath = filedialog.askopenfilename(parent=win)
  if (not filepath):
    return
  print(filepath)
  if sys.platform.startswith('win32'):
    filename = pathlib.PureWindowsPath(filepath).name
  else:
    filename = pathlib.PurePosixPath(filepath).name
  try:
    fileobj = open(filepath,'rb')
  except OSError as emsg:
    print('Error in open the file: %s' % str(emsg))
    fileobj = None
    filename = ''
  if (filename):
    showfile.set(filename)
  else:
    alertbox('Cannot open the selected file')

#################################################################################
#Do not make changes to the following code. They are for the UI                 #
#################################################################################

#
# Set up of Basic UI
#
win = Tk()
win.title("EmailApp")

#Special font settings
boldfont = font.Font(weight="bold")

#Frame for displaying connection parameters
frame1 = ttk.Frame(win, borderwidth=1)
frame1.grid(column=0,row=0,sticky="w")
ttk.Label(frame1, text="SERVER", padding="5" ).grid(column=0, row=0)
ttk.Label(frame1, text=SERVER, foreground="green", padding="5", font=boldfont).grid(column=1,row=0)
ttk.Label(frame1, text="PORT", padding="5" ).grid(column=2, row=0)
ttk.Label(frame1, text=str(SPORT), foreground="green", padding="5", font=boldfont).grid(column=3,row=0)

#Frame for From:, To:, CC:, Bcc:, Subject: fields
frame2 = ttk.Frame(win, borderwidth=0)
frame2.grid(column=0,row=2,padx=8,sticky="ew")
frame2.grid_columnconfigure(1,weight=1)
#From 
ttk.Label(frame2, text="From: ", padding='1', font=boldfont).grid(column=0,row=0,padx=5,pady=3,sticky="w")
fromfield = StringVar(value=YOUREMAIL)
ttk.Entry(frame2, textvariable=fromfield, state=DISABLED).grid(column=1,row=0,sticky="ew")
#To
ttk.Label(frame2, text="To: ", padding='1', font=boldfont).grid(column=0,row=1,padx=5,pady=3,sticky="w")
tofield = StringVar()
ttk.Entry(frame2, textvariable=tofield).grid(column=1,row=1,sticky="ew")
#Cc
ttk.Label(frame2, text="Cc: ", padding='1', font=boldfont).grid(column=0,row=2,padx=5,pady=3,sticky="w")
ccfield = StringVar()
ttk.Entry(frame2, textvariable=ccfield).grid(column=1,row=2,sticky="ew")
#Bcc
ttk.Label(frame2, text="Bcc: ", padding='1', font=boldfont).grid(column=0,row=3,padx=5,pady=3,sticky="w")
bccfield = StringVar()
ttk.Entry(frame2, textvariable=bccfield).grid(column=1,row=3,sticky="ew")
#Subject
ttk.Label(frame2, text="Subject: ", padding='1', font=boldfont).grid(column=0,row=4,padx=5,pady=3,sticky="w")
subjfield = StringVar()
ttk.Entry(frame2, textvariable=subjfield).grid(column=1,row=4,sticky="ew")

#frame for user to enter the outgoing message
frame3 = ttk.Frame(win, borderwidth=0)
frame3.grid(column=0,row=4,sticky="ew")
frame3.grid_columnconfigure(0,weight=1)
scrollbar = ttk.Scrollbar(frame3)
scrollbar.grid(column=1,row=1,sticky="ns")
ttk.Label(frame3, text="Message:", padding='1', font=boldfont).grid(column=0, row=0,padx=5,pady=3,sticky="w")
SendMsg = Text(frame3, height='10', padx=5, pady=5)
SendMsg.grid(column=0,row=1,padx=5,sticky="ew")
SendMsg.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=SendMsg.yview)

#frame for the button
frame4 = ttk.Frame(win,borderwidth=0)
frame4.grid(column=0,row=6,sticky="ew")
frame4.grid_columnconfigure(1,weight=1)
Sbutt = Button(frame4, width=5,relief=RAISED,text="SEND",command=do_Send).grid(column=0,row=0,pady=8,padx=5,sticky="w")
Atbutt = Button(frame4, width=5,relief=RAISED,text="Attach",command=do_Select).grid(column=1,row=0,pady=8,padx=10,sticky="e")
showfile = StringVar()
ttk.Label(frame4, textvariable=showfile).grid(column=1, row=1,padx=10,pady=3,sticky="e")

win.mainloop()
