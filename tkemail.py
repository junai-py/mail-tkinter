import os
import smtplib
import tkinter as tk

from tkinter import *
from email import encoders
from tkinter import messagebox
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from tkinter.filedialog import askopenfilenames
from email.mime.multipart import MIMEMultipart

#---

# Defining CreateWidgets() with argument bgcolor
def CreateWidgets(bgcolor):
    labelfromEmail = Label(root, text="EMAIL - ID : ", bg=bgcolor, font=('', 10, 'bold'))
    labelfromEmail.grid(row=0, column=0, pady=5, padx=5)
    
    root.entryfromEmail = Entry(root, width=50, textvariable=fromEmail)
    root.entryfromEmail.grid(row=0, column=1, pady=5, padx=5)

    labelpasswordEmail = Label(root, text="PASSWORD : ", bg=bgcolor, font=('', 10, 'bold'))
    labelpasswordEmail.grid(row=1, column=0, pady=5, padx=5)

    root.entrypasswordEmail = Entry(root, width=50, textvariable=passwordEmail, show="*")
    root.entrypasswordEmail.grid(row=1, column=1, pady=5, padx=5)

    labeltoEmail = Label(root, text="TO EMAIL - ID : ", bg=bgcolor, font=('', 10, 'bold'))
    labeltoEmail.grid(row=2, column=0, pady=5, padx=5)

    root.entrytoEmail = Entry(root, width=50, textvariable=toEmail)
    root.entrytoEmail.grid(row=2, column=1, pady=5, padx=5)

    labelsubjectEmail = Label(root, text="SUBJECT : ", bg=bgcolor, font=('', 10, 'bold'))
    labelsubjectEmail.grid(row=3, column=0, pady=5, padx=5)

#---

    root.entry_subjectEmail = Entry(root, width=50, textvariable=subjectEmail)
    root.entry_subjectEmail.grid(row=3, column=1, pady=5, padx=5)

    labelattachmentEmail = Label(root, text="ATTACHMENT : ", bg=bgcolor, font=('', 10, 'bold'))
    labelattachmentEmail.grid(row=4, column=0, pady=5, padx=5)

    root.entryattachmentEmail = Text(root, width=38, height = 5)
    root.entryattachmentEmail.grid(row=4, column=1, pady=5, padx=5)

    buttonattachmentEmail = Button(root, text="BROWSE", command=fileBrowse, width=20)
    buttonattachmentEmail.grid(row=4, column=2, pady=5, padx=5)

    labelbodyEmail = Label(root, text="MESSAGE : ", bg=bgcolor, font=('', 10, 'bold'))
    labelbodyEmail.grid(row=5, column=0)

    root.bodyEmail = Text(root, width=80, height=20)
    root.bodyEmail.grid(row=6, column=0, columnspan=3, pady=5, padx=5)

    buttonsendEmail = Button(root, text="SEND EMAIL", command=sendEmail, width=20, bg="limegreen")
    buttonsendEmail.grid(row=7, column=2, padx=5, pady = 5)

    buttonExit = Button(root, text="EXIT", command=emailExit, width=20, bg="red")
    buttonExit.grid(row=7, column=0, padx=5, pady = 5)

#---

# Defining fileBrowse() to browse and select attachments
def fileBrowse():
    # askopenfilenames() function is used to select multiple files
    root.filename = askopenfilenames()

    # Displaying the selected files in attacmentEntry widget
    for files in root.filename:
        file = os.path.basename(files)
        root.entryattachmentEmail.insert('1.0',file+"\n")

# Defining emailExit() to exit from the GUI
def emailExit():
    MsgBox = messagebox.askquestion('Exit Application', 'Are you sure you want to exit?')
    if MsgBox == 'yes':
        root.destroy()
        os.popen("LoginGUI.pyw")

#---

# Defining sendEmail() to send the email
def sendEmail():

    # Fetching all the necessary parameters and storing in respective variables
    fromEmail1 = fromEmail.get()
    passwordEmail1 = passwordEmail.get()
    toEmail1 = toEmail.get()
    subjectEmail1 = subjectEmail.get()
    bodyEmail1 = root.bodyEmail.get('1.0', END)

    # Creating instance of class MIMEMultipart()
    message = MIMEMultipart()

    # Storing the details in respective fields
    message['From'] = "NO NAME"
    message['To'] = toEmail1
    message['Subject'] = subjectEmail1

    # Attach message with MIME instance
    message.attach(MIMEText(bodyEmail1))

#---

    # Iterating through the files in attachment list
    for files in root.filename:

        # Opening and reading the file into attachment
        attachment = open(files, 'rb').read()

        # Creating instance of MIMEBase and naming it as emailAttach
        emailAttach = MIMEBase('application', 'octet-stream')

        # Changing the payload into encoded form
        emailAttach.set_payload(attachment)
        # Encoding the attachment into base 64
        encoders.encode_base64(emailAttach)
        # Adding headers to the files
        emailAttach.add_header('Content-Disposition', "attachment; filename= %s" % os.path.basename(files))

        # Attaching the instane emailAttach to the message instance
        message.attach(emailAttach)

#---

    try:
        # Create a smtp session
        smtp = smtplib.SMTP('smtp.gmail.com', 587)
        # Starting TLS for security
        smtp.starttls()
        # Authenticate the user
        smtp.login(fromEmail1, passwordEmail1)
        # Sending the email with Mulitpart message converted into string
        smtp.sendmail(fromEmail1, toEmail1, message.as_string())

        messagebox.showinfo("SUCCESS", "EMAIL SENT TO " + str(toEmail))
        # Terminating the session
        smtp.quit()

    # Catching authenctication error
    except smtplib.SMTPAuthenticationError:
        messagebox.showerror("ERROR", "INVALID USERNAME OR PASSWORD")
    # Catching connection error
    except smtplib.SMTPConnectError:
        messagebox.showerror("ERROR", "PLEASE TRY AGAIN LATER")

#---

# Creating object of tk class
root = tk.Tk()

# Generating random color
bgColor ='grey'


# Setting the title and background color
# disabling the resizing property
root.config(background = bgColor)
root.title("PyMail")
root.resizable(False,False)

# Creating tkinter variables
toEmail = StringVar(root)
fromEmail = StringVar(root)
passwordEmail = StringVar(root)
subjectEmail = StringVar(root)

# Calling the CreateWidgets() function with argument bgColor
CreateWidgets(bgColor)

# Defining infinite loop to run application
root.mainloop()