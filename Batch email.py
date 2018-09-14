#!Python3
#Automated batch email program (Addresses.xlsx)
#by Charles Han

import os, glob, win32com.client, logging, smtplib
import pandas as pd
from email.utils import parseaddr
from email_validator import validate_email, EmailNotValidError

emailCount = 0

logging.basicConfig(filename = 'log.txt', level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')

templateFolder = os.getcwd()+ '\\Email template\\' #set the run folder path
attachmentFolder = os.getcwd()+ '\\Attachments\\' #set attachment folder
# TODO get email template from folder

file = pd.ExcelFile('Email addresses.xlsx') #Establishes the excel file you wish to import into Pandas
logging.debug(file.sheet_names)

#--- Validate Email addresses excel file  ---

print("Do you wish to:")
print("1.Validate spreadsheet   2.Test run/draft emails   3.Send emails")
testflag= input()

try:
    testflag=int(testflag)                
except:
    print("Invalid input")
    
print("reading spreedsheet...")                
for s in file.sheet_names:
    df = file.parse(s) #Uploads Sheet1 from the Excel file into a dataframe

    #--- Iterate through all sheets ---
    
    if testflag >= 1:
        for index, row in df.iterrows(): #Loops through each row in the dataframe
            email = (row['Email Address'])  #Sets dataframe variable, 'email' to cells in column 'Email Addresss'
            subject = (row['Subject']) #Sets dataframe variable, 'subject' to cells in column 'Subject'
            body = (row['Email HTML Body']) #Sets dataframe variable, 'body' to cells in column 'Email HTML Body'
    
        #--- Print warnings ---  
            if pd.isnull(email): #Skips over rows where one of the cells in the three main columns is blank
                print('Sheet %s row %s: - Warning - email is null.' % (s, (index +2)))
            else:    
                try:
                    email2=(str(email))
                    validate_email(email2) # validate and get info   
                except EmailNotValidError as e:
                    print('Sheet %s row %s: - Warning - %s' % (s, (index +2), str(e) ))
                
            if pd.isnull(subject):
                print('Sheet %s row %s: - Warning - subject is null.' % (s, (index+2)))
                continue
            if pd.isnull(body):
                print('Sheet %s row %s: - Warning - email body is null.' % (s, (index+2)))
                continue
        
    
    if testflag > 1:
        #--- iterate through all sheets
        for s in file.sheet_names:
            df = file.parse(s) #Uploads Sheet1 from the Excel file into a dataframe
        # TODO: iterate through all sheets
    
            for index, row in df.iterrows(): #Loops through each row in the dataframe
                email = (row['Email Address'])  #Sets dataframe variable, 'email' to cells in column 'Email Addresss'
                #print(email)
                subject = (row['Subject']) #Sets dataframe variable, 'subject' to cells in column 'Subject'
                body = str((row['Email HTML Body'])) #Sets dataframe variable, 'body' to cells in column 'Email HTML Body'
    
                if (pd.isnull(email) or pd.isnull(subject) or pd.isnull(body)): #Skips over rows where one of the cells in the three main columns is blank
                    continue
    
        # --- Generate draft emails ---
                emailCount += 1
                olMailItem = 0x0 #Initiates the mail item object
                obj = win32com.client.Dispatch("Outlook.Application") #Initiates the Outlook application
                newMail = obj.CreateItem(olMailItem) #Creates an Outlook mail item
                newMail.Subject = subject #Sets the mail's subject to the 'subject' variable
                newMail.HTMLbody = (r"" +
                body +
                "") #Sets the mail's body to 'body' variable
                
                newMail.To = email #Sets the mail's To email address to the 'email' variable
    
                for filename in glob.iglob(attachmentFolder + '**/*.*', recursive=True):
                    #print(filename)
                    newMail.Attachments.Add(filename)
    
                if testflag == 2:            
                    newMail.display() #Displays the mail as a draft email
                    
                if testflag == 3:            
                    newMail.Send() #send emails
                    
print("Emails generated:%s" % emailCount)

# TODO: --- Send as another user/mailbox ---
##smtpObj = smtplib.SMTP('smtp.example.com', 587)
##smtpObj.ehlo()
##smtpObj.starttls()
##smtpObj.login('bob@example.com', ' SECRET_PASSWORD')
##smtpObj.sendmail('bob@example.com', 'alice@example.com', 'Subject: So long.\nDear Alice, so long and thanks for all the fish. Sincerely, Bob')
##
##smtpObj.quit()

# TODO: look into MAPi
#    session = mapi.MAPILogonEx(0, MAPIProfile, None, mapi.MAPI_EXTENDED |
#                           mapi.MAPI_LOGON_UI | mapi.MAPI_NEW_SESSION)
