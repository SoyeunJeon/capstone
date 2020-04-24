import email
import imaplib
import time
import html2text
import re
import logging

folderPath = 'C:/Users/Soyeun/'

"""
username = 'forms@drhugo.com.au'
password = 'Rich3121'
server = 'mail.drhugo.com.au'

username = 'bbnbebtest@gmail.com'
password = 'BearlovesAlien'
server = 'imap.gmail.com'
"""

username = 's3629975@student.rmit.edu.au'
password = 'Tkni0819678a'
server = 'smtp-mail.outlook.com'

logging.basicConfig(filename='C:/Users/Soyeun/letter.log', format='%(asctime)s-%(levelname)s-%(message)s')
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

#Function to create rtf file
def createFile(filename, content) :
    filename = folderPath + filename + '.rtf'
    f = open(filename, "w+")
    f.write(content)
    f.close()
    logger.info(filename + " has been successfully created")
    print(filename + " has been successfully created")

#Function to read Email data.
def readEmail(server, name, pwd) :
    mail = imaplib.IMAP4_SSL(server)
    mail.login(name, pwd)
    mail.select("inbox")

    #Check email inbox for an unread email.
    try :
        result, data = mail.uid('search', None, 'ALL')
        most_recent = data[0].split()[-1]
        result2, email_data = mail.uid('fetch', most_recent, '(RFC822)')
        email_message = email.message_from_string(email_data[0][1].decode("UTF-8"))
        
        emailAddress = email_message['From']
        emailSubject = email_message['Subject']

        #Gets the surname, name and DOB of the patient and names the created file in the correct format accepted by Genie.
        filename = emailSubject.split(' - ')[1].split()
        filename = filename[-2]+';'+filename[1]+';'+filename[-1].replace('-', '') # how about long name patient?

        #Checks the email to see if it is a valid email to scrape data from.
        if "noreply@jotform.com" in emailAddress and "Questionnaire" in emailSubject :
            
            #Converts raw Email to HTML format then to plain text.
            email_message = email_message.get_payload()[0].get_payload()
            
            tmp = ""
            for each in email_message :
                tmp += str(each)
            
            email_message = html2text.html2text(tmp)
            #Find History then extract the data after the '[' character.
            print(email_message)
            email_body = email_message.split('[')
            for part in email_body :
                if part.startswith("History") :
                    
                    history = part.split(']')[1].replace("** |", "")
                    print(history)
                    history = history.replace("=", "").strip()
            
            
            """
            dob = emailSubject[-1]
            emailSubject = emailSubject.split()
            for elem in emailSubject :
            """
            createFile(filename, history)

        else :
            print("No New Questionnaire..")
                                    
    except IndexError :
        print("No New Questionnaire")

'''
while True:
    readEmail(server, username, password)
    time.sleep(10)
'''
readEmail(server, username, password)
