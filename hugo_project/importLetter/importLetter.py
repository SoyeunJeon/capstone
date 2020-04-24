import email
import imaplib
import time
import html2text
import re
import logging

folderPath = 'C:/Users/Soyeun/'

username = 'forms@drhugo.com.au'
password = 'Rich3121'
server = 'mail.drhugo.com.au'
'''
username = 's3629975@student.rmit.edu.au'
password = 'Tkni0819678a'
server = 'smtp-mail.outlook.com'
'''

logging.basicConfig(filename='C:/Users/Soyeun/letter.log', format='%(asctime)s-%(levelname)s-%(message)s')
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

def readFile(filename):
    
    f = open(filename, "r")
    content = ""
    for line in f:
        line = re.sub(r'[\\x]([a-zA-Z0-9]{3})*', " ", line)
        content += line
    f.close()
    newF = open(filename, "w+")
    newF.write(content)
    print(content)
    newF.close()
    
#Function to create rtf file
def createFile(filename, content) :
    filename = folderPath + filename + '.rtf'
    f = open(filename, "w+")
    f.write(content)
    
    f.close()
    readFile(filename)
    logger.info(filename + " has been successfully created")
    print(filename + " has been successfully created")

#Function to read Email data.
def readEmail(server, name, pwd) :
    mail = imaplib.IMAP4_SSL(server)
    mail.login(name, pwd)
    mail.select("inbox")

    try :
        result, data = mail.uid('search', None, '(UNSEEN)')
        most_recent = data[0].split()[-1]
        result2, email_data = mail.uid('fetch', most_recent, '(RFC822)')
        email_message = email.message_from_string(email_data[0][1].decode("UTF-8", "ignore"))
        #print(email_message)

        #email_message = re.sub(r"x([a-z]|[A-Z]|[0-9])+", "[invalid]", email_message)
        
        emailAddress = email_message['From']
        emailSubject = email_message['Subject']

        filename = emailSubject.split(' - ')[1].split()
        filename = filename[-2]+';'+filename[1]+';'+filename[-1].replace('/', '')

        #email_message = email_message.get_payload()[0].get_payload()
        
        if "noreply@jotform.com" in emailAddress and "Questionnaire" in emailSubject:

            text = f"{email_message.get_payload(decode=True)}"
            html = text.replace("b'", "")
            h = html2text.HTML2Text()
            
            h.ignore_links = True
            output = (h.handle(f'''{html}''').replace("\\r\\n", ""))
            output = output.replace("`", "")
            #output = output.replace("'", "")
            #output = output.replace('\\', "")
            output = output.strip()

            print(output)
            #createFile(filename, output)

        else :
            print("No New Questionnaire..")
                                    
    except IndexError :
        print("No New Questionnaire")

'''
while True:
    readEmail(server, username, password)
    time.sleep(10)
'''
