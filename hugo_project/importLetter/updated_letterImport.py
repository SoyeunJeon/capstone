import email
import imaplib
import time
import html2text
from bs4 import BeautifulSoup

folderPath = 'C:/Users/Soyeun/'
username = 'bbnbebtest@gmail.com'
password = 'BearlovesAlien'
server = 'imap.gmail.com'

def createFile(filename, content) :
    filename = folderPath + filename + '.rtf'
    f = open(filename, "w+")
    f.write(content)
    f.close()
    print("Patient file has been successfully created")

def readEmail(server, name, pwd) :
    mail = imaplib.IMAP4_SSL(server)
    mail.login(name, pwd)
    mail.select("inbox")

    try :
        result, data = mail.uid('search', None, '(UNSEEN)')
        most_recent = data[0].split()[-1]
        result2, email_data = mail.uid('fetch', most_recent, '(RFC822)')
        email_message = email.message_from_string(email_data[0][1].decode("UTF-8"))

        emailAddress = email_message['From']
        emailSubject = email_message['Subject']

        if "noreply@jotform.com" in emailAddress and "Questionnaire" in emailSubject :
            for part in email_message.walk():
                if part.get_content_type() == 'text/html' :
                    soup = BeautifulSoup(part.get_payload(decode=True), features='html.parser')

                    for script in soup(["script", "style"]) :
                        script.extract()

                
                    email_body = soup.get_text().split('[')
                    for line in email_body :
                        if line.startswith("History") :
                            history = line.split(']')[1].strip()
                            

            filename = emailSubject.split(' - ')[1].split()
            filename = filename[2]+';'+filename[1]+';'+filename[3].replace('-', '')
    
            createFile(filename, history)

        else :
            print("No New Questionnaire")
                                    
    except IndexError :
        print("No New Questionnaire")



while True: 
    readEmail(server, username, password)
    time.sleep(10)
