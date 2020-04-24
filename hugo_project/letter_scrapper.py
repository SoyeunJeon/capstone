""" Auto Data Importer: The program responsible for actually importing the letters/consultation/history data. """

import re
import os
import sys
import json
import email
import imaplib
import logging
import traceback
import html2text
from pathlib import Path

def run_importer() :
    """This method is reponsible for loging into the desired email for Letters Scraping, 
    and checking if a valid email exists so that it may convert it to rtf format for 
    Genie. returns a tuple(True, None) if successful, or (False, <error message>) otherwise.
    Requires a credentials.json file with the following format in the same folder:
            
            {
                "USER": " a user name ", 
                "PASS": " a password ", 
                "DEMOGRAPHICS_KEY": " a key to identify demographics/registration emails", 
                "DEMOGRAPHICS_FILEPATH": " a filepath in which to save the files", 
                "EMAIL_SERVER": " your email server ", 
                "EMAIL_FOLDER": " the folder in your email which you want to access (not subfolder) " }           
                
    This credentials.json file can be edited via Demographics_Config_UI.exe      """

    logging.basicConfig(filename='letter_scrapper.log', filemode='a+',\
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s - %(message)s',\
         level=logging.DEBUG, datefmt='%d-%b-%y %H:%M:%S')
         
    try:
        # determine if application is a script file or frozen exe
        if getattr(sys, 'frozen', False):
            path = os.path.dirname(sys.executable)
            logging.info('<<<<<<<<<<   Running as frozen (exe)   >>>>>>>>>>')
        else:
            path = os.path.dirname(__file__)
            logging.info('<<<<<<<<<<   Running as script   >>>>>>>>>>')

        # Read in configurations.
        try:
            with open(os.path.join(path, 'credentials.json'), 'r') as f:
                configuration = json.load(f)
                f.close()
        except:
            logging.error('----- FAILED TO READ IN CONFIGURATIONS! -----')
            return (False, "----- FAILED TO READ IN CONFIGURATIONS! -----")

        #Username, Password, Access Key
        credentials = tuple((configuration['USER'], configuration['PASS'], \
            configuration['LETTERS_KEY'], configuration['EMAIL_SERVER'], \
                    configuration['LETTERS_FILEPATH'], configuration['EMAIL_FOLDER']))
        
        #Connect via imap or pop protocol-
        try:
            mail = imaplib.IMAP4_SSL(credentials[3])
            mail.login(credentials[0], credentials[1])
        except Exception as e:
            logging.error('----- FAILED TO LOGIN TO EMAIL SERVER! -----')
            return (False, "----- FAILED TO LOGIN TO EMAIL SERVER! -----")

        #Select the folder----------------
        try:    mail.select(str(credentials[5]).upper())
        except: 
            try: mail.select(str(credentials[5]).lower())
            except: 
                try:
                    mail.select(str(credentials[5]))
                except Exception as e:
                    logging.error('----- FAILED TO ACCESS EMAIL FOLDER! -----')
                    return (False, "----- FAILED TO ACCESS EMAIL FOLDER! -----")
                    
        #Search for only UNREAD emails----
        result, data = mail.uid('search', None, "UNSEEN")

        #Convert id's string to id array--
        inbox_item_list = data[0].split()

        #Fetch the emails, clean message--
        for item in inbox_item_list:

            #Read email without marking as read, change to (RFC822) or (BODY.PEEK[]) to read and mark as read or peek only
            result2, email_data = mail.uid('fetch', item, '(BODY.PEEK[])')
            raw_email = email_data[0][1].decode("utf-8")
            email_message_main = email.message_from_string(raw_email)

            #If valid email, handle it.
            if credentials[2] in email_message_main['Subject'] and \
                "noreply@jotform.com" in email_message_main['From']:
                
                #Mark email as read.
                mail.uid('fetch', item, '(RFC822)')

                #Obtain filename
                filename = email_message_main['Subject'].split(' - ')[1].split()
                filename = filename[-2]+';'+filename[1]+';'+filename[-1].replace('-', '')

                # Get payload and clean it up so that it is well formatted free flowing text.

                #Check is ascii characters exist in this email.
                test_text_for_ascii = f"{email_message_main.get_payload(decode=True)}"
                test_contains_ascii = False
                #Prepare warning message is ascii found.
                warning = ""
                try:
                    test_text_for_ascii.encode('ascii') #Check for ascii chars
                    test_contains_ascii = True # True if found.
                    warning = " <<<<<<<<<<<<<<<<<<<<    WARNING!    >>>>>>>>>>>>>>>>>>>>\n" + \
                                " INVALID CHARACTERS FOUND. EMAIL CROSS-CHECK RECOMMENDED.\n" + \
                                "_._._._._._._._._._._._._._._._._._._._._._._._._._._._._\n\n"
                except: pass #no ascii chars found... move on as per normal.

                #Now, move on as per normal.
                text = f"{email_message_main.get_payload()}"
                html = text.replace("b'", "")
                h = html2text.HTML2Text()
                h.ignore_links = True
                
                # Attach warning to the top if ascii's found
                if test_contains_ascii:
                    output = warning + "\n\n" 
                
                #produce text as per normal.
                output += (h.handle(f'''{html}''').replace("\\r\\n", ""))
                output = output.replace("'", "")
                output = output.strip()

                #if ascii chars found, replace them with 'placeholder'. This looks neater.
                if test_contains_ascii:
                    output = re.sub(r"=([a-z]|[A-Z]|[0-9]|=)+", "[IC]", output)

                #Obtain filepath and write to file
                path = Path(credentials[4]) / Path(str(filename) + ".rtf").as_posix()
                f = open(path, 'w')
                f.write(str(output))
                f.close()

                #log success
                logging.info(f'Email downloaded and saved as rtf, to {path}. ---- Subject: ' + str(email_message_main['Subject']))
                                        
        return (True, None)
    except Exception as e:
        #log failure
        tb = traceback.format_exc()

        try:    subject = "Subject: " + str(email_message_main['Subject'])
        except: subject = "Subject unobtainable. "
        
        try:    date = "Date: " + str(email_message_main['date'])
        except: date = "Date unobtainable. "
        
        try:    contents = "Contents: " + str(email_message_main.get_payload()[0].get_payload())
        except: contents = "Contents unobtainable. "
        
        logging.error('\n----- FAILED TO SCRAPE EMAIL!  -----\n' + str(e) + '\n'
            "----------- TRACEBACK -----------\n" + str(tb) + '\n' +
                "------------- DATA -------------\n" + subject + '\n' +
                    date + '\n' +contents + '\n') 
        return (False, "---- FAILED TO SCRAPE EMAIL! ----\n" + str(e))