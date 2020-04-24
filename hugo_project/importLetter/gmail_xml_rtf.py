import email
import imaplib
import html2text
import re
import os

user = 'bbnbebtest@gmail.com'
password = 'BearlovesAlien'

# read email, extract the text

mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(user, password)

mail.select("inbox")

mail.list()
result, data = mail.uid('search', None, "ALL")

inbox_list = data[0].split()
recent = inbox_list[-1]
result2, email_data = mail.uid('fetch', recent, '(RFC822)')
raw_email = email_data[0][1].decode("utf-8")

email_message = email.message_from_string(raw_email)
#email type check, if !(text/html) then wrong content
#print(email_message.get_payload()[0].get_payload())

print(email_message.get_content_type())

email_message = email_message.get_payload()[0].get_payload()
#print(type(html2text.html2text(email_message)))

email_message = html2text.html2text(email_message)

email_body = email_message.split('[')
for part in email_body :
    if part.startswith("History"):
        history = part.split('|')[1].strip()

print(history)


"""
table = etree.HTML(email_message).find("body/table")
rows = iter(table)
header = [col.text for col in next(rows)]
for row in rows :
    values = col.text for col in row]
    print(dict(zip(header, values))


if "noreply@jotform.com" in email_message['From']:
    print(email_message['From'])
    text = f"{email_message.get_payload(decode=True)}"
    html = text.replace("b'", "")
    h = html2text.HTML2Text()
    h.ignore_links = True
    output = (h.handle(f'''{html}''').replace("\\r\\n", ""))
    output = output.replace("'", "")
			
    output = output.lstrip()
    output = output.rstrip()

    print(output)
    
    
    output = output.split('</patient_summary>', 1)[0]

    output = output.split('<')

    fileformat = []
    content = ""

    # xml to better format to read

    for i in range(len(output)):
        if(output[i].startswith('/')):
            continue
        reform = re.sub('>', ' : ', output[i])
        content += reform + "\n"

        if(reform.lower().startswith('firstname') or reform.lower().startswith('surname')):
            names = reform.split(': ')[1]
            fileformat.append(names)
        elif(reform.lower().startswith('dob :')):
            birthd = reform.split(': ')[1]
            temp = birthd.split('-')
            bformat = temp[2]+temp[1]+temp[0]
            fileformat.append(str(bformat))
            
    # write in rtf file

    folderPath = 'C:/Users/Soyeun'
    fileName = fileformat[1]+";"+fileformat[0]+";"+fileformat[2] + ".rtf"
    newFile = os.path.join(folderPath, fileName)

    f = open(newFile, "w+")
    f.write(content)
    f.close()
    
    print(fileformat)

else:
    print("no email")
    
"""
