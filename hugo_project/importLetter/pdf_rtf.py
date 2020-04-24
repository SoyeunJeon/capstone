from tika import parser
import os
import pypandoc
import re
from datetime import datetime

#file = 'C:/Users/Soyeun/Desktop/2018-11-26-REG-Mrs-Suzan-Saban.pdf'

filename = 'Mrs-John-Smith-123.pdf'
path = 'C:/Users\Soyeun\Desktop\importLetter'
file = os.path.join(path, filename)

# extract data from pdf file

file_data = parser.from_file(file)
text = file_data['content']

words = text.split()
nameform = []

text = re.sub('\ufffd', '', text)
text = text.lstrip()
print(text)

# suppose the form format is name/dob/age come after each other
# name format ['Name'] [prefix] [firstname] [surname]
for i in range(len(words)):
    if words[i].lower() == 'name':
        nameform.append(words[i+2])
        nameform.append(words[i+3])
    elif words[i].lower() == 'date':
        #dob = re.sub('/', '', words[i+1])
        #nameform.append(str(dob))
        dob = words[i+4]+" "+words[i+5]+" "+words[i+6]
        dob = re.sub(',', '', dob)
        dob = datetime.strptime(dob, '%B %d %Y')
        dob = str(dob)
        dob = dob.split()[0]
        dob = dob.split('-')
        dob = dob[2]+dob[1]+dob[0]
        print(dob)
        nameform.append(dob)
        break

#re.sub("([A-Za-z\-'\. ]*) as ([A-Za-z\-'\. ]*)", '', text)


newfile = "C:/Users/Soyeun/" + nameform[1] + ";" + nameform[0] + ";" + nameform[2] + ".rtf"
print(newfile)
f = open(newfile, "w+")
f.write(text)
f.close()

