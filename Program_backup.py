#!/home/vincent/anaconda3/envs/Main/bin/python

import cgitb
import re
import pandas as pd
import xlrd
import spacy
import cgi
import playsound
from gtts import gTTS

nlp = spacy.load('en_core_web_sm')
form = cgi.FieldStorage()

def speak(text):
    tts = gTTS(text=text, lang='en')
    filename = "voice.mp3"
    tts.save(filename)
    playsound.playsound(filename)

def buildownkeywords(excel_file_path, shttarget):
    # https://stackoverflow.com/questions/20219254/how-to-write-to-an-existing-excel-file-without-overwriting-data-using-pandas/42093464
    # read a single or multi-sheet excel file
    # (returns dict of sheetname(s), dataframe(s))
    file_path = excel_file_path
    df = pd.read_excel(file_path, encoding='utf-8')  # typical for searching multiple value
    # ws_dict = df.to_dict()
    ws_dict = df.to_dict('list')
    # print(ws_dict)
    return ws_dict


def readxldb(excel_file_path, shttarget):
    # Sheet3: multiple rows with 2 columns only
    global elm
    workbook = xlrd.open_workbook(excel_file_path, on_demand=True)
    worksheet = workbook.sheet_by_index(1)
    first_row = []  # The row where we stock the name of the column
    for row in range(worksheet.nrows):
        first_row.append(worksheet.cell_value(row, 0))
    # transform the workbook to a list of dictionaries
    # data =[]
    for col in range(1, worksheet.ncols):
        elm = {}
        for row in range(worksheet.nrows):
            elm[first_row[row]] = worksheet.cell_value(row, col)
        # data.append(elm)
    return elm

def listtostring(lemmetised):  # Start the function to covert the content of the list into a string
    return ' '.join(lemmetised)


###################################################added
sourcefile = 'HCI DB (copy).xlsx'
keywords = {}
responses = readxldb(sourcefile, 'Response')
keywords_dict = buildownkeywords(sourcefile, 'Keyword')
key = 'test'
###################################################end
i = 0
keytoanswer = ""
word = ""
responese = []
lemmetised = []
temp = ""
new = ""
hello = "Hello"
#######################################################

temp = form["Response"]
temp = str(temp)

print("content-type: text/html\n\n")
print("<html><body>")
##########Inside the HTML##########
temp = temp.replace('MiniFieldStorage','')
new = temp.replace('(','')
temp = new.replace(')','')
new = temp.replace("'",'')
temp = new.replace('Response','')
new = temp.replace(',','')

#print(new)

userinput = new

reply = nlp(userinput)


for word in reply:  # Removing the stop words and putting the non-stop words into the list
    if not word.is_stop:
        if not spacy.tokens.token.Token:
            break
        responese.append(word)

for token in responese:
    #print(token, "---> " + token.lemma_)
    str = token.lemma_
    str = str.lower()
    lemmetised.append(str)

#print(lemmetised)
#print(word)

###################################################added
for word in lemmetised:
    for baseword, pattern in keywords_dict.items():
        # print(baseword, pattern)  # -> to remove in LIVE TRIAL
        for pat in pattern:
            # if a keyword matches, select the corresponding intent from the keywords_dict dictionary
            if pat == word:  # -> to show the list of words
                #print(pat)
                #print(word)
                if re.search(pat, word):
                    key = baseword
                    keytoanswer = baseword
    #print(word + " # -> from DB : " + key)
    key = 'not exist'
speak(responses[keytoanswer])
try:
    print("<img class='floating' src = 'http://www.hci.edu.sg/images/HCI_logo_2.jpg'  height='100px' width='330px'>")
    print("<p>")
    print("<input style='font-size: 20px' type='text' size='50' value='"+responses[keytoanswer]+"'readonly>")
    print("</p>")
    print("<div>")
    print("<audio controls autoplay>")
    print("<source src=voice.mp3 type='audio/mpeg'>")
    print("</audio>")
    print("</div>")
    print("<p>")
    print("To return to the previous page:")
    print("<a href = 'http://localhost://Main.html'> Click here</a>")
    print("</p>")
    print("</body></html>")
except:
    print("<h1>")
    print("Try retyping your sentence into shorter ones")
    print("</h1>")
    print("<h1>")
    print("Or you could try retyping your sentence")
    print("</h1>")
print("</body></html>")

