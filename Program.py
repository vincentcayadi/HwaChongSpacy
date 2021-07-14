#!/home/vincent/anaconda3/envs/Main/bin/python

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
    tts = gTTS(text=text, lang='en-uk')
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
#######################################################

temp = str(form["Response"].value)

print("content-type: text/html\n\n")
print("<html><body>")
##########Inside the HTML##########

userinput = nlp(temp)

for word in userinput:  # Removing the stop words and putting the non-stop words into the list
    if not word.is_stop:
        if not spacy.tokens.token.Token:
            break
        responese.append(word)

for token in responese:
    str = token.lemma_.lower()
    lemmetised.append(str)

###################################################added
for word in lemmetised:
    for baseword, pattern in keywords_dict.items():
        for pat in pattern:
            # if a keyword matches, select the corresponding intent from the keywords_dict dictionary
            if pat == word:  # -> to show the list of words
                if re.search(pat, word):
                    key = baseword
                    keytoanswer = baseword
    key = 'not exist'

try:
    print("<div>")
    print("<img class='floating' src = 'http://www.hci.edu.sg/images/HCI_logo_2.jpg'  height='100px' width='330px'>")
    print("<p>")
    print("<input style='font-size: 20px' type='text' size='50' rows='3' cols='65' value='" + responses[keytoanswer] + "'readonly>")
    speak(responses[keytoanswer])
    print("</p>")
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
    print("<b>")
    print("<p>")
    print("<ul>")
    print("<li>")
    print("Try retyping your sentence into shorter ones")
    print("</li>")
    print("<li>")
    print("Try asking more relevant questions")
    print("</li>")
    print("<li>")
    print("Check your spelling for any errors")
    print("</li>")    
    print("<li>")
    print("Or you could try retyping your sentence")
    print("</li>")
    print("</ul>")
    print("</p>")
    print("<p>")
    print("To return to the previous page:")
    print("<a href = 'http://localhost://Main.html'> Click here</a>")
    print("</p>")
    print("</b>")
print("</body></html>")

