import re
import csv
import os
import json
combined_text = ""
HTML_text = ""

DocPath = "C:\\Users\\dsmith19\\PycharmProjects\\TPdoc_extractor\\json_output\\"

fileList = [f for f in os.listdir(DocPath)]
print(fileList)

for f1 in fileList:
    with open(DocPath+"\\" + f1, "r") as text:
        for row in text:
            combined_text = combined_text +  row
        print(combined_text[:400])
        code = f1[:-5]
    print(code)
    combined_text = re.sub("\":\",", "", combined_text)
    combined_text = re.sub("\", \"", "", combined_text)
    combined_text = re.sub("\",  \"", ", ", combined_text)
    combined_text = re.sub("<i>", "<em>", combined_text)
    combined_text = re.sub("<..i>\"", "</em>", combined_text)
    combined_text = re.sub("<..i>", "</em>", combined_text)
    print(combined_text[0:400])
    combined_text = re.sub("<em>\)<.em>", "\)", combined_text)

    chopped = re.split("\":", combined_text)
    title = chopped[1]
    title = title[3:-12]
    print(title)
    input()
authors = chopped [3]

corr_author = chopped[9]

abstract = chopped [17]
# Remove tailing labels

HTML_text = HTML_text + "<strong>Title:</strong> " + title + "<strong>Authors :</strong>" + authors + "<strong>Summary: </strong>" + abstract
print(HTML_text)

input()

codelist = ["Title", "Id_code"]
for item in chopped_text:
    print(item)
    input()
    for code in codelist:
        section = re.split("\"" + code +  "\"", item)

for x in section:
    print(x)
    input()

