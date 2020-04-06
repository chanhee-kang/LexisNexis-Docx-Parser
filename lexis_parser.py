# -*- coding: utf-8 -*-
import docx2txt
import pandas as pd
import os

#files = []
#
#for file in os.listdir(r"C:\Users\Administrator\Desktop\docx\test_docx\test"):
#    if file.endswith('.docx'):
#        files.append(file)
#files
#for i in range(len(files)):
#    text = docx2txt.process(files[i])
#
#text = docx2txt.process(files)
text = docx2txt.process("test.docx")

content = []
for line in text.splitlines():
  if line != '':
    content.append(line)

title_list = []
publisher_list = []
date_list = []
body_list = []
body_codes = []
load_codes = []
tmp = []
all_index = []

for i, j in enumerate(content):
    if j.lstrip() == 'Correction Appended' or content[i].lstrip() == 'All Rights Reserved' or 'Supplied by BBC' in content[i].lstrip():
        content.pop(i)


for i, j in enumerate(content):
    if j.startswith('Copyright') and (any(char.isdigit() for char in content[i - 1])) == True:
        date_list.append(content[i - 1].lstrip())
        publisher_list.append(content[i - 2].lstrip())

        if content[i - 4] != 'End of Document':
            title_list.append(content[i - 4] + content[i - 3])
        else:
            title_list.append(content[i - 3])
    else:
        pass

    if j == 'Body':
        body_codes.append(i)
    
    if 'Load-Date' in j:
        load_codes.append(i)

for b, c in zip(body_codes, load_codes):
    tmp.append([b , c])

for i in tmp:
    for x in range (i[0], i[-1]):
        if x not in i:
            i.append(x)

for i in tmp:
    i.sort()

for a in tmp:
    body_tmp = []
    for b in a:
        body_tmp.append(content[b])
    body_tmp = " ".join(body_tmp) 
    body_tmp = body_tmp.replace("\'", "'").replace("  ", " ").replace('Body', '')

    if 'Classification Language' in body_tmp:
        body_tmp = body_tmp.split('Classification Language')
        body_list.append(body_tmp[0])
    elif 'Graphic ' in body_tmp:
        body_tmp = body_tmp.split('Graphic ')
        body_list.append(body_tmp[0])
    elif 'Correction ' in body_tmp:
        body_tmp = body_tmp.split('Correction ')
        body_list.append(body_tmp[0])
    elif ' Notes ' in body_tmp:
        body_tmp = body_tmp.split(' Notes  ')
        body_list.append(body_tmp[0])
    else:
        body_list.append(body_tmp)


df = pd.DataFrame({'Publisher': publisher_list, 'Title': title_list, 'Date': date_list, 'Body' : body_list})

country = pd.read_csv('mwp_list.csv')

df_list = df.values.tolist()

country = country.values.tolist()

publisher = []
country_list = []

for i in country:
    publisher.append(i[0])
    country_list.append(i[1])

country = []

for i in [t[0] for t in df_list]:
    if i in publisher:
        country.append(country_list[publisher.index(i)])
    else:
        country.append('Unknown')
df['Country'] = country
print(df)

