import xlrd
from collections import Counter
import os
import re
folderpath=('/Users/enes/PycharmProjects/PyDeneme/')
values = []
for filename in os.listdir(folderpath): # Searches for excel files
    if filename == 'output.xlsx': # My excel file name is output.xlsx
        workbook = xlrd.open_workbook(folderpath+filename) #My excel file is in another folder. You must show full file path
        for sheet in workbook.sheets():
            for row in range(sheet.nrows):
                    for col in range(sheet.ncols):
                        if col<3 and col>0: #I just wanna pull first 3 column values. You can change here as you wish
                            value = str(sheet.cell(row,col).value)
                            if " " in value:
                                for word in value.split(" "):
                                    values.append(word)
                            else:
                                values.append(value)

if os.path.exists('/Users/enes/PycharmProjects/PyDeneme/sonuc.txt'): #Deletes and creates sonuc.txt everytime. It prevents the file from being filled too much
    os.remove('/Users/enes/PycharmProjects/PyDeneme/sonuc.txt')

with open('sonuc.txt', 'w+') as f:
    for item in values:
        f.write("%s\n" % item)


kelimeler=re.findall(r'\w+', open('sonuc.txt').read()) #Reads values from txt file and appends kelimeler list. That 'kelimeler' list is same as 'values' list.

for i in Counter(kelimeler).most_common(15): # type of i is tuple. First element is name of value. Second element is the count of that value
    print(i)
