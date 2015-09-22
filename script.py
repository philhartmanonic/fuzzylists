import pandas as pd
import csv
from fuzzywuzzy import process

#bringing in spreadsheets
ac = pd.read_excel('AC9-21-15.xlsx')
ir = pd.read_excel('IRR9-14-15.xls')

#setting up empty lists
acl = []
irl = []

#setting up the AC list
for i in range(1, len(ac)):
  acl.append([ac.iloc[i]['First Name'], ac.iloc[i]['Last Name'], ac.iloc['Account Name'], ac.iloc[i]['Contact ID'], ac.iloc[i]['Account ID']])

#creating a combination of the fields to be used in the fuzzy matching
for i in acl:
  try:
    i.append(str(i[0])+str(i[1])+str(i[2]))
  except UnicodeEncodeError:
    i.append{'nothing')

#setting up the IR list
for i in range(1, len(ir)):
  irl.append([ir.iloc[i]['First Name'], ir.iloc[i]['Last Name'], ir.iloc[i]['Company']])

#creating combination of the fields to be used in fuzzy matching
for i in irl:
  try:
    i.append(str(i[0])+str(i[1])+str(i[2]))
  except UnicodeEncodeError:
    i.append('nothing')

#setting up tools for fuzzy matching
comp = []
for i in acl:
  comp.append(i[5])

#running fuzzy matching (takes a long time)
results = []
for i in irl:
  results.append([i[3], process.extractOne(i[3], comp)])
