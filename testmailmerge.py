from __future__ import print_function
from mailmerge import MailMerge 
from pymongo import MongoClient
import pandas as pd
import os



client = MongoClient()
db = client.merge
collection = db.data_test
datamerges = db.datamerges

# new_data = [{"Name": "Ardian",
#              "Kelas": "XI IPA 1",
#              "nilai_avg": 80},
#              {"Name": "Febri",
#              "Kelas": "XI IPA 1",
#              "nilai_avg": 83},
#              {"Name": "Della",
#              "Kelas": "XI IPA 2",
#              "nilai_avg": 81},
#              {"Name": "Devia",
#              "Kelas": "XI IPA 2",
#              "nilai_avg": 99}]

def input_new(new_data):
    new_inputs  = datamerges.insert_many(new_data)
    print(new_inputs)

def merge_documents(df, template, folder):
    for index, row in df.iterrows():
        with MailMerge(template) as document:
            document.merge(Name = str(row['Name']),
                        Kelas = str(row['Kelas']), nilai_avg = str(row['nilai_avg']))
            document.write(os.path.join(folder, str(row['Name']) + '.docx'))



merge_data = datamerges.find({})
df = pd.DataFrame(list(merge_data))
if True:
    del df['_id']

template = 'template1.docx'
folder = 'rapor'
merge_documents(df, template, folder)

