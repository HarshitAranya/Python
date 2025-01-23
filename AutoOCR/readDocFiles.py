import os
import sys
from datetime import datetime
import time
import win32com.client

# Get the directory where the .exe file is located
currentDate = datetime.now()
current_directory = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.realpath(__file__))

# List all .docx files in the current directory
docx_files = [f for f in os.listdir(current_directory) if f.endswith('.docx')]

if(len(docx_files) == 0):
    print("No file found in current directory!!")
    time.sleep(5)
    sys.exit(0)

files_with_path = []
for docFile in docx_files:
        docFile_path = os.path.join(current_directory, docFile)
        files_with_path.append(docFile_path)
        file_time = os.path.getmtime(docFile_path)
        file_Datetime = datetime.fromtimestamp(file_time)
        fileDate = file_Datetime.strftime('%d-%b-%Y')
        print(f"File Date :{fileDate} | File name :{docFile}")

oneFile1 = files_with_path[0]
oneFile2 = files_with_path[1]
# print(oneFile1)
# print(oneFile2)
todaysDate = currentDate.strftime('%d-%b-%Y')
print(f"Today's Date :{todaysDate}")

def docReader(oneFileName):
    
    word = win32com.client.Dispatch("Word.Application")
    if word.Documents.Count >0:
        word.Visible = True
        # doc.Close(False)
        # word.Quit()
        print("Please keep close all the word document and run it again.")
        time.sleep(5)
        sys.exit(0)

    try:
        doc = word.Documents.Open(oneFileName)
    except Exception as e:
        print(f"Error opening document: {e}")

    paragraph_dict = {}
    # # Loop through all paragraphs and print index and text
    # # x=1
    for index, paragraph in enumerate(doc.Paragraphs, start=1):
        # if index in prange:
        ptext = paragraph.Range.Text.strip()
        ptext = ptext.replace('\r', '').replace('\x07', '')
        paragraph_dict[index] = ptext

    # print(paragraph_dict)    
        #     # x += 1
    for key, value in paragraph_dict.items():
        # print(f"{key}: {value}")
        if "Operational Change Request" in value:
            # print(f"{key}: {value}")
            header = paragraph_dict[key]     
            print(header)
            doc.Close(False)
            word.Quit()
            # word.Visible = True
    # return ""     

docReader(oneFile1)
"""
# def get_ocr_values():
#     return ocrNO, csdNo, gitNo, rocrType, sqlFiles, desc, ocrDocType, ocrTitle, severity, fstartDate, fendDate

"""