# Importing required packages
import os
import sys
from datetime import datetime
import time
import win32com.client
import json
import subprocess
import logging

# Printing version of this utility.
print("AutoOCR | Version 7.3")
# Current date
currentDate = datetime.now()
# if exe, gets the directory containing the executable. | if not exe, gets the absolute path of the script
current_directory = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.realpath(__file__))

# List all .docx files in the current directory
# docx_files = [f for f in os.listdir(current_directory) if f.endswith('.docx')]

# Set the target directory to one level up from current directory
if getattr(sys, 'frozen', False):
    target_directory = os.path.dirname(os.path.dirname(sys.executable))  # When running as an executable
else:
    target_directory = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))  # When running as a script

# print(target_directory)

# List all .docx files in the target directory
docx_files = [f for f in os.listdir(target_directory) if f.endswith('.docx')]

print("DOCX Files:", docx_files)

# Error if no word file found.
if(len(docx_files) == 0):
    print("No OCR/ROCR/.docx file found in current directory!!")
    time.sleep(3)
    sys.exit(0)

# Error if no PAT.exe found OR empty.
pat_file_path = os.path.join(current_directory, "PAT.txt")
if not os.path.exists(pat_file_path) or os.path.getsize(pat_file_path) == 0:
    print("PAT.txt file not found or is empty!!")
    time.sleep(3)
    sys.exit(0)    

# Printing list of all available word file with file date. 
files_with_path = []
for docFile in docx_files:
        # docFile_path = os.path.join(current_directory, docFile)
        docFile_path = os.path.join(target_directory, docFile)
        files_with_path.append(docFile_path)
        file_time = os.path.getmtime(docFile_path)
        file_Datetime = datetime.fromtimestamp(file_time)
        fileDate = file_Datetime.strftime('%d-%b-%Y')
        print(f"File Date :{fileDate} | File name :{docFile}")

todaysDate = currentDate.strftime('%d-%b-%Y')
# print(f"Today's Date :{todaysDate}")

# Function to create data.json file | The data.json will be used by CreateWI.ps1 to create work item.
def jsonCreator(
    ocrFullTitle, ocrNO, desc, ocrType, ocrDocType, severity, preparedBy
):
    data = {
        "OCRTitle": ocrFullTitle,        
        "OCRNo": ocrNO,
        "Desc": desc,
        "OCRType": ocrType,
        "OCRDocType": ocrDocType,
        "Priority": severity,
        "PreparedBy": preparedBy,
    }
    # Overwrite the JSON file with an empty dictionary before updating with new data
    json_file = 'data.json'
    with open(json_file, 'w') as f:
        json.dump({}, f)  # Clear the file by writing an empty dictionary

    # Save the dictionary to a JSON file
    with open(json_file, 'w') as f:
        json.dump(data, f)

    print(f"JSON file '{json_file}' created/updated successfully.")

# Adjust paths for bundled files to execute PowerShell script
def resource_path(relative_path):
    """Get the absolute path to a resource, accounting for PyInstaller bundling."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Set PowerShell script with full path
ps_CWI = resource_path("CreateWI.ps1")
ps_GWI = resource_path("GetWI.ps1")

# Function to execute PowerShell script
def execute_powershell_script(ps1_file):
    # Prepare the PowerShell command
    ps_command = f'Set-ExecutionPolicy Bypass -Scope Process -Force; . "{ps1_file}"'

    try:
        # Execute the PowerShell script using subprocess
        result = subprocess.run(
            ['powershell', '-NoProfile', '-Command', ps_command],
            capture_output=True,
            text=True
        )

        # Handle PowerShell output
        logging.info(f"PowerShell Output:")
        logging.info(result.stdout)
        print("PowerShell Output:")
        print(result.stdout)  # Displaying the PowerShell output which user can see

        # Check if there was an error executing the script
        if result.returncode != 0:
            logging.error(f"PowerShell Script Error: {result.stderr}")
            print("PowerShell Error:")
            print(result.stderr) # Displaying the PowerShell output which user can see
    
    except Exception as e:
        # Catch and log any exceptions
        logging.error(f"Exception occurred: {str(e)}")
        print(f"Exception occurred: {str(e)}")

# Execution of GetWI.ps1 to get all available OCR from ADO. It will use PAT.txt, henv.json and create\update vailable.json
execute_powershell_script(ps_GWI)

# Load the JSON file
with open("available.json", "r") as file:
    existingData = json.load(file)

# Check if the dictionary is empty
if not existingData:
    print("Validation failed with existing OCR/ROCR  | available.json should not be empty...")
    # exit(1)  # Exit with a non-zero code
    time.sleep(3)
    sys.exit(0)

# Function to collect required information
def docReader(oneFileName):

    # Initialize all variables inside the function
    ocrNO = ""
    ocrTitle = ""
    preparedBy = ""
    csdNo = ""
    gitNo = ""
    desc = ""
    ocrDocType = ""
    severity = ""
    ocrType = ""
    executionType = []
    
    word = win32com.client.Dispatch("Word.Application")
    if word.Documents.Count >0:
        word.Visible = True
        # doc.Close(False)
        # word.Quit()
        print("Please keep close all the word document and run it again.")
        time.sleep(3)
        input("\nPress Enter to exit...")
        sys.exit(0)
    else:
        word.Visible = False
    try:
        doc = word.Documents.Open(oneFileName)
    except Exception as e:
        print(f"Error opening document: {e}")
        return

    paragraph_dict = {}
    # Loop through all paragraphs and print index and text
    for index, paragraph in enumerate(doc.Paragraphs, start=1):
        ptext = paragraph.Range.Text.strip()
        ptext = ptext.replace('\r', '').replace('\x07', '')
        paragraph_dict[index] = ptext

    # print(paragraph_dict)    
    for key, value in paragraph_dict.items():
        # print(f"{key}: {value}")
        # if "Reference" in value and "CSD" not in value:
        if value.strip() == "Reference":
            # print(f"{key}: {value}")
            ocrNO = paragraph_dict[key+1]
        if value.strip() == "Title":
            # print(f"{key}: {value}")
            ocrTitle = paragraph_dict[key+1]
        if "Change Document Prepared By" in value:
            preparedBy = paragraph_dict[key+1]
            #print(preparedBy)    
        if "Manual Treatment" in value:
            for i in range(key, key+6):
                executionType.append(paragraph_dict[i])
        # elif "Type of SCRA" in value:
        #     for i in range(key, key+5):
        #         executionType.append(paragraph_dict[i])               
        if "CSD" in value and "reference" in value.lower():
            # print(f"{key}: {value}")
            csdNo = paragraph_dict[key+1]
            csdNo = csdNo.replace(" ", "")
            # print(csdNo)
        if "Commit Number" in value:
            gitNo = paragraph_dict[key+1]
        if "Data/Code" in value:
            ocrDocType = paragraph_dict[key+1]    
        if "Severity" in value:
            severity = paragraph_dict[key+1]  
        elif "Reoccurring Operational" in value or "Standard Change Request Article (SCRA)" in value and "Data/Code" not in value:
            ocrDocType = "ROCR"
            severity = "4"
            ocrType = "NA"     
    if(ocrNO and ocrTitle):
        ocrFullTitle = f'{ocrNO} : {ocrTitle}'
    else:
        print("Error: One or more variables are not set. Check OCR_No OR Title")
        doc.Close(False)
        word.Quit()
        time.sleep(3)
        return
    if ocrDocType.lower() == "data" or ocrDocType == "ROCR":        

        if(csdNo and gitNo):
            desc = f'{ocrNO},{csdNo},{gitNo}'
        else:
            print("Error: One or more variables are not set. Check OCR_No/CSD_No OR GIT_No")
            doc.Close(False)
            word.Quit()
            time.sleep(3)
            return
        
        for box in executionType:
            # print(box)
            if "☒ - Yes" in box:
                ocrType = "Manual"
            if "☒ - No" in box:
                ocrType = "Auto"
    # For template V15 - OCR Template - System - Infrastructure                            
    elif ocrDocType.lower() == "system/infrastructure" or ocrDocType.lower() == "system" or ocrDocType.lower() == "system upgrade" or ocrDocType == "RPA":
        # print ("This is System/Infrastructure")
        ocrType = "System OCR"
        if(csdNo):
            desc = f'{ocrNO},{csdNo}'
        else:
            print("Error: One or more variables are not set. Check OCR_No/CSD_No")
            doc.Close(False)
            word.Quit()
            time.sleep(3)
            return
    # For template V13 - OCR Template - Reports    
    elif ocrDocType.lower() == "reports":
        ocrType = "Reports"
        if(csdNo):
            desc = f'{ocrNO},{csdNo}'
        else:
            print("Error: One or more variables are not set. Check OCR_No/CSD_No")
            doc.Close(False)
            word.Quit()
            time.sleep(3)
            return
    # For template V13 - OCR Template - Correspondence    
    elif ocrDocType.lower() == "correspondence":
        ocrType = "Correspondence"
        if(csdNo):
            desc = f'{ocrNO},{csdNo}'
        else:
            print("Error: One or more variables are not set. Check OCR_No/CSD_No")
            doc.Close(False)
            word.Quit()
            time.sleep(3)
            return               
    else:
        print("Document type is not matching with Data/System/Reports/RPA")
        doc.Close(False)
        word.Quit()
        return

    doc.Close(False)
    word.Quit()
    # Update data.json
    # Check if all variables are set (non-empty and not None)
    # print(f"ocrFullTitle: {ocrFullTitle}, ocrNO: {ocrNO}, desc: {desc}, ocrType: {ocrType}, ocrDocType: {ocrDocType}, severity: {severity}")
    if all([ocrFullTitle, ocrNO, desc, ocrType, ocrDocType, severity]):
        # Call jsonCreator only if all variables are set
        jsonCreator(ocrFullTitle, ocrNO, desc, ocrType, ocrDocType, severity, preparedBy)
    else:
        print("Error: One or more variables are not set.")
    # Execution of CreateWI.ps1 to create work item on ADO. It will read vailable.json, data.json and PAT.txt
    execute_powershell_script(ps_CWI)

# Calling function for each word file
for oneFile in files_with_path:
    json_file = 'data.json'
    with open(json_file, 'w') as f:
        json.dump({"temp": "temp"}, f)  # Clear the file by writing an empty dictionary

    print(f"Working on: {oneFile}")
    docReader(oneFile)   

input("\nPress Enter to exit...")
# time.sleep(10)
sys.exit(0)
