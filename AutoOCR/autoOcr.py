import os
import subprocess
import json
import sys
import logging
import time

# Configure logging
logging.basicConfig(level=logging.INFO)

# Adjust paths for bundled files
def resource_path(relative_path):
    """Get the absolute path to a resource, accounting for PyInstaller bundling."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

try:
    ocrNO, csdNo, gitNo, rocrType, sqlFiles, desc, ocrDocType, ocrTitle, severity, fstartDate, fendDate = get_ocr_values()
except Exception as e:
    print(f"Error retrieving OCR values: {e}")
    sys.exit(1)


ocrFullTitle = f"{ocrNO} : {ocrTitle}"

data = {
    "OCRNo": ocrNO,
    "CSDNo": csdNo,
    "GitNo": gitNo,
    "Desc": desc,
    "OCRType": rocrType,
    "OCRTitle": ocrFullTitle,
    "OCRDocType": ocrDocType,
    "Priority": severity,
    "StartDate": fstartDate,
    "EndDate": fendDate,
}