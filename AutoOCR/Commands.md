# Env setup for development
pythom -m venv nameOfenv
cd nameOfenv
scripts\activate
pip install -r requirements.txt

# Manual Packages installations
pip install pywin32-308-cp312-cp312-win_amd64.whl
pip install --no-index --find-links=./ .\setuptools-75.8.0-py3-none-any.whl
pip install --no-index --find-links=./ whls\pywin32-308-cp312-cp312-win_amd64.whl
pip install --no-index --find-links=./ Azure_APIM\whls\pyinstaller_hooks_contrib-2024.11-py3-none-any.whl

# Create a build
pyinstaller --onefile --add-data "GetWI.ps1;." --add-data "CreateWI.ps1;." readDocFiles.py
pyinstaller --onefile --add-data "CreateWI.ps1;." readDocFiles.py
pyinstaller --onefile --add-data "CreateWI.ps1;." --add-data "GetWI.ps1;." --add-data "henv.json;." readDocFiles.py

# PowerShell commands
$currentDir = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)