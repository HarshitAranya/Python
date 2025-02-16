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
pyinstaller --onefile --add-data "CreateWI.ps1;." --add-data "GetWI.ps1;." --add-data "henv.json;." readDocFiles.py

pyinstaller --onefile --add-data "GetWI.ps1;." --add-data "CreateWI.ps1;." readDocFiles.py
pyinstaller --onefile --add-data "CreateWI.ps1;." readDocFiles.py
pyinstaller --onefile --add-data "CreateWI.ps1;." --add-data "GetWI.ps1;." --add-data "henv.json;." --noupx readDocFiles.py
pyinstaller --onefile --add-data "CreateWI.ps1;." --add-data "GetWI.ps1;." --add-data "henv.json;." --log-level DEBUG readDocFiles.py
# List required packages
AutoOCRScripts\python-3.12.8\python.exe -m pip freeze > requirements.txt

pyi-archive_viewer dist/readDocFiles.exe
AutoOCRScripts\python-3.12.8\python.exe -m pip install --target=AutoOCRScripts\python-3.12.8\Lib\site-packages -r requirements.txt

# PowerShell commands
$currentDir = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)

# Certify
# For local system
New-SelfSignedCertificate -Type CodeSigning -Subject "CN=AutoOCR" -CertStoreLocation "Cert:\CurrentUser\My"
signtool sign /tr http://timestamp.digicert.com /td sha256 /fd sha256 /sha1 8EE391AA278C23E38A69801B306FE37D5AAF1E7F readDocFiles.exe
"C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe"
Add variable path and then run this
Get-ChildItem Cert:\CurrentUser\My\8EE391AA278C23E38A69801B306FE37D5AAF1E7F | Export-Certificate -FilePath AutoOCR.cer -Type CERT
Import-Certificate -FilePath AutoOCR.cer -CertStoreLocation Cert:\CurrentUser\Root
signtool verify /pa /v readDocFiles.exe
# For all system
signtool sign /tr http://timestamp.digicert.com /td sha256 /fd sha256 /a /n "YourCertificateName" AutoOCR_v7.2.exe

# For organisation users
New-SelfSignedCertificate -Type CodeSigning -Subject "CN=AutoOCR_Internal" -CertStoreLocation "Cert:\LocalMachine\My"
Get-ChildItem Cert:\LocalMachine\My
Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Subject -match "AutoOCR_Internal" } | Select-Object Thumbprint
Get-ChildItem Cert:\LocalMachine\My\FDC6803E88808940741EA2E5AA467EA82F9B0CEE | Export-Certificate -FilePath NAutoOCR.cer -Type CERT
Import-Certificate -FilePath NAutoOCR.cer -CertStoreLocation Cert:\LocalMachine\Root
<!-- signtool sign /tr http://timestamp.digicert.com /td sha256 /fd sha256 /sha1 FDC6803E88808940741EA2E5AA467EA82F9B0CEE AutoOCR_v7.2.exe -->
<!-- signtool sign /sm /tr http://timestamp.digicert.com /td sha256 /fd sha256 /sha1 FDC6803E88808940741EA2E5AA467EA82F9B0CEE AutoOCR_v7.2.exe -->
<!-- signtool sign /sm /tr http://timestamp.sectigo.com /td sha256 /fd sha256 /sha1 FDC6803E88808940741EA2E5AA467EA82F9B0CEE AutoOCR_v7.2.exe -->
signtool sign /sm /t http://timestamp.sectigo.com /td sha256 /fd sha256 /sha1 FDC6803E88808940741EA2E5AA467EA82F9B0CEE AutoOCR_v7.2.exe
signtool sign /sm /t http://timestamp.sectigo.com /a sha256 /fd sha256 /sha1 FDC6803E88808940741EA2E5AA467EA82F9B0CEE AutoOCR_v7.2.exe
<!-- signtool sign /sm /fd sha256 /sha1 FDC6803E88808940741EA2E5AA467EA82F9B0CEE AutoOCR_v7.2.exe -->

 <!-- $cert = Get-ChildItem Cert:\LocalMachine\My\FDC6803E88808940741EA2E5AA467EA82F9B0CEE -->
<!-- $cert | Remove-Item -Verbose -->


