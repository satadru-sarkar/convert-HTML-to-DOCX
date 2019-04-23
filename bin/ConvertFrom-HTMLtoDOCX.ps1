#**********Script written by Satadru Sarkar**********

#This is a PowerShell script to convert HTML Documents to DOCX Documents
#The script makes use of pandocs to achieve the conversion


#The following line defines the argument accepted by the PowerShell Script
param([string]$htmlPath)

#Navigating to the Base Location
Set-Location -Path ..

#Storing the Base Location to a variable
$baseLocation = Get-Location

#Navigating to the html folder
Set-Location -Path "$BaseLocation\html"

#Copying the HTML file from the location provided as an argument to the script and in to the bin folder
Copy-Item $htmlPath

#Getting the HTML file into a variable
$srcFile = Get-ChildItem "$baseLocation\html" -filter "*.html"

#Getting the File Base Name
$srcFileBaseName = $srcFile.BaseName

#Navigating back to the bin folder to execute pandoc.exe
Set-Location -Path "$baseLocation\bin"

#Storing the location of the HTML file to a varibale
$htmlLocation = "$baseLocation\html\$srcFile"

#String the location of the DOCX file to a variable
$docxLocation = "$baseLocation\docx\$srcFileBaseName.docx"

#Executing pandoc.exe to convert the HTML to DOCX
.\pandoc.exe -s $htmlLocation -o $docxLocation