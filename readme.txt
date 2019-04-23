Description: A PowerShell script to convert HTML Documents to DOCX Documents
Library Used: pandocs

Folder Description:
1. bin - Contains the pandoc.exe and the PowerShell Script
2. html - This folder will store the html documents for conversion
3. docx - This folder will store the converted docx converted html documents

Input: Path to a HTML Document
Output: Converted DOCX Document inside the docx folder

How to run the script:
1. Launch an elevated PowerShell Session
2. Navigate to the Location of the bin folder. Example cd C:\Users\admin\Desktop\ConvertFrom-HTMLtoDOCX\bin\
3. Run the cmdlet: .\ConvertFrom-HTMLtoDOCX.ps1 -htmlPath "<Path to the HTML Document>". Example: .\ConvertFrom-HTMLtoDOCX.ps1 -htmlPath "C:\Users\admin\Desktop\sample.html"
4. Wait for the cmdlet to complete
5. Look into the docx folder to get the required DOCX Document

Note: This script does not require you to set the Environment Variables.
A sample HTML Document has been provided for testing purpose.

Enjoy... :)