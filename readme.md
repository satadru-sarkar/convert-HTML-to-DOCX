
# convert-HTML-to-DOCX

Description: A PowerShell script to convert HTML Documents to DOCX Documents<br/>
Library Used: [**pandoc**](https://pandoc.org)<br/><br/>

Folder Description:<br/>
1. bin - Contains the pandoc.exe and the PowerShell Script<br/>
2. html - This folder will store the html documents for conversion<br/>
3. docx - This folder will store the docx converted html documents<br/><br/>

Input: Path to a HTML Document<br/>
Output: Converted DOCX Document inside the docx folder<br/><br/>

How to run the script:<br/>
1. Copy the 'ConvertFrom-HTMLtoDOCX' folder to a location on your PC<br/>
2. Due to Repository size restrictions, download(pandoc-X.X.X-windows-x86_64.zip) and copy the pandoc.exe file, from the [link](https://github.com/jgm/pandoc/releases/tag/2.7.2), and paste it into the bin folder<br/>
Note: pandoc.exe should be directly inside the bin folder<br/>
3. Launch an elevated PowerShell Session<br/>
4. Navigate to the Location of the bin folder. Example cd C:\Users\admin\Desktop\ConvertFrom-HTMLtoDOCX\bin\<br/>
5. Run the cmdlet: .\ConvertFrom-HTMLtoDOCX.ps1 -htmlPath "<Path to the HTML Document>". Example: .\ConvertFrom-HTMLtoDOCX.ps1 -htmlPath "C:\Users\admin\Desktop\sample.html"<br/>
6. Wait for the cmdlet to complete<br/>
7. Look into the docx folder to get the required DOCX Document<br/><br/>

Note: This script does not require you to set the Environment Variables.<br/>
A sample HTML Document has been provided for testing purpose.<br/><br/>

Link to pandoc: https://pandoc.org/<br/><br/>

Enjoy... :)
