# PowerShell MSG Scanner
This script is an updated version of my old script located https://jls3tech.wordpress.com/2013/09/21/msgextract-ps1-parses-metadata-of-msg-files/

The script expands the .msg file using 7zip then uses filters to identify key data files. That data is displayed to the user in a webpage. This is intended for SOC analysts that need to review spam emails in bulk quickly without losing fidelity. 

To use this script: 
1) Unzip and place script and folder on the c:
	C:\EmailScanner\bin
	C:\EmailScanner\msgextract
	C:\EmailScanner\ME.ps1
	C:\EmailScanner\bin\7z.exe
	C:\EmailScanner\msgextract\extractedmsg
	C:\EmailScanner\msgextract\msg
	C:\EmailScanner\msgextract\webreport.html
2) Copy emails from outlook into 'C:\EmailScanner\msgextract\msg'
3) Run C:\EmailScanner\ME.ps1 in Powershell
	May need to set-executionpolicy to bypass