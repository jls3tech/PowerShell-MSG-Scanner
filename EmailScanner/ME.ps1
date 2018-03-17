#msgextract.ps1
#Parses Metadata of .msg files
#Version 2
#@TechJLS3
#02/21/2018

<#
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
#>

$ErrorActionPreference= 'silentlycontinue'

$date = Get-Date -Format D
 
#Set directories
$bin = "C:\EmailScanner\bin"
$BaseDir = "C:\EmailScanner\msgextract\"
$CapDir = $date
$msgDir = $BaseDir + "\msg\"
$report = "C:\EmailScanner\msgextract\webreport.html"
ri $report
 
#7zip
$7z = $bin + "\7z.exe"
 
Function meta_extract($_msg){
 $OutputDir = $BaseDir + "\extractedmsg\" + $CapDir #Need to fix
 &$7z x $_msg -oC:\EmailScanner\msgextract\extractedmsg\ |out-null
}

Function cleanURL($URL){
 $CleanUrl = $URL.replace("http", "hxxp")
 $CleanUrl = $CleanUrl.replace("</a>","")
 $CleanUrl = $CleanUrl.replace("</A>","")
 $CleanUrl = $CleanUrl.replace("&gt;","")
 $CleanUrl = $CleanUrl.replace("&lt;","")
 $CleanUrl = $CleanUrl.replace(">","")
 $CleanUrl = $CleanUrl.replace("<BR>","")
 $CleanUrl = $CleanUrl.replace("<br>","")
 $CleanUrl = $CleanUrl.replace("</div>","")
 $CleanUrl = $CleanUrl.replace("</body>","")
 $CleanUrl = $CleanUrl.replace("</html>","")
 $CleanUrl = $CleanUrl.replace("<div>","")
 $CleanUrl = $CleanUrl.replace("</div>","")
 $CleanUrl = $CleanUrl.replace("<span","")
 $CleanUrl = $CleanUrl.replace("</FONT>","")
 $CleanUrl = $CleanUrl.replace("</P>","")
 return $CleanUrl
}

#Generates HTML for IR Report
Function htmlOut($html){
	if(($_.Value -match '.in/') -or ($_.Value -match '.asia/') -or ($_.Value -match '.cc/') -or ($_.Value -match '.am/') -or ($_.Value -match '.vn/') -or ($_.Value -match '.ru/') -or ($_.Value -match '.cn/') -or ($_.Value -match '.cm/') -or ($_.Value -match '.info/') -or ($_.Value -match '.ws/') -or ($_.Value -match '.su/') -or ($_.Value -match '.ch/') -or ($_.Value -match 'receipt') -or ($_.Value -match '.th/')){
	$color = 3
	}
	elseif(($_.Value -match '.tv/') -or ($_.Value -match '.us/') -or ($_.Value -match '.biz/') -or ($_.Value -match '.net/') -or ($_.Value -match '.be/') -or ($_.Value -match '.br/') -or ($_.Value -match '.pl/') -or ($_.Value -match '.eu/') -or ($_.Value -match '.il/') -or ($_.Value -match '.dk/')){
	$color = 2
	}

	$CleanUrl = cleanURL $_.Value
	$BClink = "http://sitereview.bluecoat.com/sitereview.jsp?POSTDATA=port=<localport>&host=<localserver>&url=" + $CleanUrl.replace("hxxp", "http")
	$WSLink = "http://web-sniffer.net/?url=" + $CleanUrl.replace("hxxp", "http") + "&submit=Submit&http=1.1&type=GET&uak=11"
	$MLink = "https://www.mcafee.com/threat-intelligence/site/default.aspx?url=" + $CleanUrl.replace("hxxp", "http") + ""
	$VURLLink = "http://vurl.mysteryfcm.co.uk/default.asp?url=" + $CleanUrl.replace("hxxp", "http") + "&btnvURL=Dissect&selUAStr=1&selServer=" + $VURLServer + "&ref=&cbxLinks=on&cbxSource=on&cbxBlacklist=on"
	$GLink = "https://transparencyreport.google.com/safe-browsing/search?url=" + $CleanUrl.replace("hxxp", "http") + ""
	$WOTLink = "https://www.mywot.com/scorecard/" + $CleanUrl.replace("hxxp", "http") + ""
	$UVLink = "http://www.urlvoid.com/scan/" + $CleanUrl.replace("hxxp", "http") + "/"
	$SLink = "https://sitecheck.sucuri.net/results/" + $CleanUrl.replace("hxxp", "http") + "/"
	$HLink = "http://www.isithacked.com/check/" + $CleanUrl.replace("hxxp", "http")
	#$JSLink
	#$WWLink
	#$WBLink
	#$UQLink
	
	#ac $report "<span><msg>Links</msg>"

	#Creates Links
	switch ($color)
	{
	3 {
	$temp = "<p><ht>●</ht> " + $CleanUrl + " [<a href = '$BClink' title = 'Bluecoat Site Review' target='_blank'>BC</a> | <a href = '$SLink' title = 'Sucuri' target = '_blank'>Sucuri</a> | <a href = '$UVLink' title = 'URL Void' target = '_blank'>URL Void</a> |  <a href = '$WOTLink' title = 'WOT' target = '_blank'>Web of Trust</a> | <a href = '$GLink' title = 'Google' target = '_blank'>Google</a> | <a href = '$MLink' title = 'McAfee' target = '_blank'>McAfee</a> | <a href = '$VURLLink' title = 'vURL' target = '_blank'>vURL</a> | <a href = '$WSLink' title = 'Web-sniffer.net' target='_blank'>WS</a> | <a href = '$HLink' title = 'Is It Hacked?' target = '_blank'>Is It Hacked?</a>]</p>"
	ac $report $temp}
	2 {
	$temp = "<p><mt>●</mt> " + $CleanUrl + " [<a href = '$BClink' title = 'Bluecoat Site Review' target='_blank'>BC</a> | <a href = '$SLink' title = 'Sucuri' target = '_blank'>Sucuri</a> | <a href = '$UVLink' title = 'URL Void' target = '_blank'>URL Void</a> |  <a href = '$WOTLink' title = 'WOT' target = '_blank'>Web of Trust</a> | <a href = '$GLink' title = 'Google' target = '_blank'>Google</a> | <a href = '$MLink' title = 'McAfee' target = '_blank'>McAfee</a> | <a href = '$VURLLink' title = 'vURL' target = '_blank'>vURL</a> | <a href = '$WSLink' title = 'Web-sniffer.net' target='_blank'>WS</a> | <a href = '$HLink' title = 'Is It Hacked?' target = '_blank'>Is It Hacked?</a>]</p>"
	ac $report $temp}
	default {
	$temp = "<p><lt>●</lt> " + $CleanUrl + " [<a href = '$BClink' title = 'Bluecoat Site Review' target='_blank'>BC</a> | <a href = '$SLink' title = 'Sucuri' target = '_blank'>Sucuri</a> | <a href = '$UVLink' title = 'URL Void' target = '_blank'>URL Void</a> |  <a href = '$WOTLink' title = 'WOT' target = '_blank'>Web of Trust</a> | <a href = '$GLink' title = 'Google' target = '_blank'>Google</a> | <a href = '$MLink' title = 'McAfee' target = '_blank'>McAfee</a> | <a href = '$VURLLink' title = 'vURL' target = '_blank'>vURL</a> | <a href = '$WSLink' title = 'Web-sniffer.net' target='_blank'>WS</a> | <a href = '$HLink' title = 'Is It Hacked?' target = '_blank'>Is It Hacked?</a>]</p>"
	ac $report $temp}
	}
	ac $report "</span>"
	 }
 
Function fileSorter{
 <#
 All Filters
 $msg_filters =
 @(
 @{ptrn='001A'; class="Message class"},
 @{ptrn='0037'; class="Subject"},
 @{ptrn='0040'; class="Received by name"},
 @{ptrn='0042'; class="Sent repr name"},
 @{ptrn='0044'; class="Rcvd repr name"},
 @{ptrn='004D'; class="Org author name"},
 @{ptrn='0050'; class="Reply rcipnt names"},
 @{ptrn='005A'; class="Org sender name"},
 @{ptrn='0064'; class="Sent repr adrclass"},
 @{ptrn='0065'; class="Sent repr email"},
 @{ptrn='0070'; class="Topic"},
 @{ptrn='0075'; class="Rcvd by adrclass"},
 @{ptrn='0076'; class="Rcvd by email"},
 @{ptrn='0077'; class="Repr adrclass"},
 @{ptrn='0078'; class="Repr email"},
 @{ptrn='007d'; class="Message header"},
 @{ptrn='0C1A'; class="Sender name"},
 @{ptrn='0C1E'; class="Sender adr class"},
 @{ptrn='0C1F'; class="Sender email"},
 @{ptrn='0E02'; class="Display BCC"},
 @{ptrn='0E03'; class="Display CC"},
 @{ptrn='0E04'; class="Display To"},
 @{ptrn='1000'; class="Message body"},
 @{ptrn='1046'; class="Sender email"},
 @{ptrn='3001'; class="Display name"},
 @{ptrn='3002'; class="Address class"},
 @{ptrn='3003'; class="Email address"}
 )#>
 
 $msg_filters =
 @(
 @{ptrn='0037'; class="Subject"},
 @{ptrn='0C1F'; class="Sender email"},
 @{ptrn='1000'; class="Message body"},
 @{ptrn='007d'; class="Message header"}
 )
 
$data = gci C:\EmailScanner\msgextract\extractedmsg\
 foreach($f in $data){
	 foreach($c in $msg_filters){
		 if ($f -match $c.ptrn){
		$OutName = "C:\EmailScanner\msgextract\extractedmsg\" + $c.class + ".txt"
		if ($c.ptrn -match "0037"){
			$msgdata = gc C:\EmailScanner\msgextract\extractedmsg\$f|out-file $outName
			 (gc $outName) -replace "`0", "" | sc $outName
			 #$tmphtml = "<div id=email>" + $c.class + ""
			 #ac C:\EmailScanner\msgextract\extractedmsg\webreport.html $tmphtml
			 $tmpdata = gc $outName
			 $tmphtml = "`r`n<div id=email><h1>" + $c.class + ": " + $tmpdata + "</h1>"
			 ac $report $tmphtml
		}
		else{
		 $msgdata = gc C:\EmailScanner\msgextract\extractedmsg\$f|out-file $outName
		 (gc $outName) -replace "`0", "" | sc $outName
		 $tmphtml = "<span id=msg><h2>" + $c.class + "</h2>"
		 ac $report $tmphtml
		 $tmpdata = gc $outName
		 $tmphtml = "<p>" + $tmpdata + "</p></span>"
		 ac $report $tmphtml
		 if ($c.ptrn -match "1000"){
	 
				 $data = gc $outName
				 $pattern = '(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)'
				 $resultingMatches = [Regex]::Matches($data, $pattern, "IgnoreCase")
				 ac $report "<span id=msg><h2>Links</h2>"
				 if ([string]::IsNullOrEmpty($resultingMatches)){
				 #ac $report "<span id=msg>Links</span><br><p><nt>●</nt> No Links Detected!</p></span>"
				 ac $report "<p><nt>●</nt> No Links Detected!</p>"
				 }
				 $resultingMatches|foreach {
				 htmlOut $_.Value
				 }
				 ac $report "</span>"
					 }
					 
				 }
				 }}
}}
 
Function attachmentSorter{
 $attachment_filters =
 @(
 @{ptrn='3701'; class="Attachment data"},
 #@{ptrn='3702'; class="Attachment"},
 @{ptrn='3703'; class="Attach extension"},
 @{ptrn='3704'; class="Attach filename"},
 @{ptrn='3707'; class="Attach long filename"},
 @{ptrn='370E'; class="Attach mime tag"}
 )
 
 $data = gci C:\EmailScanner\msgextract\extractedmsg\__attach*\
 $data1=$data.FullName + "\*"
 $data = gci $data1
 foreach($f in $data){
 foreach($c in $attachment_filters){
 if ($f -match $c.ptrn){
	if($c.ptrn -eq "3701"){
		$msgdata = Get-FileHash $f.FullName -Algorithm SHA256
		$msgdata = $msgdata.Hash.ToLower()
		$tmphtml = "<span id=msg><h2>" + $c.class + " SHA256 Hash</h2>"
		ac $report $tmphtml
		$GoLink = "https://www.google.com/search?q=" + $msgdata + "&cad=h"
		$VTLink = "https://www.virustotal.com/#/file/" + $msgdata + "/detection"
		$tmphtml = "<p>" + $msgdata + " [<a href = '$VTLink' title = 'Google' target = '_blank'>Virus Total</a> | <a href = '$GoLink' title = 'Google' target = '_blank'>Google Hash</a>]</p></span>"
		ac $report $tmphtml
	}
	else{
		$outName = "C:\EmailScanner\msgextract\extractedmsg\" + $c.class + ".txt"
		$msgdata = gc $f.FullName|out-file $outName
		#$msgdata = gc $OutName|out-file $outName

		 (gc $outName) -replace "`0", "" | sc $outName
		 $tmphtml = "<span id=msg><h2>" + $c.class + "</h2>"
		 ac $report $tmphtml
		 $tmpdata = gc $outName
		 $tmphtml = "<p>" + $tmpdata + "</p></span>"
		 ac $report $tmphtml
		 }
	 }
	 }
 }}
 
Function Main{
	write-host '++++++++++++++++++++++++++++++++++++++++++++++++++' -foregroundcolor black -backgroundcolor yellow
	write-host '+             Email Meta Parser V2               +' -foregroundcolor black -backgroundcolor yellow
	write-host "+    Scrapes the links so, you don't have to     +" -foregroundcolor black -backgroundcolor yellow
	write-host "+         ( O_O) Bad email, " -foregroundcolor black -backgroundcolor yellow -nonewline;
	Write-host "BAD!" -foregroundcolor red -backgroundcolor yellow -nonewline;
	Write-host " (O_O )         +" -foregroundcolor black -backgroundcolor yellow;
	write-host "++++++++++++++++ !AV Might Fire! +++++++++++++++++" -foregroundcolor black -backgroundcolor yellow

	$color = "#{0:X6}" -f (Get-Random -Maximum 0xFFFFFF)
	
 ac $report "<html><head><meta charset='utf-8'><style>ht { color: red; font-size:20px; padding-left:25px; } mt { color: `#FFD700; font-size:20px; padding-left:25px; } lt { color: green; font-size:20px; padding-left:25px; } nt { color: blue; font-size:20px; padding-left:25px; } `#msg { color:blue; background: `#FFFFFF; align-content: left;} p { color:black; } h1 { color: black; text-align:left; } `#cnt{text-align:center;} `#email {padding:0px 10px 0px 10px; border:solid; margin:4px; color: `#BDBDBD} `#accordion { margin:2px; background: `#F8F8F8; } h3 {background: `#FFFFFF;} div{background: `#FFFFFF} body{background: $color} </style><title>Email Metadata Report | $date </title></head><body><h1 id=cnt>Email Metadata Report | $date </h1>"
 
$files = Get-ChildItem $msgDir\*.msg
 $files|foreach{
 ri C:\EmailScanner\msgextract\extractedmsg\*.* -recurse -exclude *.html
 meta_extract $_
 Write-host $_.Name.replace(".msg","").replace("FW ","").replace("Fw ","") -foregroundcolor black -backgroundcolor yellow
 Write-host "Metadata Extracted!" -foregroundcolor red -backgroundcolor yellow
 fileSorter $_
 
 attachmentSorter $_
 ac $report "</div>"
 }
 ac $report "</body></html>"
 ii $report
 ri C:\EmailScanner\msgextract\extractedmsg\*.* -recurse -exclude *.html
}
 
Main