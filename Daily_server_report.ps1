<#

    .SYNOPSIS
        Daily_server_report - Performs morning server checks.   

	
    .DESCRIPTION
	Looks up all computer accounts in AD, and performes a connection test on each.  
	Then it invokes New-AssetReport.ps1 to gather an HTML formatted report for each systesm.  
	Finally it mails all this data to specified email accounts. 
        

#>

function Zip-Files([String] $aDirectory, [String] $aZipfile){
    [string]$pathToZipExe = "$($Env:ProgramFiles)\7-Zip\7z.exe";
    [Array]$arguments = "a", "-tzip", "$aZipfile", "$aDirectory", "-r";
    & $pathToZipExe $arguments;
}


$PSEmailServer = "10.2.1.9"
$recipient = "lee.buskey.ctr@ablcda.navy.mil"
$from = "SKED@sked.abl.navy.mil"
$subject = "Daily SKED Server Status Report"
$Env = "SKED"

 Import-Module Activedirectory
$computers =  Get-ADComputer -Filter 'ObjectClass -eq "Computer"' | Select -Expand Name |sort-object


#$computers = get-content c:\_Scripts\Daily_Report\servers.txt
c:\_Scripts\Daily_Report\New-AssetReport.ps1 -Computers $computers  -Verbose
$DateCode = (get-date -uformat "%Y-%m-%d_%I_%M_%p")
get-Content *.html | set-Content c:\_scripts\Daily_Report\Reports\$DateCode-$env-Server-Report.htm
$attachments = "c:\_scripts\Daily_Report\Reports\$DateCode-$env-Server-Report.htm"
#Zip-Files "c:\_scripts\Daily_Report\Reports\$DateCode-$env-Server-Report.htm" "c:\_scripts\Daily_Report\Reports\$DateCode-$env-Server-Report.zip"
#del "c:\_scripts\Daily_Report\Reports\$DateCode-$env-Server-Report.htm"
$attachments = "c:\_scripts\Daily_Report\Reports\$DateCode-$env-Server-Report.htm"


#$temp = Get-ChildItem *.html | select name
#$attachments = $temp | foreach {"$($_.Name)"}


start-transcript c:\_Scripts\Daily_Report\connection-tests.txt
write-host ""
write-host ""

$computers | foreach-object -process {If (test-connection -quiet -computername $_ -Count 1) {Write-host "$_ passed connection tests"} else {write-host " ** $_ failed connection tests"}}
 
Stop-Transcript



Get-Content c:\_Scripts\Daily_Report\connection-tests.txt | set-Content c:\_scripts\Daily_Report\Reports\$DateCode-$env-Server-Connection-Tests.txt
$temp = Get-content c:\_Scripts\Daily_Report\connection-tests.txt
$body = out-string -inputobject $temp
Send-MailMessage -From $from -To $recipient -attachment $attachments -Subject $subject -Body $body 
del c:\_Scripts\Daily_Report\*.html
del c:\_Scripts\Daily_Report\connection-tests.txt
