
<#PSScriptInfo

.VERSION 1.2

.GUID f03ddea5-f6e3-498a-b249-1ac6b7ec8f01

.AUTHOR June Castillote

.COMPANYNAME www.lazyexchangeadmin.com

.COPYRIGHT june.castillote@gmail.com

.TAGS PowerShell Script Delete Housekeeping Logs CleanUp

.LICENSEURI

.PROJECTURI https://github.com/junecastillote/Delete-FilesByAge

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 Delete files from specified paths based on age, with email summary reporting. 

#> 
Param(

        #paths to clean up (eg. "c:\Folder1","c:\folder2")
        [Parameter(Mandatory=$true)]
        [string[]]$Paths,

        #list of files or extension to INCLUDE (eg. *.blg,*.txt)
        [Parameter(Mandatory=$true)]
        [string[]]$Include,

        #list of files or extension to EXCLUDE (eg. *.blg,*.txt)
        [Parameter(Mandatory=$false)]
        [string[]]$Exclude,

        #switch to indicate recursive action
        [Parameter()]
        [switch]$Recurse,

        [Parameter(Mandatory=$true)]
        [int]$daysToKeep,

        #path to the output/Report directory (eg. c:\scripts\output)
        [Parameter(Mandatory=$true)]
		[string]$outputDirectory,

        #path to the log directory (eg. c:\scripts\logs)
        [Parameter()]
        [string]$logDirectory,

        #prefix string for the report (ex. COMPANY)
        [Parameter()]
        [string]$headerPrefix,
        
        #Switch to enable email report
        [Parameter()]
        [switch]$sendEmail,

        #Sender Email Address
        [Parameter()]
        [string]$sender,

        #Recipient Email Addresses - separate with comma
        [Parameter()]
        [string[]]$recipients,

        #smtpServer
        [Parameter()]
        [string]$smtpServer,

        #smtpPort
        [Parameter()]
        [string]$smtpPort,

        #credential for SMTP server (if applicable)
        [Parameter()]
        [pscredential]$smtpCredential,

        #switch to indicate if SSL will be used for SMTP relay
        [Parameter()]
        [switch]$smtpSSL
)

#start FUNCTIONS
#===========================================
#Function to Stop Transaction Logging
Function Stop-TxnLogging
{
	$txnLog=""
	Do {
		try {
			Stop-Transcript | Out-Null
		} 
        catch [System.InvalidOperationException]
        {
			$txnLog="stopped"
		}
    } While ($txnLog -ne "stopped")
}

#Function to Start Transaction Logging
Function Start-TxnLogging
{
    param 
    (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$logDirectory
    )
	Stop-TxnLogging
    Start-Transcript $logDirectory -Append
}

#Function to get Script Version and ProjectURI for PSv4
Function Get-ScriptInfo
{
    param 
    (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Path
    )
    $scriptInfo = "" | Select-Object Version,ProjectURI
    $scriptInfo.Version = (Select-String -Pattern ".VERSION" -Path $Path)[0].ToString().split(" ")[1]
    $scriptInfo.ProjectURI = (Select-String -Pattern ".PROJECTURI" -Path $Path)[0].ToString().split(" ")[1]
    Return $scriptInfo
}

#===========================================
#end FUNCTIONS

#$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
#Get script version and url
if ($PSVersionTable.psversion.Major -lt 5) 
{
    $scriptInfo = Get-ScriptInfo -Path $MyInvocation.MyCommand.Definition
}
else 
{
    $scriptInfo = Test-ScriptFileInfo -Path $MyInvocation.MyCommand.Definition
}
#============================

#start PARAMETER CHECK
#===========================================
$isAllGood = $true

if ($sendEmail)
{
    if (!$sender)
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: A valid sender email address is not specified." -ForegroundColor Yellow
        $isAllGood = $false
    }

    if (!$recipients)
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: No recipients specified." -ForegroundColor Yellow
        $isAllGood = $false
    }

    if (!$smtpServer )
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: No SMTP Server specified." -ForegroundColor Yellow
        $isAllGood = $false
    }

    if (!$smtpPort )
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: No SMTP Port specified." -ForegroundColor Yellow
        $isAllGood = $false
    }
}

if ($isAllGood -eq $false)
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: Exiting Script." -ForegroundColor Yellow
    EXIT
}
#===========================================
#end PARAMETER CHECK

#start Mail Header
#===========================================
$mailHeader=@'
<!DOCTYPE html>
<html>
<head>
</head>
'@
#===========================================
#end Mail Header

#start PATHS
#===========================================
$today = Get-Date
[string]$fileSuffix = '{0:dd-MMM-yyyy_hh-mm_tt}' -f $today
$logFile = "$($logDirectory)\Log_$($fileSuffix).log"
$outputCSV = "$($outputDirectory)\delete-Summary_$($fileSuffix).csv"
$outputHTML = "$($outputDirectory)\delete-Summary_$($fileSuffix).html"

#Create folders if not found
if ($logDirectory)
{
    if (!(Test-Path $logDirectory)) 
    {
        New-Item -ItemType Directory -Path $logDirectory | Out-Null
        #start transcribing----------------------------------------------------------------------------------
        Start-TxnLogging $logFile
        #----------------------------------------------------------------------------------------------------
    }
	else
	{
		Start-TxnLogging $logFile
	}
}

if (!(Test-Path $outputDirectory)) 
{
	New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}
#===========================================
#end PATHS

#start Files List
#===========================================
$fileParams = @{
    Path = $Paths
}

if ($Recurse){$fileParams+=@{Recurse=$true}}
if ($Include){$fileParams+=@{Include=$Include}}
if ($Exclude){$fileParams+=@{Exclude=$Exclude}}

$fileParams
Write-Host ""
[datetime]$oldDate = (Get-Date).AddDays(-$daysToKeep)
#$filesToDelete = Get-ChildItem @fileParams | Where-Object {$_.LastWriteTime -ge $startDate -AND $_.LastAccessTime -le $endDate -and !$_.PSIsContainer}
$filesToDelete = Get-ChildItem @fileParams | Where-Object {$_.LastWriteTime -lt $oldDate -and !$_.PSIsContainer}
#$filesToDelete = $filesToDelete | Where-Object {!$_.PSIsContainer}
#===========================================
#end Files List

#start DELETION
#===========================================
if ($filesToDelete)
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Found Total of $($filesToDelete.Count) files" -ForegroundColor Green
    $resultLog = @()
    $successful = 0
    $failed = 0
    [int64]$deletedSize = 0
    [int64]$failedSize = 0
    foreach ($file in $filesToDelete)
    {
        $temp = "" | Select-Object FileName,FileSize,Status        
        $temp.FileName = $file.FullName
        $temp.FileSize = $file.Length
        
        try {
			Remove-Item -Path ($file.FullName) -Force -Confirm:$false -ErrorAction Stop
            $temp.Status = "Success"
            $successful = $successful+1
            $deletedSize = $deletedSize + $file.Length
            Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Delete $($file.FullName) - Success " -ForegroundColor Green
		}
		catch {
            $temp.Status = "Failed"
            $failed = $failed+1
            $failedSize = $failedSize + $file.Length
			Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Delete $($file.FullName) - Failed " -ForegroundColor Red
		}        
        $resultLog += $temp
    }
    $resultLog | Export-Csv -NoTypeInformation $outputCSV
    $summary = "" | Select-Object Paths,TotalNumberOfFiles,TotalSizeOfAllFiles,SuccessfulDeletions,FailedDeletions,TotalSuccessfulDeletionSize,TotalFailedDeletionSize
    $summary.Paths = $Paths
    $summary.TotalNumberOfFiles = "{0:N0}" -f ($filesToDelete).Count
    $summary.TotalSizeOfAllFiles = "{0:N0}" -f ($filesToDelete | Measure-Object -Property Length -Sum).Sum
    $summary.SuccessfulDeletions = "{0:N0}" -f $successful
    $summary.FailedDeletions = "{0:N0}" -f $failed
    $summary.TotalSuccessfulDeletionSize = "{0:N0}" -f $deletedSize
    $summary.TotalFailedDeletionSize = "{0:N0}" -f $failedSize
    Write-Host ""
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": SUMMARY:"
    $summary

    #start HTML OUTPUT
    #===========================================

    if ($headerPrefix)
    {
        $mailSubject = "[" + $headerPrefix + "] File Deletion Task Summary: " + ('{0:dd-MMM-yyyy hh:mm:ss tt}' -f $Today)
    }
    else 
    {
        $mailSubject = "File Deletion Task Summary: " + ('{0:dd-MMM-yyyy hh:mm:ss tt}' -f $Today)
    }

    $htmlBody += $mailHeader
    #$htmlBody += "<body><table>"
    #$htmlBody += "<tr><th>--- File Deletion Task Summary ---</th></tr>"
    #$htmlBody += "<tr><th>Paths</th><td>" + ($Paths -join ";<br />")+ "</td></tr>"
    #$htmlBody += "<tr><th>Total Number of Files</th><td>" + ($summary.TotalNumberOfFiles) + " (" + ($summary.TotalSizeOfAllFiles) + " bytes)</td></tr>"
    #$htmlBody += "<tr><th>Successful Deletion</th><td>" + ($summary.SuccessfulDeletions) + " (" + ($summary.TotalSuccessfulDeletionSize) + " bytes)</td></tr>"
    #$htmlBody += "<tr><th>Failed Deletion</th><td>" + ($summary.FailedDeletions) + " (" + ($summary.TotalFailedDeletionSize) + " bytes)</td></tr>"
    #$htmlBody += "</table>"

    $htmlBody += '<body><p><font size="2" face="Tahoma">'
    $htmlBody += "<b>[--- File Deletion Task Summary ---]</b><br /><br />"
    $htmlBody += "<b>Paths:</b> " + ($Paths -join " ; ")+ "<br />"
    $htmlBody += "<b>Total Number of Files:</b> " + ($summary.TotalNumberOfFiles) + " (" + ($summary.TotalSizeOfAllFiles) + " bytes)<br />"
    $htmlBody += "<b><font color=""Green"">Successful Deletion:</b></font> " + ($summary.SuccessfulDeletions) + " (" + ($summary.TotalSuccessfulDeletionSize) + " bytes)<br />"
    $htmlBody += "<b><font color=""Red"">Failed Deletion:</b></font> " + ($summary.FailedDeletions) + " (" + ($summary.TotalFailedDeletionSize) + " bytes)<br />"
    $htmlBody += "<br /><br />"
    $htmlBody += '<p><font size="2" face="Tahoma"><u>Paremeters</u><br />'
    $htmlBody += '<b>[SELECTION]</b><br />'
    $htmlBody += "Included: " + ($Include -join ";") + "<br />"
    $htmlBody += "Excluded: " + ($Exclude -join ";") + "<br />"
    if ($Recurse)
    {
        $htmlBody += "Recursive: Yes <br /><br />"
    }
    else 
    {
        $htmlBody += "Recursive: No <br /><br />"
    }
    
    $htmlBody += '<b>[MAIL]</b><br />'

    if ($sendEmail)
    {
        $htmlBody += "Send Email Report: Yes <br />"
    }
    else 
    {
        $htmlBody += "Send Email Report: No <br />"
    }

    $htmlBody += "SMTP Server: " + $smtpServer + "<br />"
    $htmlBody += "SMTP Port: " + $smtpport + "<br />"

    if ($smtpSSL)
    {
        $htmlBody += "SMTP SSL: Yes <br />"
    }
    else 
    {
        $htmlBody += "SMTP SSL: No <br />"
    }

    if ($smtpCredential) 
    {
        $htmlBody += "SMTP Authentication: Yes <br /><br />"
    }
    else 
    {
        $htmlBody += "SMTP Authentication: Yes <br /><br />"
    }

    $htmlBody += '<b>[REPORT]</b><br />'
    $htmlBody += 'Generated from Server: ' + (Get-Content env:computername) + '<br />'
    $htmlBody += 'Script File: ' + $MyInvocation.MyCommand.Definition + '<br />'
    $htmlBody += 'CSV Summary: ' + $outputCSV + '<br />'
    $htmlBody += 'HTML Summary: ' + $outputHTML + '<br />'
    
    $htmlBody += '</p>'
    #=====


    
    $htmlBody += "<p><a href=""$($scriptInfo.ProjectURI)"">$($MyInvocation.MyCommand.Definition.ToString().Split("\")[-1].Split(".")[0]) $($scriptInfo.version)</a></p>"
    $htmlBody += "</body></html>"

    #Export HTML Report
    $htmlBody | Out-File $outputHTML

    
    #===========================================
    #end HTML OUTPUT

    #start MAIL
    #===========================================
    if ($sendEmail)
    {
        $mailParams = @{
            From = $sender
            To = $recipients
            Subject = $mailSubject
            Body = $htmlBody
            BodyAsHTML = $true
            smtpServer = $smtpServer
            Port = $smtpPort
            useSSL = $smtpSSL
            attachments = $outputCSV
        }

        #SMTP Authentication
        if ($smtpCredential){
            $mailParams += @{credential = $smtpCredential}
        }

        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Sending email to" ($recipients -join ",") -ForegroundColor Green

        #Send message
        Send-MailMessage @mailParams
    }
    #===========================================
    #end MAIL

    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": HTML Summary Report saved in $outputHTML " -ForegroundColor Cyan
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": CSV Summary Report saved in $outputCSV " -ForegroundColor Cyan
}
else 
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": No files to delete. Exiting script" -ForegroundColor Green
}
#end DELETION
#===========================================


if ($logDirectory) {Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Transcript Log saved in $logfile " -ForegroundColor Cyan}
Stop-TxnLogging