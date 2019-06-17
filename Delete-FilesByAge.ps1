
<#PSScriptInfo

.VERSION 1.3.7

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

1.3.3 (June 16, 2019)
    - Added Teams Notification Option
1.3.4 (June 17, 2019)
    - Additional fact (Source:) for MS Teams notification
1.3.5 (June 18, 2019)
    - Fixed CSS formatting of report
    - Fixed MS Teams JSON notification format
    - Code cleanup
1.3.6 (June 18, 2019)
    - Added timezone format
1.3.7 (June 18,2019)
    - Fixed files to delete collection

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
        [switch]$smtpSSL,

        #accepts Teams WebHook URI
        [Parameter()]
        [string[]]$notifyTeams
)

#...................................
#Region FUNCTION
#...................................
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
	
	$scriptFile = Get-Content $Path

	$props = @{
		Version = ""
		ProjectURI = ""
	}

	$scriptInfo = New-Object PSObject -Property $props

	# Get Version
	foreach ($line in $scriptFile)
	{	
		if ($line -like ".VERSION*")
		{
			$scriptInfo.Version = $line.Split(" ")[1]
			BREAK
		}	
	}

	# Get ProjectURI
	foreach ($line in $scriptFile)
	{
		if ($line -like ".PROJECTURI*")
		{
			$scriptInfo.ProjectURI = $line.Split(" ")[1]
			BREAK
		}		
	}
	Remove-Variable scriptFile
    Return $scriptInfo
}

#Function to get current system timezone (for PS versions below 5)
Function Get-TimeZoneInfo
{  
	$tzName = ([System.TimeZone]::CurrentTimeZone).StandardName
	$tzInfo = [System.TimeZoneInfo]::FindSystemTimeZoneById($tzName)
	Return $tzInfo	
}
#...................................
#EndRegion FUNCTION
#...................................
#...................................
#Region SCRIPT INFO
#...................................
if ($PSVersionTable.psversion.Major -lt 5) 
{
    $scriptInfo = Get-ScriptInfo -Path $MyInvocation.MyCommand.Definition
}
else 
{
    $scriptInfo = Test-ScriptFileInfo -Path $MyInvocation.MyCommand.Definition
}
#...................................

#Get TimeZone Information
$timeZoneInfo = Get-TimeZoneInfo

#EndRegion SCRIPT INFO
#...................................
#...................................
#Region PARAMETER CHECK
#...................................
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
#...................................
#EndRegion PARAMETER CHECK
#...................................
#...................................
#Region CSS
#...................................
$css_string = @'
<style type="text/css">
#HeadingInfo 
	{
		font-family:"Segoe UI";
		width:100%;
		border-collapse:collapse;
	} 
#HeadingInfo td, #HeadingInfo th 
	{
		font-size:0.8em;
		padding:3px 7px 2px 7px;
	} 
#HeadingInfo th  
	{ 
		font-size:2.0em;
		font-weight:normal;
		text-align:left;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#604767;
		color:#fff;
	} 
#SectionLabels
	{ 
		font-family:"Segoe UI";
		width:100%;
		border-collapse:collapse;
	}
#SectionLabels th.data
	{
		font-size:2.0em;
		text-align:left;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#fff;
		color:#000; 
	} 
#data 
	{
		font-family:"Segoe UI";
		width:100%;
        border-collapse:collapse;
	} 
#data td, #data th
	{ 
		font-size:0.8em;
		border:1px solid #DDD;
        padding:3px 7px 2px 7px; 
        vertical-align:top;
	} 
#data th  
	{
		font-size:0.8em;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#00B388;
        color:#fff; text-align:left;        
	} 
#data td 
	{ 	font-size:0.8em;
		padding-top:5px;
		padding-bottom:4px;
        text-align:left;
	} 
#data td.bad
	{ 	font-size:0.8em;
		font-weight: bold;
		padding-top:5px;
		padding-bottom:4px;
		color:#f04953;
	} 
#data td.good
	{ 	font-size:0.8em;
		font-weight: bold;
		padding-top:5px;
		padding-bottom:4px;
		color:#01a982;
	}

.status {
	width: 10px;
	height: 10px;
	margin-right: 7px;
	margin-bottom: 0px;
	background-color: #CCC;
	background-position: center;
	opacity: 0.8;
	display: inline-block;
}
.green {
	background: #01a982;
}
.purple {
	background: #604767;
}
.orange {
	background: #ffd144;
}
.red {
	background: #f04953;
}
</style>
'@
#...................................
#EndRegion CSS
#...................................

#...................................
#Region PATHS
#...................................
$today = Get-Date
[string]$fileSuffix = '{0:dd-MMM-yyyy_hh-mm_tt}' -f $today
$today = $today.ToString("F")
$today = "$($today) $($timeZoneInfo.DisplayName.ToString().Split(" ")[0])"
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
#...................................
#EndRegion PARAMETER CHECK
#...................................

#...................................
#Region GENERATE FILES LIST
#...................................
$fileParams = @{
    Path = $Paths
}

if ($Recurse){$fileParams+=@{Recurse=$true}}
if ($Exclude){$fileParams+=@{Exclude=$Exclude}}

$fileParams
Write-Host ""
[datetime]$oldDate = (Get-Date).AddDays(-$daysToKeep)

#$filesToDelete = Get-ChildItem @fileParams | Where-Object {$_.LastWriteTime -lt $oldDate -and !$_.PSIsContainer}
$filesToDelete = @()
foreach ($fInclude in $Include){
    $temp = Get-ChildItem @fileParams -Filter $fInclude | Where-Object {$_.LastWriteTime -lt $oldDate -and !$_.PSIsContainer}
    $filesToDelete += $temp
}

#...................................
#EndRegion GENERATE FILES LIST
#...................................

#...................................
#Region FILE DELETION
#...................................
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

    #...................................
    #Region HTML
    #...................................

    if ($headerPrefix)
    {
        $mailSubject = "[" + $headerPrefix + "] File Deletion Task Summary"
    }
    else 
    {
        $mailSubject = "File Deletion Task Summary"
    }
   
    $htmlBody += "<html><head><title>$($mailSubject) - $($Today)</title><meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />"
    $htmlBody += $css_string
    $htmlBody += '</head><body><p><font size="2" face="Tahoma">'
    $htmlBody += '<table id="HeadingInfo">'

    if ($headerPrefix)
    {
        $htmlBody += '<tr><th>'+ $headerPrefix + '<br />Delete Files Older Than ' + $daysToKeep + ' Days<br / >'+ $today +'</th></tr>'
    }
    else 
    {
        $htmlBody += '<tr><th>Delete Files Older Than ' + $daysToKeep + ' Days<br / >'+ $today +'</th></tr>'
    }
    $htmlBody += '</table><hr />'
    $htmlBody += '<table id="SectionLabels">'
    $htmlBody += '<tr><th class="data">Summary</th></tr></table>'
    $htmlBody += '<table id="data">'
    $htmlBody += '<tr><th width="15%">Computer</th><td>'+ $env:COMPUTERNAME +'</td></tr>'
    $htmlBody += '<tr><th>Paths</th><td>'+ ($Paths -join "<br />") +'</td></tr>'
    $htmlBody += '<tr><th>Successful Deletion</th><td class="good">' + ($summary.SuccessfulDeletions) + ' files (' + ($summary.TotalSuccessfulDeletionSize) + ' bytes)</td></tr>'
    $htmlBody += '<tr><th>Failed Deletion</th><td class="bad">' + ($summary.FailedDeletions) + ' files (' + ($summary.TotalFailedDeletionSize) + ' bytes)</td></tr>'
    $htmlBody += '<tr><th>Total Files</th><td>' + ($summary.TotalNumberOfFiles) + ' files (' + ($summary.TotalSizeOfAllFiles) + ' bytes)</td></tr>'
    $htmlBody += '</table><hr />'

    #start table SETTINGS
    $htmlBody += '<table id="SectionLabels">'
    $htmlBody += '<tr><th class="data">Settings</th></tr></table>'
    $htmlBody += '<table id="data">'
    $htmlBody += '<tr><th width="15%">Included</th><td>'+ ($Include -join ";") +'</td></tr>'
    $htmlBody += '<tr><th width="15%">Excluded</th><td>'+ ($Exclude -join ";") +'</td></tr>'
    $htmlBody += '<tr><th width="15%">Recursive</th><td>'+ (invoke-command {if ($Recurse) {return "Yes"} else {return "No"}}) +'</td></tr>'
    $htmlBody += '<tr><th width="15%">Send Email Report</th><td>'+ (invoke-command {if ($sendEmail) {return "Yes"} else {return "No"}}) +'</td></tr>'
    $htmlBody += '<tr><th width="15%">SMTP Server Name or IP</th><td>'+ $smtpServer +'</td></tr>'
    $htmlBody += '<tr><th width="15%">SMTP Server Port</th><td>'+ $smtpPort +'</td></tr>'
    $htmlBody += '<tr><th width="15%">SMTP SSL in Use</th><td>'+ (invoke-command {if ($smtpSSL) {return "Yes"} else {return "No"}}) +'</td></tr>'
    $htmlBody += '<tr><th width="15%">SMTP Login Required</th><td>' + (invoke-command {if ($smtpCredential) {return "Yes"} else {return "No"}}) + '</td></tr>'
    $htmlBody += '<tr><th width="15%">Script File</th><td>' + $MyInvocation.MyCommand.Definition + '</td></tr>'
    $htmlBody += '<tr><th width="15%">Csv Report File</th><td>' + $outputCSV + '</td></tr>'
    $htmlBody += '<tr><th width="15%">Html Report File</th><td>' + $outputHTML + '</td></tr>'
    $htmlBody += '<tr><th width="15%">Script Version</th><td>' + "<a href=""$($scriptInfo.ProjectURI)"">$($MyInvocation.MyCommand.Definition.ToString().Split("\")[-1].Split(".")[0]) $($scriptInfo.version)</a>" + '</td></tr>'   
    $htmlBody += '</table><hr />'
    #end table SETTINGS
    
    $htmlBody += "</body></html>"

    #Export HTML Report
    $htmlBody | Out-File $outputHTML
    
    #...................................
    #EndRegion HTML
    #...................................

    #...................................
    #Region EMAIL
    #...................................
    if ($sendEmail)
    {
        $mailParams = @{
            From = $sender
            To = $recipients
            Subject = $mailSubject + ": " + ('{0:dd-MMM-yyyy hh:mm tt}' -f $Today)
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
    #...................................
    #EndRegion EMAIL
    #...................................

    #...................................
    #Region TEAMS
    #...................................
    if ($notifyTeams)
    {
        $teamsMessage = ConvertTo-Json -Depth 4 @{
            title = $mailSubject
            text = ('{0:dd-MMM-yyyy hh:mm tt}' -f $Today)
    
            sections = @(
                @{
                    activityTitle = "Delete Files Older Than $($daysToKeep) Days"
                    activityImage = "https://raw.githubusercontent.com/junecastillote/Delete-FilesByAge/master/res/deleteFBAIcon.png"
                    activityText = ""
                },
                @{
                    title = "<h4>Summary</h4>"
                    facts = @(
                        @{
                            name = "Computer:"
                            value = "$($env:COMPUTERNAME)"
                        },
                        @{
                            name = "Paths:"
                            value = ($Paths -join ";<br />")
                        },
                        @{
                            name = "Total Number of Files: "
                            value = "$($summary.TotalNumberOfFiles) files ($($summary.TotalSizeOfAllFiles)) bytes)"
                        },
                        @{
                            name = "Successful Deletion:"
                            value = "<font color=""Green"">$($summary.SuccessfulDeletions) files ($($summary.TotalSuccessfulDeletionSize) bytes)</font>"
                        },
                        @{
                            name = "Failed Deletion:"
                            value = "<font color=""Red"">$($summary.FailedDeletions) files ($($summary.TotalFailedDeletionSize) bytes)</font>"
                        }
                    )
                },
                
                @{
                    title = "<h4>Settings</h4>"
                    facts = @(
                        @{
                            name = "Include:"
                            value = ($Include -join ";")
                        },
                        @{
                            name = "Exclude:"
                            value = ($Exclude -join ";")
                        },
                        @{
                            name = "Recurse:"
                            value = "$($Recurse)"
                        }
                        @{
                            name = "Script File:"
                            value = $MyInvocation.MyCommand.Definition
                        }
                        @{
                            name = "Csv Report File:"
                            value = $outputCSV
                        }
                        @{
                            name = "Html Report File:"
                            value = $outputHTML
                        },
                        @{
                            name = "Script Version:"
                            value = "<a href=""$($scriptInfo.ProjectURI)"">$($MyInvocation.MyCommand.Definition.ToString().Split("\")[-1].Split(".")[0]) $($scriptInfo.version)</a>"
                        }
                    )
                }
            )
        }

        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Sending Teams Notification" -ForegroundColor Green
        
        foreach ($uri in $notifyTeams)
        {
            try {
                Invoke-RestMethod -uri $uri -Method Post -body $teamsMessage -ContentType 'application/json' -ErrorAction Stop
                Write-Host "SUCCESS: $($uri)" -ForegroundColor Green
            }
            catch {
                Write-Host "FAILED: $($_.exception.message)" -ForegroundColor RED
            }
        }
    }   
    #...................................
    #EndRegion TEAMS
    #...................................

    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": HTML Summary Report saved in $outputHTML " -ForegroundColor Cyan
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": CSV Summary Report saved in $outputCSV " -ForegroundColor Cyan
}
else 
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": No files to delete. Exiting script" -ForegroundColor Green
}
#...................................
#EndRegion FILE DELETION
#...................................


if ($logDirectory) {Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Transcript Log saved in $logfile " -ForegroundColor Cyan}
Stop-TxnLogging
