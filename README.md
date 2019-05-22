<h3>
Introduction</h3>
This script can be used to perform file deletions based on age in days, and with the option to send an email summary report. In my particular case, this is deployed to performed housekeeping of IIS and Exchange Server Logs older than seven (7) days. This is a fork of the previous script published as <a href="https://www.lazyexchangeadmin.com/2016/09/DeleteFilesOlderThanXDays.html" target="_blank">DeleteFilesOlderThanXDays.ps1</a>

<h3>
Requirements</h3>
<div>
<ul>
<li>PowerShell 5.1 (recommended), but works and tested with PowerShell 4.</li>
<li>If "email summary" report will be enabled, an SMTP relay is required.</li>
</ul>
</div>
<div>

</div>
<h3>
Script Installation</h3>
<div>
<b>From GitHub</b>
<b>
</b> To get the script, you can clone or download from this link - <a href="https://github.com/junecastillote/Delete-FilesByAge" target="_blank">Delete-FilesByAge.ps1</a>&nbsp;and save it on your preferred path.

<b>From PowerShell Gallery</b>

If you have PowerShell 5.1, you can simply install from PSGallery.

<a href="https://www.powershellgallery.com/packages/Delete-FilesByAge">https://www.powershellgallery.com/packages/Delete-FilesByAge</a>

<table align="center" cellpadding="0" cellspacing="0" class="tr-caption-container" style="float: left; margin-right: 1em; text-align: left;"><tbody>
<tr><td style="text-align: center;"><a href="https://4.bp.blogspot.com/-xoBY5UYHHpE/XOTVboNqggI/AAAAAAAAGYs/F1hjKWRLVLYz6gl_yGCqCI3UpgCNrvQBgCLcBGAs/s1600/mRemoteNG_2019-05-22_12-51-36.png" imageanchor="1" style="clear: left; display: inline !important; margin-bottom: 1em; margin-left: auto; margin-right: auto; text-align: center;"><img border="0" data-original-height="41" data-original-width="447" src="https://4.bp.blogspot.com/-xoBY5UYHHpE/XOTVboNqggI/AAAAAAAAGYs/F1hjKWRLVLYz6gl_yGCqCI3UpgCNrvQBgCLcBGAs/s1600/mRemoteNG_2019-05-22_12-51-36.png" /></a></td></tr>
<tr><td class="tr-caption" style="text-align: center;">Install Script from PSGallery</td></tr>
</tbody></table>

</div>
<div>




<table cellpadding="0" cellspacing="0" class="tr-caption-container" style="float: left; margin-right: 1em; text-align: left;"><tbody>
<tr><td style="text-align: center;"><a href="https://4.bp.blogspot.com/-XDPFa1vpWY8/XOTXRZcbJHI/AAAAAAAAGY4/lHuNkw6ew_A-YPzxRu-MBtnqHNeMp8PbwCLcBGAs/s1600/mRemoteNG_2019-05-22_12-58-31.png" imageanchor="1" style="clear: left; margin-bottom: 1em; margin-left: auto; margin-right: auto;"><img border="0" data-original-height="400" data-original-width="834" src="https://4.bp.blogspot.com/-XDPFa1vpWY8/XOTXRZcbJHI/AAAAAAAAGY4/lHuNkw6ew_A-YPzxRu-MBtnqHNeMp8PbwCLcBGAs/s1600/mRemoteNG_2019-05-22_12-58-31.png" /></a></td></tr>
<tr><td class="tr-caption" style="text-align: center;">Confirmed that the script was installed</td></tr>
</tbody></table>

</div>
<h3>
How to Use</h3>
<h4>
Required Parameters</h4>
<b><u>Paths</u></b>
<b><u>
</u></b> This accepts the directory or path that will be the target of the cleanup. It accepts multiple entries separated by a comma. The paths can be a local or UNC path location.
<b>Example:</b> "C:\TEMP","\\SERVER\C$\TEMP"

<b><u>outputDirectory</u></b>
<b><u>
</u></b> Indicates the path where the output files will be saved.
<b>Example:</b> "C:\Scripts\Output"

<b><u>startDate</u></b>
<b><u>
</u></b> Indicate the oldest or earliest date of the files to include in the deletion
Example: "May 1, 2019" or "05/01/2019 2:00:00 AM"

<b><u>endDate</u></b>
<b><u>
</u></b> Indicate the newest or latest date of the files to include in the deletion
<b>Example:</b> "May 1, 2019" or "05/01/2019 2:00:00 AM"

<h4>
Optional Parameters</h4>
<div>
<b><u>logDirectory</u></b></div>
<div>
<b><u>
</u></b></div>
<div>
<div>
Indicates the path where the transcript log files will be saved.</div>
<div>
<b>Example:</b> "C:\Scripts\Logs"</div>
</div>
<div>

</div>
<b><u>include</u></b>
<b><u>
</u></b> Single or List of files or file extensions to include in the deletion.
<b>Example:</b> "*.log","file1.txt",*.blg"

<b><u>exclude</u></b>
<b><u>
</u></b> Single or List of files or file extensions to exclude from the deletion.
<b>Example:</b> "*.log","file1.txt",*.blg"

<b><u>recurse</u></b>
<b><u>
</u></b> Switch to indicate whether or not a recursive deletion will be performed. No value is required for this parameter.

<b><u>sendEmail</u></b>
<b><u>
</u></b> Switch to indicate whether or not to send the HTML summary report through email. No value is required for this parameter.

<b><u>sender</u></b>
<b><u>
</u></b> Sender Email Address to be used for the Summary Report. Required if <b>sendEmail </b>switch is used.
<b>Example:</b> sender@domain.com OR "Sender Display Name <sender domain.com="">"</sender>

<b><u>recipients</u></b>
<b><u>
</u></b> List of recipients of the email report. Multiple recipients are accepted separated with a comma. Required if&nbsp;<b>sendEmail&nbsp;</b>switch is used.
<b>Example</b>: "recipient1@domain.com","recipient2@domain.com"

<b><u>smtpServer</u></b>
<b><u>
</u></b> The SMTP Relay Server. This can be an IP Address, Hostname or FQDN. Required if&nbsp;<b>sendEmail&nbsp;</b>switch is used.
<b>Example:</b> 192.168.0.6
<b>Example:</b> smtp.server.com

<b><u>smtpPort</u></b>
<b><u>
</u></b> SMTP Server Port to use. This is normally port 25 for an unauthenticated relay, or 587 for Office 365 authenticated relay. Required if&nbsp;<b>sendEmail&nbsp;</b>switch is used.
<b>Example:</b>&nbsp;587

<b><u>smtpSSL</u></b>

Switch to indicate whether or not SSL will be used for the SMTP relay connection. No value is required for this parameter.

<b><u>smtpCredential</u></b>

Used if SMTP relay (eg. Office 365) requires authentication. Required if&nbsp;<b>sendEmail&nbsp;</b>switch is used.

<h3>
Example</h3>
<div>
In this example:</div>
<div>
<ul>
<li>Using Office 365 as SMTP Relay, with Authentication, using Port 587 with SSL enabled.</li>
<li>Stored Authentication credential in variable $credential</li>
<li>Included *.log,*.blg extensions.</li>
<li>Excluded log1.log</li>
<li>Only selected files modified between "May 1, 2019"&nbsp;and "May 29, 2019".</li>
<li>Selected Paths "\\poshlabexc\c$\inetpub\logs","c:\temp"</li>
</ul>
<div>
Command:</div>
</div>
<blockquote class="tr_bq">
.\Delete-FilesByAge.ps1 -Paths "\\poshlabexc\c$\inetpub\logs","c:\temp" -Include *.log,*.blg -Recurse -startDate "May 1, 2019" -endDate "May 29, 2019" -outputDirectory "C:\Scripts\Delete-FilesByAge\output" -headerPrefix LazyExchangeAdmin -sendEmail -sender postmaster@lazyexchangeadmin.com -recipients june@lazyexchangeadmin.com -smtpServer smtp.office365.com -smtpPort 587 -smtpCredential $credential -smtpSSL -logDirectory "C:\Scripts\Delete-FilesByAge\Log" -Exclude log1.log</blockquote>
<div>

</div>
<div>
<div class="separator" style="clear: both; text-align: center;">
<a href="https://2.bp.blogspot.com/-3-m_VXQD1EY/XOTKVFfjqII/AAAAAAAAGYE/Q20IzKTH3QoYReKDG3YKWORZets6Wqd9gCLcBGAs/s1600/mRemoteNG_2019-05-22_11-49-21.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" data-original-height="357" data-original-width="444" src="https://2.bp.blogspot.com/-3-m_VXQD1EY/XOTKVFfjqII/AAAAAAAAGYE/Q20IzKTH3QoYReKDG3YKWORZets6Wqd9gCLcBGAs/s1600/mRemoteNG_2019-05-22_11-49-21.png" /></a></div>

<div class="separator" style="clear: both; text-align: center;">
<a href="https://3.bp.blogspot.com/-d_-GPA6WueQ/XOTKZ3zJajI/AAAAAAAAGYI/P3FgQyRp9lARN_InVgsRZw58789Uov69wCLcBGAs/s1600/mRemoteNG_2019-05-22_11-52-47.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" data-original-height="377" data-original-width="861" src="https://3.bp.blogspot.com/-d_-GPA6WueQ/XOTKZ3zJajI/AAAAAAAAGYI/P3FgQyRp9lARN_InVgsRZw58789Uov69wCLcBGAs/s1600/mRemoteNG_2019-05-22_11-52-47.png" /></a></div>

<div class="separator" style="clear: both; text-align: center;">
<a href="https://2.bp.blogspot.com/-hl33yfWk6tE/XOTKeeccgpI/AAAAAAAAGYM/D5YuH1OkXdIfXIwUeJb6bUfO5wAr7DHUACLcBGAs/s1600/mRemoteNG_2019-05-22_11-53-01.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" data-original-height="242" data-original-width="859" src="https://2.bp.blogspot.com/-hl33yfWk6tE/XOTKeeccgpI/AAAAAAAAGYM/D5YuH1OkXdIfXIwUeJb6bUfO5wAr7DHUACLcBGAs/s1600/mRemoteNG_2019-05-22_11-53-01.png" /></a></div>

<h3>
Sample Output</h3>

<div class="separator" style="clear: both; text-align: center;">
<a href="https://4.bp.blogspot.com/-ozsXzUhxbmw/XOTRHUA5EKI/AAAAAAAAGYg/cJPYu3AvFfA6ttPavcpK4loaSatSZNYnwCLcBGAs/s1600/mRemoteNG_2019-05-22_12-32-55.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" data-original-height="719" data-original-width="675" src="https://4.bp.blogspot.com/-ozsXzUhxbmw/XOTRHUA5EKI/AAAAAAAAGYg/cJPYu3AvFfA6ttPavcpK4loaSatSZNYnwCLcBGAs/s1600/mRemoteNG_2019-05-22_12-32-55.png" /></a></div>


</div>
