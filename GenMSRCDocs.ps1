## Uploaded by @JohnLaTwC
## Miss security bulletins?  Create them yourself in a few lines of PowerShell:

## First, get an API key to the MSRC Portal API 
## Sign-in in here: https://portal.msrc.microsoft.com/en-us/developer, and click on the Developer tab, click the Show button on the API key.

## Install the MSRC PowerShell cmdlets, Run in an Admin PowerShell:
## Install-Module -Name MSRCSecurityUpdates -force



## Added Tomoyuki Matsuda, toocheap [at] gmail.com 10 May 2017
# To run properly on the different culture environment. Set Culture to "en-US" explicitly.
$currentThread = [System.Threading.Thread]::CurrentThread
$culture = [System.Globalization.CultureInfo]::CreateSpecificCulture("en-US")
$currentThread.CurrentCulture = $culture
$currentThread.CurrentUICulture = $culture
##

## In a normal user PowerShell:
Import-Module MSRCSecurityUpdates -Verbose:$false
Set-MSRCApiKey -ApiKey "183820dd115c415a8e9a4b1d048e9267" 
$month = Get-Date -format M |% {$_.Split(' ')[0]} 
$year = Get-Date -format yyyy
$timeperiod = $year + '-' + $month
$fname = 'MSRCSecurityUpdates' + $timeperiod + '.html'
Get-MsrcCvrfDocument -ID $timeperiod | Get-MsrcSecurityBulletinHtml | Out-File $fname
Invoke-Item $fname
