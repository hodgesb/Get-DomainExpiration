<#
.SYNOPSIS
    Gets the expiration dates of domain names and generates a report.

.DESCRIPTION
    Uses WhoisCL to query the expiration dates of domains from a specified list. Categorizes domains and sends an email report with this information.

.PARAMETER DomainList
    Path to the file containing the list of domains. Default is 'domains.txt' in the script's directory.

.PARAMETER WhoisCLIPath
    Path to the WhoisCL executable. Default is 'WhoisCL.exe' in the script's directory.

.PARAMETER EmailTo
    The recipient email address for the report.

.PARAMETER EmailFrom
    The sender email address for the report. 

.PARAMETER EmailSMTP
    The SMTP server to use for sending the email.

.PARAMETER EmailSubject
    The subject of the email report. Default includes the current date.

.EXAMPLE
    PS> Get-DomainExpiration
    Runs the function using the default domain list and WhoisCL location, with default email parameters configured in the script.

.EXAMPLE
    PS> Get-DomainExpiration -EmailTo IT@example.com -EmailFrom ScriptServer@example.com -EmailSMTP '127.0.0.1' -verbose
    Runs the function with email parameters and outputs verbose logging to console.

.NOTES
    Created By: Bobby Hodges
    Date: 5/2/2023
    Ticket Reference: WRNS-4745
    Requires: WhoisCL (https://www.nirsoft.net/utils/whoiscl.html)

.LINK
    https://www.nirsoft.net/utils/whoiscl.html - WhoisCL Tool
#>

function Get-DomainExpiration {
    [CmdletBinding()]
    param(
        [string]$DomainList = (Join-Path -Path $PSScriptRoot -ChildPath "domains.txt"),
        [string]$WhoisCLIPath = (Join-Path -Path $PSScriptRoot -ChildPath "WhoisCL.exe"),
        [string]$EmailTo = 'it@example.com',
        [string]$EmailFrom = "$($env:computername)@$($env:USERDNSDOMAIN)",
        [string]$EmailSMTP = 'smtprelay1.mydomain.local',
        [string]$EmailSubject = "Domain Expiration Report - $(Get-Date -Format 'MM-dd-yyyy')"
    )

    Begin {
        # Check if WhoisCL.exe exists
        if (!(Test-Path $WhoisCLIPath)) { 
            Write-Error "WhoisCL.exe is missing from path: $WhoisCLIPath"
            exit
        }

        # Check if domain list file exists
        if (!(Test-Path $DomainList)) { 
            Write-Error "Domain list file is missing from specified path: $DomainList"
            exit
        }

        # Initialize variables
        $today = Get-Date
        $resultsAll = @{}
        $resultsExpiring = @{}
        $resultsExpired = @{}
    }

    Process {
        # Read the domain list
        $domains = Get-Content -Path $DomainList -ErrorAction Stop

        foreach ($domain in $domains) {
            $attemptCount = 0
            $successful = $false

            while ($attemptCount -lt 3 -and -not $successful) {
                try {
                    $expiration = & $WhoisCLIPath -r $domain | Select-String -Pattern "Expir" -AllMatches
                    $datematch = [regex]::Match($expiration, '\d{4}-\d{2}-\d{2}')
                
                    if ($datematch) {
                        $expiryDate = [datetime]::ParseExact($datematch.Value, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
                        $expiryDaysLeft = ($expiryDate - $today).Days
                        $expiryDateMMDDYYYY = $expiryDate.ToString("MM-dd-yyyy")

                        $resultsAll[$domain] = "Expires in $expiryDaysLeft days - $expiryDateMMDDYYYY"
                        Write-Verbose "$domain expires in $expiryDaysLeft days"

                        if ($expiryDaysLeft -le 90) { $resultsExpiring[$domain] = "Expires in $expiryDaysLeft days - $expiryDateMMDDYYYY" }
                        if ($expiryDaysLeft -le 0) { $resultsExpired[$domain] = "Expired on $expiryDateMMDDYYYY" }

                        $successful = $true
                    } else {
                        $attemptCount++
                        Start-Sleep -Seconds 1
                    }
                } catch {
                    Write-Verbose "Attempt $attemptCount failed for $domain"
                    $attemptCount++
                    if ($attemptCount -lt 3) {
                        Start-Sleep -Seconds 2
                    } else {
                        Write-Output "ERROR returned for $domain after 3 attempts"
                        $resultsAll[$domain] = "Domain not found or query failed for this domain after 3 attempts"
                    }
                }
            }
        }
    }


    End {
        # Formatting for email
        $format= @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 2px; padding: 5px; border-style: solid; border-color: black; text-align: left}
TD {border-width: 1px; padding: 5px; border-style: solid; border-color: black; text-align: left}
</style>
"@

        $resultsAllHTML = $resultsAll.GetEnumerator() | Sort-Object -Property Name | ConvertTo-Html -Property Name,Value -Head $format | Out-String
        $resultsExpiringHTML = $resultsExpiring.GetEnumerator() | Sort-Object -Property Name | ConvertTo-Html -Property Name,Value -Head $format | Out-String
        $resultsExpiredHTML = $resultsExpired.GetEnumerator() | Sort-Object -Property Name | ConvertTo-Html -Property Name,Value -Head $format | Out-String

        $emailBody = @"
<h1>Domain Registration Report - $($today.ToString("MM-dd-yyyy"))</h1>
<h2>Expiring Domains</h2>
$resultsExpiringHTML
Total: $($resultsExpiring.count)
<BR>
<h2>Expired Domains</h2>
$resultsExpiredHTML
Total: $($resultsExpired.count)
<BR>
<h2>All Domains</h2>
$resultsAllHTML
Total: $($resultsAll.count)
"@

        # Send an email
        try {
            Send-MailMessage -From $EmailFrom -To $EmailTo -Subject $EmailSubject -SmtpServer $EmailSMTP -BodyAsHTML -Body $emailBody
        } catch {
            Write-Error "Failed to send email: $_"
        }
    }
}
