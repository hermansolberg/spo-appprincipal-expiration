class AppPrincipal
{
    [String]$DisplayName
    [String]$ObjectId
    [String]$AppId
    [DateTime]$ExpiryDate
    [int]$DaysUntilExpiry
    [Bool]$HasExpired
}

class AppPrincipalCollection
{
    hidden [Int32] $CriticalLevel = 30
    hidden [Int32] $WarningLevel = 60
    hidden [Int32] $InformationLevel = 100
    hidden [AppPrincipal[]] $AppPrincipals = @()

    [Int32] $ExpiredCount = 0
    [Int32] $CriticalCount = 0
    [Int32] $WarningCount = 0
    [Int32] $InformationCount = 0
    [Int32] $TotalCount = 0
    [String[]] $ClientIds = @()

    AppPrincipalCollection()
    {
        $this.SetLevels()
        $this.GetClientIds()
    }

    AppPrincipalCollection([Int32]$criticalDays, [Int32]$warningDays, [Int32]$informationDays)
    {
        $this.CriticalLevel = $criticalDays
        $this.WarningLevel = $warningDays
        $this.InformationLevel = $informationDays
    }

    hidden GetClientIds()
    {
        $AppPrincipalIdVariable = Get-AutomationVariable -Name "SPOAppPrincipalIdentifiers"

        if($AppPrincipalIdVariable -ne $null)
        {
            $this.ClientIds = $AppPrincipalIdVariable.Split(";")
        }

        if($this.ClientIds.Count -eq 0)
        {
            Write-Error "No App Principal Client Ids were found. Please configure the variable SPOAppPrincipalIdentifiers in your Azure Automation Account" -ErrorAction Stop
        }
    }

    hidden SetLevels()
    {
        $Information = Get-AutomationVariable -Name "SPOAppPrincipalInformationLevel" -ErrorAction SilentlyContinue
        $Warning = Get-AutomationVariable -Name "SPOAppPrincipalWarningLevel" -ErrorAction SilentlyContinue
        $Critical = Get-AutomationVariable -Name "SPOAppPrincipalCriticalLevel" -ErrorAction SilentlyContinue

        if($Information -ne $null)
        {
            $this.InformationLevel = $Information
        }

        if($Warning -ne $null)
        {
            $this.WarningLevel = $Warning
        }

        if($Critical -ne $null)
        {
            $this.CriticalLevel = $Critical
        }
    }

    hidden [Bool] IsExpired([AppPrincipal]$appPrincipal)
    {
        return ($appPrincipal.DaysUntilExpiry -lt 0)
    }

    hidden [Bool] IsCritical([AppPrincipal]$appPrincipal)
    {
        return ($appPrincipal.DaysUntilExpiry -le $this.CriticalLevel -and $_.DaysUntilExpiry -gt 0)
    }

    hidden [Bool] IsWarning([AppPrincipal]$appPrincipal)
    {
        return ($appPrincipal.DaysUntilExpiry -le $this.WarningLevel -and $appPrincipal.DaysUntilExpiry -gt $this.CriticalLevel)
    }

    hidden [Bool] IsInformation([AppPrincipal]$appPrincipal)
    {
        return ($appPrincipal.DaysUntilExpiry -le $this.InformationLevel -and $appPrincipal.DaysUntilExpiry -gt $this.WarningLevel)
    }

    Add([AppPrincipal]$appPrincipal)
    {
        if($appPrincipal.HasExpired)
        {
            $this.ExpiredCount++
            $this.TotalCount++
        }
        elseif($this.IsCritical($appPrincipal))
        {
            $this.CriticalCount++
            $this.TotalCount++
        }
        elseif($this.IsWarning($appPrincipal))
        {
            $this.WarningCount++
            $this.TotalCount++
        }
        elseif($this.IsInformation($appPrincipal))
        {
            $this.InformationCount++
            $this.TotalCount++
        }

        $this.AppPrincipals += $appPrincipal
    }

    [AppPrincipal[]]GetExpiredAppPrincipals()
    {
        return $this.AppPrincipals | Where { $_.HasExpired }
    }
    [AppPrincipal[]]GetCriticalAppPrincipals()
    {
        return $this.AppPrincipals | Where { $_.DaysUntilExpiry -le $this.CriticalLevel -and $_.DaysUntilExpiry -gt 0 }
    }

    [AppPrincipal[]]GetWarningAppPrincipals()
    {
        return $this.AppPrincipals | Where { $_.DaysUntilExpiry -le $this.WarningLevel -and $_.DaysUntilExpiry -gt $this.CriticalLevel }
    }

    [AppPrincipal[]]GetInformationAppPrincipals()
    {
        return $this.AppPrincipals | Where { $_.DaysUntilExpiry -le $this.InformationLevel -and $_.DaysUntilExpiry -gt $this.WarningLevel }
    }
}

class MonitoringSettings
{
    [String] $SmtpAddress = $null
    [String] $EmailFrom = $null
    [String] $EmailTo = $null
    [Int32] $SmtpPort = 25

    MonitoringSettings()
    {
        $this.SetSmtp()
        $this.SetEmailFrom()
        $this.SetEmailTo()
    }

    hidden SetSmtp()
    {
        $this.SmtpAddress = Get-AutomationVariable -Name "SPOMonitoringSmtp" -ErrorAction SilentlyContinue
    }

    hidden SetEmailFrom()
    {
        $this.EmailFrom = Get-AutomationVariable -Name "SPOMonitoringEmailFrom" -ErrorAction SilentlyContinue

        if([String]::IsNullOrEmpty($this.EmailFrom))
        {
            $this.EmailFrom = $this.Credentials.UserName
        }
    }

    hidden SetEmailTo()
    {
        $this.EmailTo = Get-AutomationVariable -Name "SPOMonitoringEmailTo" -ErrorAction SilentlyContinue
    }
}

class ReportingManager
{
    hidden [MonitoringSettings] $Settings = $null

    ReportingManager([MonitoringSettings]$settings)
    {
        $this.Settings = $settings
    }

    [Boolean] SendReport([AppPrincipalCollection]$data)
    {
        if($this.ReportByEmail() -and ($data.TotalCount -gt 0))
        {
            $body = $this.GetEmailBody($data)
            $parameters = $this.GetEmailReportParameters($body)
            Send-MailMessage @parameters

            return $True
        }

        return $False       
    }

    hidden [Boolean] ReportByEmail()
    {
        if([String]::IsNullorEmpty($this.Settings.SmtpAddress))
        {
            return $False
        }
        elseif([String]::IsNullorEmpty($this.Settings.EmailTo))
        {
            return $False
        }
        else
        {
            return $True
        }
    }

    hidden [String] GetEmailBody([AppPrincipalCollection]$data)
    {
        $Builder = [System.Text.StringBuilder]::new()
        $Builder.AppendLine($this.GetEmailHtmlBegin())
        $Builder.AppendLine($this.GetEmailHtmlHead())
        $Builder.AppendLine($this.GetEmailHtmlBody($data))
        $Builder.AppendLine($this.GetEmailHtmlEnd())      

        return $Builder.ToString()
    }

    hidden [String] GetEmailHtmlBegin()
    {
        $Builder = [System.Text.StringBuilder]::new()
        $Builder.AppendLine('<!DOCTYPE html>')
        $Builder.AppendLine('<html>')
        return $Builder.ToString()
    }

    hidden [String] GetEmailHtmlHead()
    {
        $Builder = [System.Text.StringBuilder]::new()
        $Builder.AppendLine('<head>')
        $Builder.AppendLine('<meta charset="UTF-8">')
        $Builder.AppendLine($this.GetEmailHtmlStyle())
        $Builder.AppendLine('</head>')
        return $Builder.ToString()
    }

    hidden [String] GetEmailHtmlStyle()
    {
        $Builder = [System.Text.StringBuilder]::new()
        $Builder.AppendLine('<style>')
        $Builder.AppendLine('table { border: 1px solid grey; text-align: left; }')
        $Builder.AppendLine('td,th { border: 1px solid grey; padding: 5px; margin: 2px; }')
        $Builder.AppendLine('</style>')
        return $Builder.ToString()
    }

    hidden [String] GetEmailHtmlBody([AppPrincipalCollection]$data)
    {
        $Builder = [System.Text.StringBuilder]::new()
        $Builder.AppendLine('<body>')
        $Builder.AppendLine($this.GetEmailHtmlBodyIntro())

        if($data.ExpiredCount -gt 0)
        {
            $Builder.AppendLine($this.GetEmailHtmlBodySection("Expired", "The following app principals have expired.", $data.GetExpiredAppPrincipals()))
        }
        
        if($data.CriticalCount -gt 0)
        {
            $Builder.AppendLine($this.GetEmailHtmlBodySection("Critical", "The following app principals are considered as critical.", $data.GetCriticalAppPrincipals()))
        }

        if($data.WarningCount -gt 0)
        {
            $Builder.AppendLine($this.GetEmailHtmlBodySection("Warning", "The following app principals are considered as warning."), $data.GetWarningAppPrincipals())
        }
        
        if($data.InformationCount -gt 0)
        {
            $Builder.AppendLine($this.GetEmailHtmlBodySection("Information", "The following app principals are considered as information.", $data.GetInformationAppPrincipals()))
        }        
        
        $Builder.AppendLine('</body>')
        return $Builder.ToString()
    }

    hidden [String] GetEmailHtmlBodyIntro()
    {
        $Builder = [System.Text.StringBuilder]::new()
        $Builder.AppendLine('<h1>SharePoint Online app principal expiration report</h1>')
        $Builder.AppendLine('<p>This report contains a list of app principals that are within one of the defined thresholds of expiration</p>')
        $Builder.AppendLine('<p>Read more about of <a href="http://aka.ms/spoappprincipalexpiration">app principal expiration</a></p>')
        return $Builder.ToString()
    }
    
    hidden [String] GetEmailHtmlBodySection([String]$heading, [String]$description, [AppPrincipal[]]$data)
    {
        $Builder = [System.Text.StringBuilder]::new()
        $Builder.AppendFormat('<h2>{0}</h2>{1}', $heading, [Environment]::NewLine)
        $Builder.AppendFormat('<p>{0}</p>{1}', $description, [Environment]::NewLine)
        $Builder.AppendLine('<table>')
        $Builder.AppendLine($this.GetEmailHtmlBodyTableRow("th", @("Display Name", "Object ID", "Client ID", "Expires in days", "Expiry date")))

        foreach($entry in $data)
        {
            $Builder.AppendLine($this.GetEmailHtmlBodyTableRow("td", @($entry.DisplayName, $entry.ObjectId, $entry.AppId, $entry.DaysUntilExpiry, $entry.ExpiryDate.ToLongDateString())))
        }

        $Builder.AppendLine('</table>')
        return $Builder.ToString()
    }

    hidden [String] GetEmailHtmlBodyTableRow([String]$tagName, [String[]]$cellText)
    {
        $Builder = [System.Text.StringBuilder]::new()        
        $Builder.AppendLine('<tr>')
        foreach($entry in $cellText)
        {
            $Builder.AppendFormat('<{0}>{1}</{0}>{2}', $tagName, $entry, [Environment]::NewLine)
        }         
        $Builder.AppendLine('</tr>')
        return $Builder.ToString()
    }

    hidden [String] GetEmailHtmlEnd()
    {
        $Builder = [System.Text.StringBuilder]::new()
        $Builder.AppendLine('</html>')
        return $Builder.ToString()
    }

    hidden [HashTable] GetEmailReportParameters([String]$body)
    {
        $parameters = @{}
        $parameters.Add("To", $this.Settings.EmailTo)
        $parameters.Add("From", $this.Settings.EmailFrom)
        $parameters.Add("Subject", "SharePoint Online App Principal Expiration Report")
        $parameters.Add("SmtpServer", $this.Settings.SmtpAddress)
        $parameters.Add("Port", $this.Settings.SmtpPort)
        $parameters.Add("UseSsl", $True)
        $parameters.Add("BodyAsHtml", $True)
        $parameters.Add("Body", $body)

        return $parameters
    }
}

$Today = Get-Date
$AppPrincipals = [AppPrincipalCollection]::new()
$Settings = [MonitoringSettings]::new()
$ReportManager = [ReportingManager]::new($Settings)
$Credentials = Get-AutomationPSCredential -Name "AzureADMonitoringAccount" -ErrorAction SilentlyContinue
$Connection = Get-AutomationConnection -Name "AzureRunAsConnection" -ErrorAction SilentlyContinue

if($Credentials -eq $null)
{
    Write-Output "Using AzureRunAsConnection to connect to AzureAD"
    Connect-AzureAD -TenantId $Connection.TenantID -ApplicationId $Connection.ApplicationID -CertificateThumbprint $Connection.CertificateThumbprint -ErrorAction Stop | Out-Null
}
else
{
    Write-Output "Using AzureADMonitoringAccount to connect to AzureAD"
    Connect-AzureAD -Credential $Credentials -ErrorAction Stop | Out-Null
}

$ServicePrincipals = Get-AzureADServicePrincipal -All:$true | Where-Object { $AppPrincipals.ClientIds.Contains($_.AppId) }

foreach($SvcPrincipal in $ServicePrincipals)
{
    # Get the current KeyCredentials and PasswordCredentials for the ServicePrincipal
    $Verify = Get-AzureADServicePrincipalKeyCredential -ObjectId $SvcPrincipal.ObjectId | Sort EndDate -Descending | Where { $_.Usage -eq "Verify" }
    $Sign = Get-AzureADServicePrincipalKeyCredential -ObjectId $SvcPrincipal.ObjectId | Sort EndDate -Descending | Where { $_.Usage -eq "Sign" }
    $Pass = Get-AzureADServicePrincipalPasswordCredential -ObjectId $SvcPrincipal.ObjectId | Sort EndDate -Descending
    
    $ExpiryDates = @()
    $ExpiryDates += $Verify.EndDate
    $ExpiryDates += $Sign.EndDate
    $ExpiryDates += $Pass.EndDate
    
    # Create a custom AppPrincipal object to represent each SPO app that should be evaluated
    $AppPrincipal = [AppPrincipal]::new()
    $AppPrincipal.ObjectId = $SvcPrincipal.ObjectId
    $AppPrincipal.DisplayName = $SvcPrincipal.DisplayName
    $AppPrincipal.AppId = $SvcPrincipal.AppId
    $AppPrincipal.ExpiryDate = $ExpiryDates | Sort -Descending | Select -First 1
    $AppPrincipal.DaysUntilExpiry = $AppPrincipal.ExpiryDate.Subtract($Today).Days
    $AppPrincipal.HasExpired = ($AppPrincipal.DaysUntilExpiry -lt 0)

    $AppPrincipals.Add($AppPrincipal)
}

Disconnect-AzureAD

$Result = $ReportManager.SendReport($AppPrincipals)

Write-Output $AppPrincipals | Select ExpiredCount,CriticalCount,WarningCount,InformationCount,TotalCount