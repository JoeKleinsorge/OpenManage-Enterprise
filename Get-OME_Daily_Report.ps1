<#
.SYNOPSIS
    A basic daily health/environment report using OME data. Best run as a scheduled task.
.DESCRIPTION
    Uses OpenManage Enterprise API to pull data, formats into HTML tables and emails report.
.EXAMPLE
    Get-OME_Daily_Report -OME 10.0.0.1 -To team@email.com -From OME@email.com -STMP stmp.company.com 
.INPUTS
    - OMEIP (IP address for OME server)
    - Credentials (OME user creds)
    - To (Recieving Email address)
    - From (Sending Email address)
    - SMTP (SMTP address)
    - OpenCasesPath (Path to OpenCases CSV)
.OUTPUTS
    Sends an HTML email report to the $to address.
.NOTES
    Author - Joey Kleinsorge
#>

[CmdletBinding()]
param (
    # OMEnterprise IP address
    [Parameter(Mandatory)] 
    [System.Net.IPAddress] $OMEIP,

    # OMEnterprise Credenials
    [Parameter(Mandatory)]
    [pscredential] $Credentials,
    
    # Email of recipent
    [Parameter(Mandatory)]
    [string] $To,

    # Email of sender
    [Parameter(Mandatory)]
    [string] $From,

    # STMP Address
    [Parameter(Mandatory)]
    [string] $STMP,

    # OpenCases.CSV Path
    [Parameter(Mandatory)]
    [string] $OpenCasesPath

)

begin {
    function Ignore-SSLCertificates {
        $Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
        $Compiler = $Provider.CreateCompiler()
        $Params = New-Object System.CodeDom.Compiler.CompilerParameters
        $Params.GenerateExecutable = $false
        $Params.GenerateInMemory = $true
        $Params.IncludeDebugInformation = $false
        $Params.ReferencedAssemblies.Add("System.DLL") > $null
        $TASource = @'
        namespace Local.ToolkitExtensions.Net.CertificatePolicy
        {
            public class TrustAll : System.Net.ICertificatePolicy
            {
                public bool CheckValidationResult(System.Net.ServicePoint sp,System.Security.Cryptography.X509Certificates.X509Certificate cert, System.Net.WebRequest req, int problem)
                {
                    return true;
                }
            }
        }
'@ 
        $TAResults = $Provider.CompileAssemblyFromSource($Params, $TASource)
        $TAAssembly = $TAResults.CompiledAssembly
        $TrustAll = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
        [System.Net.ServicePointManager]::CertificatePolicy = $TrustAll
    }

    #_Dependencies 
    Ignore-SSLCertificates
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::TLS12
    [System.Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient") | out-null

    #_Setup Vars
    $cores = 0
    $memory = 0

    #_Email settings
    $subject = "Daily Hardware Report " + $(get-date).ToShortDateString()
    $SmtpClient = new-object system.net.mail.smtpClient
    $MailMessage = New-Object system.net.mail.mailmessage

    #_Import CSV
    If ($OpenCasesPath) {
        $OpenCases = Import-Csv -Path $OpenCasesPath
    }
}

process {
    #Region Find Report Ids ( Cause they change with every update)
    $URI = "https://$OMEIP" + '/api/ReportService/ReportDefs?$top=100'
    $Response = Invoke-WebRequest -Uri $URI -Method GET -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop 
    $data = $response.Content | ConvertFrom-Json
    $values = $data.value
    ForEach ( $value in $values) {
        $temp = $value
        if ($temp.Name -match "Server Overview Report") { $overviewReportId = $temp.Id }
        if ($temp.Name -match "Warranties Expiring in Next 30 Days") { $warrantyReportId = $temp.Id }
    }
    #EndRegion

    #Region Get Alerts
    #_Get Alerts from today and yesterday
    $URI = "https://$OMEIP" + '/api/AlertService/Alerts?$top=10000'
    $response = Invoke-restmethod -Uri $URI -Credential $credential -Contenttype "application/xml" -ErrorAction:Stop

    #_Filter data from call
    $Alerts = $response.value | Select-Object AlertDeviceName, AlertDeviceIdentifier, SubCategoryName, CategoryName, TimeStamp, Message, RecommendedAction, SeverityType
    $date1 = get-date -format u
    $today = $date1.substring(0, 10)
    $date2 = (get-date (get-date).addDays(-1) -UFormat "%Y-%m-%d")
    $yesterday = $date2.substring(0, 10)
    $OMEAlerts = $Alerts | Where-Object { $_.AlertDeviceIdentifier -and $_.SeverityType -eq 16 -and ($_.TimeStamp -match $today -or $_.TimeStamp -match $yesterday) } | Sort-Object -unique AlertDeviceName
    $Alertcount = ($omealerts | measure-object).count

    #EndRegion

    #Region Check open cases
    if ($OpenCases) {
        foreach ($case in $Opencases) {
            $ServiceTag = $case.ServiceTag
            $Severity = $case.Severity
            $Description = $case.Description

            $datarow = @"
            <tr style="height: 32px;">
            <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$ServiceTag</span></span></td>
            <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$Severity</span></span></td>
            <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$Description</span></span></td>
            </tr>
"@
            $CaseReport += $datarow
        }
    }
    #EndRegion
        
    #Region Get OME Environment Info
    #_Run Env OME Report
    $body = @{
        "ReportDefId"   = $overviewReportId
        "FilterGroupId" = 0
    }
    $body = $body | ConvertTo-Json
    $URI = "https://$OMEIP/api/ReportService/Actions/ReportService.RunReport"
    $Response = Invoke-WebRequest -Uri $URI -Method POST -Body $body -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop 

    #_Get Env Report results
    $URI = "https://$OMEIP/api/ReportService/ReportDefs($overviewReportId)" + '/ReportResults/ResultRows?$top=10000'
    $Response = Invoke-WebRequest -Uri $URI -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
    $Envdata = $response.Content | ConvertFrom-Json
    $Envvalues = $Envdata.value

    #_Parse Env results
    $NumberOfServers = $Envdata."@odata.count"
    ForEach ( $value in $Envvalues ) {
        $temp = $value.Values
        #_Build list of old servers
        if ( $temp[2] -match "10" -or $temp[2] -match "00" -or $temp[2] -match "2950" ) {
            $pos = $temp[0].IndexOf("-")
            if ( $pos -gt 0 ) { $servername = $temp[0].Substring(0, $pos) }
            else { $servername = $temp[0] }
            $model = $temp[2]
            $datarow = @"
                    <tr style="height: 32px;">
                    <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$servername</span></span></td>
                    <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$Model</span></span></td>
                    </tr>
"@
            $PastRefreshReport += $datarow
        }

        #_Add up all cores and memory
        $cores += $temp[8]
        $memory += ($temp[10] / 1024)
    }

    #_Do some math
    if ($memory -ge 1024) {
        $memory = $memory / 1024
        $memory = [math]::Round($memory, 2)
        $mem = "$memory Petabytes"
    }
    else {
        $memory = [math]::Round($memory, 2)
        $mem = "$memory Terabytes"
    }
    #EndRegion 

    #Region Get OME Compliance report
    $URI = "https://$OMEIP/api/UpdateService/Baselines"
    $Response = Invoke-WebRequest -Uri $URI -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
    $data = $response.Content | ConvertFrom-Json
    $values = $data.value
    ForEach ( $value in $values) {
        $PercentComplaint = $null
        $RepositoryName = $value.RepositoryName
        $NumberOfCompliant = ($value.ComplianceSummary.NumberOfNormal + $value.ComplianceSummary.NumberOfDowngrade)
        $NumberOfNonCompliant = ($value.ComplianceSummary.NumberOfCritical + $value.ComplianceSummary.NumberOfWarning)
        $PercentComplaint = ($NumberOfCompliant / $NumberOfServers) * 100
        $PercentComplaint = [math]::Round($PercentComplaint, 2)

        if ( $PercentComplaint -ge 70 ) {
            $PercentComplaint = $PercentComplaint.ToString() + "%"
            $datarow = @"
            <tr style="height: 32px;">
            <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$RepositoryName</span></span></td>
            <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$NumberOfCompliant</span></span></td>
            <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$NumberOfNonCompliant</span></span></td>
            <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center; background-color: #ccffcc;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$PercentComplaint</span></span></td>
            </tr>
"@
        }
        else {
            $PercentComplaint = $PercentComplaint.ToString() + "%"
            $datarow = @"
                <tr style="height: 32px;">
                <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$RepositoryName</span></span></td>
                <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$NumberOfCompliant</span></span></td>
                <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$NumberOfNonCompliant</span></span></td>
                <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$PercentComplaint</span></span></td>
                </tr>
"@
        }
        $complianceReport += $datarow
    }
    #EndRegion
   
    #Region Get OME Warranty Info
    #_Run OME Warranty Report
    $body = @{
        "ReportDefId"   = $warrantyReportId
        "FilterGroupId" = 0
    }
    $body = $body | ConvertTo-Json
    $URI = "https://$OMEIP/api/ReportService/Actions/ReportService.RunReport"
    $Response = Invoke-WebRequest -Uri $URI -Method POST -Body $body -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop 

    #_Get Warranty Report results
    $URI = "https://$OMEIP/api/ReportService/ReportDefs($warrantyReportId)" + '/ReportResults/ResultRows?$top=10000'
    $Response = Invoke-WebRequest -Uri $URI -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
    $data = $response.Content | ConvertFrom-Json
    $values = $data.value

    #_Parse Warranty Results
    ForEach ( $value in $values) {
        $servername = $value.values[0]
        $model = $values.values[3]
        $daysleft = $value.values[6]
        $datarow = @"
                    <tr style="height: 32px;">
                    <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$servername</span></span></td>
                    <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$Model</span></span></td>
                    <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; line-height: 30px; height: 32px; width: 25%; text-align: center;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$DaysLeft</span></span></td>
                    </tr>
"@
        $warrantyReport += $datarow
    }
    #EndRegion

    #Region Build HTML
    $Head = @'              
            <table id="background-table-header" style="background-color: #0085c3;" width="100%" cellspacing="0" cellpadding="0">
                <tbody>
                    <tr>
                        <td style="color: #ffffff; font-size: 14px; line-height: 12px; font-family: Arial, sans-serif; width: 100%;" align="center">
                            <h1>Daily Report</h1>
                        </td>
                    </tr>
                </tbody>
            </table>
'@
    $footer = @"
                <td id="footer-pattern" style="padding: 12px 20px; border-collapse: collapse;">
                        <table id="footer-pattern-container" style="border-collapse: collapse; width: 740px;" border="0"cellspacing="0" cellpadding="0">
                            <tbody>
                                <tr>
                                    <td id="footer-pattern-text-left"
                                        style="padding: 0px; border-collapse: collapse; color: #999999; font-size: 12px; line-height: 12px; font-family: Arial, sans-serif; width: 371px;">
                                        Property of CSX Corporation. All Rights Reserved</td>
                                    <td id="footer-pattern-text-right"
                                        style="color: #999999; font-size: 12px; line-height: 12px; font-family: Arial, sans-serif; width: 445px; text-align: right;"
                                        valign="right">&nbsp;\\cjax5006fs\compops\DCHWPROV\scripts\AssetReportCombined.csv</td>
                                </tr>
                            </tbody>
                        </table>
                    </td>
"@

    $html = @()
    $html += $head

    #Region Add Alerts if any
    if ($Alertcount -ne 0) {
        $html += "<h2>&nbsp;Critical Alerts</h2>"
        foreach ($alert in $OMEAlerts) {
            $Device = $alert.AlertDeviceName
            $AssetTag = $alert.AlertDeviceIdentifier
            $Category = $alert.CategoryName
            $SubCategory = $alert.SubCategoryName
            $TimeStamp = $alert.TimeStamp
            $Message = $alert.Message
            $Action = $alert.RecommendedAction

            $table = @"
                <table style="border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;" border="1" width="100%" cellspacing="0" cellpadding="0">
                <tbody>
                <tr style="height: 32px;">
                <td style="width: 20%; vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px;"><span style="color: #007db8; text-decoration: none; font-family: Arial, sans-serif; padding: 0; font-size: 12px; line-height: 30px; mso-text-raise: 2px; mso-line-height-rule: exactly; vertical-align: middle;">Device </span></td>
                <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$Device</span></span></td>
                </tr>
                <tr style="height: 32px;">
                <td style="width: 20%; vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px;"><span style="color: #007db8; text-decoration: none; font-family: Arial, sans-serif; padding: 0; font-size: 12px; line-height: 30px; mso-text-raise: 2px; mso-line-height-rule: exactly; vertical-align: middle;">AssetTag</span></td>
                <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px;"><span style="color: #007db8; text-decoration: none; font-family: Arial, sans-serif; padding: 0; font-size: 12px; line-height: 30px; mso-text-raise: 2px; mso-line-height-rule: exactly; vertical-align: middle;">$AssetTag</span></td>
                </tr>
                <tr style="height: 32px;">
                <td style="width: 20%; vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px;"><span style="color: #007db8; text-decoration: none; font-family: Arial, sans-serif; padding: 0; font-size: 12px; line-height: 30px; mso-text-raise: 2px; mso-line-height-rule: exactly; vertical-align: middle;"> Category </span></td>
                <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px;"><span style="color: #007db8; text-decoration: none; font-family: Arial, sans-serif; padding: 0; font-size: 12px; line-height: 30px; mso-text-raise: 2px; mso-line-height-rule: exactly; vertical-align: middle;"> $Category</span></td>
                </tr>
                <tr style="height: 32px;">
                <td style="width: 20%; vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px;"><span style="color: #007db8; text-decoration: none; font-family: Arial, sans-serif; padding: 0; font-size: 12px; line-height: 30px; mso-text-raise: 2px; mso-line-height-rule: exactly; vertical-align: middle;"> Sub-category </span></td>
                <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$SubCategory</span></span></td>
                </tr>
                <tr style="height: 32px;">
                <td style="width: 20%; vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px;"><span style="color: #007db8; text-decoration: none; font-family: Arial, sans-serif; padding: 0; font-size: 12px; line-height: 30px; mso-text-raise: 2px; mso-line-height-rule: exactly; vertical-align: middle;">TimeStamp</span></td>
                <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px;"><span style="color: #007db8; text-decoration: none; font-family: Arial, sans-serif; padding: 0; font-size: 12px; line-height: 30px; mso-text-raise: 2px; mso-line-height-rule: exactly; vertical-align: middle;">$TimeStamp</span></td>
                </tr>
                <tr style="height: 224px;">
                <td style="width: 20%; vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 224px;"><span style="color: #007db8; text-decoration: none; font-family: Arial, sans-serif; padding: 0; font-size: 12px; line-height: 30px; mso-text-raise: 2px; mso-line-height-rule: exactly; vertical-align: middle;"> Message </span></td>
                <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 64px;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$Message</span></span></td>
                </tr>
                <tr style="height: 64px;">
                <td style="width: 20%; vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 64px;"><span style="color: #007db8; text-decoration: none; font-family: Arial, sans-serif; padding: 0; font-size: 12px; line-height: 30px; mso-text-raise: 2px; mso-line-height-rule: exactly; vertical-align: middle;">Recommended Action </span></td>
                <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 64px;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$Action</span></span></td>
                </tr>
                </tbody>
                </table>
                </td>
                </tr>
                <tr>
                <td style="padding: 0px 15px 0px 16px; border-collapse: collapse; color: #ffffff; height: 5px; line-height: 5px; background-color: #ffffff; border-width: 0px 1px 1px; border-bottom-right-radius: 5px; border-bottom-left-radius: 5px; width: 705px; border-color: initial #cccccc #cccccc #cccccc; border-style: initial solid solid solid;">&nbsp;</td>
                </tr>
                </tbody>
                </table>
"@
            $html += $table
        }
    }
    else {
        $html += "<h2>&nbsp;No Critical Alerts in the last 24 hours</h2>"
    }
    #EndRegion

    #Region Add Open cases if any
    if ($CaseReport) {
        $CaseTable = @"
        <h2>&nbsp;Open Cases</h2>
        <table style="border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;" border="1" width="100%" cellspacing="0" cellpadding="0">
        <tr>
        <th>Service Tag</th>
        <th>Severity</th>
        <th>Description</th>
        </tr>
        $CaseReport
        </table>
        <tr>
        <br>
"@
        $html += $CaseTable
    }
    else {
        $html += "<h2>&nbsp;No open cases with Dell</h2>"
    }
    #EndRegion

    #Region Add EnvTable
    $EnvTable = @"
        <h2>&nbsp;Environment</h2>
        <table style="border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;" border="1" width="100%" cellspacing="0" cellpadding="0">
        <tbody>
        <tr style="height: 32px;">
        <td style="width: 20%;"><strong>Number of Servers in OME</strong></td>
        <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px; width: 80%;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$NumberOfServers</span></span></td>
        </tr>
        <tr style="height: 32px;">
        <td style="width: 20%;"><strong>Number of Cores</strong></td>
        <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px; width: 80%;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$cores</span></span></td>
        </tr>
        <tr style="height: 32px;">
        <td style="width: 20%;"><strong>Total Memory</strong></td>
        <td style="vertical-align: top; padding: 0px 5px 0px 0px; border-collapse: collapse; font-size: 20px; line-height: 30px; height: 32px; width: 80%;"><span style="color: #007db8; font-family: Arial, sans-serif;"><span style="font-size: 12px;">$mem</span></span></td>
        </tr>
        </tbody>
        </table>
        <p>&nbsp;</p>
"@
    $html += $EnvTable
    #EndRegion

    #Region Add ComplianceTable
    $ComplianceTable = @"
            <h2>&nbsp;Compliance</h2>
            <table style="border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;" border="1" width="100%" cellspacing="0" cellpadding="0">
            <tr>
            <th>Repository Name</th>
            <th>Number of Compliant</th>
            <th>Number of NonCompliant</th>
            <th>Percentage Compliant</th>
            </tr>
            $complianceReport
            </table>
            <tr>
            <br>
"@
    $html += $ComplianceTable
    #EndRegion

    #Region Add WarrantyTable
    $WarrantyTable = @"
            <h2>&nbsp;Warranties Expiring Soon</h2>
            <table style="border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;" border="1" width="100%" cellspacing="0" cellpadding="0">
            <tr>
            <th>Server Name</th>
            <th>Days Left</th>
            <th>Warranty Type</th>
            </tr>
            $warrantyReport
            </table>
            <tr>
            <br>

"@
    $html += $WarrantyTable
    #EndRegion

    #Region Add PastRefreshTable
    $PastRefreshTable = @"
            <h2>&nbsp;Servers Past Refresh Date</h2>
            <table style="border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;" border="1" width="100%" cellspacing="0" cellpadding="0">
            <tr>
            <th>Server Name</th>
            <th>Model</th>
            </tr>
            $PastRefreshReport
            </table>
            <tr>
            <br>
"@
    $html += $PastRefreshTable
    #EndRegion

    #EndRegion
}

end {
    $mailbody = ConvertTo-Html -Title 'Daily Report' -Body $html -PostContent $footer  

    #_Send out email
    $SmtpClient.Host = $smtp
    $mailmessage.from = $from
    $to | Where-Object { $mailmessage.To.add($_) }
    $mailmessage.Subject = $subject
    $MailMessage.IsBodyHTML = $true
    $mailmessage.Body = $mailbody
    $smtpclient.Send($mailmessage)
    $mailmessage.Dispose()    
  
}


