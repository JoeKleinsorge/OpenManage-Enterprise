<#
.SYNOPSIS
    Wipes HDDs in Dell Poweredge servers with REDFISH enabled.
.DESCRIPTION
    Uses REDFISH REST calls to get all disks in server, uses "Drive.SecureErase" to wipe each drive. Does not work on SSDs.
.EXAMPLE
    Invoke-AllDiskSecureWipe -Server 10.0.0.1
.EXAMPLE
    Invoke-AllDiskSecureWipe -Server servernamehere
.INPUTS
    - Credentials (Must have at least read permissions on iDRAC)
    - Server (Can be IP or DNS name of iDRAC)
.OUTPUTS
    Completed jobs
.NOTES
    Author - Joey Kleinsorge
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [pscredential] $Credentials,

    [Parameter(Mandatory)] 
    [string]$server
)

Begin {
    #_Ignore-SSLCerts
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

    Ignore-SSLCertificates

    #_Create Array of jobids
    $jobs = @()
}
Process {
    #_Get list of storage controllers
    $u = "https://$idrac_ip/redfish/v1/Systems/System.Embedded.1/Storage" 
    $Response = Invoke-WebRequest -Uri $u -Credential $credential -Method Get -UseBasicParsing -Headers @{"Accept" = "application/json" }

    #_Check response to API call
    if ($Response.StatusCode -eq 200) {
        [String]::Format("`n- PASS, statuscode {0} returned successfully to get storage controller(s)", $Response.StatusCode)
    }
    else {
        [String]::Format("`n- FAIL, statuscode {0} returned", $Response.StatusCode)
        return
    }
    $Data = $Response.Content | ConvertFrom-Json
    $listofcontrollers = $Data.Members

    #_Get RAID Controller location
    $raidcontroller = $listofcontrollers | select-string "RAID"
    $raidstring = out-string -InputObject $raidcontroller 
    [regex]$regex = '/[^}]*'
    $raidcontrollerapi = $regex.Match($raidstring)
    $raidcontrollerapi = $raidcontrollerapi.Value

    #_Get list of drives attached to controller
    $u = "https://$idrac_ip" + $raidcontrollerapi
    $Response = Invoke-WebRequest -Uri $u -Credential $credential -Method Get -UseBasicParsing -Headers @{"Accept" = "application/json" }
    $Data = $Response.Content | ConvertFrom-Json
    $Drives = $Data.Drives

    #_Secure Erase each drive
    foreach ($Drive in $Drives) {
        $DriveAPI = $Drive.'@odata.id'
        $u = "https://$idrac_ip$DriveAPI"
        $Response = Invoke-WebRequest -Uri $u -Credential $credential -Method Get -UseBasicParsing -Headers @{"Accept" = "application/json" }
        $Data = $Response.Content | ConvertFrom-Json

        #_If SSDs break
        If ($Data.MediaType -eq "SSD") {
            Write-Warning "$u is an SSD and cannot be wiped"
        }

        #_Secure Erase HDDs
        Else {
            $u = "https://$idrac_ip$DriveAPI/Actions/Drive.SecureErase"
            Write-Host "Trying to erase $DriveAPI"
            try {
                $Response = Invoke-WebRequest -Uri $u -Credential $credential -Method Post -ContentType 'application/json' -ErrorVariable RespErr -Headers @{"Accept" = "application/json" }
            }
            catch {
                Write-Host
                $RespErr
                return
            }

            #_Check response to API call
            if ($Response.StatusCode -eq 202) {
                $q = $Response.RawContent | ConvertTo-Json -Compress
                $j = [regex]::Match($q, "JID_.+?r").captures.groups[0].value
                $job_id = $j.Replace("\r", "")
                [String]::Format("`n- PASS, statuscode {0} returned to successfully erase '{1}' device, {2} job ID created", $Reponse.StatusCode, $DriveAPI, $job_id)
            }
            else {
                [String]::Format("- FAIL, statuscode {0} returned", $Response.StatusCode)
                return
            }

            #_Output job status
            $u = "https://$idrac_ip/redfish/v1/Managers/iDRAC.Embedded.1/Jobs/$job_id"
            try {
                $Response = Invoke-WebRequest -Uri $u -Credential $credential -Method Get -UseBasicParsing -ContentType 'application/json' -ErrorVariable RespErr -Headers @{"Accept" = "application/json" }
            }
            catch {
                Write-Host
                $RespErr
                return
            }
            $jobs += $job_id
        }
    }
}
End {
    Write-Verbose "The following jobs were completed:"
    $jobs
}




