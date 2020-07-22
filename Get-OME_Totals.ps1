<#
.SYNOPSIS
    Outputs total number of processors and amount of memory in servers discovered by OME.
.DESCRIPTION
    Uses OpenManage Enterprise API to run report and return results.
.EXAMPLE
    Get-OME_Totals -OME 10.0.0.1 
.INPUTS
    - OME (IP address for OME server)
    - Credentials (OME user creds)
.OUTPUTS
    Total processors and memory count of servers discovered by OME.
.NOTES
    Author - Joey Kleinsorge
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)] 
    [System.Net.IPAddress] $OME,

    [Parameter(Mandatory)]
    [pscredential] $Credentials
)

Begin {
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
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::TLS12

    #_Assign vars
    $ReportId = "57508"
    $procs = 0
    $memory = 0
}
Process {
    #_Build request body
    $body = @{
        "ReportDefId"   = $ReportId
        "FilterGroupId" = 0
    }
    $body = $body | ConvertTo-Json

    #_Run report
    $URI = "https://$OME/api/ReportService/Actions/ReportService.RunReport"
    $Response = Invoke-WebRequest -Uri $URI -Method POST -Body $body -Credential $credentials -Contenttype "application/JSON" -ErrorAction:Stop 

    #_Get report results
    $URI = "https://$OME/api/ReportService/ReportDefs($ReportId)/ReportResults/ResultRows"
    $Response = Invoke-WebRequest -Uri $URI -Credential $credentials -Contenttype "application/JSON" -ErrorAction:Stop
    $data = $response.Content | ConvertFrom-Json
    $values = $data.value
    
    #_Pull values from results
    ForEach ( $value in $values) {
        $temp = $value.Values
        $procs += $temp[8]
        $memory += $temp[10]
    }

}
End {
    #_Output results
    Write-Host " Total Processors: $procs"
    if ($memory -ge 1024) {
        $memory = $memory / 1024
        
        if ($memory -ge 1024) {
            $memory = $memory / 1024
            Write-Host " Total Memory: $memory Petabytes"
        }
        else {
            Write-Host " Total Memory: $memory Terabytes"
        }
    }
    else {
        Write-Host " Total Memory: $memory Gigabytes" 
    }   
}




