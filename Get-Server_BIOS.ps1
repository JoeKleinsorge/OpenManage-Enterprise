<#
.SYNOPSIS
    Gets Dell server BIOS settings
.DESCRIPTION
    Uses REDFISH REST call to to get BIOS data. 
.EXAMPLE
    Get-Server_BIOS -server serveridracname
.INPUTS
    - Credentials (Must have at least read permissions on iDRAC)
    - Server (Can be IP or DNS name of iDRAC)
.OUTPUTS
    Server information 
.NOTES
    Author - Joey Kleinsorge
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [pscredential] $Credentials,

    [Parameter(Mandatory = $true)] 
    [string]$server
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

    #_Convert credentials
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::TLS12
    $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)
}

Process {
    write-host "BIOS Settings" -foregroundcolor "yellow"
    $biosURI = "https://$server/redfish/v1/Systems/System.Embedded.1/Bios"
    $bios = Invoke-WebRequest -Uri $biosURI -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
    $biosdata = $bios.Content | ConvertFrom-Json
    $attributes = $biosdata.Attributes
}

End {
    $attributes
}
