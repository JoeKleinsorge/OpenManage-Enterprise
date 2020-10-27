<#
.SYNOPSIS
    Returns server data from REDFISH API calls
.DESCRIPTION
    Uses REDFISH to gather server data on; System Info, Controllers, Volumes, Disks, NICs, DIMMs, and BIOS Info
.EXAMPLE
    PS C:\> Get-Data_Redfish -Server 10.0.0.1
    Uses the username and password to return data from the iDRAC with the above IP address
.EXAMPLE
    PS C:\> Get-Data_Redfish -Server servernamehere
    Uses the username and password to return data from the iDRAC with the above DNS name
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

    [Parameter(Mandatory)]
    [string]$Server
)
Begin {
    #_Convert credentials
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::TLS12
    $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)

    #_Make PowerShell window correct size
    $pshost = get-host
    $pswindow = $pshost.ui.rawui
    $newsize = $pswindow.buffersize
    $newsize.height = 5000
    $newsize.width = 5000
    $pswindow.buffersize = $newsize

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
    $Results = @()
    
}
Process {
    Write-Host "Processing System Info" -ForegroundColor Yellow
    $URI = "https://$server/redfish/v1/Systems/System.Embedded.1"
    $Response = Invoke-WebRequest -Uri $URI -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
    $data = $response.Content | ConvertFrom-Json
    $item = new-object psobject
    $item | add-member -type noteproperty -Name HostName -Value $data.HostName
    $item | add-member -type noteproperty -Name Model -Value $data.Model
    $item | add-member -type noteproperty -Name Manufacturer -Value $data.Manufacturer
    $item | add-member -type noteproperty -Name CPUModel -Value $data.ProcessorSummary.Model
    $item | add-member -type noteproperty -Name CPUCount -Value $data.ProcessorSummary.Count
    $item | add-member -type noteproperty -Name CPUHealth -Value $data.ProcessorSummary.Status.health
    $item | add-member -type noteproperty -Name MemoryGB -Value $data.MemorySummary.TotalSystemMemoryGiB
    $item | add-member -type noteproperty -Name MemoryHealth -Value $data.MemorySummary.Status.Health
    $item | add-member -type noteproperty -Name partNumber -Value $data.partnumber
    $item | add-member -type noteproperty -Name PowerState -Value $data.powerstate
    $item | add-member -type noteproperty -Name ServiceTag -Value $data.sku
    $item | add-member -type noteproperty -Name SerialNumber -Value $data.serialnumber
    $item | add-member -type noteproperty -Name BIOS -Value $data.BiosVersion
    $item | add-member -type noteproperty -Name Status -Value $data.status
    $results += $item
    $system = $results | Sort-Object HostName | Format-Table -a *
    $Results = @()

    Write-Host "Processing Controller Info" -ForegroundColor Yellow
    $Cont = "https://$server/redfish/v1/Systems/System.Embedded.1/Storage"
    $ContResponse = Invoke-WebRequest -Uri $Cont -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
    $ContData = $ContResponse.Content | ConvertFrom-Json
    $contMembers = $ContData.Members
    Foreach ( $_ in $contMembers) {
        $MemData = $_.'@odata.id'
        if ($MemData -match "RAID") { $controller = $MemData }
    }
    

    write-host "Processing Volumes Info " -foregroundcolor "yellow"
    $volumes = "https://" + $server + $controller + "/Volumes"
    $volresponse = Invoke-WebRequest -Uri $volumes -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
    $voldata = $volresponse.Content | ConvertFrom-Json
    $voldata.members | ForEach-Object {
        $volid = $_.'@odata.id'
        $volURI = "https://$server$volid"
        $volResponse2 = Invoke-WebRequest -Uri $volURI -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
        $voldata2 = $volresponse2.Content | ConvertFrom-Json
        $gb = [math]::round($voldata2.CapacityBytes / 1GB)
        $item = New-Object PSObject 
        $item | Add-Member -type NoteProperty -Name Name  -Value $voldata2.Name 
        $item | Add-Member -type NoteProperty -Name DiskCount -Value $voldata2.links.'Drives@odata.count' 
        $item | Add-Member -type NoteProperty -Name blocksizebytes  -Value $voldata2.blocksizebytes 
        $item | Add-Member -type NoteProperty -Name status  -Value $voldata2.status 
        $item | Add-Member -type NoteProperty -Name Size  -Value $gb 
        $results += $item
    } 
    $Volumes = $results | Sort-Object Name | Format-Table -a *
    $Results = @()

    write-host "Processing Disk Info " -foregroundcolor "yellow"
    $uri = "https://" + $server + $controller
    $disks = Invoke-WebRequest -Uri $URI -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
    $disksdata = $disks.Content | ConvertFrom-Json
    if ($disksdata.'Drives@odata.count' -lt 1) { Write-Host "No Drives found" }
    Else { 
        $disksdata.drives | ForEach-Object {
            $deviceid = $_.'@odata.id'
            $deviceURI = "https://$server$deviceid"
            $deviceResponse = Invoke-WebRequest -Uri $deviceURI -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
            $devicedata2 = $deviceresponse.Content | ConvertFrom-Json
            $diskgb = [math]::round($devicedata2.CapacityBytes / 1GB)
            $item = New-Object PSObject 
            $item | Add-Member -type NoteProperty -Name percID -Value $disksdata.id 
            $item | Add-Member -type NoteProperty -Name percName -Value $disksdata.Name 
            $item | Add-Member -type NoteProperty -Name DiskManufacturer -Value $devicedata2.manufacturer 
            $item | Add-Member -type NoteProperty -Name DiskModel  -Value $devicedata2.Model 
            $item | Add-Member -type NoteProperty -Name DiskName -Value $devicedata2.Name 
            $item | Add-Member -type NoteProperty -Name Status -Value $devicedata2.status 
            $item | Add-Member -type NoteProperty -Name DiskBay -Value $devicedata2.Id 
            $item | Add-Member -type NoteProperty -Name Speed -Value $devicedata2.NegotiatedSpeedGbs 
            $item | Add-Member -type NoteProperty -Name Size -Value $diskgb 
            $item | Add-Member -type NoteProperty -Name FailurePredicted -Value $devicedata2.FailurePredicted 
            $item | Add-Member -type NoteProperty -Name MediaType -Value $devicedata2.MediaType 
            $item | Add-Member -type NoteProperty -Name PartNumber -Value $devicedata2.PartNumber 
            $item | Add-Member -type NoteProperty -Name Protocol -Value $devicedata2.Protocol 
            $item | Add-Member -type NoteProperty -Name RPM -Value $devicedata2.RotationSpeedRPM 
            $item | Add-Member -type NoteProperty -Name Serial -Value $devicedata2.SerialNumber
            $results += $item
        }
        $Disks = $results | Sort-Object Name | Format-Table -a * 
        $Results = @()
    }

    write-host "Processing NIC Info " -foregroundcolor "yellow"
    $EthernetURI = "https://$server/redfish/v1/Systems/System.Embedded.1/NetworkAdapters"
    $NicsResponse = Invoke-WebRequest -Uri $EthernetURI -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
    $nicsdata = $Nicsresponse.Content | ConvertFrom-Json
    $nicsdata.members | ForEach-Object {
        $netuRI = $_.'@odata.id'
        $EthURI = "https://$server$neturi"
        $NicsResponse2 = Invoke-WebRequest -Uri $EthURI -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
        $nicsdata2 = $Nicsresponse2.Content | ConvertFrom-Json
        $item = New-Object PSObject 
        $item | Add-Member -type NoteProperty -Name Id -Value $nicsdata2.Id 
        $item | Add-Member -type NoteProperty -Name Manufacturer -Value $nicsdata2.Manufacturer 
        $item | Add-Member -type NoteProperty -Name Model -Value $nicsdata2.Model 
        $item | Add-Member -type NoteProperty -Name Partnumber -Value $nicsdata2.partnumber  
        $item | Add-Member -type NoteProperty -Name SerialNumber -Value $nicsdata2.SerialNumber  
        $item | Add-Member -type NoteProperty -Name Status  -Value $nicsdata2.Status
        $results += $item
    } 
    $nics = $results | Sort-Object Id | Format-Table -a *
    $Results = @()

    write-host "Processing DIMM Details" -foregroundcolor "yellow"
    $memURI = "https://$server/redfish/v1/Systems/System.Embedded.1/Memory"
    $mem = Invoke-WebRequest -Uri $memURI -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
    $memdata = $mem.Content | ConvertFrom-Json
    $memdata.members | ForEach-Object {
        $memid = $_.'@odata.id'
        $memURI2 = "https://$server$memid"
        $memResponse2 = Invoke-WebRequest -Uri $memuri2 -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
        $memdata2 = $memResponse2.Content | ConvertFrom-Json
        $item = New-Object PSObject 
        $item | Add-Member -type NoteProperty -Name Name   -Value  $memdata2.Id 
        $item | Add-Member -type NoteProperty -Name Capacity -Value  $memdata2.CapacityMiB 
        $item | Add-Member -type NoteProperty -Name Type   -Value  $memdata2.memorydevicetype 
        $item | Add-Member -type NoteProperty -Name Speed   -Value  $memdata2.operatingspeedmhz 
        $item | Add-Member -type NoteProperty -Name partnumber   -Value  $memdata2.partnumber 
        $item | Add-Member -type NoteProperty -Name Rank   -Value  $memdata2.RankCount 
        $item | Add-Member -type NoteProperty -Name Serial   -Value  $memdata2.serialnumber 
        $item | Add-Member -type NoteProperty -Name Status   -Value  $memdata2.status
        $item | Add-Member -type NoteProperty -Name Manufacturer   -Value  $memdata2.manufacturer
        $results += $item
    } 
    $DIMMS = $results | Sort-Object Name | Format-Table -a *
    $Results = @()


    write-host "Processing BIOS Settings" -foregroundcolor "yellow"
    $biosURI = "https://$server/redfish/v1/Systems/System.Embedded.1/Bios"
    $bios = Invoke-WebRequest -Uri $biosURI -Credential $credential -Contenttype "application/JSON" -ErrorAction:Stop
    $biosdata = $bios.Content | ConvertFrom-Json
    $BIOS = $biosdata.Attributes


}
End {
    Write-Host ""
    Write-Host "System Info" -ForegroundColor Yellow
    $system
    Write-Host ""

    Write-Host "Controller Info" -ForegroundColor Yellow
    $controller
    Write-Host ""

    Write-Host "Volume Info" -ForegroundColor Yellow
    $volumes

    if ($disksdata.'Drives@odata.count' -lt 1) { Write-Host "No Drives found" }
    else {
        Write-Host "Disk Info" -ForegroundColor Yellow
        $disks
    }

    Write-Host "NIC Info" -ForegroundColor Yellow
    $nics

    Write-Host "DIMM Info" -ForegroundColor Yellow
    $DIMMS

    Write-Host "BIOS Info" -ForegroundColor Yellow
    $BIOS
}