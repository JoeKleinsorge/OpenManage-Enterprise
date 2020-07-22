<#
.SYNOPSIS
    Outputs currently open Dell cases.
.DESCRIPTION
    Uses TechDirect production API to get check if cases are open for each service tag in OME.
.EXAMPLE
    Should only be run as a scheduled task.
.NOTES
    CSV output is to be used for Daily report. 
#>
[CmdletBinding()]
param (
    # Path to OMEnterpiseData.csv
    [Parameter()]
    [string]
    $OMEnterpiseCSV,

    # TechDirect Client ID
    [Parameter()]
    [string]
    $ClientID,

    # TechDirect Client Secret
    [Parameter()]
    [string]
    $ClientSecret,

    # Path for CSV output
    [Parameter()]
    [string]
    $CSVOutputPath

)
begin {
    #_Import Asset Tags
    $ListofTags = @()
    $OMEnterpiseCSV | ForEach-Object { 
        $ListofTags += $_.ServiceTag 
    }

    #_Get Access Token
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add( "Content-Type", "application/x-www-form-urlencoded" )
    
    $body = "client_id=$ClientID&client_secret=$ClientSecret" + "grant_type=client_credentials"
    
    $response = Invoke-RestMethod 'https://apigtwb2c.us.dell.com/auth/oauth/v2/token' -Method 'POST' -Headers $headers -Body $body
    $token = $Response.access_token

    #_Setup varible
    $OpenCases = @()
}

process {
    #_Get SR data for each service tag 
    ForEach ( $AssetTag in $ListofTags ) {
        Write-Host "Checking $AssetTag" 

        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add( "Authorization", "Bearer $token" )
        $headers.Add( "Content-Type", "text/xml" )
            
        $body = "<soapenv:Envelope xmlns:soapenv=`"http://schemas.xmlsoap.org/soap/envelope/`" 
                `nxmlns:dell=`"http://api.dell.com/dell.externalapi.casequery`" 
                `nxmlns:dell1=`"http://schemas.datacontract.org/2004/07/Dell.ExternalApi.CaseQuery.Domain.SearchCase`" 
                `nxmlns:dell2=`"http://schemas.datacontract.org/2004/07/Dell.ExternalApi.CaseQuery.Domain`" 
                `nxmlns:dell3=`"http://schemas.datacontract.org/2004/07/Dell.ExternalApi.CaseQuery.Domain.Enum`">          
                `n              <soapenv:Body>
                `n            <dell:SearchCase>
                `n                <dell:SearchCasesRequest>
                `n                <dell1:MessageHeader>
                `n                    <dell1:SourceSystem>
                `n                        <dell1:Name>API Inc</dell1:Name>
                `n                    </dell1:SourceSystem>
                `n                </dell1:MessageHeader>
                `n                <dell1:QueryFilter>
                `n                    <dell1:DataSourceInclusionList>
                `n                        <dell2:DataSourceType>DELTA</dell2:DataSourceType>
                `n                    </dell1:DataSourceInclusionList>
                `n                    <dell1:PageNumber>1</dell1:PageNumber>
                `n                    <dell1:PageSize>50</dell1:PageSize>
                `n                </dell1:QueryFilter>
                `n                <dell1:Asset>
                `n                    <dell1:ServiceTag>$AssetTag</dell1:ServiceTag>
                `n                </dell1:Asset>
                `n                </dell:SearchCasesRequest>
                `n            </dell:SearchCase>
                `n        </soapenv:Body>
                `n</soapenv:Envelope>"
            
        $response = Invoke-RestMethod 'https://apigtwb2c.us.dell.com/support/case/v3/searchcaselite' -Method 'POST' -Headers $headers -Body $body
        
        $cases = $response.Envelope.Body.SearchCaseResponse.SearchCaseResult.Cases.Case.CaseHeader

        #_Loop through ease case for the ST/ add to list if open
        $OpenCase = $Cases | ForEach-Object {
            If ( $cases.length -gt 0 -and $_.StatusDescription -notmatch "Closed" ) {
                New-Object PSObject |
                Add-Member -pass NoteProperty ServiceTag $AssetTag |
                Add-Member -pass NoteProperty Severity  $_.SeverityDescription |
                Add-Member -pass NoteProperty Description $_.Title
            }
        }
        #_Add open case to master list
        If ( $OpenCase ) {
            $OpenCases += $OpenCase
        }
    }
}

end {
    If ($OpenCases) {
        $OpenCases | Format-Table -a
        If ($CSVOutputPath) {
            $FilePath = $CSVOutputPath + "/OpenCases.csv"
            $OpenCases | export-csv $FilePath -NoTypeInformation -UseCulture
        }
    }
    Else {
        Write-Host "No open cases found."
    }
}
