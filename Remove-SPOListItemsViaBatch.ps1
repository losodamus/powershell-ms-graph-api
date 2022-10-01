<#
    Remove-SPOListItemsViaBatch.ps1
    ------------------------------


    - Populate script variables.
    - Query target site GUID.
    - Query target list GUID.
    - Query target list items.
    - Delete in 20 item batches.
#>
Clear-Host


##  Variable(s)
$global:AppGuid = "<app_id>"
$global:AppSecret = "<app_secret>"
$global:TenantName = "<tenant_name>"


[System.String] $sitePath = "<relative_site_path>"  ##  E.g., /sites/intranet
[System.String] $listPath = "<relative_list_path>"  ##  E.g., /Lists/news


function Get-RestAPIHeader {
    $Body = @{
        client_id = "$($global:AppGuid)"
	    client_secret = "$($global:AppSecret)"
	    scope = "https://graph.microsoft.com/.default"
	    grant_type = "client_credentials"
    }


    $PostSplat = @{
        ContentType = "application/x-www-form-urlencoded"
        Method = "POST"
        Body = $Body
        Uri = "https://login.microsoftonline.com/$($global:TenantName).onmicrosoft.com/oauth2/v2.0/token"
    }
    $Request = Invoke-RestMethod @PostSplat


    return (@{
        Authorization = "$($Request.token_type) $($Request.access_token)"
    })
}


function Get-RestAPIResonse {
    param(
        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [ValidateSet('Get')] 
        [System.String] $APIMethod,


        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [ValidateSet('Beta', 'v1.0')] 
        [System.String] $APIVersion,


        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [System.String] $APIResource
    )


    try {
        $r = Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/$($APIVersion)/$($APIResource)" `
            -Headers (Get-RestAPIHeader) `
            -Method $APIMethod `
            -ContentType "application/json"
        return $r
    }
    catch [Exception] {
        return $null
    }
}


function Get-RestAPIBatchResponse {
    param(
        [Parameter(Mandatory)] 
        [System.Array] $APIBatch
    )


    try {
        $jsonBody = @{
            "requests" = $APIBatch
        } | ConvertTo-Json -Depth 5
        $jsonBody | Out-Host


        $r = (Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/v1.0/`$batch" `
            -Headers (Get-RestAPIHeader) `
            -Body $jsonBody `
            -Method POST `
            -ContentType "application/json")
        return $r
    }
    catch [Exception] {
        return $null
    }
}


##  E.g., /sites/marketing
function Get-SPOSiteByRelativePath {
    param(
        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()]
        [System.String] $PathOfSite
    )


    $t = (Get-RestAPIResonse `
                -APIMethod Get `
                -APIVersion v1.0 `
                -APIResource "sites/$($global:TenantName).sharepoint.com:/$($PathOfSite)/")
    return $t
}


## E.g., /Lists/collateral
function Get-SPOSiteListByRelativePath {
    param(
        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()]
        [System.String] $GuidOfSite,


        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()]
        [System.String] $PathOfList
    )


    $t = (Get-RestAPIResonse `
                -APIMethod Get `
                -APIVersion v1.0 `
                -APIResource "sites/$($GuidOfSite)/lists/").value `
                | ? { $_.webUrl -like "*$($PathOfList)" }
    return $t
}


function Remove-SPOSiteListItemsByBatch {
    param(
        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()]
        [System.String] $GraphAPIEndpoint
    )


    [System.Int32] $indexOf = 0
    [System.Array] $batchOf = @()
    [System.String] $nextLink = ""


    try {
        [System.String] $e = "$($GraphAPIEndpoint)".Replace("/sites/", "|").Replace("/lists/", "|").Replace("/items", "|")
        [System.String] $guidOfThisSite = $e.Split("|")[1]
        [System.String] $guidOfThisList = $e.Split("|")[2]


        $GraphReq = Invoke-RestMethod `
            -Uri $GraphAPIEndpoint `
            -Headers (Get-RestAPIHeader) `
            -Method GET `
            -ContentType "application/json"


        if ($GraphReq.'@odata.nextLink'.Length -ne 0) {
            $nextLink = $GraphReq.'@odata.nextLink'
        }


        $GraphReq.value | % {
            $indexOf += 1
            $batchOf += @{
                "url" = "/sites/$($guidOfThisSite)/lists/$($guidOfThisList)/items/$($_.id)"
                "method" = "DELETE"
                "id" = "$($indexOf)"
                "headers" = @{
                    "Content-Type" = "application/json"
                }
            }


            if ($batchOf.Count -eq 20) {
                Get-RestAPIBatchResponse `
                    -APIBatch $batchOf
                $indexOf = 0
                $batchOf = @()
            }
        }


        if ($batchOf.Count -ne 0) {
            Get-RestAPIBatchResponse `
                -APIBatch $batchOf
        }
    }
    catch [Exception] {
        Write-Host -F Red $_.Exception.Message
    }


    ##  Process next batch
    if ($nextLink -ne "") {
        Write-Host -F Yellow "Next:`t$($nextLink)"
        Start-Sleep -Milliseconds 2500
        Remove-SPOSiteListItemsByBatch `
            -GraphAPIEndpoint "$nextLink"
    }
}
Clear-Host


try {


    ##  E.g., <tenant_name>.sharepoint.com,076bf298-a1b2-49b8-9658-9c9796e3623c,7e80afae-d0f6-abcd-1234-773b9fe11381
    ##  --------------------------------------------------
    foreach($site in (Get-SPOSiteByRelativePath `
        -PathOfSite "$($sitePath)")) {
        Write-Host -F Cyan "Site:`t$($site.id)"


        ##  E.g., b!mPJrB6couEmWWJyXluNiPK6vgH720IpBh_R404_hE4H0-fEDWHMKQ6oR_pSl_atl
        ##  --------------------------------------------------
        foreach($list in (Get-SPOSiteListByRelativePath `
            -GuidOfSite "$($site.id)" `
            -PathOfList "$($listPath)")) {
            Write-Host -F Magenta "List:`t$($list.id)"


            Remove-SPOSiteListItemsByBatch `
                -GraphAPIEndpoint "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($list.id)/items/?`$select=id,webUrl&`$top=2000"
        }
    }
}
catch [Exception] {
    Write-Error "`nError Message:"
    Write-Error $_.Exception.Message
}