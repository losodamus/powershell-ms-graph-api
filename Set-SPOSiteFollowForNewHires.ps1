<#
    Set-SPOSiteFollowOfNewHires.ps1
    ------------------------------


    Query and return list of recent new hires.
    Query and return SharePoint intranet site id.
    Follow SPO site on behalf of new hires. 
#>
cls


##  Variable(s)
$global:AppGuid = "<app_id>"
$global:AppSecret = "<app_secret>"
$global:TenantName = "<tenant_name>"


[System.String] $sitePath = "/sites/Intranet"   ##  SharePoint Intranet path
[System.String] $groupName = "<name_of_group>"  ##  Azure AD group of new hires


[System.Int32] $indexOf = 0
[System.Array] $batchOf = @()


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


function Get-RestAPIResonse () {
    param(
        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [ValidateSet('Get', 'Post', 'Patch')] 
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
cls


try {


    ##  E.g., <tenant_name>.sharepoint.com,076bf298-a1b2-49b8-9658-9c9796e3623c,7e80afae-d0f6-abcd-1234-773b9fe11381
    ##  --------------------------------------------------
    foreach($site in (Get-RestAPIResonse `
                        -APIMethod Get `
                        -APIVersion v1.0 `
                        -APIResource "sites/$($global:TenantName).sharepoint.com:/$($sitePath)/?`$select=id,name,webUrl")) {


        ##  E.g., 8317d5d7-84a4-a404-bb00-1f865ae63203
        ##  --------------------------------------------------
        foreach($group in (Get-RestAPIResonse `
                            -APIMethod Get `
                            -APIVersion v1.0 `
                            -APIResource "groups/?`$select=id,displayName,description&`$filter=startswith(displayName,'$($groupName)')").value) {


            ##  E.g., ed348d15-4fff-404a-bfce-26b6e19e9d4b
            ##  --------------------------------------------------
            foreach($member in (Get-RestAPIResonse `
                                -APIMethod Get `
                                -APIVersion v1.0 `
                                -APIResource "groups/$($group.id)/members/?`$select=id,displayName").value) {


                $indexOf += 1
                $batchOf += @{
                    "url" = "/users/$($member.id)/followedSites/add"
                    "method" = "POST"
                    "id" = "$($indexOf)"
                    "body" = @{
                        "value" = @(
                            @{
                                "id" = "$($site.id)"
                            }
                        )
                    }
                    "headers" = @{
                        "Content-Type" = "application/json"
                    }
                }


                if ($indexOf -eq 20) {
                    $indexOf = 0
                
                
                    Get-RestAPIBatchResponse `
                        -APIBatch $batchOf
                    $batchOf = @()
                }
            }


            if ($indexOf -ne 0) {
                Get-RestAPIBatchResponse `
                    -APIBatch $batchOf
                $batchOf = @()
            }
        }
    }
}
catch [Exception] {
    Write-Error $_.Exception.Message
}
