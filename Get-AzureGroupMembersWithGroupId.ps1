<#
    Get-AzureGroupMembersWithGroupId.ps1
    ------------------------------
    

    - Populate script variable(s).
    - Query Azure AD group id using display name.
    - Query Azure AD group members using group id.
#>


##  Variable(s)
[System.String] $groupDispName = "<group_name>"


$global:AppGuid = "<app_id>"
$global:AppSecret = "<app_secret>"
$global:TenantName = "<tenant_name>"


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
        [System.String] $APIResource,


        [Parameter(Mandatory)]
        [System.Collections.Hashtable] $APIBody
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


##  E.g., $select=id,displayName,mail,resourceProvisioningOptions
##  --------------------------------------------------
function Get-AzureGroupByName {
    param(
        [Parameter(Mandatory)] 
        [System.String] $NameOfGroup
    )
    $t = (Get-RestAPIResonse `
            -APIMethod Get `
            -APIVersion v1.0 `
            -APIBody @{} `
            -APIResource "groups/?`$filter=startswith(displayName,'$($NameOfGroup)')").value
    return $t
}


##  E.g., $select=id,displayName,userPrincipalName
##  --------------------------------------------------
function Get-AzureGroupMembersListById {
    param(
        [Parameter(Mandatory)] 
        [System.String] $GuidOfGroup
    )
    $t = (Get-RestAPIResonse `
            -APIMethod Get `
            -APIVersion v1.0 `
            -APIBody @{} `
            -APIResource "groups/$($GuidOfGroup)/members").value
    return $t
}
Clear-Host


try {
    foreach($g in (Get-AzureGroupByName `
            -NameOfGroup $groupDispName)) {
        Write-Host -F Cyan "$($g.displayName)"
        Write-Host -F Cyan "$($g.id)`n"


        foreach($m in (Get-AzureGroupMembersListById `
                -GuidOfGroup $g.id)) {
            Write-Host -F Magenta "$($m.displayName)"
            Write-Host -F Magenta "$($m.userPrincipalName)`n"
        }
    }
}
catch [Exception] {
    Write-Error $_.Exception.Message
}