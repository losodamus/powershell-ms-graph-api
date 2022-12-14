<#
    Get-AzureGroupIdByName.ps1
    ##  --------------------------------------------------


    Query an Azure group by name and return the group ID.
#>


##  Variable(s)
$global:AppGuid = "<app_id>"
$global:AppSecret = "<app_secret>"
$global:TenantName = "<tenant_name>"


[System.String] $groupGuid = ""
[System.String] $groupName = "<group_name>"


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
cls


try {
    ##  E.g., 8317d5d7-84a4-a404-bb00-1f865ae63203
    ##  --------------------------------------------------
    foreach($g in (Get-RestAPIResonse `
                    -APIMethod Get `
                    -APIVersion v1.0 `
                    -APIResource "groups/?`$select=id,displayName&`$filter=startswith(displayName,'$($groupName)')").value) {
        $groupGuid = "$($g.id)"
    }


    ##  Found or NOT Found
    if ($groupGuid -ne "") {
        Write-Host -F Green "Group Found: $($groupGuid)."
    }
    else {
        Write-Host -F Red "Group NOT Found!"
    }
}
catch [Exception] {
    Write-Error $_.Exception.Message
}
