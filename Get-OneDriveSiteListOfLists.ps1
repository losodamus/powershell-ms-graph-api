<#
    Get-OneDriveSiteListOfLists.ps1
    ------------------------------


    Output all lists of a specific 
    user's OneDrive site. 
#>
cls


##  Variable(s)
$global:AppGuid = "<app_id>"
$global:AppSecret = "<app_secret>"
$global:TenantName = "<tenant_name>"


[System.String] $sitePath = "/personal/charles_theiilakesgroup_com"
[System.String] $siteGuid = ""


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
cls


try {
    ##  E.g., <tenant_name>-my.sharepoint.com,076bf298-a1b2-49b8-9658-9c9796e3623c,7e80afae-d0f6-418a-87f4-773b9fe11381
    ##  --------------------------------------------------
    foreach($site in (Get-RestAPIResonse `
                        -APIMethod Get `
                        -APIVersion v1.0 `
                        -APIResource "sites/$($global:TenantName)-my.sharepoint.com:/$($sitePath)")) {


        foreach($list in (Get-RestAPIResonse `
                            -APIMethod Get `
                            -APIVersion v1.0 `
                            -APIResource "sites/$($site.id)/lists/?`$select=id,name,webUrl,description").value) {
            $list | fl
        }
    }
}
catch [Exception] {
    Write-Error $_.Exception.Message
}