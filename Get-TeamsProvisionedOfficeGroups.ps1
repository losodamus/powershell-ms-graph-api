<#
    Get-TeamsProvisionedOfficeGroups.ps1
    ------------------------------


   Query, filter and output list of Office 365 
   groups provisioned through Microsoft Teams. 
#>


$global:AppGuid = "<app_id>"
$global:AppSecret = "<app_secret>"
$global:TenantName = "<tenant_name>.onmicrosoft.com"


function Get-RequestAPIHeader {
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
        Uri = "https://login.microsoftonline.com/$($global:TenantName)/oauth2/v2.0/token"
    }
    $Request = Invoke-RestMethod @PostSplat


    return (@{
        Authorization = "$($Request.token_type) $($Request.access_token)"
    })
}
cls


try {
    $v = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/groups/?`$select=displayName,creationOptions&`$orderby=displayName" `
        -Headers (Get-RequestAPIHeader) `
        -Method GET `
        -ContentType "application/json"


    $v.value | ? { $_.creationOptions -like "*Team*" } | % {
        $_
    }
}
catch [Exception] {
    Write-Host -F Red $_.Exception.Message
}