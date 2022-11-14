<#
    Get-AzureUserListWithRecurision.ps1
    ------------------------------


    - Populate script variables.
    - Query Azure AD users.
    - Use the '@odata.nextLink' property value for recursion.
#>


##  Variable(s)
$global:AppGuid = "<app_id>"          ##  E.g., "468a8e5d-4040-1a3c-5e7g-9a668ebeef07"
$global:AppSecret = "<app_secret>"    ##  E.g., "b-fj928rjf02hm8r2y398r,2^~448jgergkgj34g"
$global:TenantName = "<tenant_name>"  ##  E.g., "contoso"


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
        Uri = "https://login.microsoftonline.com/$($global:TenantName).onmicrosoft.com/oauth2/v2.0/token"
    }
    $Request = Invoke-RestMethod @PostSplat


    return (@{
        Authorization = "$($Request.token_type) $($Request.access_token)"
    })
}


function Get-AzureADListOfUsers {
    param(
        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [System.String] $GraphAPIRequest
    )


    [System.String] $nextLink = ""
    try {
        $GraphReq = Invoke-RestMethod `
            -Uri "$GraphAPIRequest" `
            -Headers (Get-RequestAPIHeader) `
            -Method GET `
            -ContentType "application/json"


        ##  Get next batch link
        if ($GraphReq.'@odata.nextLink'.Length -ne 0) {
            $nextLink = $GraphReq.'@odata.nextLink'
        }


        ##  Output users
        $GraphReq.value | % {
            Write-Host -F Green "$($_.displayName)"
        }
    }
    catch [Exception] {
        Write-Host -F Red $_.Exception.Message
    }


    ##  Return next batch
    if ($nextLink -ne "") {
        Get-AzureADListOfUsers `
            -GraphAPIRequest "$nextLink"
    }
}
Clear-Host


try {
    Get-AzureADListOfUsers `
        -GraphAPIRequest "https://graph.microsoft.com/v1.0/users/?`$orderby=displayName&`$top=100"
}
catch [Exception] {
    Write-Host -F Red $_.Exception.Message
}
