<#
    Get-AzureUserListWithRecurision.ps1
    ------------------------------


    Loop through and output tenant users using the 
    Microsoft Graph API and the '@odata.nextLink' parameter. 
#>


##  Variable(s)
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


function Get-RecursiveListOfUsers {
    param(
        [System.String] $GraphAPIEndpoint
    )


    try {
        $GraphReq = Invoke-RestMethod `
            -Uri $GraphAPIEndpoint `
            -Headers (Get-RequestAPIHeader) `
            -Method GET `
            -ContentType "application/json"


        $GraphAPIEndpoint = ""
        if ($GraphReq.'@odata.nextLink'.Length -ne 0) {
            $GraphAPIEndpoint = $GraphReq.'@odata.nextLink'
        }


        $GraphReq.value | % {
            Write-Host -F Green "$($_.displayName)"
        }
    }
    catch [Exception] {
        Write-Host -F Red $_.Exception.Message
    }


    ##  Return next batch of users
    if ($GraphAPIEndpoint -ne "") {
        Get-RecursiveListOfUsers `
            -GraphAPIEndpoint "$GraphAPIEndpoint"
    }
}
cls


try {
    Get-RecursiveListOfUsers `
        -GraphAPIEndpoint "https://graph.microsoft.com/v1.0/users/?`$select=displayName&`$orderby=displayName&`$top=100"
}
catch [Exception] {
    Write-Host -F Red $_.Exception.Message
}
