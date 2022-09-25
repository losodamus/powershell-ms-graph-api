<#
    Add-SPOLibrFolderSingle.ps1
    ------------------------------


    - Populate script variables.
    - Query target SharePoint site GUID.
    - Query target document library GUID.
    - Create folder at root of SPO library. 
#>
Clear-Host


##  Variable(s)
$global:AppGuid = "<app_id>"
$global:AppSecret = "<app_secret>"
$global:TenantName = "<tenant_name>"


[System.String] $sitePath = "/sites/Portal"
[System.String] $librPath = "/Shared%20Documents"
[System.String] $nameOfFolder = "TEST"


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
        [ValidateSet('Get', 'Post')] 
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
        if ($APIBody.Count -eq 0) {
            $r = Invoke-RestMethod `
                -Uri "https://graph.microsoft.com/$($APIVersion)/$($APIResource)" `
                -Headers (Get-RestAPIHeader) `
                -Method $APIMethod `
                -ContentType "application/json"
        }
        else {
            $jsonBody = $APIBody | ConvertTo-Json -Depth 5
            $jsonBody | Out-Host


            $r = Invoke-RestMethod `
                -Uri "https://graph.microsoft.com/$($APIVersion)/$($APIResource)" `
                -Headers (Get-RestAPIHeader) `
                -Method $APIMethod `
                -ContentType "application/json" `
                -Body $jsonBody
        }
        return $r
    }
    catch [Exception] {
        return $null
    }
}
Clear-Host


try {
    ##  E.g., <tenant_name>.sharepoint.com,076bf298-a1b2-49b8-9658-9c9796e3623c,7e80afae-d0f6-abcd-1234-773b9fe11381
    ##  --------------------------------------------------
    foreach($s in (Get-RestAPIResonse `
                        -APIMethod Get `
                        -APIVersion v1.0 `
                        -APIBody @{} `
                        -APIResource "sites/$($global:TenantName).sharepoint.com:/$($sitePath)/?`$select=id,name,webUrl")) {
        Write-Host -F Cyan "Site:`t$($s.id)"


        ##  E.g., b!mPJrB6couEmWWJyXluNiPK6vgH720IpBh_R404_hE4H0-fEDWHMKQ6oR_pSl_atl
        ##  --------------------------------------------------
        foreach($l in (Get-RestAPIResonse `
                            -APIMethod Get `
                            -APIVersion v1.0 `
                            -APIBody @{} `
                            -APIResource "sites/$($s.id)/drives/?`$select=id,name,webUrl").value `
                            | ? { $_.webUrl -like "*$($librPath)" }) {
            Write-Host -F Magenta "Libr:`t$($l.id)"


            ##  Create Folder
            Get-RestAPIResonse `
                -APIMethod Post `
                -APIVersion v1.0 `
                -APIResource "sites/$($s.id)/drives/$($l.id)/root/children" `
                -APIBody (@{
                    "name" = "$nameOfFolder"
                    "folder" = @{ }
                    "@microsoft.graph.conflictBehavior" = "rename"
                })
        }
    }
}
catch [Exception] {
    Write-Host -F Red "`nError Message:"
    Write-Host -F Red $_.Exception.Message
}