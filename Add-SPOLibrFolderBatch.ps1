<#
    Add-SPOLibrFolderBatch.ps1
    ------------------------------


    - Populate script variables.
    - Query target SharePoint site GUID.
    - Query target document library GUID.
    - Create batch of folders at root of SPO library.
#>
cls


##  Variable(s)
$global:AppGuid = "<app_id>"
$global:AppSecret = "<app_secret>"
$global:TenantName = "<tenant_name>"


[System.String] $sitePath = "/sites/Portal"
[System.String] $librPath = "/Shared%20Documents"


[System.Int32] $batchIndex = 0
[System.Array] $batchOfFolders = @()
[System.Array] $listOfFolders = @(
    "Batch Folder 001", "Batch Folder 002", "Batch Folder 003", "Batch Folder 004", "Batch Folder 005",
    "Batch Folder 006", "Batch Folder 007", "Batch Folder 008", "Batch Folder 009", "Batch Folder 010",
    "Batch Folder 011", "Batch Folder 012", "Batch Folder 013", "Batch Folder 014", "Batch Folder 015",
    "Batch Folder 016", "Batch Folder 017", "Batch Folder 018", "Batch Folder 019", "Batch Folder 020"
)


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
cls


try {
    ##  E.g., <tenant_name>.sharepoint.com,076bf298-a1b2-49b8-9658-9c9796e3623c,7e80afae-d0f6-abcd-1234-773b9fe11381
    ##  --------------------------------------------------
    foreach($s in (Get-RestAPIResonse `
                        -APIMethod Get `
                        -APIVersion v1.0 `
                        -APIResource "sites/$($global:TenantName).sharepoint.com:/$($sitePath)/?`$select=id,name,webUrl")) {
        Write-Host -F Cyan "Site:`t$($s.id)"


        ##  E.g., b!mPJrB6couEmWWJyXluNiPK6vgH720IpBh_R404_hE4H0-fEDWHMKQ6oR_pSl_atl
        ##  --------------------------------------------------
        foreach($l in (Get-RestAPIResonse `
                            -APIMethod Get `
                            -APIVersion v1.0 `
                            -APIResource "sites/$($s.id)/drives/?`$select=id,name,webUrl").value `
                            | ? { $_.webUrl -like "*$($librPath)" }) {
            Write-Host -F Magenta "Libr:`t$($l.id)"


            ##  Enumerate folder array
            foreach($folder in $listOfFolders) {


                $batchIndex += 1
                $batchOfFolders += @{
                    "url" = "/sites/$($s.id)/drives/$($l.id)/root/children"
                    "method" = "POST"
                    "id" = "$($batchIndex)"
                    "body" = @{
                        "name" = "$($folder)"
                        "folder" = @{}
                        "@microsoft.graph.conflictBehavior" = "rename"
                    }
                    "headers" = @{
                        "Content-Type" = "application/json"
                    }
                }
            }


            Get-RestAPIBatchResponse `
                -APIBatch $batchOfFolders
        }
    }
}
catch [Exception] {
    Write-Error "`nError Message:"
    Write-Error $_.Exception.Message
}
