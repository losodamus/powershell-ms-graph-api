<#
    Add-SPOLibraryBatchOfFoldersNested.ps1
    ------------------------------


    Create series of nested folders via batch submission. 
#>
cls


##  Variable(s)
$global:AppGuid = "<app_id>"
$global:AppSecret = "<app_secret>"
$global:TenantName = "<tenant_name>"


[System.String] $sitePath = "/sites/Portal"
[System.String] $librPath = "/Shared%20Documents"


[System.Array] $batchOfFolders = @()
[System.Array] $listOfFolders = @(
    "Folder 001", "Folder 002", "Folder 003", 
    "Folder 004", "Folder 005"
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
        Write-Host "Site GUID:`t$($site.id)"


        ##  E.g., b!mPJrB6couEmWWJyXluNiPK6vgH720IpBh_R404_hE4H0-fEDWHMKQ6oR_pSl_atl
        ##  --------------------------------------------------
        foreach($libr in (Get-RestAPIResonse `
                            -APIMethod Get `
                            -APIVersion v1.0 `
                            -APIResource "sites/$($site.id)/drives/?`$select=id,name,webUrl").value | ? { $_.webUrl -like "*$($librPath)" }) {
            Write-Host "Libr GUID:`t$($libr.id)"


            ##  Root folder
            $batchOfFolders += @{
                "url" = "/sites/$($site.id)/drives/$($libr.id)/root/children"
                "method" = "POST"
                "id" = "1"
                "body" = @{
                    "name" = "$($listOfFolders[0])"
                    "folder" = @{}
                    "@microsoft.graph.conflictBehavior" = "rename"
                }
                "headers" = @{
                    "Content-Type" = "application/json"
                }
            }


            ##  Child-folder #1
            $batchOfFolders += @{
                "url" = "/sites/$($site.id)/drives/$($libr.id)/root:/$($listOfFolders[0]):/children".Replace(" ", "%20")
                "dependsOn" = @("1")
                "method" = "POST"
                "id" = "2"
                "body" = @{
                    "name" = "$($listOfFolders[1])"
                    "folder" = @{}
                    "@microsoft.graph.conflictBehavior" = "rename"
                }
                "headers" = @{
                    "Content-Type" = "application/json"
                }
            }


            ##  Child-folder #2
            $batchOfFolders += @{
                "url" = "/sites/$($site.id)/drives/$($libr.id)/root:/$($listOfFolders[0])/$($listOfFolders[1]):/children".Replace(" ", "%20")
                "dependsOn" = @("2")
                "method" = "POST"
                "id" = "3"
                "body" = @{
                    "name" = "$($listOfFolders[2])"
                    "folder" = @{}
                    "@microsoft.graph.conflictBehavior" = "rename"
                }
                "headers" = @{
                    "Content-Type" = "application/json"
                }
            }


            ##  Child-folder #3
            $batchOfFolders += @{
                "url" = "/sites/$($site.id)/drives/$($libr.id)/root:/$($listOfFolders[0])/$($listOfFolders[1])/$($listOfFolders[2]):/children".Replace(" ", "%20")
                "dependsOn" = @("3")
                "method" = "POST"
                "id" = "4"
                "body" = @{
                    "name" = "$($listOfFolders[3])"
                    "folder" = @{}
                    "@microsoft.graph.conflictBehavior" = "rename"
                }
                "headers" = @{
                    "Content-Type" = "application/json"
                }
            }


            ##  Child-folder #4
            $batchOfFolders += @{
                "url" = "/sites/$($site.id)/drives/$($libr.id)/root:/$($listOfFolders[0])/$($listOfFolders[1])/$($listOfFolders[2])/$($listOfFolders[3]):/children".Replace(" ", "%20")
                "dependsOn" = @("4")
                "method" = "POST"
                "id" = "5"
                "body" = @{
                    "name" = "$($listOfFolders[4])"
                    "folder" = @{}
                    "@microsoft.graph.conflictBehavior" = "rename"
                }
                "headers" = @{
                    "Content-Type" = "application/json"
                }
            }


            Get-RestAPIBatchResponse `
                -APIBatch $batchOfFolders
        }
    }
}
catch [Exception] {
    Write-Error $_.Exception.Message
}