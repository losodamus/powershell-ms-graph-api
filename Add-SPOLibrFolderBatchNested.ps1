<#
    Add-SPOLibrFolderBatchNested.ps1
    ------------------------------


    - Populate script variables.
    - Query target SharePoint site GUID.
    - Query target document library GUID.
    - Create batch of nested folders at root with "dependsOn" property.
#>


##  Variable(s)
$global:AppGuid = "<app_id>"
$global:AppSecret = "<app_secret>"
$global:TenantName = "<tenant_name>"


[System.String] $sitePath = "/sites/Portal"
[System.String] $librPath = "/Shared%20Documents"
[System.String] $folderPath = ""


[System.Int32] $indexOf = 1
[System.Array] $batchOf = @()
[System.Array] $listOfFolders = @(
    "Folder 001", "Folder 002", "Folder 003", 
    "Folder 004", "Folder 005", "Folder 006",
    "Folder 007", "Folder 008", "Folder 009", 
    "Folder 010"
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
Clear-Host


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
            foreach($folder in ($listOfFolders.GetEnumerator() `
                | Sort -Property Name)) {
                Write-Host -F Yellow "Path:`t$($folder)`t|`t$($folderPath)"


                $temp = @{
                    "url" = "/sites/$($s.id)/drives/$($l.id)/root/children"
                    "dependsOn" = @("$($indexOf - 1)")
                    "method" = "POST"
                    "id" = "$($indexOf)"
                    "body" = @{
                        "name" = "$($folder)"
                        "folder" = @{}
                        "@microsoft.graph.conflictBehavior" = "rename"
                    }
                    "headers" = @{
                        "Content-Type" = "application/json"
                    }
                }


                ##  If Folder Nested
                if ($folderPath -ne "") { $temp["url"] = "/sites/$($s.id)/drives/$($l.id)/root:$($folderPath):/children".Replace(" ", "%20") }
                $folderPath += "/$($folder)"


                ##  Request #1 - No Dependency
                if ($indexOf -eq 1) {
                    $temp.Remove("dependsOn")
                }
                $batchOf += $temp
                $indexOf++
            }


            Get-RestAPIBatchResponse `
                -APIBatch $batchOf
        }
    }
}
catch [Exception] {
    Write-Host -F Red "`nError Message:"
    Write-Host -F Red $_.Exception.Message
}
