<#
    Copy-SPOFileFromAndTo.ps1
    ------------------------------


    Copy file from one SPO document 
    library to another document 
    library in the site. 
#>


##  Variable(s)
$global:AppGuid = "<app_id>"
$global:AppSecret = "<app_secret>"
$global:TenantName = "<tenant_name>"


[System.String] $sourceSitePath = "/sites/Portal"
[System.String] $sourceLibrPath = "/Source"  ##  Relative path of document library
[System.String] $targetLibrPath = "/Target"  ##  Relative path of document library


[System.String] $sourceSiteGuid = ""
[System.String] $sourceLibrGuid = ""
[System.String] $targetLibrGuid = ""


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


function Get-SPOSiteGUID () {
    param(
        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [System.String] $RelativeSitePath
    )


    try {
        $r = (Get-RestAPIResonse `
            -APIMethod Get `
            -APIVersion v1.0 `
            -APIResource "sites/$($global:TenantName).sharepoint.com:/$($RelativeSitePath)?`$top=1")
        $r.id
    }
    catch [Exception] {
        return ""
    }
}


function Get-SPOSiteLibrGUID () {
    param(
        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [System.String] $SPOSiteGuid,


        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [System.String] $SPOLibrPath
    )


    try {
        $g = ""
        $r = (Get-RestAPIResonse `
            -APIMethod Get `
            -APIVersion v1.0 `
            -APIResource "sites/$($SPOSiteGuid)/drives/?`$select=id,name,webUrl")


        $r.value | ? { $_.webUrl -like "*$($SPOLibrPath)" } | % {
            $g = "$($_.id)"
        }
        return $g
    }
    catch [Exception] {
        return ""
    }
}


function Get-SPOSiteLibrPathGUID () {
    param(
        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [System.String] $SPOSiteGuid,


        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [System.String] $SPOLibrGuid
    )


    try {
        $g = ""
        $r = (Get-RestAPIResonse `
            -APIMethod Get `
            -APIVersion v1.0 `
            -APIResource "sites/$($SPOSiteGuid)/drives/$($SPOLibrGuid)/root?`$select=id,name,webUrl")


        $r | % {
            $g = "$($_.id)"
        }
        return $g
    }
    catch [Exception] {
        return ""
    }
}


function Copy-SPOSiteLibrItemsFromAndTo () {
    param(
        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [System.String] $SourceSPOSiteGuid,


        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [System.String] $SourceSPOLibrGuid,


        [Parameter(Mandatory)] 
        [ValidateNotNullOrEmpty()] 
        [System.String] $TargetSPOLibrGuid
    )


    try {


        ##  E.g., 01DA4TVFN6Y2GOVW7725BZO770PWSELRRZ
        ##  --------------------------------------------------
        $g = (Get-SPOSiteLibrPathGUID `
            -SPOSiteGuid $SourceSPOSiteGuid `
            -SPOLibrGuid $TargetSPOLibrGuid)


        if ($g -ne "") {
            $r = (Get-RestAPIResonse `
                -APIMethod Get `
                -APIVersion v1.0 `
                -APIResource "sites/$($SourceSPOSiteGuid)/drives/$($SourceSPOLibrGuid)/root/children/?`$select=id,name")


            ##  ------------------------------
            ##  If not COPY, but MOVE:
            ##  "method" = "PATCH"
            ##  "url" = "/sites/$($SourceSPOSiteGuid)/drives/$($SourceSPOLibrGuid)/items/$($SPOItemID)/"
            ##  ------------------------------
            $r.value | % {


                $listOf = @()
                $listOf += @{
                    "id" = "$($listOf.Count + 1)"
                    "method" = "POST"
                    "url" = "/sites/$($SourceSPOSiteGuid)/drives/$($SourceSPOLibrGuid)/items/$($_.id)/copy"
                    "headers" = @{
                        "Content-Type" = "application/json"
                    }
                    "body" = @{
                        "name" = "$($_.name)"
                        "parentReference"= @{
                            "driveId" = "$($TargetSPOLibrGuid)"
                            "id" = "$($g)";
                        }
                    }
                }


                Get-RestAPIBatchResponse `
                    -APIBatch $listOf
            }
        }
    }
    catch [Exception] {
        Write-Error $_.Exception.Message
    }
}
cls


try {


    ##  E.g., <tenant_name>.sharepoint.com,076bf298-a1b2-49b8-9658-9c9796e3623c,7e80afae-d0f6-418a-87f4-773b9fe11381
    ##  --------------------------------------------------
    $sourceSiteGuid = (Get-SPOSiteGUID -RelativeSitePath $sourceSitePath)
    $sourceSiteGuid


    if ($sourceSiteGuid -ne "") {


        ##  E.g., b!mPJrB6couEmWWJyXlzZiPK6vgH720IpBh_R404_hE4H4zFEHPdfVQIEVyqdL_iyY
        ##  --------------------------------------------------
        $sourceLibrGuid = "$(Get-SPOSiteLibrGUID -SPOSiteGuid $sourceSiteGuid -SPOLibrPath $sourceLibrPath)"
        $targetLibrGuid = "$(Get-SPOSiteLibrGUID -SPOSiteGuid $sourceSiteGuid -SPOLibrPath $targetLibrPath)"
        $sourceLibrGuid
        $targetLibrGuid


        if ($sourceLibrGuid -ne "" -and `
            $targetLibrGuid -ne "") {
        

            Copy-SPOSiteLibrItemsFromAndTo `
                -SourceSPOSiteGuid $sourceSiteGuid `
                -SourceSPOLibrGuid $sourceLibrGuid `
                -TargetSPOLibrGuid $targetLibrGuid
        }
    }
}
catch [Exception] {
    Write-Error $_.Exception.Message
}
