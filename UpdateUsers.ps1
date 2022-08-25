# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format.
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' property is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"

################### Secure Application Model Information ##### 
$ApplicationId = $ENV:ApplicationId
$ApplicationSecret = $ENV:ApplicationSecret
$RefreshToken = $ENV:RefreshToken
$MyTenant = $ENV:MyTenant
# set the name of the report alongside the path ".\ relative path"
$customerExclude =@("ExampleDomain","ExmapleDomain2","ExampleDomain3")
$ReportName = "cippredirect\Office365UsersReport.csv"
######## Secrets #########

Write-Host "My Tenant $MyTenant"

function Connect-graphAPI {
    [CmdletBinding()]
    Param
    (
        [parameter(Position = 0, Mandatory = $false)]
        [ValidateNotNullOrEmpty()][String]$ApplicationId,

        [parameter(Position = 1, Mandatory = $false)]
        [ValidateNotNullOrEmpty()][String]$ApplicationSecret,

        [parameter(Position = 2, Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]$TenantId,

        [parameter(Position = 3, Mandatory = $false)]
        [ValidateNotNullOrEmpty()][String]$RefreshToken

    )
    Write-Verbose 'Removing old token if it exists'
    $Script:GraphHeader = $null
    Write-Verbose 'Logging into Graph API'
    try {
        if ($ApplicationId) {
            Write-Verbose '   using the entered credentials'
            $script:ApplicationId = $ApplicationId
            $script:ApplicationSecret = $ApplicationSecret
            $script:RefreshToken = $RefreshToken
            $AuthBody = @{
                client_id     = $ApplicationId
                client_secret = $ApplicationSecret
                scope         = 'https://graph.microsoft.com/.default'
                refresh_token = $RefreshToken
                grant_type    = 'refresh_token'
            }
        } else {
            Write-Verbose '   using the cached credentials'
            $AuthBody = @{
                client_id     = $script:ApplicationId
                client_secret = $Script:ApplicationSecret
                scope         = 'https://graph.microsoft.com/.default'
                refresh_token = $script:RefreshToken
                grant_type    = 'refresh_token'
            }
        }
        $AccessToken = (Invoke-RestMethod -Method post -Uri "https://login.microsoftonline.com/$($TenantId)/oauth2/v2.0/token" -Body $Authbody -ErrorAction Stop).access_token
        $Script:GraphHeader = @{ Authorization = "Bearer $($AccessToken)" }
    } catch {
        Write-Host "Could not log into the Graph API for tenant $($TenantID): $($_.Exception.Message)" -ForegroundColor Red
    }
}
Write-Host 'Starting test of the standard Refresh Token' -ForegroundColor Green
try {
    Write-Host 'Attempting to retrieve an Access Token' -ForegroundColor Green
    Connect-graphAPI -ApplicationId $ApplicationId -ApplicationSecret $ApplicationSecret -RefreshToken $RefreshToken -TenantID $MyTenant
} catch {
    $ErrorDetails = if ($_.ErrorDetails.Message) {
        $ErrorParts = $_.ErrorDetails.Message | ConvertFrom-Json
        "[$($ErrorParts.error)] $($ErrorParts.error_description)"
    } else {
        $_.Exception.Message
    }
    Write-Host "Unable to generate access token. The detailed error information, if returned was: $($ErrorDetails)" -ForegroundColor Red
}
try {
    Write-Host 'Attempting to retrieve all tenants you have delegated permission to' -ForegroundColor Green
    $Tenants = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/contracts?`$top=999" -Method GET -Headers $script:GraphHeader).value    
} catch {
    $ErrorDetails = if ($_.ErrorDetails.Message) {
        $ErrorParts = $_.ErrorDetails.Message | ConvertFrom-Json
        "[$($ErrorParts.error)] $($ErrorParts.error_description)"
    } else {
        $_.Exception.Message
    }
    Write-Host "Unable to retrieve tenants. The detailed error information, if returned was: $($ErrorDetails)" -ForegroundColor Red
}

$TenantCount = $Tenants.Count

$ErrorCount = 0
$ExportArray = @()
Write-Host "$TenantCount tenants found, attempting to loop through each to test access to each individual tenant" -ForegroundColor Green
# Loop through every tenant we have, and attempt to interact with it with Graph
foreach ($Tenant in $Tenants) {
    if (-Not ($customerExclude -contains $Tenant.displayName)){    

    try {
        Connect-graphAPI -ApplicationId $ApplicationId -ApplicationSecret $ApplicationSecret -RefreshToken $RefreshToken -TenantID $Tenant.customerid
        write-host $Tenant.displayName
    } catch {
        $ErrorDetails = if ($_.ErrorDetails.Message) {
            $ErrorParts = $_.ErrorDetails.Message | ConvertFrom-Json
            "[$($ErrorParts.error)] $($ErrorParts.error_description)"
        } else {
            $_.Exception.Message
        }
        Write-Host "Unable to connect to graph API for $($Tenant.defaultDomainName). The detailed error information, if returned was: $($ErrorDetails)" -ForegroundColor Red
        $ErrorCount++
        continue
    }
    try {

        $UsersDetails = (Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/users?$select=userPrincipalName,id,proxyaddresses' -Method GET -Headers $script:GraphHeader).value        
        Foreach ($UserDetail in $UsersDetails) {
            #Write-Host $UserDetail.userPrincipalName           
            $ExportArray += [PSCustomObject][Ordered]@{
                # get the office 365 user's userprincipalname
                "UserPrincipalName"              = $UserDetail.userPrincipalName            

                # get the office 365 user's display name
                #"Customer Name"                   = $Tenant.displayName

                # get the office 365 Customers Tenant
                #"Customer Tenant"                   = $Tenant.customerId

                # get the office 365 Customers Default Domain
                "DefaultDomain"                   = $tenant.defaultDomainName

                # get the office 365 user's Object ID
                "UserObjectId"                   = $UserDetail.id

                # get the office 365 user's display name
                #"Display Name"                   = $UserDetail.DisplayName

                # get the office 365 user's creation date
                #"Mail"                           = $userdetail.mail

                # get the office 365 user's proxy addresses
                #"Proxy Addresses"                = ($UserDetail.ProxyAddresses | Out-String).Trim()
                "Proxy Addresses"                = $UserDetail.ProxyAddresses -join ', '

            }
        }        
    } catch {
        $ErrorDetails = if ($_.ErrorDetails.Message) {
            $ErrorParts = $_.ErrorDetails.Message | ConvertFrom-Json
            "[$($ErrorParts.error)] $($ErrorParts.error_description)"
        } else {
            $_.Exception.Message
        }
        Write-Host "Unable to get users from $($Tenant.defaultDomainName) in Refresh Token Test. The detailed error information, if returned was: $($ErrorDetails)" -ForegroundColor Red
        $ErrorCount++
    }
}
}
Write-Host "Standard Graph Refresh Token Test: $TenantCount total tenants, with $ErrorCount failures"
Write-Host 'All Tests Finished'
$ExportArray | Export-Csv -Path $ReportName -Delimiter "," -NoTypeInformation  