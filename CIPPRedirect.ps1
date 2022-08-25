using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Interact with query parameters or the body of the request.
$email = $Request.Query.email
if (-not $email) {
    $email = $Request.Body.email
}
$Compromise = $Request.Query.Compromise
if (-not $Compromise) {
    $Compromise = $Request.Body.Compromise
}

if (!$email) {
    $body = "Error no Email Address."
    # Associate values to output bindings by calling 'Push-OutputBinding'.
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = $body
    })
    exit 0
}

if ($email) {
    $UPN = $null
    $CIPPURL = $env:CIPPURL
    $data = import-csv "cippredirect\Office365UsersReport.csv" -delimiter ","
    $table = $data | Group-Object -AsHashTable -AsString -Property UserPrincipalName
    $UPN = $table.Get_Item($email)

    if (!$upn) {
        $originalEmail = $email
        $table = $data | Group-Object -AsHashTable -AsString -Property "Proxy Addresses"
        $table.keys | % { if($_.contains($email)){$search = $_}}

        if ($search) {
            $UPN = $table.Get_Item($search)
            $email = $UPN.UserPrincipalName
        }
    }
    
    #CIPP - View User
    if ($email -eq $UPN.UserPrincipalName) {
        write-host Sending the URL to View the user in CIPP
        $RidirectURL = "$($CIPPURL).auth/login/aad?post_login_redirect_uri=$($CIPPURL)identity/administration/users/view?userId=$($UPN.ObjectId)%26tenantDomain%3D$($UPN.DefaultDomain)"    
    }    
    
    #CIPP - Research Compromise
    if ($Compromise) {
        write-host User Exists in this table        
        $RidirectURL = "$($CIPPURL).auth/login/aad?post_login_redirect_uri=$($CIPPURL)identity/administration/ViewBec?userId=$($UPN.ObjectId)%26tenantDomain%3D$($UPN.DefaultDomain)"
    }        
}


if ($RidirectURL){
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{ 
        StatusCode = [HttpStatusCode]::MovedPermanently
        Headers = @{
            'Location' = "$($RidirectURL)"
        }
        Body = "$body"
    })
} else {
    $body = "We cannot access this contact.  They may not have 365 or there may be an issue. If you think this is by mistake contact Support."

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = $body
    })
}


