
#PowerShell-Graph-Appflow-Send-10-emails.ps1 

<# 
-----------------------------------------------------------------------------
Raw Graph + App-only OAuth (client_credentials)
Sends 10 emails, adds counter + timestamp to subject.
You can adjust the number of emails to send by changing $SendTotal.

Generated initialy with CoPilot then modified.
-----------------------------------------------------------------------------
Create an appliciton flow application registration with no callback.
For permissions try:
    Mail.ReadWrite
    Mail.Send
Note: Be sure to do an Admin grant on these permissions.
Note: You will need to gather the Tenant/Direcytory ID, Client/Application ID and Client Secret Value (not it's ID) to use in this sample.
-----------------------------------------------------------------------------
References:
- OAuth2 client credentials flow (app-only): https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow
- Graph sendMail endpoint + JSON body shape: https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0
- Send Outlook messages from another user:  https://learn.microsoft.com/en-us/graph/outlook-send-mail-from-other-user    
- Get Outlook messages in a shared or delegated folder: https://learn.microsoft.com/en-us/graph/outlook-share-messages-folders
-----------------------------------------------------------------------------
#>

# ----------------------------
# CONFIG (fill these in)
# ----------------------------
$TenantId     = "<TENANT_ID_GUID_OR_DOMAIN>"    # TODO: Change to the Tenant/Directory ID
$ClientId     = "<APP_CLIENT_ID>"               # TODO: Change to the Client/Application ID
$ClientSecret = "<APP_CLIENT_SECRET>"           # TODO: Change to the VALUE of the Client Secret

# Set the total number of emails to send
$SendTotal = 10                                # TODO: Change to how many emails you want to send

# The mailbox you want to send *as* (must exist and be allowed for the app)
$SenderUpn    = "sender@contoso.com"            # TODO: Change to sending UPN

# Recipient
$ToAddress    = "recipient@contoso.com"        # TODO: Change to recipient SMTP Address

# Optional: slow down to avoid bursts
$SleepMsBetweenSends = 0                       # TODO: Change to time between calls- e.g. 250  
 
# ----------------------------
# Get app-only access token
# ----------------------------
function Get-GraphAppToken {
    param(
        [Parameter(Mandatory)] [string] $TenantId,
        [Parameter(Mandatory)] [string] $ClientId,
        [Parameter(Mandatory)] [string] $ClientSecret
    )

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" # [1](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow)

    # client_credentials with .default scope is the common pattern for Graph app-only tokens [2](https://teams.microsoft.com/l/message/19:4c34f18b-61ac-4976-b641-dcf33d3002d8_d34de66d-3d2c-4c6b-bec5-3bc5ddb0f6d5@unq.gbl.spaces/1738258015213?context=%7B%22contextType%22:%22chat%22%7D)[1](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow)
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
    }

    $tokenResponse = Invoke-RestMethod -Method POST -Uri $tokenEndpoint `
        -ContentType "application/x-www-form-urlencoded" -Body $body

    return $tokenResponse.access_token
}

# ----------------------------
# Send mail via Graph (raw REST)
# POST /users/{id|userPrincipalName}/sendMail [3](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0)
# ----------------------------
function Send-GraphMail {
    param(
        [Parameter(Mandatory)] [string] $AccessToken,
        [Parameter(Mandatory)] [string] $SenderUpn,
        [Parameter(Mandatory)] [string] $ToAddress,
        [Parameter(Mandatory)] [string] $Subject,
        [Parameter(Mandatory)] [string] $BodyText
    )

    $uri = "https://graph.microsoft.com/v1.0/users/$SenderUpn/sendMail" # [3](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0)

    $payload = @{
        message = @{
            subject = $Subject
            body    = @{
                contentType = "Text"
                content     = $BodyText
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $ToAddress
                    }
                }
            )
        }
        # saveToSentItems defaults to true (only specify if false) [3](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0)
    } | ConvertTo-Json -Depth 10

    $headers = @{
        Authorization = "Bearer $AccessToken"  # required [3](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0)
        "Content-Type" = "application/json"
    }

    # On success, Graph returns 202 Accepted and no body [3](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0)
    Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $payload | Out-Null
}

# ----------------------------
# MAIN
# ----------------------------
$token = Get-GraphAppToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

 
for ($i = 1; $i -le $SendTotal; $i++) {

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $subject = "Test Email [$i/$SendTotal] $ts"
    $body    = "Hello! This is test message $i of $SendTotal sent at $ts."

    try {
        Send-GraphMail -AccessToken $token -SenderUpn $SenderUpn -ToAddress $ToAddress -Subject $subject -BodyText $body
        Write-Host "[$(Get-Date -Format HH:mm:ss)] Sent $i/$SendTotal : $subject"
    }
    catch {
        Write-Warning "Failed $i/$SendTotal : $($_.Exception.Message)"
    }

    if ($SleepMsBetweenSends -gt 0) { Start-Sleep -Milliseconds $SleepMsBetweenSends }
}
