# Use MS O365 to send email.   
# Requires MS Graph API and Mail.Send permission with type=application (not delegated) and admin constent for all access.
# O365 Azure/Entra Admin UI:  https://entra.microsoft.com/ Applications/App_Registration/New Registration
# $URLsend = "https://graph.microsoft.com/v1.0/users/$MailSender/sendMail"
# Ref.  https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http#request-headers

# Update Variables for OAuth token request
$tenantID = "1ae30954-4635-4348-86a0-72dc6134497e"            # Replace with your Tenant ID - ANA Tenant
$clientID = "dd1c94ec-7a98-4cd6-b64d-c318646fed8d"            # Replace with your Client ID - Tied to Application.
$clientSecret = "XXXXXXXXXXXXXXXXXXXXXXXX"                    # Replace with your Client Secret - Generate a new one. Expire every 6 mo

# Define the sender email address.
$MailSender = "alan.baugher@anapartner.com"
$MailRecipient = "alan.baugher@anapartner.com"


$HiddenHTMLMessage = "
<!-- Hidden metadata starts here -->
<!-- 
Metadata:
- X-AAAA: CustomValue1
- X-BBBB: CustomValue2
-->
<!-- Hidden metadata ends here -->
"

$CustomHeaderHere = "ALAN_WAS_HERE_CustomHeaderValue"  # Custom header X-AAAA

###################################################################

# Connect to GRAPH API
$tokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientID
    Client_Secret = $clientSecret
}
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $tokenBody
$accessToken = $tokenResponse.access_token


# Custom headers, including X-AAAA
$headers = @{
    "Authorization" = "Bearer $accessToken"
    "Content-type"  = "application/json"
    "X-AAAA"        = "$CustomHeaderHere"  # Custom header X-AAAA
}


##########################################################################
# Send Mail    
$URLsend = "https://graph.microsoft.com/v1.0/users/$MailSender/sendMail"
$BodyJsonsend = @"
{
    "message": {
      "subject": "Hello World from Microsoft Graph API",
      "body": {
        "contentType": "HTML",
        "content": "This Mail is sent via Microsoft <br> GRAPH <br> API<br>
		
$HiddenHTMLMessage
        "
      },
      "toRecipients": [
        {
          "emailAddress": {
            "address": "$MailRecipient"
          }
        }
      ],
	     "internetMessageHeaders": [
      {
        "name": "x-custom-header-group-name",
        "value": "Nevada"
      },
      {
        "name": "x-custom-header-group-id",
        "value": "NV001"
      }
    ]
	  	  
	  
    },	
	
    "saveToSentItems": "false"
}
"@


##########################################################################
# Invoke REST method to send the email
try {
    Write-Host "Attempting to send the email..."
    $response = Invoke-RestMethod -Method POST -Uri $URLsend -Headers $headers -Body $BodyJsonsend
    Write-Host "Email sent successfully!"
} catch {
    Write-Host "API Request Failed!"
    Write-Host "Error: " $_.Exception.Message

    # Capture and log response if available
    if ($_.Exception.Response) {
        $responseStream = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($responseStream)
        $responseBody = $reader.ReadToEnd()
        Write-Host "Error Response: $responseBody"
    }
}
