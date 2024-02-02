# Useful script

## Register application in Entra ID
![image](https://github.com/guguji666666/O365/assets/96930989/3077e12e-6b48-4451-bda6-878945c7f262)


## List all email subjects and output in grid view
```powershell
# Replace these values with your app details
$client_id = '<your client id>'
$client_secret = '<your client secret value>'
$tenant_id = '<your tenant id>'


# Get access token
$tokenEndpoint = "https://login.microsoftonline.com/$tenant_id/oauth2/v2.0/token"
$tokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $client_id
    client_secret = $client_secret
    scope         = "https://graph.microsoft.com/.default"
}

try {
    $tokenResponse = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $tokenBody -ErrorAction Stop

    # Get all users in the tenant
    $usersEndpoint = "https://graph.microsoft.com/v1.0/users"
    $users = Invoke-RestMethod -Uri $usersEndpoint -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} | Select-Object -ExpandProperty value

    # Create an array to store email subjects
    $emailSubjects = @()

    # Display email subjects for each user
    foreach ($user in $users) {
        $userId = $user.id
        $userEmail = $user.userPrincipalName

        # Initialize variables for pagination
        $pageSize = 50
        $pageNumber = 1
        $allMessagesRetrieved = $false

        do {
            # Get a page of emails for the user
            $searchEndpoint = "https://graph.microsoft.com/v1.0/users/$userId/messages"
            $searchParams = @{
                '$top' = $pageSize
                '$skip' = ($pageNumber - 1) * $pageSize
            }

            try {
                $searchResponse = Invoke-RestMethod -Uri $searchEndpoint -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get -Body $searchParams -ErrorAction Stop

                # Add email subjects to the array
                foreach ($email in $searchResponse.value) {
                    $emailSubjects += [PSCustomObject]@{
                        UserEmail = $userEmail
                        Subject = $email.subject
                    }
                }

                # Check if there are more messages
                if ($searchResponse.'@odata.nextLink') {
                    $pageNumber++
                } else {
                    $allMessagesRetrieved = $true
                }
            }
            catch {
                Write-Host "Error during email retrieval for user $userEmail. Error: $_"
                $allMessagesRetrieved = $true  # Stop the loop on error
            }
        } while (-not $allMessagesRetrieved)
    }

    # Display email subjects in a grid view
    $emailSubjects | Out-GridView -Title "Email Subjects"
}
catch {
    Write-Host "Failed to obtain access token. Error: $_"
}
```

## Delete emails that contains specified keyword in subject
```powershell
# Replace these values with your app details
$client_id = '<your client id>'
$client_secret = '<your client secret value>'
$tenant_id = '<your tenant id>'
$keyword = '<YourKeyword in email subject>'

# Get access token
$tokenEndpoint = "https://login.microsoftonline.com/$tenant_id/oauth2/v2.0/token"
$tokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $client_id
    client_secret = $client_secret
    scope         = "https://graph.microsoft.com/.default"
}

try {
    $tokenResponse = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $tokenBody -ErrorAction Stop

    # Get all users in the tenant
    $usersEndpoint = "https://graph.microsoft.com/v1.0/users"
    $users = Invoke-RestMethod -Uri $usersEndpoint -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} | Select-Object -ExpandProperty value

    # Delete emails with subject containing the specified keyword for each user
    foreach ($user in $users) {
        $userId = $user.id
        $userEmail = $user.userPrincipalName

        # Initialize variables for pagination
        $pageSize = 50
        $pageNumber = 1
        $allMessagesRetrieved = $false

        do {
            # Get a page of emails for the user
            $searchEndpoint = "https://graph.microsoft.com/v1.0/users/$userId/messages"
            $searchParams = @{
                '$top' = $pageSize
                '$skip' = ($pageNumber - 1) * $pageSize
                '$filter' = "contains(subject,'$keyword')"
            }

            try {
                $searchResponse = Invoke-RestMethod -Uri $searchEndpoint -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get -Body $searchParams -ErrorAction Stop

                # Delete emails
                foreach ($email in $searchResponse.value) {
                    $emailId = $email.id
                    $deleteEndpoint = "https://graph.microsoft.com/v1.0/users/$userId/messages/$emailId"
                    
                    try {
                        $deleteResponse = Invoke-RestMethod -Uri $deleteEndpoint -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Delete -ErrorAction Stop

                        if ($deleteResponse.StatusCode -eq 204) {
                            Write-Host "Email with ID $emailId deleted successfully for user $userEmail."
                        } else {
                            Write-Host "Failed to delete email with ID $emailId for user $userEmail. Status code: $($deleteResponse.StatusCode)"
                        }
                    }
                    catch {
                        Write-Host "Error during email deletion for user $userEmail. Email ID: $emailId. Error: $_"
                    }
                }

                # Check if there are more messages
                if ($searchResponse.'@odata.nextLink') {
                    $pageNumber++
                } else {
                    $allMessagesRetrieved = $true
                }
            }
            catch {
                Write-Host "Error during email retrieval for user $userEmail. Error: $_"
                $allMessagesRetrieved = $true  # Stop the loop on error
            }
        } while (-not $allMessagesRetrieved)
    }
}
catch {
    Write-Host "Failed to obtain access token. Error: $_"
}
```
