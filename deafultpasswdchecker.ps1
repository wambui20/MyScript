# Define the list of default passwords
$defaultPasswords = @("password@12345", "Password@12345", "password@23456") 

# Specify the OU distinguished name (DN) where you want to check users
$ouDN = "DC=ELMA,DC=LOCAL"  # Update with your OU's DN

# Specify the output file path
$outputFilePath = "C:\Users\chelsea.wambui\Documents\default_passwords2.txt"

try {
    # Get a list of user accounts from the specified OU
    $userAccounts = Get-ADUser -Filter * -SearchBase $ouDN -ErrorAction Stop | Where-Object { $_.Enabled -eq $true }

    $defaultPasswordsFound = $false
    $userInformation = @()

    # Loop through each user account
    foreach ($user in $userAccounts) {
        $userName = $user.SamAccountName
        $userSID = $user.SID

        # Check if the password is in the list of default passwords
        foreach ($defaultPassword in $defaultPasswords) {
            $credential = New-Object System.Management.Automation.PSCredential ($userName, (ConvertTo-SecureString -String $defaultPassword -AsPlainText -Force))
            if ($user | Get-ADUser -Credential $credential -ErrorAction SilentlyContinue) {
                $defaultPasswordsFound = $true

                # Add user information to the array
                $userInfo = "User: $userName ($userSID)"
                $userInformation += $userInfo
                break
            }
        }
    }

    # Output user information to the file
    if ($defaultPasswordsFound) {
        $userInformation | Out-File -FilePath $outputFilePath -Encoding utf8
        Write-Host "User information has been saved to $outputFilePath"
    } else {
        Write-Host "No users found with default passwords."
    }
	
	
	 # Output the number of user accounts processed
    $usersWithDefaultPasswordsCount = $userInformation.Count
    Write-Host "Number of users with default passwords: $usersWithDefaultPasswordsCount"
	
} catch {
    Write-Host "Error occurred: $_"
}
