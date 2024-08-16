# Define the output file path
$outputFile = "location output file"

# Clear the file if it exists (optional)
Clear-Content -Path $outputFile -ErrorAction SilentlyContinue

# Create the TaskScheduler COM object
$taskService = New-Object -ComObject "Schedule.Service"
$taskService.Connect()

# Get the root folder (tasks at root level)
$rootFolder = $taskService.GetFolder("\")
$folders = $rootFolder.GetFolders(0)

# Define a function to retrieve task details
function Get-TaskDetails($folder) {
    $tasks = $folder.GetTasks(0)
    foreach ($task in $tasks) {
        $taskName = $task.Name
        $taskPath = $folder.Path
        
        # Using COM object to get Author and UserId
        $taskAuthor = $task.RegistrationInfo.Author
        $taskUserAccount = $task.Principal.UserId
        $taskState = $task.State.ToString()
        $lastRunTime = $task.LastRunTime
        $nextRunTime = $task.NextRunTime
        $lastResult = $task.LastTaskResult
        $enabled = $task.Enabled
        $description = $task.RegistrationInfo.Description
        
        # Fallback to Get-ScheduledTask if Author or UserId is missing
        if (-not $taskAuthor -or -not $taskUserAccount) {
            $nativeTask = Get-ScheduledTask | Where-Object { $_.TaskName -eq $taskName -and $_.TaskPath -eq $taskPath }
            $taskAuthor = $nativeTask.RegistrationInfo.Author
            $taskUserAccount = $nativeTask.Principal.UserId
        }

        # Prepare the output string
        $output = @"
Task Name: $taskName
Task Path: $taskPath
State: $taskState
Last Run Time: $lastRunTime
Next Run Time: $nextRunTime
Last Result: $lastResult
Enabled: $enabled
Description: $description
Author: $taskAuthor
User Account: $taskUserAccount
---------------------------
"@

        # Optionally output to the file
        Add-Content -Path $outputFile -Value $output
    }

    # Recurse into subfolders
    foreach ($subFolder in $folder.GetFolders(0)) {
        Get-TaskDetails $subFolder
    }
}

# Start retrieving details from the root folder
Get-TaskDetails $rootFolder

# Notify the user
Write-Output "Scheduled tasks and their details have been written to $outputFile"
