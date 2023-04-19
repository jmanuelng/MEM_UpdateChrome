<# 
    .SYNOPSIS
        Updates Google Chrome using Winget.

    .DESCRIPTION
        Checks for the latest version of Google Chrome available in the Winget repository,
        compares it with the currently installed version, and updates Google Chrome if a newer version is available.

#>

# Initialize variables
$Error.Clear()
$detectSummary = ""
$result = 0

# Check if Google Chrome is installed
$chrome = Get-ItemProperty "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome" -ErrorAction SilentlyContinue

if ($null -ne $chrome)  {
    # Get the current version of Google Chrome
    $installedVersion = $chrome.DisplayVersion
    Write-Host "Found Chrome version $installedVersion"
    $detectSummary += "Chrome Installed version = $installedVersion. " 
}
else {
    Write-Host "Google Chrome not installed on device."
    $detectSummary += "Chrome not found on device. "
    $result = 2
}

# Check if Winget (Windows Package Manager) is installed
$wingetPath = (Get-Command -Name winget -ErrorAction SilentlyContinue).Source

# The above might not work in System context, so if not found lets try to find Winget in PogramFile\WindowsApps.... for which system should have access.
if (-not $wingetPath) {
    Write-Host "Winget not found under user Path, searching for Winget under Program Files\WindowesApps..."
    $WingetSearchPath = "${env:ProgramW6432}\WindowsApps\Microsoft.DesktopAppInstaller*_x64*\winget.exe"
    $WingetPath = (Get-ChildItem -Path $WingetSearchPath -File -ErrorAction SilentlyContinue).FullName
}

if (-not $wingetPath) {
    Write-Host "Winget (Windows Package Manager) not installed on device. Please install it and run script again." 
    $detectSummary += "Winget NOT found. "
    $result = 3
}

# If Google Chrome is installed and no previous errors occurred
if (($null -ne $chrome) -and ($result -eq 0)) {

    # Get the latest version of Google Chrome from Winget repository
    try {
        # Search for Google Chrome in the repository and parse its version number...
        #    the complicated version as it will run in system context.

        # Create a temporary file to store the output of the winget command
        $tempFile = New-TemporaryFile

        # Execute the winget command using Start-Process, redirect output to the temporary file,
        # process is executed without creating a new window (-NoNewWindow), script waits (-Wait)
        Start-Process -FilePath "$wingetPath" -ArgumentList "search --id ""Google.Chrome"" --exact" -NoNewWindow -Wait -RedirectStandardOutput $tempFile.FullName
        # Read the contents of the temporary file into the $output variable
        $latestChrome = Get-Content $tempFile.FullName
        # Remove the temporary file as it is no longer needed
        Remove-Item $tempFile.FullName

        # Find pattern in $latestChrome and store the result in the same variable
        # The updated pattern will now look for a version number with at least one dot
        $latestChrome = $latestChrome | Select-String -Pattern "\d+(\.\d+)+" -ErrorAction SilentlyContinue

        # If the version number is found, store it in the $targetVersion variable
        if ($latestChrome -match "\d+(\.\d+)+") {
            $targetVersion = $matches[0]
            $detectSummary += "Available Winget version $targetVersion. "
        } else {
            # If the version number is not found, throw an error
            throw "Error: Could not find version number in the output."
        }
    }
    catch {
        # Handle the error while fetching the latest version
        Write-Host "Error fetching the latest version of Google Chrome from Winget repository: $_"
        $detectSummary += "Chrome latest version unknown. "
        $result = 0
    }

    # If no errors occurred while fetching the latest version
    if ($result -eq 0) {

        # Compare the installed version with the latest version
        $comparisonResult = [version]$installedVersion -lt [version]$targetVersion

        # If a newer version is available
        if ($comparisonResult) {

            #Chrome should be updated
            Write-Host "Chrome needs to be updated to $targetVersion."
            $detectSummary += "Chrome needs update to $targetVersion. "
            $result = 1

        }
        else {
            Write-Host "Google Chrome is up-to-date."
            $detectSummary += "Chrome up-to-date. "
            $result = 0
        }

    }

}


#Return result
if ($result -eq 0) {
    Write-Host "OK $([datetime]::Now) : $detectSummary"
    Exit 0
}
elseif ($result -eq 1) {
    Write-Host "WARNING $([datetime]::Now) : $detectSummary"
    Exit 1
}
else {
    Write-Host "NOTE $([datetime]::Now) : $detectSummary"
    Exit 0
}


