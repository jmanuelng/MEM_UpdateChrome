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

# Check if Winget (Windows Package Manager) is installed
$wingetPath = (Get-Command -Name winget -ErrorAction SilentlyContinue).Source

if (-not $wingetPath) {
    Write-Host "Winget (Windows Package Manager) not installed on device. Please install it and run script again." 
    $detectSummary += "No Winget found. "
    $result = 1
}

# Check if Google Chrome is installed
$chrome = Get-ItemProperty "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome" -ErrorAction SilentlyContinue

# If Google Chrome is installed and no previous errors occurred
if (($null -ne $chrome) -and ($result -eq 0)) {
    # Get the current version of Google Chrome
    $installedVersion = $chrome.DisplayVersion
    $detectSummary += "Chrome Installed version = $installedVersion. "

    # Get the latest version of Google Chrome from Winget repository
    try {
        # Search for Google Chrome in the repository and parse its version number
        $latestChrome = winget search --id "Google.Chrome" --exact | Select-String -Pattern "\d+\.\d+\.\d+\.\d+" -ErrorAction SilentlyContinue
        
        # If the version number is found, store it in the $targetVersion variable
        if ($latestChrome -match "\d+\.\d+\.\d+\.\d+") {
            $targetVersion = $matches[0]
            $detectSummary += "Found newer version $targetVersion. "
        } else {
            # If the version number is not found, throw an error
            throw "Error: Could not find version number in the output."
        }
    }
    catch {
        # Handle the error while fetching the latest version
        Write-Host "Error fetching the latest version of Google Chrome from Winget repository: $_"
        $detectSummary += "Chrome latest version unknown. "
        $result = 1
    }

    # If no errors occurred while fetching the latest version
    if ($result -eq 0) {

        # Compare the installed version with the latest version
        $comparisonResult = [version]$installedVersion -lt [version]$targetVersion

        # If a newer version is available
        if ($comparisonResult) {
            # Check if Google Chrome is running
            $chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue

            # If Google Chrome is running, display a message, log.
            if ($null -ne $chromeProcess) {
                Write-Host "Google Chrome is currently running."
                $detectSummary += "Chrome open/running. "
            }

            # Update Google Chrome using Winget
            try {
                # Run the Winget upgrade command and wait for it to complete
                $installInfo = Start-Process -FilePath "winget" -ArgumentList "upgrade --id ""Google.Chrome"" --exact --force" -NoNewWindow -Wait -PassThru
                
                # Check if the installation/upgrade was successful
                if ($installInfo.ExitCode -eq 0) {
                    Write-Host "Google Chrome installation/upgrade completed successfully."
                    $detectSummary += "Chrome successfully updated. "
                    $result = 0
                }
                else {
                    $detectSummary += "Error updating Chrome $($installInfo.ExitCode). "
                    $result = 1
                }
            }
            catch {
                Write-Host "Error updating Google Chrome: $_"
                $detectSummary += "Error updating Chrome $($installInfo.ExitCode). "
                $result = 1
            }
        }
        else {
            Write-Host "Google Chrome is up-to-date."
            $detectSummary += "Chrome up-to-date. "
            $result = 0
        }

    }

}

else {
    Write-Host "Google Chrome not installed on device."
    $detectSummary += "Chrome not found on device."
    $result = 0
}

#Return result
if ($result -eq 0) {
    Write-Host "OK $([datetime]::Now) : $detectSummary `n`n"
    Exit 0
}
else {
    Write-Host "WARNING $([datetime]::Now) : $detectSummary `n`n"
    Exit 1
}


