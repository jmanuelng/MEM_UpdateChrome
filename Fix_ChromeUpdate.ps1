<# 
    .SYNOPSIS
        Updates Google Chrome using Winget, or if that does not work use GoogleUpdate.exe.
        Winget allows for easy check of current available version.

    .DESCRIPTION
        Checks for the latest version of Google Chrome available in the Winget repository,
        compares it with the currently installed version, and updates Google Chrome if a newer version is available.

#>


# Initialize variables
$Error.Clear()
$detectSummary = ""
$result = 0

#To make it easier to read in AgentExecutor Log.
Write-Host `n`n

# Check if Winget (Windows Package Manager) is installed
$wingetPath = (Get-Command -Name winget -ErrorAction SilentlyContinue).Source

# The above might not work in System context, so if not found lets try to find Winget in PogramFile\WindowsApps.... for which system should have access.
if (-not $wingetPath) {
    Write-Host "Winget not found under user Path, searching for Winget under Program Files\WindowesApps..."
    $wingetSearchPath = "${env:ProgramW6432}\WindowsApps\Microsoft.DesktopAppInstaller*_x64*\winget.exe"
    $wingetPath = (Get-ChildItem -Path $WingetSearchPath -File -ErrorAction SilentlyContinue).FullName
}

if (-not $wingetPath) {
    Write-Host "Winget (Windows Package Manager) not installed on device. Please install it and run script again." 
    $detectSummary += "No Winget found. "
    $result = 1
}
else {
    Write-Host "Winget Path = $wingetPath"
}

# Check if Google Chrome is installed
$chrome = Get-ItemProperty "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome" -ErrorAction SilentlyContinue

# If Google Chrome is installed and no previous errors occurred
if (($null -ne $chrome) -and ($result -eq 0)) {
    # Get the current version of Google Chrome
    $installedVersion = $chrome.DisplayVersion
    $detectSummary += "Chrome Installed version = $installedVersion. "
    $googleUpdateExe = [System.IO.Path]::GetFullPath((Join-Path -Path $($chrome.InstallLocation) -ChildPath "..\..\Update\ChromeUpdate.exe"))

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
                # Run the Winget upgrade command and wait for it to complete. Also considering System context execution
                $tempFile = New-TemporaryFile
                $processResult = Start-Process -FilePath "$wingetPath" -ArgumentList "upgrade --id ""Google.Chrome"" --exact --force" -NoNewWindow -Wait -RedirectStandardOutput $tempFile.FullName -PassThru

                # Get the exit code and description
                $exitCode = $processResult.ExitCode

                # Read the output from the temporary file
                $installInfo = Get-Content $tempFile.FullName
                Remove-Item $tempFile.FullName

                # Display the exit code and output
                Write-Host "Exit Code: $exitCode"
                Write-Host "Output: $installInfo"
                
                # Check if the installation/upgrade was successful
                if ($exitCode -eq 0) {
                    Write-Host "Winget Google Chrome installation/upgrade completed successfully."
                    $detectSummary += "Chrome successfull Winget update. "
                    $result = 0
                }
                # If installed version not compatible with update then use Google Update.exe, just for that case.
                elseif ($exitCode -eq -1978335189) {
                    Write-Host "Error trying to update: Install technology different vs installed version. "
                    Write-Host "Trying update with GoogleUpdate.exe"

                    if (Test-Path $googleUpdateExe) {
                        $detectSummary += "Trying with GoogleUpdate.exe. "

                        try {
                        # Run GoogleUpdate.exe with arguments to update Chrome
                        Start-Process -FilePath $googleUpdateExe -ArgumentList "/update", "appguid={8A69D345-D564-463c-AFF1-A69D9E530F96}" -NoNewWindow -Wait -PassThru
                        }
                        catch {
                            $detectSummary += "Error trying update with GoogleUpdate.exe. "
                        }
                        
                    }
                    else {
                        Write-Host "GoogleUpdate.exe not found. "
                    }

                }
                else {$
                    $detectSummary += "Error updating Chrome: $installInfo, code: $exitCode. "
                    $result = 1
                }
            }
            catch {
                Write-Host "Error updating Google Chrome: $_"
                $detectSummary += "Error updating Chrome $($processResult.ExitCode). "
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

# To make it easier to read in AgentExecutor Log.
Write-Host `n`n

#Return result
if ($result -eq 0) {
    Write-Host "OK $([datetime]::Now) : $detectSummary"
    Exit 0
}
else {
    Write-Host "WARNING $([datetime]::Now) : $detectSummary"
    Exit 1
}

