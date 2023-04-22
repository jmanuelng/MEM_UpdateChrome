<# 
    .SYNOPSIS
        Updates Google Chrome using Winget, or if that does not work use GoogleUpdate.exe.
        Winget allows for easy check of current available version.

    .DESCRIPTION
        Checks for the latest version of Google Chrome available in the Winget repository,
        compares it with the currently installed version, and updates Google Chrome if a newer version is available.

    .NOTES
        To Do:
        If Chrome running, show toast message to user, wait 3 minutes, restart Chrome.

#>


#region Settings
$Error.Clear()
$detectSummary = ""
$result = 0

#To make it easier to read in AgentExecutor Log.
Write-Host `n`n

#endregion Settings

#region Functions

function Find-WingetPath {
    # Define the possible locations for winget.exe
    $possibleLocations = @(
        "${env:ProgramFiles}\WindowsApps\Microsoft.DesktopAppInstaller_*\app\winget.exe",
        "${env:ProgramFiles(x86)}\WindowsApps\Microsoft.DesktopAppInstaller_*\app\winget.exe",
        "${env:LOCALAPPDATA}\Microsoft\WindowsApps\winget.exe",
        "${env:USERPROFILE}\AppData\Local\Microsoft\WindowsApps\winget.exe"
    )

    # Try to find winget.exe in the possible locations
    foreach ($location in $possibleLocations) {
        try {
            # Get the items that match the current location
            $items = Get-ChildItem -Path $location -ErrorAction Stop

            # If an item is found, return the full path of winget.exe
            if ($items) {
                $wingetPath = $items[0].FullName
                return $wingetPath
            }
        }
        # If an error occurs, continue searching in the next location
        catch {
            Write-Warning "Unable to search for winget.exe in the following location: $location"
        }
    }

    # If winget.exe is not found in any location, display an error message and return $null
    Write-Error "Winget not found in any of the checked locations."
    return $null
}

function Get-ChromeExeDetails {
    <#
    .DESCRIPTION
        Searches for Google Chrome's installation location and display version
        by checking known file paths and registry paths.

    .EXAMPLE
        $chromeDetails = Get-ChromeExeDetails
        Write-Host "Google Chrome is installed at $($chromeDetails.InstallLocation) and the version is $($chromeDetails.DisplayVersion)"
    #>    
    # Define the known file paths and registry paths for Google Chrome
    $chromePaths = [System.IO.Path]::Combine($env:ProgramW6432, "Google\Chrome\Application\chrome.exe"),
                   [System.IO.Path]::Combine(${env:ProgramFiles(x86)}, "Google\Chrome\Application\chrome.exe")
    $registryPaths = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome",
                     "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome",
                     "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"

    # Check the known file paths
    $installedPath = $chromePaths | Where-Object { Test-Path $_ }
    if ($installedPath) {
        $chromeDetails = New-Object PSObject -Property @{
            InstallLocation = $installedPath
            DisplayVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($installedPath).FileVersion
        }
        return $chromeDetails
    }

    # Check the known registry paths
    $registryInstalled = $registryPaths | Where-Object { Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue }
    if ($registryInstalled) {
        $chromeRegistryPath = $registryInstalled | ForEach-Object { Get-ItemProperty -Path $_ }
        foreach ($registryPath in $chromeRegistryPath) {
            if ($registryPath.InstallLocation) {
                $chromeDetails = New-Object PSObject -Property @{
                    InstallLocation = [System.IO.Path]::GetFullPath((Join-Path -Path $registryPath.InstallLocation -ChildPath "..\Application\chrome.exe"))
                    DisplayVersion = $registryPath.DisplayVersion
                }
                if (Test-Path $chromeDetails.InstallLocation) {
                    return $chromeDetails
                }
            } elseif ($registryPath.'(Default)') {
                $chromeDetails = New-Object PSObject -Property @{
                    InstallLocation = $registryPath.'(Default)'
                    DisplayVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($registryPath.'(Default)').FileVersion
                }
                if (Test-Path $chromeDetails.InstallLocation) {
                    return $chromeDetails
                }
            }
        }
    }

    # If Google Chrome is not found, return $null
    return $null
}


function Find-GoogleUpdateExe {
    param (
        [Parameter(Mandatory=$true)]
        [string]$chromeInstallLocation
    )

    $programW6432 = $env:ProgramW6432
    $programX86 = ${env:ProgramFiles(x86)}
    # Define the file name to search for
    $googleUpdateExeFile = "GoogleUpdate.exe"

    $googleUpdateExePath = [System.IO.Path]::GetFullPath((Join-Path -Path $chromeInstallLocation -ChildPath "..\..\..\"))
    $googleUpdateExe = [System.IO.Path]::GetFullPath((Join-Path -Path $googleUpdateExePath -ChildPath $googleUpdateExeFile))

    # Search for the file in the current directory and all its subdirectories
    $googleUpdateExefound = Get-ChildItem -Path $googleUpdateExePath -Filter $googleUpdateExeFile -Recurse -ErrorAction SilentlyContinue -Force

    # Check if the file is found.
    if ($googleUpdateExefound) {
        $googleUpdateExe = $googleUpdateExefound[0].FullName
    } 


    if (-not(Test-Path $googleUpdateExe)) {

        # GoogleUpdate.exe not found, check if $originalPath, lets try to find it in $programW6432 and $programX86.
        #   This part necessary for environments with multiple OS languages. Actually... Do I really need this? not completely sure. 
        #   Tries to find GoogleUpdate.exe in paths based on environment variables "ProgramW6432" and "ProgramFiles(x86)."
        # Define a list of environment variables to try
        $envVariables = @($ProgramW6432, $ProgramX86)
        $googleUpdateExe = $null

        foreach ($envVar in $envVariables) {
            $envValue = $envVar.ToString()
            $googleUpdateExePathStr = $googleUpdateExePath.ToString()
        
            if (-not($googleUpdateExePathStr -like $envValue)) {
                # Split the original path by the first backslash after the drive letter
                $pathParts = $googleUpdateExePathStr -split '\\', 3
        
                # Replace the second part of the path (Program Files) with the value of the current environment variable
                $pathParts[1] = $envValue.Trim("C:\")
        
                # Join the parts of the path back together
                $potentialGUpdate = Join-Path $pathParts[0] -ChildPath ($pathParts[1..($pathParts.Length - 1)] -join '\')
                
                # Go search for GoogleUpdate.exe in this path
                $googleUpdateExe = Get-ChildItem -Path $potentialGUpdate -Filter $googleUpdateExeFile -Recurse -ErrorAction SilentlyContinue -Force

                # If GoogleUpdate.exe is found, break the loop
                # Check if the file is found.
                if ($googleUpdateExe) {
                    $googleUpdateExe = $googleUpdateExe[0].FullName
                } 

                if (Test-Path $googleUpdateExe) {
                    break
                }
            }
        }
        
        if ($null -eq $googleUpdateExe) {
            Write-Host "GoogleUpdate.exe not found."
            return $null
        } else {
            Write-Host "GoogleUpdate.exe found at: $googleUpdateExe."
            return $googleUpdateExe
        }
    } else {
        Write-Host "GoogleUpdate.exe found at: $googleUpdateExe."
        return $googleUpdateExe
    }
}

function Get-WingetLatestChromeVersion {
    param (
        [Parameter(Mandatory = $true)]
        [string]$WingetFilePath
    )

    # Get the latest version of Google Chrome from Winget repository
    try {
        # Search for Google Chrome in the repository and parse its version number...
        #    the complicated version as it will run in system context.

        # Create a temporary file to store the output of the winget command
        $tempFile = New-TemporaryFile

        # Execute the winget command using Start-Process, redirect output to the temporary file,
        # process is executed without creating a new window (-NoNewWindow), script waits (-Wait)
        Start-Process -FilePath $WingetFilePath -ArgumentList "search --id ""Google.Chrome"" --exact" -NoNewWindow -Wait -RedirectStandardOutput $tempFile.FullName

        # Add a 10-second sleep after winget search
        Start-Sleep -Seconds 10

        # Read the contents of the temporary file into the $output variable
        $latestChrome = Get-Content $tempFile.FullName
        # Remove the temporary file as it is no longer needed
        Remove-Item $tempFile.FullName

        # Find pattern in $latestChrome and store the result in the same variable
        # The updated pattern will now look for a version number with at least one dot
        $latestChrome = $latestChrome | Select-String -Pattern "\d+(\.\d+)+" -ErrorAction SilentlyContinue

        # If the version number is found, store it in the $chromeVersion variable
        if ($latestChrome -match "\d+(\.\d+)+") {
            $chromeVersion = $matches[0]
        } else {
            # If the version number is not found, throw an error
            throw "Error: Could not find version number in the Winget output."
        }
    }
    catch {
        # Handle the error while fetching the latest version
        Write-Host "Error fetching the latest version of Google Chrome from Winget repository: $_"
        return $null
    }

    return $chromeVersion
}

function Get-OmahaLatestChromeVersion {
    # Get the latest Google Chrome version from OmahaProxy API. In case Winget does not exist.
    # $latestChromeVersion = Get-LatestChromeVersion

    # Set the URL for the OmahaProxy API
    $url = "https://omahaproxy.appspot.com/all.json"

    # Try fetching the latest version from the API
    try {
        # Perform the request and store the response
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Stop
        # Convert the response content to JSON
        $json = ConvertFrom-Json -InputObject $response.Content

        # Find the latest stable version for Windows
        $latestVersionInfo = $json |
            Where-Object { $_.os -eq 'win' } |
            ForEach-Object { $_.versions } |
            Where-Object { $_.channel -eq 'stable' } |
            Select-Object -First 1

        # If a version is found, display it and return the value
        if ($latestVersionInfo) {
            $chromeVersion = $latestVersionInfo.version -split ' ' | Select-Object -First 1
            return $chromeVersion
        } else {
            # If no version is found, display an error message and return $null
            Write-Error "Failed to fetch the latest version of Google Chrome. No version found."
            return $null
        }
    }
    # If an error occurs, display the error message and return $null
    catch {
        Write-Error "Failed to fetch the latest version of Google Chrome. Error: $_"
        return $null
    }
}

function GUpdate {
    param (
        [Parameter(Mandatory=$true)]
        [string]$googleUpdatePath
    )

    try {
        # Run GoogleUpdate.exe with arguments to update Chrome
        $processResult = Start-Process -FilePath $googleUpdatePath -ArgumentList "/ua /installsource scheduler" -NoNewWindow -Wait -PassThru
        
        $exitCode = $processResult.ExitCode
        Write-Host "GUpdate Exit Code: $exitCode"

        # Check if successful upgrade with Gupdate
        if ($exitCode -eq 0) {
            Write-Host "GoogleUpdate.exe upgraded successfully."
            $result = 0
        }
        else {
            Write-Host "GoogleUpdate.exe error upgrading."
            $result = 1
        }

    }
    catch {
        $detectSummary += "Error trying update with GoogleUpdate.exe. "
        $result = 1
    }

    return $result
    
}

function RestartChrome {
    [CmdletBinding()]
    param (
        [int]$WaitSeconds = 10,
        [string]$ChromePath
    )

    try {
        # If ChromePath is not provided or invalid, attempt to find it
        if (-not $ChromePath -or -not (Test-Path $ChromePath)) {
            Write-Host "Finding Google Chrome path..."
            $ChromePath = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe').'(Default)'
            if (-not $ChromePath -or -not (Test-Path $ChromePath)) {
                throw "Google Chrome path not found. Please provide a valid ChromePath or install Google Chrome."
            }
        }

        # Get the running Google Chrome process
        $chromeProcess = Get-Process -Name "chrome" -ErrorAction Stop

        # Check if Google Chrome is open
        if ($chromeProcess) {
            Write-Host "Google Chrome is open. Waiting $WaitSeconds seconds before restarting..."

            # Wait for the specified amount of time
            Start-Sleep -Seconds $WaitSeconds

            # Close and restart Google Chrome
            Stop-Process -Name "chrome" -Force

            # Wait 3 seconds before strating Chrome
            Start-Sleep -Seconds 3

            # Starting Chrome
            Start-Process -FilePath $ChromePath

            Write-Host "Google Chrome has been restarted."
            return 0
        } else {
            Write-Warning "Google Chrome is not open. Please launch the browser before running this script."
            return 1
        }
    } catch {
        if ($_.Exception.GetType().Name -eq "ProcessCommandException") {
            Write-Error "Google Chrome is not open. Please launch the browser before running this script."
            return -1
        } else {
            Write-Error "An error occurred: $($_.Exception.Message)"
            return 1
        }
    }
}

#endregion Functions


#region Main

# Check if Winget (Windows Package Manager) is installed
$wingetPath = (Get-Command -Name winget -ErrorAction SilentlyContinue).Source

# The above might not work in System context, so if not found lets try to find Winget in known places.... for which system should have access.
if (-not $wingetPath) {
    Write-Host "Winget not found under user Path, searching for Winget in system..."
    $wingetPath = Find-WingetPath
}

if (-not $wingetPath) {
    Write-Host "Winget (Windows Package Manager) not found on device." 
    $detectSummary += "Winget NOT found. "
    $result = -1
}
else {
    $detectSummary += "Winget found $wingetPath. "
}

# Check if Google Chrome is installed
$chrome = Get-ChromeExeDetails

# If Google Chrome is installed and no previous important errors occurred
if (($null -ne $chrome) -and ($result -ne 1)) {
    # Get the current version of Google Chrome
    $installedVersion = $chrome.DisplayVersion
    $detectSummary += "Chrome Installed version = $installedVersion. "

    # Find GoogleUpdate.exe, in case its needed.
    $gUpdateExe = Find-GoogleUpdateExe -chromeInstallLocation $($chrome.InstallLocation)
    if ($null -eq $gUpdateExe) {
        $detectSummary += "GoogleUpdate.exe not found. "
    }
    else {
        $detectSummary += "GoogleUpdate.exe found at $gUpdateExe. "
    }


    # Get the latest version of Google Chrome from Winget repository (if Winget exists)
    if ($wingetPath){
        try {

            # Get the latest version of Google Chrome from Winget repository
            $targetVersion = Get-WingetLatestChromeVersion -WingetFilePath $wingetPath
            if ($null -ne $targetVersion) {
                Write-Host "Latest Google Chrome version from Winget repository: $targetVersion"
            } 
            else {
                Write-Host "No Chrome version found using Winget."
                Write-Host "Trying to find latest Google Chrome version using Omaha URL"
                # In case Winget does not return a Chrome targe version.
                $targetVersion = Get-OmahaLatestChromeVersion
            }
        
            if ($null -eq $targetVersion) {
                Write-Host "Unable to fetch information on latest version of Google Chrome from Winget and Omaha"
                $detectSummary += "Chrome latest version unknown. "
                # Unable to get latest version does not allow to identify if an update is needed. 
                # Setting $result 0 to reduce number of errors and/or issues on Proactive Remediation.
                $result = 0    
            }
            else {
                $detectSummary += "Chrome latest version = $targetVersion. "
            }

        }
        catch {
            # Handle the error while fetching the latest version
            Write-Host "Error fetching the latest version of Google Chrome from Winget repository: $_"
            $detectSummary += "Chrome latest version unknown. "
            $result = 1
        }
    }
    else {
        Write-Host "Trying to find latest Google Chrome version using Omaha"
        $targetVersion = Get-OmahaLatestChromeVersion   
        
        # If we don't find the latest Google Chrome version just say its 999.0.0.0
        if ($null -eq $targetVersion) {
            Write-Host "Unable to fetch information on latest version of Google Chrome from Omaha"
            $detectSummary += "Chrome latest version unknown. "
            # If no $targetVersión available, assume we found a super mega new version like 999.0.0.0, 
            #   this will make the script try to upgrade. 
            $targetVersion = "999.0.0.0"
            # Unable to get latest version does not allow to identify if an update is needed.
            #   This is the remediation part of the script, so, trying to update should be the right thing to do.
            $result = 0    
        }
    }

    # If no important errors occurred while fetching the latest version
    if ($result -ne 1) {
        # Display latest Chrome version found
        Write-Host "Lastest Chrome version: $targetVersion."
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


            if (-not $wingetPath) {

                if (Test-Path $gUpdateExe) {
                    # No Winget found, so update using GoogleUpdate.exe
                    $result = GUpdate -googleUpdatePath $gUpdateExe

                    if ($result -eq 0) {
                        $detectSummary += "Gupdate success. "
                    }
                    else {
                        $detectSummary += "Gupdate error. "
                    }
                }
            }

            else {
            
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
                    Write-Host "Winget updgrade exit code: $exitCode"
                    Write-Host "Winget upgrade output: $installInfo"
                    
                    # Check if the installation/upgrade was successful
                    if ($exitCode -eq 0) {
                        Write-Host "Winget Google Chrome installation/upgrade completed successfully."
                        $detectSummary += "Chrome successful Winget update. "
                        $result = 0
                    }
                    # If installed version not compatible with update then use Google Update.exe, just for that case.
                    elseif ($exitCode -eq -1978335189) {
                        Write-Host "Winget error trying to update: Install technology different vs installed version. "
                        Write-Host "... Now trying update with GoogleUpdate.exe"

                        if (Test-Path $gUpdateExe) {
                            $detectSummary += "Backup plan w/GoogleUpdate.exe. "

                            $result = GUpdate -googleUpdatePath $gUpdateExe

                            if ($result -eq 0) {
                                $detectSummary += "Backup Gupdate success. "
                            }
                            else {
                                $detectSummary += "Backup Gupdate error. "
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
elseif ($result -eq 1)  {
    Write-Host "WARNING $([datetime]::Now) : $detectSummary"
    Exit 1
}
else {
    Write-Host "NOTE $([datetime]::Now) : $detectSummary"
    Exit 0
}

#endregion Main