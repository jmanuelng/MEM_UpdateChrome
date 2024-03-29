<# 
    .SYNOPSIS
        Updates Google Chrome using Winget.

    .DESCRIPTION
        Checks for the latest version of Google Chrome available in the Winget repository,
        compares it with the currently installed version, and updates Google Chrome if a newer version is available.

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
    <#
    .SYNOPSIS
        Locates the winget.exe executable within a system.

    .DESCRIPTION
        Finds the path of the `winget.exe` executable on a Windows system. 
        Aimed at finding Winget when main script is executed as SYSTEM, but will also work under USER
        
        Windows Package Manager (`winget`) is a command-line tool that facilitates the 
        installation, upgrade, configuration, and removal of software packages. Identifying the 
        exact path of `winget.exe` allows for execution (installations) under SYSTEM context.

        METHOD
        1. Defining Potential Paths:
        - Specifies potential locations of `winget.exe`, considering:
            - Standard Program Files directory (64-bit systems).
            - 32-bit Program Files directory (32-bit applications on 64-bit systems).
            - Local application data directory.
            - Current user's local application data directory.
        - Paths may utilize wildcards (*) for flexible directory naming, e.g., version-specific folder names.

        2. Iterating Through Paths:
        - Iterates over each potential location.
        - Resolves paths containing wildcards to their actual path using `Resolve-Path`.
        - For each valid location, uses `Get-ChildItem` to search for `winget.exe`.

        3. Returning Results:
        - If `winget.exe` is located, returns the full path to the executable.
        - If not found in any location, outputs an error message and returns `$null`.

    .EXAMPLE
        $wingetLocation = Find-WingetPath
        if ($wingetLocation) {
            Write-Output "Winget found at: $wingetLocation"
        } else {
            Write-Error "Winget was not found on this system."
        }

    .NOTES
        While this function is designed for robustness, it relies on current naming conventions and
        structures used by the Windows Package Manager's installation. Future software updates may
        necessitate adjustments to this function.

    .DISCLAIMER
        This function and script is provided as-is with no warranties or guarantees of any kind. 
        Always test scripts and tools in a controlled environment before deploying them in a production setting.
        
        This function's design and robustness were enhanced with the assistance of ChatGPT, it's important to recognize that 
        its guidance, like all automated tools, should be reviewed and tested within the specific context it's being 
        applied. 

    #>
    # Define possible locations for winget.exe
    $possibleLocations = @(
        "${env:ProgramFiles}\WindowsApps\Microsoft.DesktopAppInstaller*_x64__8wekyb3d8bbwe\winget.exe", 
        "${env:ProgramFiles(x86)}\WindowsApps\Microsoft.DesktopAppInstaller*_8wekyb3d8bbwe\winget.exe",
        "${env:LOCALAPPDATA}\Microsoft\WindowsApps\winget.exe",
        "${env:USERPROFILE}\AppData\Local\Microsoft\WindowsApps\winget.exe"
    )

    # Iterate through the potential locations and return the path if found
    foreach ($location in $possibleLocations) {
        try {
            # Resolve path if it contains a wildcard
            if ($location -like '*`**') {
                $resolvedPaths = Resolve-Path $location -ErrorAction SilentlyContinue
                # If the path is resolved, update the location for Get-ChildItem
                if ($resolvedPaths) {
                    $location = $resolvedPaths.Path
                }
                else {
                    # If path couldn't be resolved, skip to the next iteration
                    Write-Warning "Couldn't resolve path for: $location"
                    continue
                }
            }
            
            # Try to find winget.exe using Get-ChildItem
            $items = Get-ChildItem -Path $location -ErrorAction Stop
            if ($items) {
                Write-Host "Found Winget at: $items"
                return $items[0].FullName
                break
            }
        }
        catch {
            Write-Warning "Couldn't search for winget.exe at: $location"
        }
    }

    Write-Error "Winget wasn't located in any of the specified locations."
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


#enregion Functions

#region Main

# Check if Google Chrome is installed
$chrome = Get-ChromeExeDetails

if ($null -ne $chrome)  {
    # Get the current version of Google Chrome
    $installedVersion = $chrome.DisplayVersion
    Write-Host "Found Installed Chrome version $installedVersion"
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


# If Google Chrome is installed and no previous important errors occurred
if (($null -ne $chrome) -and ($result -le 0)) {

    # Get the latest version of Google Chrome from Winget repository
    $targetVersion = Get-WingetLatestChromeVersion -WingetFilePath $wingetPath

    if ($null -ne $targetVersion) {
        Write-Host "Latest Google Chrome version from Winget repository: $targetVersion"
    } else {
        Write-Host "No Chrome version found using Winget."
        Write-Host "Trying to find latest Google Chrome version using Omaha URL"
        $targetVersion = Get-OmahaLatestChromeVersion
    }

    if ($null -eq $targetVersion) {
        Write-Host "Unable to fetch information on latest version of Google Chrome from Winget nor Omaha"
        $detectSummary += "Chrome latest version unknown. "
        # Unable to get latest version does not allow to identify if an update is needed. 
        # Setting $result 0 to reduce number of errors and/or issues on Proactive Remediation.
        $result = 0    
    }
    
    else {
        # Display latest Chrome version found
        Write-Host "Lastest Chrome version: $targetVersion."
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

#To make it easier to read in AgentExecutor Log.
Write-Host `n`n

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


#enregion Main
