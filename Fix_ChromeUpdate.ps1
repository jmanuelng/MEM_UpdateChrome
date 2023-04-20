<# 
    .SYNOPSIS
        Updates Google Chrome using Winget, or if that does not work use GoogleUpdate.exe.
        Winget allows for easy check of current available version.

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

#region Funcitions

function Get-ChromeExeDetails {
    $chromePaths = [System.IO.Path]::Combine($env:ProgramW6432, "Google\Chrome\Application\chrome.exe"),
                   [System.IO.Path]::Combine(${env:ProgramFiles(x86)}, "Google\Chrome\Application\chrome.exe")
    $registryPaths = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"

    $installedPath = $chromePaths | Where-Object { Test-Path $_ }
    if ($installedPath) {
        $chromeDetails = New-Object PSObject -Property @{
            InstallLocation = $installedPath
            DisplayVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($installedPath).FileVersion
        }
        return $chromeDetails
    }

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
            }
        }
    }

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

function GUpdate {

    try {
        # Run GoogleUpdate.exe with arguments to update Chrome
        $gUpdateExe = "C:\Path\To\GoogleUpdate.exe"
        $processResult = Start-Process -FilePath $gUpdateExe -ArgumentList "/c" -NoNewWindow -Wait -PassThru
        
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

#endregion Functions


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
    $result = 2
}
else {
    Write-Host "Winget Path = $wingetPath"
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
            # Search for Google Chrome in the repository and parse its version number...
            #    the complicated version as it will run in system context.

            # Create a temporary file to store the output of the winget command
            $tempFile = New-TemporaryFile

            # Execute the winget command using Start-Process, redirect output to the temporary file,
            # process is executed without creating a new window (-NoNewWindow), script waits (-Wait)
            Start-Process -FilePath "$wingetPath" -ArgumentList "search --id ""Google.Chrome"" --exact" -NoNewWindow -Wait -RedirectStandardOutput $tempFile.FullName
            # Wait 10 seconds, found, after multiple test, thet when run the first time it does take a bit more of time.
            Start-Sleep -Seconds 10 
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
    }
    else {
        # If no Winget, assume a super mega old version like 50.0.0.0
        $targetVersion = "50.0.0.0"
    }

    # If no important errors occurred while fetching the latest version
    if ($result -ne 1) {

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


            if ($wingetPath) {

                if (Test-Path $gUpdateExe) {
                    # No Winget found, so update using GoogleUpdate.exe
                    $result = GUpdate

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
                    Write-Host "Winget Exit Code: $exitCode"
                    Write-Host "Winget Output: $installInfo"
                    
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

                            $result = GUpdate

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

