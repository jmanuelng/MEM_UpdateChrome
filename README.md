# MEM_UpdateChrome
<h1>Keep Google Chrome Up-to-Date in Your Client's Environment using Microsoft Intune's Proactive Remediation</h1>

<p>Can be used if you're in a hurry to upgrade Google Chrome on an environment. I've used int in scenarioes where Chrome continues to be flagged as a top recommendation in Microsoft Defender for Endpoint, despite having update policies in place. This scriptsw help address this issue. It leverages Winget and GoogleUpdate.exe to ensure Google Chrome installation is always up-to-date. This script is to be used as Proactive Remediation via Microsoft Intune.</p>

<h2>Key Features</h2>
    <ol>
        <li><b>Winget Integration:</b> The script takes advantage of the Windows Package Manager, Winget, to fetch the latest version of Google Chrome and update it seamlessly (if possible).</li>
        <li><b>GoogleUpdate.exe Fallback:</b> If Winget is not present or fails to update the application, the script automatically switches to using GoogleUpdate.exe to ensure a successful update.</li>
        <li><b>Version Comparison:</b> The script intelligently compares the installed version of Google Chrome with the latest version available, updating only if necessary.</li>
        <li><b>Error Handling:</b> Comprehensive error handling is built into the script to provide informative feedback and appropriate exit codes.</li>
    </ol>
    
 <h2>How to Use the Script with Microsoft Intune</h2>
    <ol>
        <li>In the Microsoft Endpoint Manager admin center, create a new Proactive Remediation script.</li>
        <li>Upload Detect_ChromeUpdate.ps1 file as the detection script and Fix_ChromeUpdate.ps1.ps1 as remediation script.</li>
        <li>Configure the script settings, such as the script execution schedule.</li>
        <li>Assign the Proactive Remediation script to the appropriate device groups.</li>
        <li>Monitor the script execution status and results in the Microsoft Endpoint Manager admin center.</li>
    </ol>
    
<p>Managing Google Chrome updates in a client's environment can be a headache, this PowerShell can help. Give it a try and let me know your thoughts, comments and/or feedback!</p>



