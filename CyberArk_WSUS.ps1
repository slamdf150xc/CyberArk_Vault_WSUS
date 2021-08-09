################################### GET-HELP #############################################
<#
.SYNOPSIS
 	All in one utility for WSUS updates on CyberArk Vault Server
 
.EXAMPLE
 	.\CyberArk_WSUS.ps1
 
.INPUTS  
	None via command line
	
.OUTPUTS
	None
	
.NOTES
	AUTHOR:  
	Randy Brown

	VERSION HISTORY:
	See GitHub
#>
##########################################################################################
######################### GLOBAL VARIABLE DECLARATIONS ###################################

$regPaths = @{}

# Variables for the registery path
$WsusRegistryPath = "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate"
$WsusAURegistryPath = "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU"
$WsusUpdateRegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"
$MSIServer = "HKLM:\SYSTEM\CurrentControlSet\Services\msiserver"

# Names of the registery variable
$WUServer = "WUServer"
$WUStatusServer = "WUStatusServer"
$AUOptions = "AUOptions"
$RecommendUpdates = "IncludeRecommendedUpdates"
$UseWUServer = "UseWUServer"
$AutoInstallMinorUpdates = "AutoInstallMinorUpdates"
$DetectionFrequencyEnabled = "DetectionFrequencyEnabled"
$DetectionFrequency = "DetectionFrequency"
$NoAutoRebootWithLoggedOnUsers = "NoAutoRebootWithLoggedOnUsers"
$NoAutoUpdate = "NoAutoUpdate"
$ScheduledInstallDay = "ScheduledInstallDay"
$ScheduledInstallTime = "ScheduledInstallTime"
$AcceptTrustedPublisherCerts = "AcceptTrustedPublisherCerts"
$ElevateNonAdmins = "ElevateNonAdmins"
$TargetGroupEnabled = "TargetGroupEnabled"
$MSIServerKey = "Start"
$DisableWindowsUpdateAccess = "DisableWindowsUpdateAccess"

# Names of the services
$wuauservName = "wuauserv"
$TrustedInstallerName = "TrustedInstaller"
$UpdateOrchestratorName = "UsoSvc"

# Setup for user input
$yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'Yes'
$no = New-Object System.Management.Automation.Host.ChoiceDescription '&No', 'No'
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$message = "Is this the correct WSUS address?"

########################## START FUNCTIONS ###############################################

function writeRed($text) {
	Write-Host $text -ForegroundColor Red
}

function writeGreen($text) {
	Write-Host $text -ForegroundColor Green
}

function wuauReport {
    Write-Host "Reporting into WSUS server..." -NoNewline
    wuauclt.exe /ReportNow

    Start-Sleep 10

    writeGreen "OK"
}

function configureWSUS {
    $wsusURL = getWSUSURL
    $wsusPrompt = "Please enter the WSUS IP/URL and port. (http://10.1.20.12:8530)"
    ""
    do {
        if ($wsusURL) {
            Write-Host "Your current WSUS address is: $wsusURL" -ForegroundColor Yellow
            ""
            $wsusURL = Read-Host $wsusPrompt
        } else {
            $wsusURL = Read-Host $wsusPrompt
        }
        $result = $host.ui.PromptForChoice($wsusURL, $message, $options, 0)
        switch ($result) {
            '0' {}
            '1' {}
        }
    } until ($result -eq '0')

    try {
        touchService $wuauservName "Automatic" "Running" "Enabling and Starting $wuauservName..." $false


        if (!(Test-Path $WsusRegistryPath)) {
            try {
                Write-Host "Creating the $WsusRegistryPath path"
                New-Item -Path $WsusRegistryPath -Force | Out-Null
            } catch {
                writeRed "An error occured creating the $WsusRegistryPath key"
                writeRed $_
                Exit 1
            }
        }

        # Setting the WSUS server to $wsusURL
        $regPaths.Add($WUServer,  $wsusURL)
        $regPaths.Add($WUStatusServer, $wsusURL)

        foreach ($key in $regPaths.keys) {
            createRegistryKey $WsusRegistryPath $key $regPaths[$key] "String"
        }

        $regPaths.Clear()

        if (!(Test-Path $WsusAURegistryPath)) {
            try {
                "Creating the $WsusAURegistryPath key"
                New-Item -Path $WsusAURegistryPath -Force | Out-Null
            } catch {
                writeRed "An error occured creating the $WsusAURegistryPath key"
                writeRed $_
                Exit 1
            }
        }

        # Adding the required attributes
        $regPaths.Add($AutoInstallMinorUpdates, 0)
        $regPaths.Add($DetectionFrequencyEnabled, 1)
        $regPaths.Add($DetectionFrequency, 14)
        $regPaths.Add($NoAutoRebootWithLoggedOnUsers, 1)
        $regPaths.Add($NoAutoUpdate, 1)
        $regPaths.Add($AUOptions, 5)
        $regPaths.Add($ScheduledInstallDay, 0)
        $regPaths.Add($ScheduledInstallTime, 14)
        $regPaths.Add($UseWUServer, 1)

        foreach ($key in $regPaths.keys) {
            createRegistryKey $WsusAURegistryPath $key $regPaths[$key] "DWORD"
        }

        $regPaths.Clear()

        if(!(Test-Path $WsusUpdateRegistryPath)) {
            try {
                "Creating the $WsusUpdateRegistryPath key"
                New-Item -Path $WsusUpdateRegistryPath -Force | Out-Null
            } catch {
                writeRed "An error occured creating the $WsusUpdateRegistryPath key"
                writeRed $_
                Exit 1
            }
        }

        $regPaths.Add($AUOptions, 1)
        $regPaths.Add($RecommendUpdates, 0)

        foreach ($key in $regPaths.keys) {
            createRegistryKey $WsusUpdateRegistryPath $key $regPaths[$key] "DWORD"
        }

        $regPaths.Clear()

        $regPaths.Add($AcceptTrustedPublisherCerts, 0)
        $regPaths.Add($ElevateNonAdmins, 0)
        $regPaths.Add($TargetGroupEnabled, 1)
        $regPaths.Add($DisableWindowsUpdateAccess, 0)
        
        foreach ($key in $regPaths.keys) {
            createRegistryKey $WsusRegistryPath $key $regPaths[$key] "DWORD"
        }

        createRegistryKey $WsusRegistryPath TargetGroup "" "String"

        createRegistryKey $MSIServer $MSIServerKey 3 "DWORD"

        # Stoping & Disabling "Windows Update" service
        touchService $wuauservName "Disabled" "Stopped" "Stopping and Disabling $wuauservName..." $true

        writeGreen "WSUS was integrated succesfully"
        Write-Host "If this was the first time you have run this script, please reboot the server now." -ForegroundColor Yellow
        Pause
    } catch {
        writeRed "Error: Failed to complete the WSUS server integration."
        writeRed $_
        ""
        stopServices
        Exit 1
    }
}

function createRegistryKey {
    param (
        $Path,
        $Name,
        $Value,
        $Type
    )
    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
}

function touchService {
    param (
        $svcName,
        $svcStartupType,
        $svcStatus,
        $message,
        $force
    )
    Write-Host $message -NoNewline

    if ($force -eq $true -and $svcStartupType -eq "Disabled") {
        Get-Service -Name $svcName | Stop-Service -Force | Out-Null
        Set-Service -Name $svcName -StartupType Disabled | Out-Null
    } else {
        Set-Service -Name $svcName -StartupType $svcStartupType -Status $svcStatus | Out-Null
    }

    writeGreen "OK"
}

function showMenu {
    $title = 'WSUS Options'
    $wsusURL = getWSUSURL

    Write-Host "================ $title ================"
    ""

    if ($wsusURL) {
        Write-Host ("Your WSUS server is currently set to: " + $wsusURL)
    } else {
        Write-Host "WSUS server has not been configured, please run option 1 to create inital configuration." -ForegroundColor Yellow
    }

    ""
    Write-Host "1: Configure WSUS server URL:Port"
    Write-Host "2: Start WSUS services and open the firewall"
    Write-Host "3: Stop WSUS servies and close the firewall"
    Write-Host "4: Download updates from WSUS"
    Write-Host "5: Install updates that have been downloaded"
    Write-Host "6: Download then Install updates from WSUS"
    Write-Host "7: Reboot Server"
    Write-Host "8: Force Vault to check in with WSUS"
    Write-Host "Q: Press 'Q' to quit."
    Write-Host ""
}

function getWSUSURL {
    try {
        if(Test-Path $WsusRegistryPath) {
            $WSUSServer = Get-ItemProperty -Path $WsusRegistryPath | Select-Object -ExpandProperty $WUServer
            return "$WSUSServer"
        } else {
            return ""
        }
    } catch {
        writeRed "An error occurred:"
        writeRed $_
        Exit 1
    }
}

function stopServices {
    touchService $UpdateOrchestratorName "Disabled" "Stopped" "Stopping and Disabling $UpdateOrchestratorName..." $true
    touchService $wuauservName "Disabled" "Stopped" "Stopping and Disabling $wuauservName..." $true

    Write-Host "Stopping and Disabling $TrustedInstallerName..." -NoNewline
    
	$i = 0
	$maxTries = 60
	$waitSeconds = 5
	while ($i -le $maxTries) {
		try {
			Get-Service -Name $TrustedInstallerName | Stop-Service -Force -WarningAction Stop -ErrorAction Stop
            Write-Host "OK" -ForegroundColor Green
            break
		} catch {
			if($i -eq 0) {
                ""
				Write-Host ("Waiting " + ($waitSeconds * $maxTries) + " seconds for TrustedInstaller to stop.") -ForegroundColor Yellow -NoNewline
			} else {
				Write-Host "." -ForegroundColor Yellow -NoNewline
			}

			if($i -eq $maxTries) {
				""
				WriteRed "Failed to stop the TrustedInstaller service."
				WriteRed "This may be related to Windows updates that need a restart."
				WriteRed "Restart the server and verify that the TrustedInstaller service is stopped."
				break
			}

			$i++

			Start-Sleep -s $waitSeconds
		}
    }

    configureFirewall "delete" "WSUS Outbound"
}

function startServices {
    configureFirewall "create" "WSUS Outbound"

    touchService $wuauservName "Automatic" "Running" "Enabling and Starting $wuauservName..." $false
    touchService $TrustedInstallerName "Manual" "Stopped" "Setting $TrustedInstallerName to Manual..." $false
    touchService $UpdateOrchestratorName "Manual" "Stopped" "Setting $UpdateOrchestratorName to Manual..." $false
}

function downloadUpdates {
    param (
        [bool]$install
    )
    startServices

    $Session = New-Object -ComObject Microsoft.Update.Session
    Write-Host "Searching for Windows Updates..." -NoNewline
    $SearchStr = "IsInstalled=0"
    $Searcher = $Session.CreateUpdateSearcher()

    try {
        $SearcherResult = $Searcher.Search($SearchStr).Updates
    } catch {
        ""
        WriteRed ("Search for Windows Updates failed with error: " + $_.Exception.Message + " Failed Item: " + $_.Exception.ItemName)
        stopServices
    }

    if($SearcherResult.Count -eq 0) {
        WriteGreen "OK"
        Write-Host "No updates found" -ForegroundColor Yellow
    } else {
        writeGreen "OK"
        Write-Host ("Found " + $SearcherResult.Count + " updates...") -ForegroundColor Yellow

        for ($i=0; $i -lt $SearcherResult.Count; $i++){
            "" + ($i + 1) + ": " + $SearcherResult.Item($i).Title
        }

        Write-Host "Downloading Windows Updates..." -NoNewline

        $Downloader = $Session.CreateUpdateDownloader()
        $Downloader.Updates = $SearcherResult
        $DownloadResult = $Downloader.Download()
        $DFailed = 1
        
        switch ($DownloadResult.ResultCode) {
            0 {
                ""
                WriteRed ("Download failed with Error: NotStarted, HResult: " + $DownloadResult.HResult) 
            } 1 {
                "" 
                WriteRed ("Download failed with Error: InProgress, HResult: " + $DownloadResult.HResult)
            } 2 { 
                if($DownloadResult.HResult -eq 0){
                    $DFailed = 0
                    WriteGreen "OK"
                } else {
                    ""
                    WriteRed ("Download failed with Error: HResult: " + $DownloadResult.HResult)
                }
            } 3 {
                ""
                WriteRed ("Download finished with errors, HResult: " + $DownloadResult.HResult)
            } 4 {
                ""
                WriteRed ("Download failed with Error: Failed, HResult: " + $DownloadResult.HResult)
            } 5 {
                ""
                WriteRed ("Download failed with Error: Aborted, HResult: " + $DownloadResult.HResult)
            } default {
                ""
                WriteRed "Download failed with an unknown error"
            }
        }

        if($DFailed){
            Write-Host "List of the Windows updates that failed to download:"
            for ($i=0; $i -lt $Downloader.Updates.Count; $i++) {
                if(-Not $Downloader.Updates.Item($i).IsDownloaded) {
                    "" + ($i + 1) +": " + $Downloader.Updates.Item($i).Title
                }
            }
        }
    }

    if ($install -and $SearcherResult.Count -ge 1) {
        installUpdates $true
    } else {
        wuauReport

        stopServices
    }     
}

function installUpdates {
    param (
        [bool]$fromDownload
    )

    if (!$fromDownload) {
        startServices
    }

    Write-Host "Installing updates..."
    $Session = New-Object -ComObject Microsoft.Update.Session
    Write-Host "Searching for windows updates..." -NoNewline
    $SearchStr = "IsInstalled=0"
    $Searcher = $Session.CreateUpdateSearcher()
    
    try{
        $SearcherResult = $Searcher.Search($SearchStr).Updates
    } catch {
        WriteRed ("Search for Windows Updates failed with error: " + $_.Exception.Message + " Failed Item: " + $_.Exception.ItemName)        
        stopServices
    }

    if($SearcherResult.Count -eq 0) {
        writeGreen "OK"
        Write-Host "No updates found"
    } else {
        $downloadCount = 0
        for ($i=0; $i -lt $SearcherResult.Count; $i++) {
            if ($SearcherResult.Item($i).IsDownloaded)
            {
                $downloadCount++
            }
        }

        ""
        Write-Host ("Found " + $SearcherResult.Count + " updates...") -ForegroundColor Yellow
        Write-Host ($downloadCount.ToString() + " updates are downloaded and ready to be installed.")
        
        for ($i=0; $i -lt $SearcherResult.Count; $i++) {
            if ($SearcherResult.Item($i).IsDownloaded) {
                Write-Host ("" + ($i + 1) + ": " + $SearcherResult.Item($i).Title)
            }
        }
        ""

        $NumberOfUpdate = 0
        $NotInstalledCount = 0
        $ErrorCount = 0
        $NeedsReboot = $false

        Write-Host "Installing "$downloadCount" updates..."
        ""
        for ($i=0; $i -lt $SearcherResult.Count; $i++) {
            $Update = $SearcherResult.Item($i)
            if ($Update.IsDownloaded) {
                Write-Host ("Installing update: " + $Update.Title + "...")

                $NumberOfUpdate++                
                $objCollectionTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
                $objCollectionTmp.Add($Update) | Out-Null
                
                $objInstaller = $Session.CreateUpdateInstaller()
                $objInstaller.Updates = $objCollectionTmp

                try {
                    $InstallResult = $objInstaller.Install()
                } catch {
                    ""
                    $ErrorCount++
                    
                    if($_ -match "HRESULT: 0x80240044") {
                        WriteRed "Your security policy doesn't allow a non-administator idendity to perform this task"
                    } else {
                        WriteRed $_
                    }
                    continue
                }
                
                if (!$NeedsReboot) { 
                    $NeedsReboot = $installResult.RebootRequired 
                }
                
                switch -exact ($InstallResult.ResultCode)
                {
                    0   { $Status = "NotStarted"}
                    1   { $Status = "InProgress"}
                    2   { $Status = "Installed"}
                    3   { $Status = "InstalledWithErrors"}
                    4   { $Status = "Failed"}
                    5   { $Status = "Aborted"}
                }
                    
                switch ($Update.MaxDownloadSize)
                {
                    {[System.Math]::Round($_/1KB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1KB,0))+" KB"; break }
                    {[System.Math]::Round($_/1MB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1MB,0))+" MB"; break }  
                    {[System.Math]::Round($_/1GB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1GB,0))+" GB"; break }    
                    {[System.Math]::Round($_/1TB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1TB,0))+" TB"; break }
                    default { $size = $_+"B" }
                }
                
                $log = New-Object PSObject -Property @{
                    Title = $Update.Title
                    Number = "" + $NumberOfUpdate + "/" + $downloadCount
                    KB = "KB" + $Update.KBArticleIDs
                    Size = $size
                    Status = $Status
                }
                
                if (-NOT $Status -eq "Installed") {
                    $NotInstalledCount++
                    WriteRed $Status
                }

                $log
            }
        }

        WriteGreen ("" + ($NumberOfUpdate - $NotInstalledCount - $ErrorCount) + " updates installed")

        if ($NotInstalledCount -gt 0) {
            WriteRed ("" + $NotInstalledCount + " updates not installed")
        }

        if($ErrorCount -gt 0) {
            WriteRed ("" + $ErrorCount + " updates with errors")
        }

        if ($NeedsReboot) { 
            WriteRed "Requires a restart"
        }
    }

    wuauReport

    stopServices
}

function configureFirewall {
    param (
        $createDelete,
        $WSUSRuleName
    )
    try {
        if ($createDelete -eq "create") {
            Write-Host "Opening Firewall for WSUS..." -NoNewline
        } elseif ($createDelete -eq "delete") {
            Write-Host "Closing Firewall for WSUS..." -NoNewline
        }

        $WSUSServer = getWSUSURL

        if($WSUSServer.Contains(":")) {
            $SplitIndex = $WSUSServer.LastIndexOf(":")
            $WSUSPort = $WSUSServer.Substring($SplitIndex + 1)
            $WSUSDNS = $WSUSServer.Substring(0, $SplitIndex)
            if($WSUSDNS.ToLower().Contains("http://")) {
                $WSUSDNS = $WSUSDNS.Substring(7)
            } else {
                if($WSUSDNS.ToLower().Contains("https://")) {
                    $WSUSDNS = $WSUSDNS.Substring(8)
                }
            }
            
            try {
                $WSUSIp = [System.Net.Dns]::GetHostByName($WSUSDNS).AddressList.IPAddressToString
                if(-NOT $WSUSIp) {
                    $WSUSIp = [System.Net.Dns]::GetHostByName($WSUSDNS).HostName
                }
            } catch {
                WriteRed "Couldn't resolve the WSUS IP address. If you are using DNS, Update the hosts file"
                writeRed $_
                stopServices
            }
        } else {
            WriteRed "Couldn't find the port for the WSUS server"
            writeRed $_
            stopServices
        }

        if ($createDelete -eq "create") {
            New-NetFirewallRule -DisplayName $WSUSRuleName -Direction "Out" -Action "Allow" -Protocol "TCP" -RemotePort $WSUSPort -RemoteAddress $WSUSIp | Out-Null
        } elseif ($createDelete -eq "delete") {
            Remove-NetFirewallRule -DisplayName $WSUSRuleName
        }

        writeGreen "OK"
    } catch {
        writeRed "Failed to $createDelete firewall rule..."
        writeRed $_
    }    
}
########################## END FUNCTIONS #################################################

########################## MAIN SCRIPT BLOCK #############################################

Clear-Host

if (!(getWSUSURL)) {
    do {
        Write-Host "It looks like you've not setup your WSUS URL. Would you like to do that now?"
        $response = Read-Host "(Y/N)"
        switch ($response.ToLower()) {
            'y' { configureWSUS }
            'n' {}
        }
    } until ($response.ToLower() -eq 'y' -or $response.ToLower() -eq 'n')
}

do {
    showMenu
    $selection = Read-Host "Please make a selection"
    ""
    switch ($selection) {
        '1' { configureWSUS }
        '2' { startServices }
        '3' { stopServices }
        '4' { downloadUpdates $false }
        '5' { installUpdates $false }
        '6' { downloadUpdates $true }
        '7' { shutdown.exe -r -t 01
            exit 0 }
        '8' { 
            startServices
            wuauReport
            stopServices
        }
    }
    ""
} until ($selection.ToLower() -eq 'q')

########################### END SCRIPT ###################################################