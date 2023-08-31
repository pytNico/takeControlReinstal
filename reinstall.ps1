# '==================================================================================================================================================================
# 'Script to Cleanup and Uninstall Take Control
#'
# 'Disclaimer
# 'The sample scripts are not supported under any SolarWinds support program or service.
# 'The sample scripts are provided AS IS without warranty of any kind.
# 'SolarWinds further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose.
# 'The entire risk arising out of the use or performance of the sample scripts and documentation stays with you.
# 'In no event shall SolarWinds or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever
# '(including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss)
# 'arising out of the use of or inability to use the sample scripts or documentation.
# '==================================================================================================================================================================
function getAgentPath() {
	$script:agentLocationGP = "\Advanced Monitoring Agent GP\"
	$script:agentLocation = "\Advanced Monitoring Agent\"

    try {
        $Keys = Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall -ErrorAction Stop
    } catch {
		$msg = $_.Exception.Message
		$line = $_.InvocationInfo.ScriptLineNumber
		writeToLog F "Error occurred during the lookup of the CurrentVersion\Uninstall Path in the registry, due to:`r`n$msg"
		writeToLog F "This occurred on line number: $line"
		writeToLog F "Failing script."
		Exit 1001
    }

    $Items = $Keys | Foreach-Object {
		Get-ItemProperty $_.PsPath
	}

    ForEach ($Item in $Items) {
        If ($Item.DisplayName -like "Advanced Monitoring Agent" -or $Item.DisplayName -like "Advanced Monitoring Agent GP"){
            $script:localFolder = $Item.installLocation
            #$script:registryPath = $Item.PsPath
            #$script:registryName = $Item.PSChildName
            break
        }
    }

    try {
        $Keys = Get-ChildItem HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall -ErrorAction Stop
    } catch {
		$msg = $_.Exception.Message
		$line = $_.InvocationInfo.ScriptLineNumber
		writeToLog F "Error during the lookup of the WOW6432Node - CurrentVersion\Uninstall Path in the registry, due to:`r`n$msg"
		writeToLog F "This occurred on line number: $line"
		writeToLog F "Failing script."
		Exit 1001
    }
    
    $Items = $Keys | Foreach-Object {
		Get-ItemProperty $_.PsPath
	}
    
    ForEach ($Item in $Items) {
        If ($Item.DisplayName -like "Advanced Monitoring Agent" -or $Item.DisplayName -like "Advanced Monitoring Agent GP"){
            $script:localFolder = $Item.installLocation
            #$script:registryPath = $Item.PsPath
            #$script:registryName = $Item.PSChildName
            break
        }
    }
    
    If (!$script:localFolder) {
        writeToLog F "Installation path for the Advanced Monitoring Agent location was not found."
		Exit 1001
    }
    
    If (($script:localFolder -match '.+?\\$') -eq $false) {
        $script:localFolder = $script:localFolder + "\"
    }

    #writeToLog "INFO: Determined registry path as:`r`n$registryPath"
    #writeToLog "INFO: Determined name as:`r`n$registryName"
    writeToLog I "Agent install location:`r`n$script:localFolder"    
}
function getMSPAID() {
	try {
		[string]$fileContents = Get-Content ($localFolder + "settings.ini") -ErrorAction Stop
	}
	catch {
		$msg = $_.Exception.Message
		$line = $_.InvocationInfo.ScriptLineNumber
		writeToLog F "Error occurred when attempting to get information from the settings.ini file, due to:`r`n$msg"
		writeToLog F "This occurred on line number: $line"
		writeToLog F "Failing script."
		Exit 1001
	}

	If ($fileContents -match "mspid=([0-9A-z]+)") {
		$script:mspaID = $Matches[1]
		writeToLog I "The mspaID has been found. mspaID = $mspaID"
	}
	else {
		writeToLog F "Was unable to detect mspaID from settings.ini file."
		writeToLog F "Failing script."
		Exit 1001
	}
}
function downloadInstaller() {

	$URL = "https://swi-rc.cdn-sw.net/logicnow/Updates/7.00.11/TakeControlAgentInstall-7.00.11-20191126.exe"

	$script:tempFolder = "C:\Temp\mspaInstaller\"
	$script:mspaInstaller = "installer.exe"

	New-Item -Path $tempFolder -ItemType directory -Force | out-null
	$source = $URL
	$dest = $tempFolder + $mspaInstaller

	writeToLog I "Downloading new Take Control agent to $tempFolder."
	
	$wc = New-Object System.Net.WebClient
	try {
  		$wc.DownloadFile($source, $dest)
	} catch {
		$msg = $_.Exception.Message
		writeToLog F "Failed to download installer, due to:`r`n$msg"
		writeToLog F "Failing script."
		Exit 1001
	}
}
function runMSPAUninstaller() {

	$array = @()
	$array += "C:\Program Files (x86)\Take Control Agent\uninstall.exe"
	$array += "C:\Program Files\Take Control Agent\uninstall.exe"

	foreach ($path in $array) {
		If (Test-Path $path) {
			writeToLog I "Detected following path exists:`r`n$path"
			writeToLog I "Now running uninstaller."

			$pinfo = New-Object System.Diagnostics.ProcessStartInfo
			$pinfo.FileName = $path
			$pinfo.RedirectStandardError = $true
			$pinfo.RedirectStandardOutput = $true
			$pinfo.UseShellExecute = $false
			$pinfo.Arguments = "/S /R"
			$p = New-Object System.Diagnostics.Process
			$p.StartInfo = $pinfo
			$p.Start() | Out-Null
			$p.WaitForExit()
			$script:exitCode = $p.exitCode

			writeToLog I "The Exit Code is:" $exitCode
			Start-Sleep -Seconds 10
		}
		If ($exitCode -ne 0) {
			writeToLog F "Uninstaller failed to complete successfully."
			writeToLog F "Failing script."
			Exit 1001
		}
	}
}
function mspaCleanup() {

	$array = @()
    $array += "C:\Program Files (x86)\Take Control Agent"
    $array += "C:\Program Files (x86)\Take Control Agent_&lt;#INSTANCE_NAME#&gt;"
    $array += "C:\Program Files\Take Control Agent"
    $array += "C:\ProgramData\GetSupportService_LOGICnow"
    $array += "C:\ProgramData\GetSupportService_LOGICNow_&lt;#INSTANCE_NAME#&gt;"
    
	foreach ($folderLocation in $array) {
		If (Test-Path $folderLocation) {
			writeToLog I "$folderLocation exists. Now forcing removal of folder."
			try {
				Remove-Item $folderLocation -recurse -force -ErrorAction Stop
			} catch {
				$msg = $_.Exception.Message
				writeToLog W "Failed to remove $folderLocation, due to:`r`n$msg"
				writeToLog W "Continuing with script."
			}
		}
	}
}
function killMSPAProcesses() {

	$array = @()
	$array += "BASupSrvc"
	$array += "BASupSrvcCnfg"
	$array += "BASupSrvcUpdater"
	$array += "BASupTSHelper"
	$array += "BASupClpHlp"

	foreach ($processName in $array) {
		
		$processObj = Get-Process -Name $processName -ErrorAction SilentlyContinue

		If ($processObj) {
			writeToLog I "Detected the $processName process, will now attempt to kill process."
			try {
				$processObj | Stop-Process -Force -ErrorAction Stop
			} catch {
				$msg = $_.Exception.Message
				writeToLog F "Error attempting to kill the $processName process, due to:`r`n$msg"
				writeToLog F "$processName cannot be killed automatically, please perform this manually."
				writeToLog F "Failing script."
				Exit 1001
			}
		}
	}
}
function removeMSPAServices() {

	$array = @()
	$array += "BASupportExpressStandaloneService_LOGICnow"
	$array += "BASupportExpressSrvcUpdater_LOGICnow"

	foreach ($serviceName in $array) {
		If (Get-Service $serviceName -ErrorAction SilentlyContinue) {
			writeToLog I "Detected the $serviceName service, will attempt to remove."
			try {
				Stop-Service -Name $serviceName -ErrorAction Stop
				sc.exe delete $serviceName -ErrorAction Stop
			} catch {
				$msg = $_.Exception.Message
				writeToLog F "Error attempting to stop the $serviceName service, due to:`r`n$msg"
				writeToLog F "$serviceName cannot be stopped automatically, please perform this manually."
				writeToLog F "Failing script."
				Exit 1001
			}
		}
	}
}
function runInstaller() {

	$switches = "/S /R /L /MSPID $mspaID"
	
	$limit = 5
	$stage = 1

	Do {
		writeToLog I "Current installation iteration number:$stage"

		$pinfo = New-Object System.Diagnostics.ProcessStartInfo
		$pinfo.FileName = $extractLocation+$extractedfile
		$pinfo.RedirectStandardError = $true
		$pinfo.RedirectStandardOutput = $true
		$pinfo.UseShellExecute = $false
		$pinfo.Arguments = $switches
		$p = New-Object System.Diagnostics.Process
		$p.StartInfo = $pinfo
		$p.Start() | Out-Null
		$p.WaitForExit()
		$script:exitCode = $p.ExitCode
	
		If ($exitCode -ne 0) {
			writeToLog E "Did not get exitcode = 0, on iteration number: $stage"
			writeToLog E "Exit Code returned:$exitCode"
			$stage++
		} ElseIf (($exitCode -ne 0) -and ($stage -eq $limit)) {
			writeToLog F "Was unable to perform installation. Failing script."
			Exit 1001
		} ElseIf ($exitCode -eq 0) {
			writeToLog I "Successfully returned exitcode 0 from installation."
		}
	}
	Until (($exitCode -eq 0) -or ($stage -gt $limit))
	
	writeToLog I "Now going to remove the files and folders from the temporary directory."
	
	Remove-Item "C:\Temp\mspaInstaller\" -Recurse -Force
}
function getTimeStamp() {
	return "[{0:dd/MM/yy} {0:HH:mm:ss}]" -f (Get-Date)
}
function writeToLog($state, $message) {

	switch -regex -Wildcard ($state) {
		"-" {
			$state = "-"
		}
		"I" {
			$state = "INFO"
		}
		"E" {
			$state = "ERROR"
		}
		"W" {
			$state = "WARNING"
		}
		"F"  {
			$state = "FAILURE"
		}
		""  {
			$state = "INFO"
		}
		Default {
			$state = "INFO"
		}
	 }
	Write-Host "$(getTimeStamp) - [$state]: $message"
}
function main() {
	writeToLog - "Started running getAgentPath function."
	getAgentPath
	writeToLog - "Completed running getAgentPath function."

	writeToLog - "Started running getMSPAID function."
	getMSPAID
	writeToLog - "Completed running getMSPAID function."

	writeToLog - "Started running downloadInstaller function."
	downloadInstaller
	writeToLog - "Completed running downloadInstaller function."

	writeToLog - "Started running runMSPAUninstaller function."
	runMSPAUninstaller
	writeToLog - "Completed running runMSPAUninstaller function."
	
	writeToLog - "Started running mspaCleanup function."
	mspaCleanup
	writeToLog - "Completed running mspaCleanup function."

	writeToLog - "Started running killMSPAProcesses function."
	killMSPAProcesses
	writeToLog - "Completed running killMSPAProcesses function."
	
	writeToLog - "Started running removeMSPAServices function."
	removeMSPAServices
	writeToLog - "Completed running removeMSPAServices function."

	writeToLog - "Started running mspaCleanup function."
	runInstaller
	writeToLog - "Started running mspaCleanup function."

	writeToLog I "Script has completed successfully."
	Exit 0
}
main
