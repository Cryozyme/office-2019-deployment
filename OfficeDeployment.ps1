<#

Program Name: Automated Office Deployment Script
Copyright (C) 2024 Alan R. Markley

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published
by the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.

Developer Contact: alternate8592@duck.com

function installO365() {

    Write-Output -InputObject "Downloading Office 365"

    for($i=0;$i -le 4;$i++) {

        try {
            
            $o365_download = Start-Process -FilePath "$env:HOMEDRIVE\OfficeDeploymentTool\setup.exe" -ArgumentList "/download $o365_modified" -PassThru -Wait

            if($o365_download.ExitCode -eq 0) {

                return

            } else {

                throw "Office Deployment Toolkit Could Not Download Microsoft Office`nCHECK YOUR INTERNET CONNECTION"

            }

        } catch {

            if($i -eq 4) {
            
                Write-Output -InputObject "$_"

            } else {

                continue

            }

        }

    }

    Write-Output -InputObject "Installing Office 365"

    try {

        $o365_install = Start-Process -FilePath "$env:HOMEDRIVE\OfficeDeploymentTool\setup.exe" -ArgumentList "/configure $o365_modified" -PassThru -Wait

        if($o365_install.ExitCode -eq 0) {

            return

        } else {

            throw "Office Deployment Toolkit Could Not Install Microsoft Office"

        }
        
    } catch {

        Write-Output -InputObject "$_"

    }

    return

}

function installO2021() {

    Write-Output -InputObject "Downloading Office 2021..."

    for($i=0;$i -le 4;$i++) {

        try {
            
            $o2021_download = Start-Process -FilePath "$env:HOMEDRIVE\OfficeDeploymentTool\setup.exe" -ArgumentList "/download $o2021_modified" -PassThru -Wait
            
            if($o2021_download.ExitCode -eq 0) {

                return

            } else {

                throw "Office Deployment Toolkit Could Not Download Microsoft Office`nCHECK YOUR INTERNET CONNECTION"

            }

        } catch {

            if($i -eq 4) {
            
                Write-Output -InputObject "$_"

            } else {

                continue

            }

        }

    }

    Write-Output -InputObject "Installing Office 2021..."

    try {
        
        $o2021_install = Start-Process -FilePath "$env:HOMEDRIVE\OfficeDeploymentTool\setup.exe" -ArgumentList "/configure $o2021_modified" -PassThru -Wait

        if($o2021_install.ExitCode -eq 0) {

            return

        } else {

            throw "Office Deployment Toolkit Could Not Install Microsoft Office"

        }

    } catch {
        
        Write-Output -InputObject "$_"

    }
    
    return

}

function installO2019() {

    Write-Output -InputObject "Downloading Office 2019..."

    for($i=0;$i -le 4;$i++) {
        
        try {

            $o2019_download = Start-Process -FilePath "$env:HOMEDRIVE\OfficeDeploymentTool\setup.exe" -ArgumentList "/download $o2019_modified" -PassThru -Wait

            if($o2019_download.ExitCode -eq 0) {
                
                return
            
            } else {
                
                throw "Office Deployment Toolkit Could Not Download Microsoft Office`nCHECK YOUR INTERNET CONNECTION"
            
            }
            
        } catch {
            
            if($i -eq 4) {
            
                Write-Output -InputObject "$_"

            } else {

                continue

            }

        }

    }

    Write-Output -InputObject "Installing Office 2019..."

    try {
        
        $o2019_install = Start-Process -FilePath "$env:HOMEDRIVE\OfficeDeploymentTool\setup.exe" -ArgumentList "/configure $o2019_modified" -PassThru -Wait

        if($o2019_install.ExitCode -eq 0) {
            
            return
        
        } else {
            
            throw "Office Deployment Toolkit Could Not Install Microsoft Office"
        
        }

    } catch {
        
        Write-Output -InputObject "$_"

    }
    
    return

}

function installO2016() {

    Write-Output -InputObject "Downloading Office 2016..."

    for($i=0;$i -le 4;$i++) {
        
        try {

            $o2016_download = Start-Process -FilePath "$env:HOMEDRIVE\OfficeDeploymentTool\setup.exe" -ArgumentList "/download $o2016_modified" -PassThru -Wait

            if($o2016_download.ExitCode -eq 0) {
                
                return
            
            } else {
                
                throw "Office Deployment Toolkit Could Not Download Microsoft Office`nCHECK YOUR INTERNET CONNECTION"
            
            }
        
        } catch {
        
            if($i -eq 4) {

                Write-Output -InputObject "$_"

            } else {

                continue

            }

        }

    }

    Write-Output -InputObject "Installing Office 2016..."

    try {
        
        $o2016_install = Start-Process -FilePath "$env:HOMEDRIVE\OfficeDeploymentTool\setup.exe" -ArgumentList "/configure $o2016_modified" -PassThru -Wait

        if($o2016_install.ExitCode -eq 0) {
            
            return
        
        } else {

            throw "Office Deployment Toolkit Could Not Install Microsoft Office"

        }

    } catch {

        Write-Output -InputObject "$_"

    }
    
    return

}

function uninstallPrevious() {

    try {

        $odt_office_scrub = Start-Process -FilePath "$env:HOMEDRIVE\OfficeDeploymentTool\setup.exe" -ArgumentList "/configure $uninstall" -PassThru -Wait

        if($odt_office_scrub.ExitCode -eq 0) {

            return

        } else {
            
            throw "Office Deployment Toolkit Could Not Uninstall Office"            

        }

    } catch {

        Write-Output -InputObject "$_"

        try {

            $office_scrub = Start-Process -FilePath "$env:HOMEDRIVE\OfficeDeploymentTool\SaRACMD.exe" -ArgumentList "-S OfficeScrubScenario -AcceptEula -OfficeVersion All" -PassThru -Wait
            
            if($office_scrub.ExitCode -eq 0) {

                return
    
            } else {
                
                throw "SaRACMD Could Not Uninstall Office"            
    
            }

        } catch {
            
            Write-Output -InputObject "$_"

        }
    
    }
    
    return

}

function configDownload() {

    try {
        
        $appID = 49117
        $urlBase = "https://www.microsoft.com/en-us/download/confirmation.aspx"
        $odtVersion = "$urlBase`?id=$appID"
        $LatestVersion = Invoke-WebRequest -UseBasicParsing $odtVersion
        $filePattern = "officedeploymenttool_"
        $version = (($LatestVersion.Links | Where-Object {$_.href -match "$filePattern"}).href).Split("`n")[0].Trim()
        $fileName = "odt.exe"
        
        Write-Output -InputObject "Downloading Office Deployment Tool from $version"
        Start-BitsTransfer -Source $version -Destination "$env:HOMEDRIVE\OfficeDeploymentTool\$fileName" -TransferType Download -Priority Foreground

    } catch {

        Write-Output -InputObject "$_"

    } try {

        $saracmd_version = "https://download.microsoft.com/download/5/0/d/50dd45c9-f465-402e-92d2-537871f1f106/SaRACmd_17_01_1440_0.zip"
        $saracmd_filename = "sara.zip"
        
        Write-Output -InputObject "SaRACMD from $saracmd_version"
        Start-BitsTransfer -Source $saracmd_version -Destination "$env:HOMEDRIVE\OfficeDeploymentTool\$saracmd_filename" -TransferType Download -Priority Foreground

    } catch {
        
        Write-Output -InputObject "$_"

    }
    
    Start-Process -FilePath "$env:HOMEDRIVE\OfficeDeploymentTool\$fileName" -ArgumentList "/extract:$env:HOMEDRIVE\OfficeDeploymentTool /quiet /passive /norestart" -Wait
    Remove-Item -Path "$env:HOMEDRIVE\OfficeDeploymentTool\configuration-*.xml", "$env:HOMEDRIVE\OfficeDeploymentTool\$fileName" -Force
    Expand-Archive -Path "$env:HOMEDRIVE\OfficeDeploymentTool\sara.zip" -DestinationPath "$env:HOMEDRIVE\OfficeDeploymentTool\" -Force
    Remove-Item -Path "$env:HOMEDRIVE\OfficeDeploymentTool\sara.zip" -Force
    
    return

}

function serviceHandling() {

    sc.exe stop "ClickToRunSvc" | Out-Null
    sc.exe stop "ose64" | Out-Null

    Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "POWERPNT" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "MSACCESS" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "MSPUB" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "OUTLOOK" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "WINPROJ" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "VISIO" -Force -ErrorAction SilentlyContinue

    return

}

#Annoying Apps Removed Config Links
$o365_modified = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-365/office-365-modified.xml"
$o2021_modified = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-2021/office-2021-modified.xml"
$o2019_modified = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-2019/office-2019-modified.xml"
$o2016_modified = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-2016/office-2016-modified.xml"

Full Installation Config Links
$o365_full = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-365/office-365-full.xml"
$o2021_full = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-2021/office-2021-full.xml"
$o2019_full = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-2019/office-2019-full.xml"
$o2016_full = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-2016/office-2016-full.xml"

#Uninstall Config Link
$uninstall = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/uninstall/office-uninstall.xml"

#>








<#---------------------------------------------------------------------------------------------------------------------------------------#>
[CmdletBinding()]
param()

<#
.SYNOPSIS
    Program Name: Automated Office Deployment Script
    Copyright (C) 2024 Alan R. Markley

    This program aims to aid in deploying Click-to-Run(C2R) versions of Microsoft Office.
    Supports versions 2016, 2019, 2021, and 365.
    
.NOTES
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU Affero General Public License as published
    by the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU Affero General Public License for more details.
    You should have received a copy of the GNU Affero General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.

    Developer Contact: alternate8592@duck.com
#>

#Start Global Variables

#start folder paths
[String[]]$folderPaths = "$($env:HOMEDRIVE)\OfficeDeployment", "$($env:HOMEDRIVE)\OfficeDeployment\OfficeDeploymentTool", "$($env:HOMEDRIVE)\OfficeDeployment\Logs", "$($env:HOMEDRIVE)\OfficeDeployment\SaraCmd", "$($env:HOMEDRIVE)\OfficeDeployment\Configs"
#end folder paths

#start link variables
[String]$SaraCmdLink = "https://aka.ms/SaRA_EnterpriseVersionFiles"
[String]$urlBase = "https://github.com/Cryozyme/office-deployment/raw/refactor/XmlConfig/"
[String[]]$remoteFilenames = "enterprise-retail-officeplus.xml", "business-retail-officeplus.xml", "professional-vl-officeplus.xml", "enterprise-retail-office.xml", "business-retail-office.xml", "standard-vl-officeplus.xml", "professional-vl-office.xml", "standard-vl-office.xml"
[String[]]$O2019Links = "$($urlBase)Office2019/office/$($remoteFilenames[0])", "$($urlBase)Office2019/office/$($remoteFilenames[2])", "$($urlBase)Office2019/office/$($remoteFilenames[1])", "$($urlBase)Office2019/office/$($remoteFilenames[3])"
[String[]]$O2021Links = "$($urlBase)Office2021/office/$($remoteFilenames[0])", "$($urlBase)Office2021/office/$($remoteFilenames[2])", "$($urlBase)Office2021/office/$($remoteFilenames[1])", "$($urlBase)Office2021/office/$($remoteFilenames[3])"
[String[]]$O365Links = "$($urlBase)Office365/office/$($remoteFilenames[0])", "$($urlBase)Office365/office/$($remoteFilenames[2])", "$($urlBase)Office365/office/$($remoteFilenames[1])", "$($urlBase)Office365/office/$($remoteFilenames[3])"
#end link variables

#End Global Variables

function Start-OfficeScrub() {

    [CmdletBinding()]
    param()

    try {
        
        [String]$version = "$([System.Net.HttpWebRequest]::Create(`"$($SaraCmdLink)`").GetResponse().ResponseUri.AbsoluteUri)"
        [String]$fileName = "$($version.Split("/")[8])"
        
        Write-Output -InputObject "Downloading SaRACMD Tool from $($version)"

        if($PSBoundParameters["Verbose"] -eq $true) {

            Start-BitsTransfer -Source "$($version)" -Destination "$($folderPaths[3])\$($fileName)" -TransferType Download -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp, RedirectPolicyAllowReport, RedirectPolicyAllowSilent -Verbose

            Expand-Archive -Path "$($folderPaths[3])\$($fileName)" -DestinationPath "$($folderPaths[3])\" -Force -Verbose

            Remove-Item -Path "$($folderPaths[3])\$($fileName)" -Recurse -Force -Verbose

            $SaRACMDStart = Start-Process -FilePath "$($folderPaths[3])\SaRACmd.exe" -ArgumentList "-S OfficeScrubScenario -AcceptEula -CloseOffice" -NoNewWindow -PassThru -Wait -Verbose

            if($SaRACMDStart.ExitCode -ne 0) {

                throw -InputObject "SaRACMD could not uninstall Office or Office was not found on the machine. Alternatively, you did not run this script as administrator"

            } else {

                Write-Output -InputObject "SaRACMD scrubbed Office off the machine successfully"

            }

        } else {

            Start-BitsTransfer -Source "$($version)" -Destination "$($folderPaths[3])\$($fileName)" -TransferType Download -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp, RedirectPolicyAllowReport, RedirectPolicyAllowSilent

            Expand-Archive -Path "$($folderPaths[3])\$($fileName)" -DestinationPath "$($folderPaths[3])\" -Force

            Remove-Item -Path "$($folderPaths[3])\$($fileName)" -Recurse -Force

            $SaRACMDStart = Start-Process -FilePath "$($folderPaths[3])\SaRACmd.exe" -ArgumentList "-S OfficeScrubScenario -AcceptEula -CloseOffice" -NoNewWindow -PassThru -Wait

            if($SaRACMDStart.ExitCode -ne 0) {

                throw "SaRACMD could not uninstall Office or Office was not found on the machine"

            } else {

                Write-Output -InputObject "SaRACMD scrubbed Office off the machine successfully"

            }

        }

    } catch {

        Write-Output -InputObject "$_"
        Pause

    } finally {

        Clear-Host
        Start-MainMenu

    }

    return

}

function Start-ODTDownload() {

    [CmdletBinding()]
    param()

    try {
        
        [String]$appID = "49117"
        [String]$urlBase = "https://www.microsoft.com/en-us/download/confirmation.aspx"
        [String]$odtVersion = "$($urlBase)`?id=$($appID)"
        [String]$filePattern = "officedeploymenttool_"
        [String]$version = "$((($((Invoke-WebRequest -Uri "$($odtVersion)").Links) | Where-Object {$_.href -match "$($filePattern)"}).href).Split("`n")[0].Trim())"
        [String]$fileName = "$($version.Split("/")[8])"
        
        Write-Output -InputObject "Downloading Office Deployment Tool from $($version)"

        if($PSBoundParameters["Verbose"] -eq $true){

            Start-BitsTransfer -Source "$($version)" -Destination "$($folderPaths[1])\$($fileName)" -TransferType Download -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp, RedirectPolicyAllowReport, RedirectPolicyAllowSilent -Verbose
            
            $startODTExtractor = Start-Process -FilePath "$($folderPaths[1])\$($fileName)" -ArgumentList "/extract:$($folderPaths[1])\ /norestart /passive /quiet" -NoNewWindow -PassThru -Wait -Verbose

            if($startODTExtractor.ExitCode -ne 0) {

               throw "The Office Deployment Toolkit either didn't start correctly or it had an error during the extraction process"

            } else {

                Remove-Item -Path "$($folderPaths[1])\$($fileName), $($folderPaths[1])\configuration-Office365-x64.xml" -Recurse -Force -Verbose
                Remove-Item -Path "$($folderPaths[1])\configuration-Office365-x64.xml" -Recurse -Force -Verbose

                $startODTSetupDownload = Start-Process -FilePath "$($folderPaths[1])\setup.exe" -ArgumentList "/download $($folderPaths[4])\O365-$($remoteFilenames[0])" -NoNewWindow -PassThru -Wait -WorkingDirectory "$($folderPaths[1])\" -Verbose

                if($startODTSetupDownload.ExitCode -ne 0) {
                    
                    throw "The setup.exe could not download office using the specified XML file"

                } else {

                    $startODTSetupConfigure = Start-Process -FilePath "$($folderPaths[1])\setup.exe" -ArgumentList "/configure $($folderPaths[4])\O365-$($remoteFilenames[0])" -NoNewWindow -PassThru -Wait -WorkingDirectory "$($folderPaths[1])" -Verbose

                    if($startODTSetupConfigure.ExitCode -ne 0) {

                       throw "The setup.exe could not install office using the specified XML file"

                    } else {
                    }

                }

            }

        } else {

            Start-BitsTransfer -Source "$($version)" -Destination "$($folderPaths[1])\$($fileName)" -TransferType Download -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp, RedirectPolicyAllowReport, RedirectPolicyAllowSilent

            $startODTExtractor = Start-Process -FilePath "$($folderPaths[1])\$($fileName)" -ArgumentList "/extract:$($folderPaths[1])\ /norestart /passive /quiet" -NoNewWindow -PassThru -Wait

            if($startODTExtractor.ExitCode -ne 0) {

                throw "The Office Deployment Toolkit either didn't start correctly or it had an error during the extraction process"

            } else {

                Remove-Item -Path "$($folderPaths[1])\$($fileName)" -Recurse -Force
                Remove-Item -Path "$($folderPaths[1])\configuration-Office365-x64.xml" -Recurse -Force

                $startODTSetupDownload = Start-Process -FilePath "$($folderPaths[1])\setup.exe" -ArgumentList "/download $($folderPaths[4])\O365-$($remoteFilenames[0])" -NoNewWindow -PassThru -Wait -WorkingDirectory "$($folderPaths[1])"

                if($startODTSetupDownload.ExitCode -ne 0) {

                    throw "The setup.exe could not download office using the specified XML file"

                } else {

                    $startODTSetupConfigure = Start-Process -FilePath "$($folderPaths[1])\setup.exe" -ArgumentList "/configure $($folderPaths[4])\O365-$($remoteFilenames[0])" -NoNewWindow -PassThru -Wait -WorkingDirectory "$($folderPaths[1])"

                    if($startODTSetupConfigure.ExitCode -ne 0) {

                        throw "The setup.exe could not install office using the specified XML file"

                    } else {
                    }

                }

            }

        }

    } catch {

        Write-Output -InputObject "$_"
        Pause

    } finally {

        Clear-Host
        Start-MainMenu

    }

    return

}

function Start-DownloadConfigurations() {

    [CmdletBinding()]
    param(
        
        [Parameter(Mandatory=$true)][Int32]$menuSelection

    )

    switch ($menuSelection) {

        2 {

            Start-BitsTransfer -Source "$($O365Links[0])" -Destination "$($folderPaths[4])\O365-$($remoteFilenames[0])" -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp

            return

        } 3 {

            return

        } 4 {

            return

        } 5 {

            return

        } Default {

            return

        }

    }

    return

}

function Start-MainMenu() {

    [CmdletBinding()]
    param()

    Write-Verbose -Message "Generating menu and menu options"
    Write-Verbose -Message ""
    
    Write-Output -InputObject "`nMain Menu`n"
    Write-Output -InputObject "----------------`n"
    Write-Output -InputObject "1:Uninstall Office`n"
    Write-Output -InputObject "2:Install Office 365`n"
    Write-Output -InputObject "3:Install Office 2021`n"
    Write-Output -InputObject "4:Install Office 2019`n"
    Write-Output -InputObject "5:Install Office 2016`n"
    Write-Output -InputObject "6:Exit`n"
    
    try {
        
        Write-Verbose -Message "Reading user input to determine the selected option"
        Write-Verbose -Message ""

        [String]$option = "$(Read-Host -Prompt "Select an option")"

        if($option -eq "1") {

            Write-Verbose -Message "Uninstall Office"
            Start-OfficeScrub

        } elseif($option -eq "2") {

            Write-Verbose -Message "Install Office 365"
            Start-DownloadConfigurations 2
            Start-ODTDownload

        } elseif($option -eq "3") {

            Write-Verbose -Message "Install Office 2021"

            Start-DownloadConfigurations 3
            Start-ODTDownload

        } elseif($option -eq "4") {

            Write-Verbose -Message "Install Office 2019"

            Start-DownloadConfigurations 4
            Start-ODTDownload

        } elseif($option -eq "5") {

            Write-Verbose -Message "Install Office 2016"
            
            Start-DownloadConfigurations 5
            Start-ODTDownload
            
        } elseif($option -eq "6") {

            exit 6

        } else {

            throw "Invalid Response"

        }

    } catch {
        
        Write-Verbose -Message "$_"
        Pause
        Clear-Host
        Start-MainMenu

    }

    return

}

function Start-Main() {

    [CmdletBinding()]
    param()

    Write-Verbose -Message ""
    Write-Verbose -Message "Verifying if the script is running on Windows"
    Write-Verbose -Message ""

    if($IsWindows -eq $true) {

        Write-Verbose -Message "Operating system is Windows!"
        Write-Verbose -Message ""
        Write-Verbose -Message "Starting file operations"
        Write-Verbose -Message ""

        [Int32]$i = 0
            
        if(-not(Test-Path -Path "$($folderPaths[$i])" -PathType Container)) {

            Write-Verbose -Message "The script has not been run on this machine before"

            [Boolean]$firstRun = $true

        } else {

            Write-Verbose -Message "The script has been run on this machine before. It will delete previous files and folders"

            [Boolean]$firstRun = $false

        } if($firstRun -eq $false) {

            if($PSBoundParameters["Verbose"] -eq $true) {

                Remove-Item -Path "$($folderPaths[0])" -Recurse -Force -Verbose

            } else {

                Remove-Item -Path "$($folderPaths[0])" -Recurse -Force

            }
            
        }

        $folderPaths.ForEach(

            {

                try {

                    if(-not(Test-Path -Path "$($folderPaths[$i])" -PathType Container)) {

                        if($PSBoundParameters["Verbose"] -eq $true) {

                            New-Item -Path "$($folderPaths[$i])" -ItemType Directory -Force -Verbose

                        } else {

                            New-Item -Path "$($folderPaths[$i])" -ItemType Directory -Force

                        }
                    
                    } else {

                        if($folderPaths[$i] -eq $folderPaths[1]) {

                            throw "Deployment folder exists"

                        } elseif($folderPaths[$i] -eq $folderPaths[2]) {

                            throw "Log folder exists"

                        } elseif($folderPaths[$i] -eq $folderPaths[3]) {

                            throw "SaraCmd folder exists"

                        } elseif($folderPaths[$i] -eq $folderPaths[4]) {

                            throw "Configuration folder exists"

                        } else {

                            throw "Index out of range or there is an array and global file-path variable mismatch"

                        }
                        
                    }
            
                } catch {

                    Write-Verbose -Message "$_"
                    Write-Verbose -Message ""
            
                } finally {

                    $i++

                }

            }

        )

        Start-MainMenu

    } else {

        Write-Verbose -Message "Operating system is not Windows"
        Write-Verbose -Message ""
        
        Write-Output -InputObject ""
        Write-Output -InputObject "Either the script is not running in a Windows environment (Please run this script in Windows 10+)"
        Write-Output -InputObject "OR"
        Write-Output -InputObject "The script is being not being run in PowerShell version 7 or greater"
        Write-Output -InputObject ""

        Pause

    }

    return

}

Clear-Host

if (-not([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {

    Write-Output -InputObject "The script will not continue because it must be launched in administrator mode to work properly"
    Write-Verbose -Message "Elevation is required to run SaRACMD and Office Deployment Toolkit"
    exit

} else {
}

if($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true) {

    Start-Main -Verbose

} else {

    Start-Main

}