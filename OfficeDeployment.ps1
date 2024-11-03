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
[String[]]$folderPaths = "$($env:TEMP)\OfficeDeployment", "$($env:TEMP)\OfficeDeployment\OfficeDeploymentTool", "$($env:TEMP)\OfficeDeployment\Logs", "$($env:TEMP)\OfficeDeployment\SaraCmd", "$($env:TEMP)\OfficeDeployment\Configs"
#end folder paths

#start link variables
[String]$SaraCmdLink = "https://aka.ms/SaRA_EnterpriseVersionFiles"
[String]$urlBase = "https://github.com/Cryozyme/office-deployment/raw/refactor/XmlConfig/"
[String[]]$remoteFilenames = "enterprise-retail-officeplus.xml", "business-retail-officeplus.xml", "professional-vl-officeplus.xml", "enterprise-retail-office.xml", "business-retail-office.xml", "standard-vl-officeplus.xml", "professional-vl-office.xml", "standard-vl-office.xml"
[String[]]$O2016Links = "$($urlBase)Office2016/office/$($remoteFilenames<#[2]#>)", "$($urlBase)Office2016/office/$($remoteFilenames<#[5]#>)", "$($urlBase)Office2016/office/$($remoteFilenames<#[6]#>)", "$($urlBase)Office2019/office/$($remoteFilenames<#[7]#>)"
[String[]]$O2019Links = "$($urlBase)Office2019/office/$($remoteFilenames[2])", "$($urlBase)Office2019/office/$($remoteFilenames[5])", "$($urlBase)Office2019/office/$($remoteFilenames[6])", "$($urlBase)Office2019/office/$($remoteFilenames[7])"
[String[]]$O2021Links = "$($urlBase)Office2021/office/$($remoteFilenames[2])", "$($urlBase)Office2021/office/$($remoteFilenames[5])", "$($urlBase)Office2021/office/$($remoteFilenames[6])", "$($urlBase)Office2021/office/$($remoteFilenames[7])"
[String[]]$O365Links = "$($urlBase)Office365/office/$($remoteFilenames[0])", "$($urlBase)Office365/office/$($remoteFilenames[1])", "$($urlBase)Office365/office/$($remoteFilenames[3])", "$($urlBase)Office365/office/$($remoteFilenames[4])"
#end link variables

#End Global Variables

function Start-OfficeScrub() {

    [CmdletBinding()]
    param()

    try {
        
        [String]$version = "$([System.Net.HttpWebRequest]::Create(`"$($SaraCmdLink)`").GetResponse().ResponseUri.AbsoluteUri)"
        [String]$fileName = "$($version.Split("/")[8])"
        
        Write-Output -InputObject "Downloading SaRACMD Tool from $($version)"

        Start-BitsTransfer -Source "$($version)" -Destination "$($folderPaths[3])\$($fileName)" -TransferType Download -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp, RedirectPolicyAllowReport, RedirectPolicyAllowSilent

        Expand-Archive -Path "$($folderPaths[3])\$($fileName)" -DestinationPath "$($folderPaths[3])\" -Force

        Remove-Item -Path "$($folderPaths[3])\$($fileName)" -Recurse -Force

        $SaRACMDStart = Start-Process -FilePath "$($folderPaths[3])\SaRACmd.exe" -ArgumentList "-S OfficeScrubScenario -AcceptEula -CloseOffice" -NoNewWindow -PassThru -Wait

        if($SaRACMDStart.ExitCode -ne 0) {

            throw "SaRACMD could not uninstall Office or Office was not found on the machine"

        } else {

            Write-Output -InputObject "SaRACMD scrubbed Office off the machine successfully"

        }

    } catch {

        Write-Output -InputObject "$($_)"
        Pause

    } finally {

        Clear-Host
        Start-MainMenu

    }

    return

}

function Start-ODTDownload() {

    [CmdletBinding()]
    param(

        [Parameter(Mandatory=$true)][Int32]$remoteFileNameIndex,
        [Parameter(Mandatory=$true)][String]$configurationFilePrefix

    )

    try {
        
        [String]$appID = "49117"
        [String]$urlBase = "https://www.microsoft.com/en-us/download/confirmation.aspx"
        [String]$odtVersion = "$($urlBase)`?id=$($appID)"
        [String]$filePattern = "officedeploymenttool_"
        [String]$version = "$((($((Invoke-WebRequest -Uri "$($odtVersion)").Links) | Where-Object {$_.href -match "$($filePattern)"}).href).Split("`n")[0].Trim())"
        [String]$fileName = "$($version.Split("/")[8])"
        
        Write-Output -InputObject "Downloading Office Deployment Tool from $($version)"

        Start-BitsTransfer -Source "$($version)" -Destination "$($folderPaths[1])\$($fileName)" -TransferType Download -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp, RedirectPolicyAllowReport, RedirectPolicyAllowSilent

        $startODTExtractor = Start-Process -FilePath "$($folderPaths[1])\$($fileName)" -ArgumentList "/extract:$($folderPaths[1])\ /norestart /passive /quiet" -NoNewWindow -PassThru -Wait

        if($startODTExtractor.ExitCode -ne 0) {

            throw "The Office Deployment Toolkit either didn't start correctly or it had an error during the extraction process"

        } else {

            Remove-Item -Path "$($folderPaths[1])\$($fileName)" -Recurse -Force
            Remove-Item -Path "$($folderPaths[1])\configuration-Office365-x64.xml" -Recurse -Force

            $startODTSetupDownload = Start-Process -FilePath "$($folderPaths[1])\setup.exe" -ArgumentList "/download $($folderPaths[4])\$($configurationFilePrefix)-$($remoteFilenames[$($remoteFileNameIndex)])" -WorkingDirectory "$($folderPaths[1])" -NoNewWindow -PassThru -Wait

            if($startODTSetupDownload.ExitCode -ne 0) {

                throw "The setup.exe could not download office using the specified XML file"

            } else {

                $startODTSetupConfigure = Start-Process -FilePath "$($folderPaths[1])\setup.exe" -ArgumentList "/configure $($folderPaths[4])\$($configurationFilePrefix)-$($remoteFilenames[$($remoteFileNameIndex)])" -WorkingDirectory "$($folderPaths[1])" -NoNewWindow -PassThru -Wait

                if($startODTSetupConfigure.ExitCode -ne 0) {

                    throw "The setup.exe could not install office using the specified XML file"

                } else {
                }

            }

        }

    } catch {

        Write-Output -InputObject "$($_)"
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
        
        [Parameter(Mandatory=$true)][Int32]$menuSelection,
        [Parameter(Mandatory=$true)][Int32]$remoteFileNameIndex,
        [Parameter(Mandatory=$true)][String]$configurationFilePrefix

    )

    try {

        switch ($menuSelection) {

            1 {
    
                if(-not(Test-Path -Path "$(${env:CommonProgramFiles})\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" -PathType Leaf)) {
    
                    Start-BitsTransfer -Source "$($O365Links[0])" -Destination "$($folderPaths[4])\$($configurationFilePrefix)-$($remoteFilenames[$($remoteFileNameIndex)])" -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp
    
                } else {
    
                    throw "Office is already installed, please uninstall Office before continuing"
    
                }
    
                break
    
            } 2 {
    
                if(-not(Test-Path -Path "$(${env:CommonProgramFiles})\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" -PathType Leaf)) {
    
                    Start-BitsTransfer -Source "$($O365Links[1])" -Destination "$($folderPaths[4])\$($configurationFilePrefix)-$($remoteFilenames[$($remoteFileNameIndex)])" -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp
    
                } else {
    
                    throw "Office is already installed, please uninstall Office before continuing"
    
                }
    
                break
    
            } 3 {
    
                if(-not(Test-Path -Path "$(${env:CommonProgramFiles})\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" -PathType Leaf)) {
    
                    Start-BitsTransfer -Source "$($O365Links[1])" -Destination "$($folderPaths[4])\$($configurationFilePrefix)-$($remoteFilenames[$($remoteFileNameIndex)])" -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp
    
                } else {
    
                    throw "Office is already installed, please uninstall Office before continuing"
    
                }
    
                break
    
            } 4 {
    
                if(-not(Test-Path -Path "$(${env:CommonProgramFiles})\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" -PathType Leaf)) {
    
                    Start-BitsTransfer -Source "$($O365Links[1])" -Destination "$($folderPaths[4])\$($configurationFilePrefix)-$($remoteFilenames[$($remoteFileNameIndex)])" -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp
    
                } else {
    
                    throw "Office is already installed, please uninstall Office before continuing"
    
                }
    
                break
    
            } 5 {
    
                if(-not(Test-Path -Path "$(${env:CommonProgramFiles})\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" -PathType Leaf)) {
    
                    Start-BitsTransfer -Source "$($O365Links[1])" -Destination "$($folderPaths[4])\$($configurationFilePrefix)-$($remoteFilenames[$($remoteFileNameIndex)])" -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp
    
                } else {
    
                    throw "Office is already installed, please uninstall Office before continuing"
    
                }
    
                break
    
            } Default {
    
                break
    
            }
    
        }
        
    } catch {

        Write-Output -InputObject "$($_)"
        Pause
        Clear-Host
        Start-Office365SubMenu

    }

    return

}

function Start-Office365SubMenu() {

    [CmdletBinding()]
    param()

    Write-Verbose -Message "Generating menu and menu options"
    Write-Verbose -Message ""

    Write-Output -InputObject "`nOffice 365 Main Menu`n"
    Write-Output -InputObject "--------------------`n"
    Write-Output -InputObject "1:enterprise-retail-officeplus.xml`n"
    Write-Output -InputObject "2:business-retail-officeplus.xml`n"
    Write-Output -InputObject "3:enterprise-retail-office.xml`n"
    Write-Output -InputObject "4:business-retail-office.xml`n"
    Write-Output -InputObject "5:Back`n"

    try {
        
        Write-Verbose -Message "Reading user input to determine the selected option"
        Write-Verbose -Message ""

        [String]$option = "$(Read-Host -Prompt "Select an option")"

        switch($option) {

            1 {

                Write-Verbose -Message "enterprise-retail-officeplus.xml"

                Start-DownloadConfigurations 1 0 "O365"
                Start-ODTDownload 0 "O365"

                break

            } 2 {

                Write-Verbose -Message "business-retail-officeplus.xml"

                Start-DownloadConfigurations 2 1 "O365"
                Start-ODTDownload 1 "O365"

                break

            } 3 {

                Write-Verbose -Message "enterprise-retail-office.xml"

                Start-DownloadConfigurations 3 3 "O365"
                Start-ODTDownload 3 "O365"

                break

            } 4 {

                Write-Verbose -Message "business-retail-office.xml"

                Start-DownloadConfigurations 4 4 "O365"
                Start-ODTDownload 4 "O365"

                break

            } 5 {

                Clear-Host
                Start-MainMenu

                break

            } Default {

                throw "Invalid Response"
                break

            }

        }

    } catch {

        Write-Verbose -Message "$($_)"
        Pause
        Clear-Host
        Start-Office365SubMenu

    }

    return

}

function Start-MainMenu() {

    [CmdletBinding()]
    param()

    Write-Verbose -Message "Generating menu and menu options"
    Write-Verbose -Message ""
    
    Write-Output -InputObject "`nMain Menu`n"
    Write-Output -InputObject "---------`n"
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

        switch($option) {

            1 {

                Write-Verbose -Message "Uninstall Office"
                Start-OfficeScrub

                break

            } 2 {

                Write-Verbose -Message "Install Office 365"
                Clear-Host
                Start-Office365SubMenu

                break

            } 3 {

                Write-Verbose -Message "Install Office 2021"
                Clear-Host
                Start-Office2021SubMenu

                break

            } 4 {

                Write-Verbose -Message "Install Office 2019"
                Clear-Host
                Start-Office2019SubMenu

                break

            } 5 {

                Write-Verbose -Message "Install Office 2016"
                Clear-Host
                Start-Office2016SubMenu

                break
                
            } 6 {

                if($PSBoundParameters["Verbose"] -eq $true) {

                    Remove-Item -Path "$($folderPaths[0])" -Recurse -Force -Verbose
    
                } else {
    
                    Remove-Item -Path "$($folderPaths[0])" -Recurse -Force
    
                }

                exit 06
                break

            } Default {

                throw "Invalid Response"
                break

            }

        }

    } catch {
        
        Write-Verbose -Message "$($_)"
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

                    Write-Verbose -Message "$($_)"
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

    if($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true) {

        Write-Verbose -Message "Running in `"Verbose`" Mode"
    
        Start-Main -Verbose
    
    } else {
    
        Start-Main
    
    }

}