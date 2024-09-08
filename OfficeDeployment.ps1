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
        
        $AppId = 49117
        $UrlBase = "https://www.microsoft.com/en-us/download/confirmation.aspx"
        $OdtVersion = "$UrlBase`?id=$AppId"
        $LatestVersion = Invoke-WebRequest -UseBasicParsing $OdtVersion
        $FilePattern = "officedeploymenttool_"
        $version = (($LatestVersion.Links | Where-Object {$_.href -match "$FilePattern"}).href).Split("`n")[0].Trim()
        $file_name = "odt.exe"
        
        Write-Output -InputObject "Downloading Office Deployment Tool from $version"
        Start-BitsTransfer -Source $version -Destination "$env:HOMEDRIVE\OfficeDeploymentTool\$file_name" -TransferType Download -Priority Foreground

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
    
    Start-Process -FilePath "$env:HOMEDRIVE\OfficeDeploymentTool\$file_name" -ArgumentList "/extract:$env:HOMEDRIVE\OfficeDeploymentTool /quiet /passive /norestart" -Wait
    Remove-Item -Path "$env:HOMEDRIVE\OfficeDeploymentTool\configuration-*.xml", "$env:HOMEDRIVE\OfficeDeploymentTool\$file_name" -Force
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
[String[]]$FolderPaths = "$($env:HOMEDRIVE)\OfficeDeployment", "$($env:HOMEDRIVE)\OfficeDeployment\OfficeDeploymentTool", "$($env:HOMEDRIVE)\OfficeDeployment\Logs", "$($env:HOMEDRIVE)\OfficeDeployment\SaraCmd", "$($env:HOMEDRIVE)\OfficeDeployment\Configs"
#end folder paths

#start link variables
[String]$SaraCmdLink = "https://aka.ms/SaRA_EnterpriseVersionFiles"
[String]$UrlBase = "https://github.com/Cryozyme/office-deployment/raw/refactor/XmlConfig/"
[String[]]$RemoteFilenames = "professional-vl-officeplus.xml, professional-vl-office.xml, standard-vl-officeplus.xml, standard-vl-office.xml"
[String[]]$O2019Links = "$($UrlBase)Office2019/office/$($RemoteFilenames[0]), $($UrlBase)Office2019/office/$($RemoteFilenames[2]), $($UrlBase)Office2019/office/$($RemoteFilenames[1]), $($UrlBase)Office2019/office/$($RemoteFilenames[3])"
[String[]]$O2021Links = "$($UrlBase)Office2021/office/$($RemoteFilenames[0]), $($UrlBase)Office2021/office/$($RemoteFilenames[2]), $($UrlBase)Office2021/office/$($RemoteFilenames[1]), $($UrlBase)Office2021/office/$($RemoteFilenames[3])"
[String[]]$O365Links = "$($UrlBase)Office365/office/$($RemoteFilenames[0]), $($UrlBase)/Office365/office/$($RemoteFilenames[2]), $($UrlBase)/Office365/office/$($RemoteFilenames[1]), $($UrlBase)/Office365/office/$($RemoteFilenames[3])"
#end link variables

#End Global Variables

function Start-OfficeScrub() {

    [CmdletBinding()]
    param()

    try {
        
        [String]$version = "$([System.Net.HttpWebRequest]::Create(`"https://aka.ms/SaRA_EnterpriseVersionFiles`").GetResponse().ResponseUri.AbsoluteUri)"
        [String]$file_name = "$($version.Split("/")[8])"
        
        Write-Output -InputObject "Downloading SaRACMD Tool from $($version)"

        if($PSBoundParameters["Verbose"] -eq $true) {

            Start-BitsTransfer -Source "$($version)" -Destination "$($FolderPaths[3])\$($file_name)" -TransferType Download -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp, RedirectPolicyAllowReport, RedirectPolicyAllowSilent -Verbose

            Expand-Archive -Path "$($FolderPaths[3])\$($file_name)" -DestinationPath "$($FolderPaths[3])\" -Force -Verbose

            Remove-Item -Path "$($FolderPaths[3])\$($file_name)" -Recurse -Force -Verbose

            $SaRACMDStart = Start-Process -FilePath "$($FolderPaths[3])\SaRACmd.exe" -ArgumentList "-S OfficeScrubScenario -AcceptEula -CloseOffice" -NoNewWindow -PassThru -Wait -Verbose

            if($SaRACMDStart.ExitCode -ne 0) {

                throw -InputObject "SaRACMD could not uninstall Office or Office was not found on the machine"

            } else {

                Write-Output -InputObject "SaRACMD scrubbed Office off the machine successfully"

            }

        } else {

            Start-BitsTransfer -Source "$($version)" -Destination "$($FolderPaths[3])\$($file_name)" -TransferType Download -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp, RedirectPolicyAllowReport, RedirectPolicyAllowSilent

            Expand-Archive -Path "$($FolderPaths[3])\$($file_name)" -DestinationPath "$($FolderPaths[3])\" -Force

            Remove-Item -Path "$($FolderPaths[3])\$($file_name)" -Recurse -Force

            $SaRACMDStart = Start-Process -FilePath "$($FolderPaths[3])\SaRACmd.exe" -ArgumentList "-S OfficeScrubScenario -AcceptEula -CloseOffice" -NoNewWindow -PassThru -Wait

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
        
        [String]$AppId = "49117"
        [String]$UrlBase = "https://www.microsoft.com/en-us/download/confirmation.aspx"
        [String]$OdtVersion = "$($UrlBase)`?id=$($AppId)"
        [String]$FilePattern = "officedeploymenttool_"
        [String]$version = "$((($((Invoke-WebRequest -Uri "$($OdtVersion)").Links) | Where-Object {$_.href -match "$($FilePattern)"}).href).Split("`n")[0].Trim())"
        [String]$file_name = "$($version.Split("/")[8])"
        
        Write-Output -InputObject "Downloading Office Deployment Tool from $($version)"

        Start-BitsTransfer -Source "$($version)" -Destination "$($FolderPaths[1])\$($file_name)" -TransferType Download -Priority Foreground -SecurityFlags IgnoreCertCNInvalid, IgnoreCertDateInvalid, IgnoreCertWrongUsage, IgnoreUnknownCA, RedirectPolicyAllowHttpsToHttp, RedirectPolicyAllowReport, RedirectPolicyAllowSilent

    } catch {

        Write-Output -InputObject "$_"
        Pause

    } finally {

        Clear-Host
        Start-MainMenu

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

    Write-Verbose -Message "Reading user input to determine the selected option"
    Write-Verbose -Message ""
    
    try {
        
        [String]$option = "$(Read-Host -Prompt "Select an option")"

        if($option -eq "1") {

            Write-Output -InputObject "Uninstall Office"
            Start-OfficeScrub

        } elseif($option -eq "2") {

            Write-Output -InputObject "Install Office 365"
            Start-ODTDownload

        } elseif($option -eq "3") {

            Write-Output -InputObject "Install Office 2021"

            Start-ODTDownload

        } elseif($option -eq "4") {

            Write-Output -InputObject "Install Office 2019"

            Start-ODTDownload

        } elseif($option -eq "5") {

            Write-Output -InputObject "Install Office 2016"
            
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
            
        if(-not(Test-Path -Path "$($FolderPaths[$i])" -PathType Container)) {

            Write-Verbose -Message "The script has not been run on this machine before"

            [Boolean]$first_run = $true

        } else {

            Write-Verbose -Message "The script has been run on this machine before. It will delete previous files and folders"

            [Boolean]$first_run = $false

        } if($first_run -eq $false) {

            if($PSBoundParameters["Verbose"] -eq $true) {

                Remove-Item -Path "$($FolderPaths[0])" -Recurse -Force -Verbose

            } else {

                Remove-Item -Path "$($FolderPaths[0])" -Recurse -Force

            }
            
        }

        $FolderPaths.ForEach(

            {

                try {

                    if(-not(Test-Path -Path "$($FolderPaths[$i])" -PathType Container)) {

                        if($PSBoundParameters["Verbose"] -eq $true) {

                            New-Item -Path "$($FolderPaths[$i])" -ItemType Directory -Force -Verbose

                        } else {

                            New-Item -Path "$($FolderPaths[$i])" -ItemType Directory -Force

                        }
                    
                    } else {

                        if($FolderPaths[$i] -eq $FolderPaths[1]) {

                            throw "Deployment folder exists"

                        } elseif($FolderPaths[$i] -eq $FolderPaths[2]) {

                            throw "Log folder exists"

                        } elseif($FolderPaths[$i] -eq $FolderPaths[3]) {

                            throw "SaraCmd folder exists"

                        } elseif($FolderPaths[$i] -eq $FolderPaths[4]) {

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

if($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true) {

    Start-Main -Verbose

} else {

    Start-Main

}