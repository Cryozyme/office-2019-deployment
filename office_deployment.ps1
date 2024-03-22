function installO365() {

    Write-Host "Downloading Office 365"

    for($i=0;$i -le 4;$i++) {

        try {
            
            $o365_download = Start-Process -FilePath "$env:homedrive\odt\setup.exe" -ArgumentList "/download $o365_modified" -PassThru -Wait

            if($o365_download.ExitCode -eq 0) {

                return

            } else {

                throw "Office Deployment Toolkit Could Not Download Microsoft Office`nCHECK YOUR INTERNET CONNECTION."

            }

        } catch {

            if($i -eq 4) {
            
                Write-Host "$_"

            } else {

                continue

            }

        }

    }

    Write-Host "Installing Office 365"

    try {

        $o365_install = Start-Process -FilePath "$env:homedrive\odt\setup.exe" -ArgumentList "/configure $o365_modified" -PassThru -Wait

        if($o365_install.ExitCode -eq 0) {

            return

        } else {

            throw "Office Deployment Toolkit Could Not Install Microsoft Office"

        }
        
    } catch {

        Write-Host "$_"

    }

    return

}

function installO2021() {

    Write-Host "Downloading Office 2021..."

    for($i=0;$i -le 4;$i++) {

        try {
            
            $o2021_download = Start-Process -FilePath "$env:homedrive\odt\setup.exe" -ArgumentList "/download $o2021_modified" -PassThru -Wait
            
            if($o2021_download.ExitCode -eq 0) {

                return

            } else {

                throw "Office Deployment Toolkit Could Not Download Microsoft Office`nCHECK YOUR INTERNET CONNECTION."

            }

        } catch {

            if($i -eq 4) {
            
                Write-Host "$_"

            } else {

                continue

            }

        }

    }

    Write-Host "Installing Office 2021..."

    try {
        
        $o2021_install = Start-Process -FilePath "$env:homedrive\odt\setup.exe" -ArgumentList "/configure $o2021_modified" -PassThru -Wait

        if($o2021_install.ExitCode -eq 0) {

            return

        } else {

            throw "Office Deployment Toolkit Could Not Install Microsoft Office"

        }

    } catch {
        
        Write-Host "$_"

    }
    
    return

}

function installO2019() {

    for($i=0;$i -le 4;$i++) {
        
        try {

            $o2019_download = Start-Process -FilePath "$env:homedrive\odt\setup.exe" -ArgumentList "/download $o2019_modified" -PassThru -Wait

            if($o2019_download.ExitCode -eq 0) {
                
                return
            
            } else {
                
                throw "Office Deployment Toolkit Could Not Download Microsoft Office`nCHECK YOUR INTERNET CONNECTION."
            
            }
            
        } catch {
            
            if($i -eq 4) {
            
                Write-Host "$_"

            } else {

                continue

            }

        }

    }

    try {
        
        $o2019_install = Start-Process -FilePath "$env:homedrive\odt\setup.exe" -ArgumentList "/configure $o2019_modified" -PassThru -Wait

        if($o2019_install.ExitCode -eq 0) {
            
            return
        
        } else {
            
            throw "Office Deployment Toolkit Could Not Install Microsoft Office"
        
        }

    } catch {
        
        Write-Host "$_"

    }
    
    return

}

function installO2016() {

    $o2016_download = Start-Process -FilePath "$env:homedrive\odt\setup.exe" -ArgumentList "/download $o2016_modified" -PassThru -Wait
    $o2016_install = Start-Process -FilePath "$env:homedrive\odt\setup.exe" -ArgumentList "/configure $o2016_modified" -PassThru -Wait
    return

}

function uninstallPrevious() {

    try {

        $odt_office_scrub = Start-Process -FilePath "$env:homedrive\odt\setup.exe" -ArgumentList "/configure $uninstall" -PassThru -Wait

        if($odt_office_scrub.ExitCode -eq 0) {

            return

        } else {
            
            throw "Office Deployment Toolkit Could Not Uninstall Office"            

        }

    } catch {

        Write-Host "$_"

        try {

            $office_scrub = Start-Process -FilePath "$env:homedrive\odt\SaRACMD.exe" -ArgumentList "-S OfficeScrubScenario -AcceptEula -OfficeVersion All" -PassThru -Wait
            
            if($office_scrub.ExitCode -eq 0) {

                return
    
            } else {
                
                throw "SaRACMD Could Not Uninstall Office"            
    
            }

        } catch {
            
            Write-Host "$_"

        }
    
    }
    
    return

}

function configDownload() {

    try {
        
        $app_id = 49117
        $url_base = "https://www.microsoft.com/en-us/download/confirmation.aspx"
        $odt_version = "$url_base`?id=$app_id"
        $latest = Invoke-WebRequest -UseBasicParsing $odt_version
        $file_pattern = "officedeploymenttool_"
        $version = (($latest.Links | Where-Object {$_.href -match "$file_pattern"}).href).Split("`n")[0].Trim()
        $file_name = "odt.exe"
        
        Write-Host "Downloading Office Deployment Tool from $version"
        Start-BitsTransfer -Source $version -Destination "$env:homedrive\odt\$file_name" -TransferType Download -Priority Foreground

    } catch {

        Write-Host "$_"

    } try {

        $saracmd_version = "https://download.microsoft.com/download/5/0/d/50dd45c9-f465-402e-92d2-537871f1f106/SaRACmd_17_01_1440_0.zip"
        $saracmd_filename = "sara.zip"
        
        Write-Host "SaRACMD from $saracmd_version"
        Start-BitsTransfer -Source $saracmd_version -Destination "$env:homedrive\odt\$saracmd_filename" -TransferType Download -Priority Foreground

    } catch {
        
        Write-Host "$_"

    }
    
    Start-Process -FilePath "$env:homedrive\odt\$file_name" -ArgumentList "/extract:$env:homedrive\odt /quiet /passive /norestart" -Wait
    Remove-Item -Path "$env:homedrive\odt\configuration-*.xml", "$env:homedrive\odt\$file_name" -Force
    Expand-Archive -Path "$env:homedrive\odt\sara.zip" -DestinationPath "$env:homedrive\odt\" -Force
    Remove-Item -Path "$env:homedrive\odt\sara.zip" -Force
    
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

function fileHandling() {

    $deployment_tool_path = "$env:homedrive\odt"
    $logging_path_uninstall = "C:\office_deployment\odt_uninstall"
    $logging_path_365 = "C:\office_deployment\odt_365"
    $logging_path_2021 = "C:\office_deployment\odt_2021"
    $logging_path_2019 = "C:\office_deployment\odt_2019"
    $logging_path_2016 = "C:\office_deployment\odt_2016"

    try {
        
        if(-not(Test-Path -Path $deployment_tool_path -PathType Container)) {
        
            New-Item -Path "$deployment_tool_path" -ItemType Directory -Force
        
        } else {

            throw "Deployment Path Already Exists"
            
        }

    } catch {

        Write-Host "$_"

    }

    try {

        if(-not(Test-Path -Path $logging_path_uninstall -PathType Container)) {

            New-Item -Path "$logging_path_uninstall" -ItemType Directory -Force

        } else {

            throw "Uninstall Logging Path Already Exists"

        }

    } catch {

        Write-Host "$_"

    }
    
    try {

        if(-not(Test-Path -Path $logging_path_365 -PathType Container)) {

            New-Item -Path "$logging_path_365" -ItemType Directory -Force

        } else {

            throw "Office 365 Logging Path Already Exists"

        }

    } catch {

        Write-Host "$_"

    }

    try {

        if(-not(Test-Path -Path $logging_path_2021 -PathType Container)) {

            New-Item -Path "$logging_path_2021" -ItemType Directory -Force

        } else {

            throw "Office 2021 Logging Path Already Exists"

        }

    } catch {

        Write-Host "$_"

    }

    try {

        if(-not(Test-Path -Path $logging_path_2019 -PathType Container)) {

            New-Item -Path "$logging_path_2019" -ItemType Directory -Force

        } else {

            throw "Office 2019 Logging Path Already Exists"

        }

    } catch {

        Write-Host "$_"

    }

    try {

        if(-not(Test-Path -Path $logging_path_2016 -PathType Container)) {

            New-Item -Path "$logging_path_2016" -ItemType Directory -Force

        } else {

            throw "Office 2016 Logging Path Already Exists"

        }

    } catch {

        Write-Host "$_"

    }

    return

}

function optionSelection() {

    Clear-Host

    $option = "$(Read-Host -Prompt "Select an option`n----------------`n1:Uninstall Office`n2:Install Office 365`n3:Install Office 2021`n4:Install Office 2019`n5:Install Office 2016`n6:Exit`n")"

    try {

        if($option.Trim() -eq "1") {

            serviceHandling
            configDownload
            uninstallPrevious
            optionSelection

        } elseif($option.Trim() -eq "2") {

            serviceHandling
            configDownload
            uninstallPrevious
            installO365
            optionSelection

        } elseif($option.Trim() -eq "3") {

            serviceHandling
            configDownload
            uninstallPrevious
            installO2021
            optionSelection

        } elseif($option.Trim() -eq "4") {

            serviceHandling
            configDownload
            uninstallPrevious
            installO2019
            optionSelection

        } elseif($option.Trim() -eq "5") {

            serviceHandling
            configDownload
            uninstallPrevious
            installO2016
            optionSelection

        } elseif($option.Trim() -eq "6") {

            Set-Location -Path "..\"
            Remove-Item -Path "$env:homedrive\odt" -Recurse -Force
            Remove-Item -Path "$env:homedrive\office_deployment" -Recurse -Force
            Write-Host "Finished"
            Exit

        } else {

            throw "Invalid Response"

        }

    } catch {

        optionSelection

    }

    return
    
}

function main() {

    fileHandling
    Set-Location -Path "$env:homedrive\odt"
    optionSelection

    return

}

#Annoying Apps Removed Config Links
$o365_modified = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-365/office-365-modified.xml"
$o2021_modified = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-2021/office-2021-modified.xml"
$o2019_modified = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-2019/office-2019-modified.xml"
$o2016_modified = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-2016/office-2016-modified.xml"

<#Full Installation Config Links
$o365_full = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-365/office-365-full.xml"
$o2021_full = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-2021/office-2021-full.xml"
$o2019_full = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-2019/office-2019-full.xml"
$o2016_full = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/office-2016/office-2016-full.xml"#>

#Uninstall Config Link
$uninstall = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/uninstall/office-uninstall.xml"

main