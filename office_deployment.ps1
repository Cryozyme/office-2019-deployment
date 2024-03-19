function installO365() {

    Start-Process -FilePath "$(fileHandling)\setup.exe" -ArgumentList "/download $(fileHandling)\O365.xml" -Wait
    Start-Process -FilePath "$(fileHandling)\setup.exe" -ArgumentList "/configure $(fileHandling)\O365.xml" -Wait
    return

}

function installO2021() {

    Start-Process -FilePath "$(fileHandling)\setup.exe" -ArgumentList "/download $(fileHandling)\O2021.xml" -Wait
    Start-Process -FilePath "$(fileHandling)\setup.exe" -ArgumentList "/configure $(fileHandling)\O2021.xml" -Wait
    return

}

function installO2019() {

    Start-Process -FilePath "$(fileHandling)\setup.exe" -ArgumentList "/download $(fileHandling)\O2019.xml" -Wait
    Start-Process -FilePath "$(fileHandling)\setup.exe" -ArgumentList "/configure $(fileHandling)\O2019.xml" -Wait
    return

}

function installO2016() {

    Start-Process -FilePath "$(fileHandling)\setup.exe" -ArgumentList "/download $(fileHandling)\O2016.xml" -Wait
    Start-Process -FilePath "$(fileHandling)\setup.exe" -ArgumentList "/configure $(fileHandling)\O2016.xml" -Wait
    return

}

function uninstallPrevious() {

    
    Start-Process -FilePath "$(fileHandling)\setup.exe" -ArgumentList "/configure $(fileHandling)\uninstall.xml" -Wait
    return

}

function configDownload() {

    $config_url = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/install.csv"

    try {
        
        Write-Host "Downloading Configuration Files from $config_url"
        Invoke-WebRequest -UseBasicParsing "$config_url" -OutFile "$(fileHandling)\install.csv"
        Import-Csv -Path "$(fileHandling)\install.csv" -Delimiter "," | Start-BitsTransfer -TransferType Download -Priority Foreground
        Remove-Item -Path "$(fileHandling)\install.csv" -Force

    } catch {

        Write-Host "$_"

    } try {
        
        $app_id = 49117
        $url_base = "https://www.microsoft.com/en-us/download/confirmation.aspx"
        $odt_version = "$url_base`?id=$app_id"
        $latest = Invoke-WebRequest -UseBasicParsing $odt_version
        $file_pattern = "officedeploymenttool_"
        $version = (($latest.Links | Where-Object {$_.href -match "$file_pattern"}).href).Split("`n")[0].Trim()
        $file_name = "odt.exe"
        
        Write-Host "Downloading Office Deployment Tool from $version"
        Start-BitsTransfer -Source $version -Destination "$(fileHandling)\$file_name" -TransferType Download -Priority Foreground

    } catch {}

    Start-Process -FilePath "$(fileHandling)\$file_name" -ArgumentList "/extract:$(fileHandling) /quiet /passive /norestart" -Wait
    Remove-Item -Path "$(fileHandling)\configuration-*.xml", "$(fileHandling)\$file_name" -Force
    
    return

}

function serviceHandling() {

    sc.exe stop "BITS" | Out-Null
    sc.exe stop "ClickToRunSvc" | Out-Null
    sc.exe stop "wuauserv" | Out-Null
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

    } catch {}

    try {

        if(-not(Test-Path -Path $logging_path_uninstall -PathType Container)) {

            New-Item -Path "$logging_path_uninstall" -ItemType Directory -Force

        } else {

            throw "Uninstall Log Path Already Exists"

        }

    } catch {}
    
    try {

        if(-not(Test-Path -Path $logging_path_365 -PathType Container)) {

            New-Item -Path "$logging_path_365" -ItemType Directory -Force

        } else {

            throw "Uninstall Log Path Already Exists"

        }

    } catch {}

    try {

        if(-not(Test-Path -Path $logging_path_2021 -PathType Container)) {

            New-Item -Path "$logging_path_2021" -ItemType Directory -Force

        } else {

            throw "Uninstall Log Path Already Exists"

        }

    } catch {}

    try {

        if(-not(Test-Path -Path $logging_path_2019 -PathType Container)) {

            New-Item -Path "$logging_path_2019" -ItemType Directory -Force

        } else {

            throw "Uninstall Log Path Already Exists"

        }

    } catch {}

    try {

        if(-not(Test-Path -Path $logging_path_2016 -PathType Container)) {

            New-Item -Path "$logging_path_2016" -ItemType Directory -Force

        } else {

            throw "Uninstall Log Path Already Exists"

        }

    } catch {}

    return $deployment_tool_path

}

function optionSelection() {

    Clear-Host

    $option = "$(Read-Host -Prompt "Select an option`n----------------`n1:Uninstall All`n2:Install Office 365`n3:Install Office 2021`n4:Install Office 2019`n5:Install Office 2016`n6:Exit`n")"

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
            Remove-Item -Path "$(fileHandling)" -Recurse -Force
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

    fileHandling | Out-Null
    Set-Location -Path "$(fileHandling)"

    optionSelection
    return

}

main