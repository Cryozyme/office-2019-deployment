function dependenciesDownload() {

    try {
        
        $app_id = 49117
        $url_base = "https://www.microsoft.com/en-us/download/confirmation.aspx"
        $odt_version = "$url_base`?id=$app_id"
        $latest = Invoke-WebRequest -UseBasicParsing $odt_version
        $file_pattern = "officedeploymenttool_"
        $version = (($latest.Links | Where-Object {$_.href -match "$file_pattern"}).href).Split("`n")[0].Trim()
        $file_name = "odt2019.exe"
        
        Write-Host "Downloading Office Deployment Tool from $version"
        Start-BitsTransfer -Source $version -Destination "$deployment_tool_path\$file_name" -TransferType Download -Priority Foreground

    } catch {
        
        Write-Host "$_"

    } try {
        
        if(-not(Test-Path -Path "$env:ProgramFiles\7-Zip" -PathType Container)) {
            
            $url_base = "https://7-zip.org/download.html"
            $latest = Invoke-WebRequest -UseBasicParsing $url_base
            $file_pattern = "-x64.exe"
            $version = (($latest.Links | Where-Object {$_.href -match "$file_pattern"}).href).Split(" ")[0].Trim()
            $version = ($url_base.Split("download.html")[0].Trim() + $version)
            $file_name = "7zip.exe"
            
            Write-Host "Downloading 7-zip from $version"
            Start-BitsTransfer -Source $version -Destination "$deployment_tool_path\$file_name" -TransferType Download -Priority Foreground

        } else {

            throw "7-Zip is Already Installed"

        }

    } catch {

        Write-Host "$_"

    } finally {

        Exit-PSHostProcess
        Exit-PSSession

    }

    return

}
function configDownload() {

    $bits_transfer = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/office-2019/xml-config/configuration.csv"

    try {
        
        Write-Host "Downloading Configuration Files from $bits_transfer"
        Invoke-WebRequest -UseBasicParsing "$bits_transfer" -OutFile "$deployment_tool_path\configuration.csv"
        Set-Location -Path "$deployment_tool_path"
        Import-Csv -Path "$deployment_tool_path\configuration.csv" -Delimiter "," | Start-BitsTransfer -TransferType Download -Priority Foreground
        Set-Location -Path "-"
        Remove-Item -Path "$deployment_tool_path\configuration.csv" -Force

    } catch {

        Write-Host "$_"

    }

    dependenciesDownload
    return

}


function fileHandling() {

    #Variables
    $deployment_tool_path = "$env:homedrive\office_deployment\odt_2019"

    try {
        
        if(-not(Test-Path -Path $deployment_tool_path -PathType Container)) {
        
            New-Item -Path "$deployment_tool_path" -ItemType Directory -Force
        
        } else {

            throw "Installation Path Already Exists"
            
        }

    } catch {

        Write-Host "$_"
    
    } finally {

        configDownload

    }
    
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
    fileHandling
    return

}

function installContinue() {

    try {
        
        $continue = Read-Host -Prompt "Do you want to continue? (y/n)" -MaskInput
        
        if($continue.Trim() -eq "y") {

            serviceHandling

        } elseif ($continue.Trim() -eq "n") {

            return

        } else {

            throw "Invalid Response"

        }

    } catch {

        Write-Host "$_"
        installContinue
    
    }

    return

}

function main() {

    Write-Host "------------------------------------------------------------------------------------------"
    Write-Host "`| THIS SCRIPT WILL INSTALL THE LATEST VERSION OF MICROSOFT OFFICE 2019 PROFESSIONAL PLUS `|"
    Write-Host "------------------------------------------------------------------------------------------"
    Write-Host ""

    installContinue
    Write-Host "Finished"
    
    return

}

main