function uninstallPrevious() {

    Start-Process -FilePath "$deployment_tool_path\setup.exe" -ArgumentList "/configure $deployment_tool_path\uninstall.xml" -Wait
    return

}

function configDownload() {

    $config_url = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/xml-config/install.csv"

    try {
        
        Write-Host "Downloading Configuration Files from $config_url"
        Invoke-WebRequest -UseBasicParsing "$config_url" -OutFile "$deployment_tool_path\install.csv"
        Set-Location -Path "$deployment_tool_path"
        Import-Csv -Path "$deployment_tool_path\install.csv" -Delimiter "," | Start-BitsTransfer -TransferType Download -Priority Foreground
        Remove-Item -Path "$deployment_tool_path\install.csv" -Force

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
        Start-BitsTransfer -Source $version -Destination "$deployment_tool_path\$file_name" -TransferType Download -Priority Foreground

    } catch {
        
        Write-Host "$_"

    }

    Start-Process -FilePath "$deployment_tool_path\$file_name" -ArgumentList "/extract:$deployment_tool_path /quiet /passive /norestart" -Wait
    Remove-Item -Path "$deployment_tool_path\configuration-*.xml", "$deployment_tool_path\$file_name" -Force
    uninstallPrevious
    return

}


function fileHandling() {

    #Variables
    $deployment_tool_path = "$env:homedrive\office_deployment\odt"

    try {
        
        if(-not(Test-Path -Path $deployment_tool_path -PathType Container)) {
        
            New-Item -Path "$deployment_tool_path" -ItemType Directory -Force
        
        } else {

            throw "Deployment Path Already Exists"
            
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
    Stop-Process -Name "WINWORD" -Force -ErrorAction Continue
    Stop-Process -Name "POWERPNT" -Force -ErrorAction Continue
    Stop-Process -Name "EXCEL" -Force -ErrorAction Continue
    Stop-Process -Name "MSACCESS" -Force -ErrorAction Continue
    Stop-Process -Name "MSPUB" -Force -ErrorAction Continue
    return

}

function installContinue() {

    $continue = Read-Host -Prompt "Do you want to continue? (y/n)";Clear-Host
        
    if($continue.Trim() -eq "y") {

        serviceHandling

    } elseif ($continue.Trim() -eq "n") {

        Exit-PSHostProcess

    } else {

        throw "Invalid Response"

    }

    return

}

function main() {

    Clear-Host
    Write-Host "---------------------------------------------------------------------"
    Write-Host "`| THIS SCRIPT WILL INSTALL THE LATEST VERSION OF MICROSOFT OFFICE `|"
    Write-Host "---------------------------------------------------------------------"
    Write-Host "`n"
    Write-Host "---------------------------------------------------------------------------------------------"
    Write-Host "`| THIS SCRIPT WILL UNINSTALL ALL PREVIOUS CLICK-TO-RUN VERSIONS OF OFFICE ON THE COMPUTER `|"
    Write-Host "---------------------------------------------------------------------------------------------"
    Write-Host "`n"

    Set-Location -Path "$deployment_tool_path" -ErrorAction Continue

    try {

        installContinue
    
    } catch {

        Write-Host "$_"
        installContinue

    } try {
        
        serviceHandling

    } catch {

        Write-Host "$_"

    } try {

        fileHandling

    } catch {

        Write-Host "$_"

    }

    Write-Host "Finished"
    Pause
    return

}

main