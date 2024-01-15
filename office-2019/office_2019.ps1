function configDownload() {

    $script_base_url = "https://raw.githubusercontent.com/Cryozyme/office-deployment/main/office-2019"

    try {
        
        Invoke-RestMethod -Uri "$script_base_url/xml-config/bits-transfer.csv" -UseBasicParsing -OutFile "$deployment_tool_path\bits-transfer.csv"
        Import-Csv -Path "$deployment_tool_path\bits-transfer.csv" -Delimiter "," | Start-BitsTransfer -TransferType Download -Priority Foreground

    } catch {

        Write-Host "$_"
        Pause

    }

    return

}


function fileHandling() {

    #Variables
    $deployment_tool_path = "$env:homedrive\office_deployment\odt_2019"
    #$cpu_architecture = "$env:processor_architecture"

    try {
        
        if(-not(Test-Path -Path $deployment_tool_path -PathType Container)) {
        
            New-Item -Path "$deployment_tool_path" -ItemType Directory -Force
        
        } else {

            throw "Path Already Exists"
            
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
    return

}

function installContinue() {

    try {
        
        $continue = Read-Host -Prompt "Do you want to continue? (y/n)" -MaskInput
        
        if($continue.Trim() -eq "y") {

            serviceHandling
            fileHandling

        } elseif ($continue.Trim() -eq "n") {

            Write-Host "Finished"
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
    
    return

}

main