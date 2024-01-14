function configDownload() {

    $script_base_url = "https://www.github.com/Cryozyme/office-deployment/raw/main/office_2019"

    try {
        
        Import-CSV -Path "$script_base_url/xml-config/bits-transfer.txt" | Start-BitsTransfer -TransferType Download

    }
    catch {
        Write-Host "$_" -ErrorAction SilentlyContinue
        Pause
    }
    return

}


function fileHandling() {

    #Variables
    $deployment_tool_path = "$env:homedrive\office_deployment\odt_2019"
    $cpu_architecture = "$env:processor_architecture"

    try {
        
        if(-not(Test-Path -Path $deployment_tool_path -PathType Container -ErrorAction SilentlyContinue)) {
        
            New-Item -Path "$deployment_tool_path" -ItemType Directory -Force -ErrorAction SilentlyContinue
        
        } else {

            throw "Path Already Exists"
            
        }

    } catch {

        Write-Host "$_" -ErrorAction SilentlyContinue
    
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
        
        $continue = Read-Host -Prompt "Do you want to continue? (y/n)" -MaskInput -ErrorAction SilentlyContinue
        
        if($continue.Trim() -eq "y") {

            serviceHandling
            fileHandling

        } elseif ($continue.Trim() -eq "n") {

            Write-Host "Finished" -ErrorAction SilentlyContinue
            return

        } else {

            throw "Invalid Response"

        }

    } catch {

        Write-Host "$_" -ErrorAction SilentlyContinue
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