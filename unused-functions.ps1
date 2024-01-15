function dependenciesDownload() {

    try {
        
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