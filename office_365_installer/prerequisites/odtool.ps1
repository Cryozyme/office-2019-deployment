#Downloads latest office deployment tool from microsoft website
function Get-OdToolVersion {
    $appID=49117
    $urlBase="https://www.microsoft.com/en-us"
    $releaseBase="download/confirmation.aspx"
    $version="$urlBase/$releaseBase`?id=$appID"
        #can be latest or another more specific version (ex: $urlBase/$releaseBase/v0.54.0)
    $latest=Invoke-WebRequest -useb $version
    $filePattern="officedeploymenttool_"
        #end of file pattern (ex: in file v0.54.6_linux_ost.tar.gz, the file pattern is _linux_ost.tar.gz)
    $versionBootStrapper=($latest.Links | Where-Object { $_.href -match "$filePattern" }).href
    return $versionBootStrapper=$versionBootStrapper.Split('`n')[0].Trim(" ")
        #splits string at the first newline, and then uses the 0th index (first part of the newly split string) as the url
}

function Get-OdToolExecutable {
    Write-Information "Downloading Office Deployment Tool from $(Get-OdToolVersion)" -InformationAction Continue
    $wc=New-Object net.webclient
    $wc.UseDefaultCredentials=$true
    $wc.Proxy.Credentials=$wc.Credentials
    $fileName="office_2019.exe"
    $cwdLocation="$env:USERPROFILE\custom-programs\" # '.\' means current working directory for the terminal
    $wc.DownloadFile($(Get-OdToolVersion), (Join-Path (Resolve-Path "$cwdLocation") "$fileName"))
}
Clear-Host
Get-OdToolExecutable
pause