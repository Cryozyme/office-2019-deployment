#Downloads latest pscore7 from github
function Get-PwshVersion {
    $urlBase = "https://github.com"
    $releaseBase="PowerShell/PowerShell/releases"
    $version="$urlBase/$releaseBase/latest" #can be latest or another more specific version (ex: $urlBase/$releaseBase/v0.54.0)
    $latest=Invoke-WebRequest -useb $version
    $filePattern="-win-x64.msi" #end of file pattern (ex: in file v0.54.6_linux_ost.tar.gz, the file pattern is _linux_ost.tar.gz)
    $versionBootStrapper=($latest.Links | Where-Object { $_.href -match "$filePattern" }).href
    $versionBootStrapper=$versionBootStrapper.Trim(" ")
    $dlUrl="$urlBase$versionBootStrapper"
    return $dlUrl
}

function Get-PwshExecutable {
    $url=Get-PwshVersion
    Write-Information "Downloading PSCORE7 from $url" -InformationAction Continue
    $wc=New-Object net.webclient
    $wc.UseDefaultCredentials=$true
    $wc.Proxy.Credentials=$wc.Credentials
    $fileName="pscore7.msi"
    $cwdLocation="$env:USERPROFILE\custom-programs\" # '.\' means current working directory for the terminal
    $wc.DownloadFile($url, (Join-Path (Resolve-Path "$cwdLocation") "$fileName"))
}
Clear-Host
Get-PwshExecutable
pause