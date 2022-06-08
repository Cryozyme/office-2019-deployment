#Downloads latest pscore7 from github
function Get-7zVersion {
    $urlBase = "https://7-zip.org"
    $releaseBase="download.html"
    $version="$urlBase/$releaseBase" #can be latest or another more specific version (ex: $urlBase/$releaseBase/v0.54.0)
    $latest=Invoke-WebRequest -useb $version
    $filePattern="-x64.exe" #end of file pattern (ex: in file v0.54.6_linux_ost.tar.gz, the file pattern is _linux_ost.tar.gz)
    $versionBootStrapper=($latest.Links | Where-Object { $_.href -match "$filePattern" }).href
    $versionBootStrapper=$versionBootStrapper.Split(' ')[0].Trim(" ") #splits string at the first space, and then uses the 0th index (first part of the newly split string) as the url
    $dlUrl="$urlBase/$versionBootStrapper"
    return $dlUrl
}

function Get-7zExecutable {
    $url=Get-7zVersion
    Write-Information "Downloading 7-zip from $url" -InformationAction Continue
    $wc=New-Object net.webclient
    $wc.UseDefaultCredentials=$true
    $wc.Proxy.Credentials=$wc.Credentials
    $fileName="7-zip.exe"
    $cwdLocation="$env:USERPROFILE\custom-programs\" # '.\' means current working directory for the terminal
    $wc.DownloadFile($url, (Join-Path (Resolve-Path "$cwdLocation") "$fileName"))
}
Clear-Host
Get-7zExecutable
pause