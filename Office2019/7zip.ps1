#Downloads latest pscore7 from github

cls

nal print echo

nal nobj New-Object

nal jnp Join-Path

$urlBase = "https://7-zip.org"

$releaseBase="download.html"

$version="$urlBase/$releaseBase" #can be latest or another more specific version (ex: $urlBase/$releaseBase/v0.54.0)

$latest=iwr -UseBasicParsing $version

$filePattern="-x64.exe" #end of file pattern (ex: in file v0.54.6_linux_ost.tar.gz, the file pattern is _linux_ost.tar.gz)

$versionBootStrapper=($latest.Links | where { $_.href -match "$filePattern" }).href

$versionBootStrapper=$versionBootStrapper.Split(' ')[0].Trim(" ") #splits string at the first space, and then uses the 0th index (first part of the newly split string) as the url

$dlUrl="$urlBase/$versionBootStrapper"

Write-Host "Downloading 7-zip from $dlUrl"

$wc=nobj net.webclient

$wc.UseDefaultCredentials=$true

$wc.Proxy.Credentials=$wc.Credentials

$fileName="7-zip.exe"

$cwdLocation="$env:USERPROFILE\custom-programs\" # '.\' means current working directory for the terminal

$wc.DownloadFile($dlurl, (jnp (rvpa "$cwdLocation") "$fileName"))

& "$env:USERPROFILE\custom-programs\7-zip.exe" /S

exit