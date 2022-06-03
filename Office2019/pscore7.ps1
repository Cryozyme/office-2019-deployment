#Downloads latest pscore7 from github

nal nobj New-Object

nal jnp Join-Path

$urlBase = "https://github.com"

$releaseBase="PowerShell/PowerShell/releases"

$version="$urlBase/$releaseBase/latest" #can be latest or another more specific version (ex: $urlBase/$releaseBase/v0.54.0)

$latest=iwr -useb $version

$filePattern="-win-x64.msi" #end of file pattern (ex: in file v0.54.6_linux_ost.tar.gz, the file pattern is _linux_ost.tar.gz)

$versionBootStrapper=($latest.Links | where { $_.href -match "$filePattern" }).href

$versionBootStrapper=$versionBootStrapper.Trim(" ")

$dlUrl="$urlBase$versionBootStrapper"

Write-Host "Downloading PSCORE7 from $dlUrl"

$wc=nobj net.webclient

$wc.UseDefaultCredentials=$true

$wc.Proxy.Credentials=$wc.Credentials

$fileName="pscore7.msi"

$cwdLocation="$env:USERPROFILE\custom-programs\" # '.\' means current working directory for the terminal

$wc.DownloadFile($dlurl, (jnp (rvpa "$cwdLocation") "$fileName"))

exit