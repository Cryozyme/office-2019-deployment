@echo off & cls

echo ------------------------------------------------------------------------------------------

echo ^| THIS SCRIPT WILL INSTALL THE NEWEST VERSION OF MICROSOFT OFFICE ^2019 PROFESSIONAL PLUS ^|

echo ------------------------------------------------------------------------------------------

echo: & echo Cancel if you don't want to run this operation ^<ctrl^> + ^<c^> & echo: & timeout 3

cls

echo ----------------------------------------

echo ^| STOPPING OFFICE SERVICES ^& PROCESSES ^|

echo ----------------------------------------

sc stop "BITS" >nul & sc stop "ClickToRunSvc" >nul & taskkill /f /im "OfficeClickToRun.exe" /t >nul

cls

echo --------------------------------------------

echo ^| REMOVING PREEXISTING FILES ^& DIRECTORIES ^|

echo --------------------------------------------

del /f /s /q "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Skype for Business.ink" >nul & rd /s /q "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Office Tools" >nul & del /f /s /q "%ProgramData%\Microsoft\Windows\Start Menu\Programs\PowerPoint.ink" >nul & del /f /s /q "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Publisher.ink" >nul & del /f /s /q "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Outlook.ink" >nul & del /f /s /q "%ProgramData%\Microsoft\Windows\Start Menu\Programs\OneNote.ink" >nul & del /f /s /q "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Access.ink" >nul & del /f /s /q "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Excel.ink" >nul & del /f /s /q "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Word.ink" >nul & rd /s /q "%ProgramFiles%\common files\OfficeSoftwareProtectionPlatform" >nul & rd /s /q "%ProgramFiles%\common files\ClickToRun" >nul & rd /s /q "%ProgramFiles%\common files\OFFICE16" >nul & rd /s /q "%ProgramFiles%\Microsoft Office 15" >nul & rd /s /q "%ProgramData%\Microsoft\ClickToRun" >nul & rd /s /q "%LocalAppData%\Microsoft\OneNote" >nul & rd /s /q "%LocalAppData%\Microsoft\Office" >nul & rd /s /q "%ProgramFiles%\Microsoft Office" >nul & rd /s /q "%ProgramData%\Microsoft\Office" >nul & rd /s /q "%UserProfile%\custom-programs" >nul & rd /s /q "%Windir%\Temp" >nul & rd /s /q "%Temp%" >nul

cls

echo ------------------------

echo ^| DOWNLOADING ^& INSTALLING PSCORE 7 ^|

echo ------------------------

mkdir "%UserProfile%\custom-programs\odt2019" >nul & set odtPath=%UserProfile%\custom-programs\odt2019 & call "%~dp0\pscore7.ps1"

cls

echo ----------------------------------

echo ^| DOWNLOADING ^& INSTALLING 7-ZIP ^|

echo ----------------------------------

call "%~dp0\7zip.ps1"

cls

echo -------------------

echo ^| REFRESHING ^PATH ^|

echo -------------------

REM echo RefreshEnv.cmd only works from cmd.exe

echo: & echo | set /p dummy="Refreshing environment variables from registry for cmd.exe. Please wait..." & goto main

:SetFromReg
    "%WinDir%\System32\Reg" QUERY "%~1" /v "%~2" > "%TEMP%\_envset.tmp" 2>NUL & 
    (for /f "usebackq skip=2 tokens=2,*" %%A IN ("%TEMP%\_envset.tmp") do (echo/set "%~3=%%B")) & goto :EOF

:GetRegEnv
    "%WinDir%\System32\Reg" QUERY "%~1" > "%TEMP%\_envget.tmp" & (for /f "usebackq skip=2" %%A IN ("%TEMP%\_envget.tmp") do (if /I not "%%~A"=="Path" (call :SetFromReg "%~1" "%%~A" "%%~A"))) & goto :EOF

:main
    echo/@echo off >"%TEMP%\_env.cmd" & call :GetRegEnv "HKLM\System\CurrentControlSet\Control\Session Manager\Environment" >> "%TEMP%\_env.cmd" & call :GetRegEnv "HKCU\Environment">>"%TEMP%\_env.cmd" >> "%TEMP%\_env.cmd"
    call :SetFromReg "HKLM\System\CurrentControlSet\Control\Session Manager\Environment" Path Path_HKLM >> "%TEMP%\_env.cmd" & call :SetFromReg "HKCU\Environment" Path Path_HKCU >> "%TEMP%\_env.cmd" & echo/set "Path=%%Path_HKLM%%;%%Path_HKCU%%" >> "%TEMP%\_env.cmd" & del /f /q "%TEMP%\_envset.tmp" 2>nul & del /f /q "%TEMP%\_envget.tmp" 2>nul & SET "OriginalUserName=%USERNAME%" & SET "OriginalArchitecture=%PROCESSOR_ARCHITECTURE%" & call "%TEMP%\_env.cmd" & del /f /q "%TEMP%\_env.cmd" 2>nul & SET "USERNAME=%OriginalUserName%" & SET "PROCESSOR_ARCHITECTURE=%OriginalArchitecture%" & echo | set /p dummy="Finished." & echo:

cls

echo -----------------------------------

echo ^| DOWNLOADING CONFIGURATION FILES ^|

echo -----------------------------------

pwsh -command "$baseUrl='https://www.github.com/Cryozyme/office-deployment/raw/main';$odtPath='%UserProfile%\custom-programs\odt2019';iwr -UseBasicParsing $baseUrl/Office2019.xml -out $odtPath\office2019.xml;iwr -UseBasicParsing $baseUrl/Office2019_Uninstall.xml -out $odtPath\office2019-uninstall.xml;iwr -UseBasicParsing $baseUrl/Office365_Uninstall.xml -out $odtPath\office365-uninstall.xml"

cls

echo -------------------------------------------------------

echo ^| DOWNLOADING MICROSOFT OFFICE ^2019 PROFESSIONAL PLUS ^|

echo -------------------------------------------------------

pwsh -command "iwr -UseBasicParsing 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_14527-20178.exe' -out '%UserProfile%\custom-programs\odt2019\odt2019.exe'"

cls

echo --------------------------

echo ^| EXTRACTING ODT2019.EXE ^|

echo --------------------------

call "%ProgramFiles%\7-Zip\7z.exe" x "%odtPath%\odt2019.exe" -y -o"%odtPath%" -r >nul

cls

echo ------------------------------------------------

echo ^| UNINSTALLING PREVIOUS OFFICE INSTALLATION(S) ^|

echo ------------------------------------------------

start /w /b "" "%odtPath%\setup.exe" /configure "%odtPath%\office2019-uninstall.xml" & start /w /b "" "%odtPath%\setup.exe" /configure "%odtPath%\office365-uninstall.xml"

cls

echo --------------------------------------------

echo ^| INSTALLING OFFICE ^2019 PROFESSIONAL PLUS ^|

echo --------------------------------------------

start /w /b "" "%odtPath%\setup.exe" /download "%odtPath%\office2019.xml" & start /w /b "" "%odtPath%\setup.exe" /configure "%odtPath%\office2019.xml"

cls

echo ----------------

echo  ^| ACTIVATING ^|

echo ----------------

REM Office2019 ACTIVATING

(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16") & (if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16") & 
(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do (cscript ospp.vbs /inslic:"..\root\Licenses16\%%x") >nul) & cscript //nologo slmgr.vbs /ckms >nul & cscript ospp.vbs /setprt:1688 >nul & cscript ospp.vbs /unpkey:6MWKP >nul & cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul & cscript ospp.vbs /sethst:s8.uk.to >nul & cscript ospp.vbs /act >nul

REM PROJECT2019 ACTIVATING

(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16") & (if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16") & 
(for /f %%x in ('dir /b ..\root\Licenses16\projectpro2019vl_kms*.xrm-ms') do (cscript ospp.vbs /inslic:"..\root\Licenses16\%%x") >nul) & cscript ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B >nul & cscript ospp.vbs /sethst:s8.uk.to >nul & cscript ospp.vbs /setprt:1688 >nul & cscript ospp.vbs /act >nul

REM VISIO2019 ACTIVATING

(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16") & (if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16") & (for /f %%x in ('dir /b ..\root\Licenses16\visiopro2019vl_kms*.xrm-ms') do (cscript ospp.vbs /inslic:"..\root\Licenses16\%%x") >nul) & cscript ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB >nul & cscript ospp.vbs /sethst:s8.uk.to >nul & cscript ospp.vbs /setprt:1688 >nul & cscript ospp.vbs /act >nul

cls

echo ------------------

echo ^| CLEANING UP... ^|

echo ------------------

del /f /s /q "%odtPath%\configuration-Office2019Enterprise.xml" >nul & del /f /s /q "%odtPath%\configuration-Office365-x64.xml" >nul & del /f /s /q "%odtPath%\configuration-Office365-x86.xml" >nul & rd /s /q "%UserProfile%\custom-programs" >nul & sc start "BITS" >nul & sc start "ClickToRunSvc" >nul

cls

echo -------------

echo ^| FINISHED! ^|

echo -------------

pause