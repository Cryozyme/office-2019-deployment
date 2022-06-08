call :main
pause
exit /b %errorlevel%

:main
    @echo off & cls
    set "officeBaseURL=https://www.github.com/Cryozyme/office-deployment/raw/main"
    set "odtoolPath=%userprofile%\custom-programs\odt2019"
    set "customPrograms=%userprofile%\custom-programs"
    set "arch=%processor_architecture%"
    set "officePath64=%programfiles%\Microsoft Office\Office16"
    set "officePath32=%programfiles(x86)%\Microsoft Office\Office16"
    set "7zFileName=7z.ps1"
    set "pwshFileName=pwsh.ps1"
    set "odtoolFileName=odtool.ps1"
    mkdir "%odtoolPath%" >nul 2>nul
    call :intro
    call :services
    call :config
    call :download
    call :odt_extract
    call :uninstall_previous
    call :install_prereqs
    call :install_2019
    if /i %arch%==amd64 (call :activate64) else (call :activate32)
    call :clean
    exit /b 0

:intro
    echo ------------------------------------------------------------------------------------------
    echo ^| THIS SCRIPT WILL INSTALL THE NEWEST VERSION OF MICROSOFT OFFICE ^2019 PROFESSIONAL PLUS ^|
    echo ------------------------------------------------------------------------------------------
    echo:
    echo Cancel if you don't want to run this operation ^<ctrl^> + ^<c^>
    echo:
    timeout 3
    exit /b 0

:services
    echo ----------------------------------------
    echo ^| STOPPING OFFICE SERVICES ^& PROCESSES ^|
    echo ----------------------------------------
    sc stop "BITS" >nul 2>nul
    sc stop "ClickToRunSvc" >nul 2>nul
    sc stop "wuauserv" >nul 2>nul
    sc stop "ose64" >nul 2>nul
    taskkill /f /im "OfficeClickToRun.exe" /t >nul 2>nul
    taskkill /f /im "OSE.exe" /t >nul 2>nul
    exit /b 0

:config
    echo -----------------------------------
    echo ^| DOWNLOADING CONFIGURATION FILES ^|
    echo -----------------------------------
    set "config_download=^
        #Invoke-WebRequest -UseBasicParsing %baseURL%/Office2019.xml -out %odtoolPath%\Office2019.xml;^
        #Invoke-WebRequest -UseBasicParsing %baseURL%/Office2019_Uninstall.xml -out %odtoolPath%\Office2019_Uninstall.xml;^
        #Invoke-WebRequest -UseBasicParsing %baseURL%/Office365_Uninstall.xml -out %odtoolPath%\Office365_Uninstall.xml;^
        #Invoke-WebRequest -UseBasicParsing %baseURL%/Office2021_Uninstall.xml -out %odtoolPath%\Office2021_Uninstall.xml;^
        Invoke-WebRequest -UseBasicParsing https://github.com/Cryozyme/office-deployment/raw/main/odtool.ps1 | Invoke-Expression"
    pwsh -NoProfile -NoLogo -NonInteractive -Command %config_download% >nul 2>nul
    pwsh -NoProfile -NoLogo -NonInteractive -ExecutionPolicy RemoteSigned -File %~dp0%7zFileName% >nul 2>nul
    pwsh -NoProfile -NoLogo -NonInteractive -ExecutionPolicy RemoteSigned -File %~dp0%pwshFileName% >nul 2>nul
    exit /b 0

:download
    echo -------------------------------------------------------
    echo ^| DOWNLOADING MICROSOFT OFFICE ^2019 PROFESSIONAL PLUS ^|
    echo -------------------------------------------------------
    set deployment_tool=^
        Invoke-WebRequest -uri %odtoolUrl% -out %odtoolPath%\odt2019.exe
    pwsh -NoProfile -NoLogo -NonInteractive -Command %deployment_tool% >nul 2>nul
    exit /b 0

:odt_extract
    echo --------------------------
    echo ^| EXTRACTING ODT2019.EXE ^|
    echo --------------------------
    7z.exe x "%odtoolPath%\odt2019.exe" -y -o"%odtoolPath%" -r >nul 2>nul
    exit /b 0

:uninstall_previous
    echo ------------------------------------------------
    echo ^| UNINSTALLING PREVIOUS OFFICE INSTALLATION(S) ^|
    echo ------------------------------------------------
    start "" /realtime /node 0 /affinity 0xff /wait "%odtoolPath%\setup.exe" /configure %odtoolPath%\Office2019_Uninstall.xml
    start "" /realtime /node 0 /affinity 0xff /wait "%odtoolPath%\setup.exe" /configure %odtoolPath%\Office365_Uninstall.xml
    start "" /realtime /node 0 /affinity 0xff /wait "%odtoolPath%\setup.exe" /configure %odtoolPath%\Office2021_Uninstall.xml
    exit /b 0

:install_prereqs
    echo ----------------------------
    echo ^| INSTALLING PREREQUISITES ^|
    echo ----------------------------
    start "" /realtime /node 0 /affinity 0xff /wait "%odtPrograms%\7-zip.exe" /S >nul 2>nul
    start "" /realtime /node 0 /affinity 0xff /wait "msiexec.exe" /package "%odtPrograms%\pscore7.msi" /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1 USE_MU=1 ENABLE_MU=1 >nul 2>nul

:install_2019
    echo --------------------------------------------
    echo ^| INSTALLING OFFICE ^2019 PROFESSIONAL PLUS ^|
    echo --------------------------------------------
    start "" /realtime /node 0 /affinity 0xff /wait "%odtoolPath%\setup.exe" /download "%odtoolPath%\Office2019.xml" >nul 2>nul
    start "" /realtime /node 0 /affinity 0xff /wait "%odtoolPath%\setup.exe" /configure "%odtoolPath%\Office2019.xml" >nul 2>nul
    exit /b 0

:activate64
    echo ----------------
    echo  ^| ACTIVATING ^|
    echo ----------------

    REM Office2019 ACTIVATING
    (if exist "%officePath64%\ospp.vbs" (cd /d "%officePath64%") >nul 2>nul)
    (for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do (cscript ospp.vbs /inslic:"..\root\Licenses16\%%x") >nul 2>nul)
    cscript //nologo slmgr.vbs /ckms >nul 2>nul
    cscript //nologo ospp.vbs /setprt:1688 >nul 2>nul
    cscript //nologo ospp.vbs /unpkey:6MWKP >nul 2>nul
    cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul 2>nul
    cscript //nologo ospp.vbs /sethst:s8.uk.to >nul 2>nul
    cscript //nologo ospp.vbs /act >nul 2>nul

    REM PROJECT2019 ACTIVATING
    (if exist "%officePath64%\ospp.vbs" (cd /d "%officePath64%") >nul 2>nul)
    (for /f %%x in ('dir /b ..\root\Licenses16\projectpro2019vl_kms*.xrm-ms') do (cscript ospp.vbs /inslic:"..\root\Licenses16\%%x") >nul 2>nul)
    cscript //nologo ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B >nul 2>nul
    cscript //nologo ospp.vbs /sethst:s8.uk.to >nul 2>nul
    cscript //nologo ospp.vbs /setprt:1688 >nul 2>nul
    cscript //nologo ospp.vbs /act >nul 2>nul

    REM VISIO2019 ACTIVATING
    (if exist "%officePath64%\ospp.vbs" (cd /d "%officePath64%") >nul 2>nul)
    (for /f %%x in ('dir /b ..\root\Licenses16\visiopro2019vl_kms*.xrm-ms') do (cscript ospp.vbs /inslic:"..\root\Licenses16\%%x") >nul 2>nul)
    cscript //nologo ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB >nul 2>nul
    cscript //nologo ospp.vbs /sethst:s8.uk.to >nul 2>nul
    cscript //nologo ospp.vbs /setprt:1688 >nul 2>nul
    cscript //nologo ospp.vbs /act >nul 2>nul
    exit /b 0

:activate32
    echo ----------------
    echo  ^| ACTIVATING ^|
    echo ----------------

    REM Office2019 ACTIVATING
    (if exist "%officePath32%\ospp.vbs" (cd /d "%officePath32%" >nul 2>nul))
    (for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do (cscript ospp.vbs /inslic:"..\root\Licenses16\%%x") >nul 2>nul)
    cscript //nologo slmgr.vbs /ckms >nul 2>nul
    cscript //nologo ospp.vbs /setprt:1688 >nul 2>nul
    cscript //nologo ospp.vbs /unpkey:6MWKP >nul 2>nul
    cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul 2>nul
    cscript //nologo ospp.vbs /sethst:s8.uk.to >nul 2>nul
    cscript //nologo ospp.vbs /act >nul 2>nul

    REM PROJECT2019 ACTIVATING
    (if exist "%officePath32%\ospp.vbs" (cd /d "%officePath32%" >nul 2>nul))
    (for /f %%x in ('dir /b ..\root\Licenses16\projectpro2019vl_kms*.xrm-ms') do (cscript ospp.vbs /inslic:"..\root\Licenses16\%%x") >nul 2>nul)
    cscript //nologo ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B >nul 2>nul
    cscript //nologo ospp.vbs /sethst:s8.uk.to >nul 2>nul
    cscript //nologo ospp.vbs /setprt:1688 >nul 2>nul
    cscript //nologo ospp.vbs /act >nul 2>nul

    REM VISIO2019 ACTIVATING
    (if exist "%officePath32%\ospp.vbs" (cd /d "%officePath32%" >nul 2>nul))
    (for /f %%x in ('dir /b ..\root\Licenses16\visiopro2019vl_kms*.xrm-ms') do (cscript ospp.vbs /inslic:"..\root\Licenses16\%%x") >nul 2>nul)
    cscript //nologo ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB >nul 2>nul
    cscript //nologo ospp.vbs /sethst:s8.uk.to >nul 2>nul
    cscript //nologo ospp.vbs /setprt:1688 >nul 2>nul
    cscript //nologo ospp.vbs /act >nul 2>nul
    exit /b 0

:clean
    echo ------------------
    echo ^| CLEANING UP... ^|
    echo ------------------
    rd /s /q "%odtoolPath%" >nul 2>nul
    exit /b 0